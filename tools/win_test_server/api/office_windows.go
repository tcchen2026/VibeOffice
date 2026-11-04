package api

import (
	"context"
	"crypto/sha256"
	_ "embed"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"io"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"sync"
	"syscall"
	"time"
	"unsafe"
)

//go:embed worker.ps1
var worker string

//go:embed capture.ps1
var capture string

var user32 = syscall.NewLazyDLL("user32.dll")
var enumWindows = user32.NewProc("EnumWindows")
var enumChildren = user32.NewProc("EnumChildWindows")
var windowPID = user32.NewProc("GetWindowThreadProcessId")
var windowText = user32.NewProc("GetWindowTextW")
var windowClass = user32.NewProc("GetClassNameW")
var windowVisible = user32.NewProc("IsWindowVisible")
var getWindow = user32.NewProc("GetWindow")
var openDesktop = user32.NewProc("OpenInputDesktop")
var desktopInfo = user32.NewProc("GetUserObjectInformationW")
var closeDesktop = user32.NewProc("CloseDesktop")

func nativeText(proc *syscall.LazyProc, handle uintptr) string {
	b := make([]uint16, 4096)
	proc.Call(handle, uintptr(unsafe.Pointer(&b[0])), uintptr(len(b)))
	return syscall.UTF16ToString(b)
}

func checkDesktop() error {
	handle, _, err := openDesktop.Call(0, 0, 1) // DESKTOP_READOBJECTS
	if handle == 0 {
		return fmt.Errorf("Interactive desktop unavailable (logged out or locked): %v", err)
	}
	defer closeDesktop.Call(handle)
	var needed uint32
	b := make([]uint16, 256)
	ok, _, err := desktopInfo.Call(handle, 2, uintptr(unsafe.Pointer(&b[0])), uintptr(len(b)*2), uintptr(unsafe.Pointer(&needed)))
	if ok == 0 {
		return fmt.Errorf("Cannot inspect input desktop: %v", err)
	}
	if !strings.EqualFold(syscall.UTF16ToString(b), "Default") {
		return fmt.Errorf("Desktop is locked; keep the test VM signed in and unlocked")
	}
	return nil
}

func officeBusy() (bool, error) {
	snapshot, err := syscall.CreateToolhelp32Snapshot(syscall.TH32CS_SNAPPROCESS, 0)
	if err != nil {
		return false, err
	}
	defer syscall.CloseHandle(snapshot)
	var entry syscall.ProcessEntry32
	entry.Size = uint32(unsafe.Sizeof(entry))
	err = syscall.Process32First(snapshot, &entry)
	for err == nil {
		name := strings.ToLower(syscall.UTF16ToString(entry.ExeFile[:]))
		if name == "winword.exe" || name == "excel.exe" || name == "powerpnt.exe" {
			return true, nil
		}
		err = syscall.Process32Next(snapshot, &entry)
	}
	if err != syscall.ERROR_NO_MORE_FILES {
		return false, err
	}
	return false, nil
}

type observedWindow struct {
	handle uintptr
	dialog Dialog
}

// NewCallback allocates a permanent runtime trampoline. Allocate these two once, never per poll.
var observationMu sync.Mutex
var observedPID int
var observedWindows []observedWindow
var observedDialog *Dialog
var childCallback = syscall.NewCallback(func(h, unused uintptr) uintptr {
	if label := nativeText(windowText, h); label != "" {
		observedDialog.Text += "\n" + label
	}
	return 1
})
var windowCallback = syscall.NewCallback(func(handle, unused uintptr) uintptr {
	var ownerPID uint32
	windowPID.Call(handle, uintptr(unsafe.Pointer(&ownerPID)))
	visible, _, _ := windowVisible.Call(handle)
	if int(ownerPID) != observedPID || visible == 0 {
		return 1
	}
	class := nativeText(windowClass, handle)
	owner, _, _ := getWindow.Call(handle, 4) // GW_OWNER
	main := class == "OpusApp" || class == "XLMAIN" || class == "PPTFrameClass" || class == "wndclass_desked_gsk"
	if main || !(class == "#32770" || strings.HasPrefix(class, "bosa_sdm") || class == "NUIDialog" || class == "NetUIHWND" || owner != 0) {
		return 1
	}
	text := nativeText(windowText, handle)
	// Ignore tooltips and untitled, non-dialog helper windows.
	if strings.Contains(strings.ToLower(class), "tooltip") || (text == "" && class != "#32770" && !strings.HasPrefix(class, "bosa_sdm")) {
		return 1
	}
	d := Dialog{Title: text, Class: class, Text: text}
	observedDialog = &d
	enumChildren.Call(handle, childCallback, 0)
	observedWindows = append(observedWindows, observedWindow{handle, d})
	return 1
})

func observe(pid int) ([]observedWindow, error) {
	observationMu.Lock()
	defer observationMu.Unlock()
	observedPID, observedWindows = pid, nil
	ok, _, err := enumWindows.Call(windowCallback, 0)
	if ok == 0 {
		return nil, fmt.Errorf("Cannot observe Office windows: %v", err)
	}
	return observedWindows, nil
}

func powershell() string {
	base := os.Getenv("SystemRoot")
	// This VM has 32-bit Office; using the 32-bit COM client also supports its registration view.
	ps := filepath.Join(base, "SysWOW64", "WindowsPowerShell", "v1.0", "powershell.exe")
	if _, err := os.Stat(ps); err == nil {
		return ps
	}
	return filepath.Join(base, "System32", "WindowsPowerShell", "v1.0", "powershell.exe")
}

func hashFile(name string) (string, error) {
	f, err := os.Open(name)
	if err != nil {
		return "", err
	}
	defer f.Close()
	h := sha256.New()
	if _, err := io.Copy(h, f); err != nil {
		return "", err
	}
	return hex.EncodeToString(h.Sum(nil)), nil
}

type workerState struct {
	FailedPhase   string `json:"failedPhase"`
	Phase         string `json:"phase"`
	Opened        bool   `json:"opened"`
	Exported      bool   `json:"exported"`
	OfficePID     int    `json:"officePid"`
	OfficeVersion string `json:"officeVersion"`
	OfficeBuild   string `json:"officeBuild"`
	Application   string `json:"application"`
	Error         string `json:"error"`
}

func RunOffice(ctx context.Context, job Job) (out Outcome) {
	out = Outcome{Status: "error", PDFStatus: "not_attempted", Dialogs: []Dialog{}}
	if err := checkDesktop(); err != nil {
		out.Error = err.Error()
		return
	}
	busy, err := officeBusy()
	if err != nil {
		out.Error = err.Error()
		return
	}
	if busy {
		out.Status = "busy"
		out.Error = "An Office process is already running; close it before submitting. No existing process was touched."
		return
	}
	for name, text := range map[string]string{"worker.ps1": worker, "capture.ps1": capture} {
		if err := os.WriteFile(filepath.Join(job.Dir, name), []byte(text), 0600); err != nil {
			out.Error = err.Error()
			return
		}
	}
	log, err := os.Create(filepath.Join(job.Dir, "worker.log"))
	if err != nil {
		out.Error = err.Error()
		return
	}
	defer log.Close()
	cmd := exec.Command(powershell(), "-NoProfile", "-NonInteractive", "-STA", "-ExecutionPolicy", "Bypass", "-File", filepath.Join(job.Dir, "worker.ps1"), "-Directory", job.Dir)
	cmd.Stdout, cmd.Stderr = log, log
	if err := cmd.Start(); err != nil {
		out.Error = err.Error()
		return
	}
	done := make(chan error, 1)
	go func() { done <- cmd.Wait() }()
	var state workerState
	var officeHandle syscall.Handle
	workerExited := false
	readState := func() {
		if b, err := os.ReadFile(filepath.Join(job.Dir, "state.json")); err == nil {
			var next workerState
			if json.Unmarshal(b, &next) == nil {
				state = next
				if officeHandle == 0 && state.OfficePID > 0 && state.Phase != "done" {
					officeHandle, _ = syscall.OpenProcess(syscall.PROCESS_TERMINATE|syscall.SYNCHRONIZE, false, uint32(state.OfficePID))
				}
			}
		}
	}
	defer func() {
		// Only this worker and its fresh Office process. A held handle cannot target a reused PID.
		_ = cmd.Process.Kill()
		if !workerExited {
			select {
			case <-done:
			case <-time.After(2 * time.Second):
			}
		}
		readState()
		if officeHandle != 0 {
			_ = syscall.TerminateProcess(officeHandle, 1)
			_, _ = syscall.WaitForSingleObject(officeHandle, 2000)
			_ = syscall.CloseHandle(officeHandle)
		}
		out.Phase, out.Application = state.Phase, state.Application
		out.OfficePID, out.OfficeVersion, out.OfficeBuild = state.OfficePID, state.OfficeVersion, state.OfficeBuild
		hash, err := hashFile(filepath.Join(job.Dir, "input"+strings.ToLower(filepath.Ext(job.Filename))))
		out.InputUnchanged = err == nil && hash == job.InputSHA256
		if !out.InputUnchanged {
			out.Status = "error"
			out.CanOpenWithoutRepair = nil
			out.Error += " Input hash changed or could not be verified."
		}
	}()
	ticker := time.NewTicker(100 * time.Millisecond)
	defer ticker.Stop()
	firstSeen := map[uintptr]time.Time{}
	for {
		select {
		case <-ctx.Done():
			out.Status = "timeout"
			out.Error = "Office exceeded the per-job deadline; no acceptance was inferred."
			return
		case processErr := <-done:
			workerExited = true
			readState()
			if processErr != nil {
				out.Error = "Office worker failed: " + processErr.Error()
				return
			}
			if state.Error != "" {
				out.Error = state.Error
				if state.Opened {
					out.Status, out.PDFStatus = "export_failed", "failed"
					if out.ObservationCount > 0 {
						yes := true
						out.CanOpenWithoutRepair = &yes
					}
				} else if state.FailedPhase == "opening" {
					out.Status = "open_failed"
					no := false
					out.CanOpenWithoutRepair = &no
				}
				if repairWords.MatchString(state.Error) {
					out.Status = "repair_required"
					no := false
					out.CanOpenWithoutRepair = &no
				}
				return
			}
			if !state.Opened || state.Phase != "done" || out.ObservationCount == 0 {
				out.Error = "Worker did not complete an observed normal open"
				return
			}
			yes := true
			out.CanOpenWithoutRepair = &yes
			if !state.Exported {
				out.Status, out.PDFStatus, out.Error = "export_failed", "failed", "Office did not finish PDF export"
				return
			}
			pdf := filepath.Join(job.Dir, "output.pdf")
			f, err := os.Open(pdf)
			if err == nil {
				var header [5]byte
				_, err = io.ReadFull(f, header[:])
				f.Close()
				if string(header[:]) != "%PDF-" {
					err = fmt.Errorf("Invalid PDF signature")
				}
			}
			if err != nil {
				out.Status, out.PDFStatus, out.Error = "export_failed", "failed", fmt.Sprint(err)
				return
			}
			out.PDFSHA256, err = hashFile(pdf)
			if err != nil {
				out.Status, out.PDFStatus, out.Error = "export_failed", "failed", err.Error()
				return
			}
			out.Status, out.PDFStatus, out.PDF = "ok", "ok", "/jobs/"+job.ID+"/output.pdf"
			return
		case <-ticker.C:
			readState()
			if err := checkDesktop(); err != nil {
				out.Error = err.Error()
				return
			}
			if state.OfficePID == 0 {
				continue
			}
			windows, err := observe(state.OfficePID)
			if err != nil {
				out.Error = err.Error()
				return
			}
			out.ObservationCount++
			for _, window := range windows {
				if ClassifyDialog(window.dialog) == "progress" {
					continue
				}
				if firstSeen[window.handle].IsZero() {
					firstSeen[window.handle] = time.Now()
				}
				// Brief publishing/progress windows may disappear normally. A repair prompt is immediate.
				if !repairWords.MatchString(window.dialog.Text) && time.Since(firstSeen[window.handle]) < 1500*time.Millisecond {
					continue
				}
				d := window.dialog
				captureCtx, cancel := context.WithTimeout(ctx, 8*time.Second)
				c := exec.CommandContext(captureCtx, powershell(), "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", filepath.Join(job.Dir, "capture.ps1"), "-Handle", fmt.Sprint(window.handle), "-Directory", job.Dir)
				c.Stdout, c.Stderr = log, log
				_ = c.Run()
				cancel()
				if text, err := os.ReadFile(filepath.Join(job.Dir, "dialog.txt")); err == nil {
					d.Text += "\n" + string(text)
				}
				if _, err := os.Stat(filepath.Join(job.Dir, "dialog.png")); err == nil {
					d.Screenshot = "/jobs/" + job.ID + "/dialog.png"
				}
				out.Dialogs = append(out.Dialogs, d)
				out.Status, out.Error = "blocked", "Office displayed a dialog; it was recorded, never accepted."
				if ClassifyDialog(d) == "repair_required" {
					out.Status = "repair_required"
					no := false
					out.CanOpenWithoutRepair = &no
				} else if ClassifyDialog(d) == "open_failed" && state.Phase == "opening" {
					out.Status = "open_failed"
					no := false
					out.CanOpenWithoutRepair = &no
				}
				return
			}
		}
	}
}

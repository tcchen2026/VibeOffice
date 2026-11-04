// Test infrastructure only; never served by public/ or called by the web apps.
package api

import (
	"context"
	"crypto/rand"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net/http"
	"os"
	"path/filepath"
	"regexp"
	"runtime"
	"strconv"
	"strings"
	"sync/atomic"
	"time"
)

const MaxUpload = 100 << 20
const acceptance = "The file can be opened with Office 2016 without repair prompts."

type Dialog struct {
	Title      string `json:"title"`
	Class      string `json:"class"`
	Text       string `json:"text"`
	Screenshot string `json:"screenshot,omitempty"`
}

type Outcome struct {
	Status               string   `json:"status"`
	CanOpenWithoutRepair *bool    `json:"canOpenWithoutRepair"`
	Message              string   `json:"message,omitempty"`
	Application          string   `json:"application,omitempty"`
	OfficeVersion        string   `json:"officeVersion,omitempty"`
	OfficeBuild          string   `json:"officeBuild,omitempty"`
	OfficePID            int      `json:"officePid,omitempty"`
	Phase                string   `json:"phase,omitempty"`
	PDFStatus            string   `json:"pdfStatus"`
	PDF                  string   `json:"pdf,omitempty"`
	PDFSHA256            string   `json:"pdfSha256,omitempty"`
	InputUnchanged       bool     `json:"inputUnchanged"`
	Dialogs              []Dialog `json:"dialogs"`
	Error                string   `json:"error,omitempty"`
	ObservationCount     int      `json:"observationCount"`
}

type Job struct {
	ID             string     `json:"id"`
	Filename       string     `json:"filename"`
	InputSHA256    string     `json:"inputSha256"`
	TimeoutSeconds int        `json:"timeoutSeconds"`
	CreatedAt      time.Time  `json:"createdAt"`
	StartedAt      *time.Time `json:"startedAt,omitempty"`
	ElapsedMS      int64      `json:"elapsedMs,omitempty"`
	Outcome
	Dir string `json:"-"`
}

type Runner func(context.Context, Job) Outcome
type Server struct {
	root    string
	run     Runner
	queue   chan Job
	pending chan struct{}
	active  atomic.Int32
}

var idPattern = regexp.MustCompile(`^[0-9a-f]{32}$`)

func New(root string, run Runner) (*Server, error) {
	root, err := filepath.Abs(root)
	if err != nil {
		return nil, err
	}
	if err = os.MkdirAll(filepath.Join(root, "jobs"), 0700); err != nil {
		return nil, err
	}
	s := &Server{root: root, run: run, queue: make(chan Job, 16), pending: make(chan struct{}, 16)}
	// Never let a restart turn an interrupted job into a successful result.
	entries, err := os.ReadDir(filepath.Join(root, "jobs"))
	if err != nil {
		return nil, err
	}
	for _, e := range entries {
		if !e.IsDir() || !idPattern.MatchString(e.Name()) {
			continue
		}
		job, err := s.read(e.Name())
		if err == nil && (job.Status == "queued" || job.Status == "running") {
			job.Status, job.Error = "error", "Server restarted before this job finished; submit again."
			job.CanOpenWithoutRepair = nil
			if err := writeJSON(filepath.Join(job.Dir, "result.json"), job); err != nil {
				return nil, err
			}
		}
	}
	go s.work()
	return s, nil
}

func writeJSON(name string, value any) error {
	b, err := json.MarshalIndent(value, "", "  ")
	if err != nil {
		return err
	}
	if err = os.WriteFile(name+".tmp", append(b, '\n'), 0600); err != nil {
		return err
	}
	return os.Rename(name+".tmp", name)
}

func (s *Server) read(id string) (Job, error) {
	var job Job
	job.Dir = filepath.Join(s.root, "jobs", id)
	b, err := os.ReadFile(filepath.Join(job.Dir, "result.json"))
	if err == nil {
		err = json.Unmarshal(b, &job)
	}
	return job, err
}

func (s *Server) work() {
	for job := range s.queue {
		s.active.Store(1)
		started := time.Now().UTC()
		job.StartedAt, job.Status = &started, "running"
		err := writeJSON(filepath.Join(job.Dir, "result.json"), job)
		if err == nil {
			ctx, cancel := context.WithTimeout(context.Background(), time.Duration(job.TimeoutSeconds)*time.Second)
			job.Outcome = s.run(ctx, job)
			cancel()
		} else {
			job.Status, job.Error = "error", err.Error()
		}
		job.ElapsedMS = time.Since(started).Milliseconds()
		if job.Dialogs == nil {
			job.Dialogs = []Dialog{}
		}
		if job.CanOpenWithoutRepair != nil && *job.CanOpenWithoutRepair {
			job.Message = acceptance
		}
		if err := writeJSON(filepath.Join(job.Dir, "result.json"), job); err != nil {
			// Preserve the failure in the process log if storage has become unavailable.
			fmt.Fprintln(os.Stderr, "cannot persist job", job.ID, err)
		}
		s.active.Store(0)
		<-s.pending
	}
}

func send(w http.ResponseWriter, code int, v any) {
	w.Header().Set("Content-Type", "application/json; charset=utf-8")
	w.Header().Set("Cache-Control", "no-store")
	w.WriteHeader(code)
	_ = json.NewEncoder(w).Encode(v)
}

func failure(w http.ResponseWriter, code int, message string) {
	send(w, code, map[string]string{"error": message})
}

func (s *Server) ServeHTTP(w http.ResponseWriter, r *http.Request) {
	// This private test API has no CORS. Do not allow a web page to submit a user's files.
	if r.Header.Get("Origin") != "" {
		failure(w, 403, "Browser cross-origin requests are not supported")
		return
	}
	if r.Method == "GET" && r.URL.Path == "/health" {
		send(w, 200, map[string]any{"status": "ok", "apiVersion": 1, "platform": runtime.GOOS, "workDir": s.root, "active": s.active.Load(), "pending": len(s.pending)})
		return
	}
	if r.Method == "POST" && (r.URL.Path == "/jobs" || r.URL.Path == "/convert") {
		s.submit(w, r)
		return
	}
	parts := strings.Split(strings.Trim(r.URL.Path, "/"), "/")
	if r.Method != "GET" || len(parts) < 2 || parts[0] != "jobs" || !idPattern.MatchString(parts[1]) {
		failure(w, 404, "Not found")
		return
	}
	job, err := s.read(parts[1])
	if err != nil {
		failure(w, 404, "Job not found")
		return
	}
	if len(parts) == 2 {
		send(w, 200, job)
		return
	}
	if len(parts) != 3 {
		failure(w, 404, "Not found")
		return
	}
	file := parts[2]
	if file == "output.pdf" && job.PDFStatus == "ok" {
		w.Header().Set("Content-Type", "application/pdf")
		w.Header().Set("Content-Disposition", `attachment; filename="output.pdf"`)
	} else if file == "dialog.png" && len(job.Dialogs) > 0 && job.Dialogs[0].Screenshot != "" {
		w.Header().Set("Content-Type", "image/png")
	} else {
		failure(w, 404, "Artifact not available")
		return
	}
	w.Header().Set("Cache-Control", "no-store")
	http.ServeFile(w, r, filepath.Join(job.Dir, file))
}

func (s *Server) submit(w http.ResponseWriter, r *http.Request) {
	filename := r.URL.Query().Get("filename")
	ext := strings.ToLower(filepath.Ext(filename))
	if filename == "" || len(filename) > 255 || strings.ContainsAny(filename, "/\\\x00\r\n") || (ext != ".docx" && ext != ".xlsx" && ext != ".pptx") {
		failure(w, 400, "filename must be a DOCX, XLSX or PPTX basename")
		return
	}
	seconds := 120
	if value := r.URL.Query().Get("timeout"); value != "" {
		var err error
		seconds, err = strconv.Atoi(value)
		if err != nil || seconds < 5 || seconds > 600 {
			failure(w, 400, "timeout must be 5..600 seconds")
			return
		}
	}
	select {
	case s.pending <- struct{}{}:
	default:
		failure(w, 429, "Queue full (16 jobs)")
		return
	}
	submitted := false
	defer func() {
		if !submitted {
			<-s.pending
		}
	}()
	var random [16]byte
	if _, err := rand.Read(random[:]); err != nil {
		failure(w, 500, err.Error())
		return
	}
	job := Job{ID: hex.EncodeToString(random[:]), Filename: filename, TimeoutSeconds: seconds, CreatedAt: time.Now().UTC(), Outcome: Outcome{Status: "queued", PDFStatus: "pending", Dialogs: []Dialog{}}}
	job.Dir = filepath.Join(s.root, "jobs", job.ID)
	if err := os.Mkdir(job.Dir, 0700); err != nil {
		failure(w, 500, err.Error())
		return
	}
	defer func() {
		if !submitted {
			_ = os.RemoveAll(job.Dir)
		}
	}()
	f, err := os.OpenFile(filepath.Join(job.Dir, "input"+ext), os.O_CREATE|os.O_EXCL|os.O_WRONLY, 0600)
	if err != nil {
		failure(w, 500, err.Error())
		return
	}
	hash := sha256.New()
	n, err := io.Copy(io.MultiWriter(f, hash), http.MaxBytesReader(w, r.Body, MaxUpload))
	closeErr := f.Close()
	if err != nil {
		var tooLarge *http.MaxBytesError
		if errors.As(err, &tooLarge) {
			failure(w, 413, "Upload limit is 100 MiB")
		} else {
			failure(w, 400, err.Error())
		}
		return
	}
	if closeErr != nil {
		failure(w, 500, closeErr.Error())
		return
	}
	if n == 0 {
		failure(w, 400, "Empty upload")
		return
	}
	job.InputSHA256 = hex.EncodeToString(hash.Sum(nil))
	if err = writeJSON(filepath.Join(job.Dir, "result.json"), job); err != nil {
		failure(w, 500, err.Error())
		return
	}
	submitted = true
	s.queue <- job
	w.Header().Set("Location", "/jobs/"+job.ID)
	if r.URL.Path == "/jobs" {
		send(w, 202, job)
		return
	}
	// Disconnecting a synchronous caller does not cancel or lose its recorded job.
	ticker := time.NewTicker(200 * time.Millisecond)
	defer ticker.Stop()
	for {
		select {
		case <-r.Context().Done():
			return
		case <-ticker.C:
			result, err := s.read(job.ID)
			if err != nil {
				failure(w, 500, err.Error())
				return
			}
			if result.Status != "queued" && result.Status != "running" {
				send(w, 200, result)
				return
			}
		}
	}
}

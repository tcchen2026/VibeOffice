package office_test

import (
	"bytes"
	"context"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"io"
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"sync/atomic"
	"testing"
	"time"

	"vibeoffice/win_test_server/api"
)

func request(t *testing.T, method, url string, body []byte, want int) []byte {
	t.Helper()
	r, err := http.NewRequest(method, url, bytes.NewReader(body))
	if err != nil {
		t.Fatal(err)
	}
	resp, err := http.DefaultClient.Do(r)
	if err != nil {
		t.Fatal(err)
	}
	defer resp.Body.Close()
	b, err := io.ReadAll(resp.Body)
	if err != nil {
		t.Fatal(err)
	}
	if resp.StatusCode != want {
		t.Fatalf("%s: %d, want %d: %s", url, resp.StatusCode, want, b)
	}
	return b
}

func readJob(t *testing.T, b []byte) api.Job {
	t.Helper()
	var job api.Job
	if err := json.Unmarshal(b, &job); err != nil {
		t.Fatal(err)
	}
	return job
}

func TestConversionAndArtifacts(t *testing.T) {
	input := []byte("OOXML fixture")
	pdf := []byte("%PDF-1.7\nfixture")
	digest := sha256.Sum256(input)
	run := func(ctx context.Context, j api.Job) api.Outcome {
		data, err := os.ReadFile(filepath.Join(j.Dir, "input.docx"))
		if err != nil || !bytes.Equal(data, input) {
			t.Errorf("input changed: %q %v", data, err)
		}
		deadline, ok := ctx.Deadline()
		if !ok || time.Until(deadline) > 5*time.Second {
			t.Error("job timeout missing")
		}
		if j.InputSHA256 != hex.EncodeToString(digest[:]) {
			t.Error("input hash mismatch")
		}
		if err := os.WriteFile(filepath.Join(j.Dir, "output.pdf"), pdf, 0600); err != nil {
			t.Error(err)
		}
		yes := true
		return api.Outcome{Status: "ok", CanOpenWithoutRepair: &yes, InputUnchanged: true, PDFStatus: "ok", PDF: "/jobs/" + j.ID + "/output.pdf"}
	}
	h, err := api.New(t.TempDir(), run)
	if err != nil {
		t.Fatal(err)
	}
	s := httptest.NewServer(h)
	defer s.Close()
	job := readJob(t, request(t, "POST", s.URL+"/convert?filename=spaces%20%26%20quotes.docx&timeout=5", input, 200))
	if job.Status != "ok" || job.CanOpenWithoutRepair == nil || !*job.CanOpenWithoutRepair || job.Message != "The file can be opened with Office 2016 without repair prompts." {
		t.Fatalf("bad result: %+v", job)
	}
	if got := request(t, "GET", s.URL+job.PDF, nil, 200); !bytes.Equal(got, pdf) {
		t.Fatal("PDF bytes changed")
	}
	saved := readJob(t, request(t, "GET", s.URL+"/jobs/"+job.ID, nil, 200))
	if saved.InputSHA256 != job.InputSHA256 || saved.Status != "ok" {
		t.Fatal("result not persisted")
	}
	request(t, "GET", s.URL+"/jobs/"+job.ID+"/input.docx", nil, 404)
	request(t, "GET", s.URL+"/jobs/"+job.ID+"/worker.ps1", nil, 404)
}

func TestSerialQueueAndFailureResults(t *testing.T) {
	var active, maximum atomic.Int32
	root := t.TempDir()
	run := func(ctx context.Context, j api.Job) api.Outcome {
		n := active.Add(1)
		if n > maximum.Load() {
			maximum.Store(n)
		}
		defer active.Add(-1)
		time.Sleep(30 * time.Millisecond)
		if j.Filename == "repair.xlsx" {
			no := false
			return api.Outcome{Status: "repair_required", CanOpenWithoutRepair: &no, PDFStatus: "not_attempted", Dialogs: []api.Dialog{{Text: "Repair content?"}}}
		}
		return api.Outcome{Status: "timeout", PDFStatus: "not_attempted"}
	}
	h, err := api.New(root, run)
	if err != nil {
		t.Fatal(err)
	}
	s := httptest.NewServer(h)
	defer s.Close()
	var jobs []api.Job
	for _, name := range []string{"repair.xlsx", "blocked.pptx", "timeout.docx"} {
		jobs = append(jobs, readJob(t, request(t, "POST", s.URL+"/jobs?filename="+name, []byte("input"), 202)))
	}
	for i, job := range jobs {
		deadline := time.Now().Add(3 * time.Second)
		for {
			job = readJob(t, request(t, "GET", s.URL+"/jobs/"+job.ID, nil, 200))
			if job.Status != "queued" && job.Status != "running" {
				break
			}
			if time.Now().After(deadline) {
				t.Fatal("job stuck")
			}
			time.Sleep(10 * time.Millisecond)
		}
		if job.Message != "" || job.PDF != "" {
			t.Fatalf("failure became success: %+v", job)
		}
		if i == 0 && (job.CanOpenWithoutRepair == nil || *job.CanOpenWithoutRepair) {
			t.Fatal("repair accepted")
		}
		if i != 0 && job.CanOpenWithoutRepair != nil {
			t.Fatal("timeout inferred acceptance")
		}
		request(t, "GET", s.URL+"/jobs/"+job.ID+"/output.pdf", nil, 404)
	}
	if maximum.Load() != 1 {
		t.Fatal("Office workers overlapped")
	}
}

func TestValidationAndInterruptedJob(t *testing.T) {
	root := t.TempDir()
	id := "0123456789abcdef0123456789abcdef"
	dir := filepath.Join(root, "jobs", id)
	if err := os.MkdirAll(dir, 0700); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, "result.json"), []byte(`{"id":"`+id+`","status":"running","pdfStatus":"pending"}`), 0600); err != nil {
		t.Fatal(err)
	}
	h, err := api.New(root, func(context.Context, api.Job) api.Outcome {
		t.Error("invalid input reached Office")
		return api.Outcome{}
	})
	if err != nil {
		t.Fatal(err)
	}
	s := httptest.NewServer(h)
	defer s.Close()
	for _, query := range []string{"filename=x.exe", "filename=..%2Fx.docx", "filename=C%3A%5Cx.docx", "filename=x.xlsx&timeout=0", "filename=x.xlsx&timeout=601", ""} {
		request(t, "POST", s.URL+"/jobs?"+query, []byte("input"), 400)
	}
	request(t, "POST", s.URL+"/jobs?filename=empty.docx", nil, 400)
	r := httptest.NewRequest("POST", "/jobs?filename=x.docx", bytes.NewReader([]byte("input")))
	r.Header.Set("Origin", "https://example.invalid")
	w := httptest.NewRecorder()
	h.ServeHTTP(w, r)
	if w.Code != 403 {
		t.Fatal("cross-origin submission allowed")
	}
	job := readJob(t, request(t, "GET", s.URL+"/jobs/"+id, nil, 200))
	if job.Status != "error" || job.CanOpenWithoutRepair != nil {
		t.Fatal("interrupted job became accepted")
	}
}

func TestDialogEvidence(t *testing.T) {
	for _, c := range []struct{ title, text, want string }{
		{"Microsoft PowerPoint", "PowerPoint found a problem with content. Repair?", "repair_required"},
		{"Microsoft Excel", "我们发现部分内容有问题。是否让我们尽量尝试恢复？", "repair_required"},
		{"Microsoft Word", "Word 發現無法讀取的內容", "repair_required"},
		{"密码", "Password required", "blocked"},
		{"Microsoft Word", "很抱歉，无法打开 input，因为内容有问题。", "open_failed"},
		{"Unknown Office dialog", "", "blocked"},
		{"Publishing...", "Writing page 300", "progress"},
		{"Publishing...", "Repair required", "repair_required"},
	} {
		if got := api.ClassifyDialog(api.Dialog{Title: c.title, Text: c.text}); got != c.want {
			t.Errorf("%q: %s, want %s", c.title, got, c.want)
		}
	}
}

//go:build !windows

package api

import "context"

func RunOffice(ctx context.Context, job Job) Outcome {
	return Outcome{Status: "error", PDFStatus: "not_attempted", Error: "Office conversion requires Windows and an interactive Office 2016 desktop."}
}

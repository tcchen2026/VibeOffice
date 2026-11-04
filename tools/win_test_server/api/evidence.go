package api

import "regexp"

// This VM's Office UI is Chinese; unknown/localized dialogs still block acceptance.
var repairWords = regexp.MustCompile(`(?i)repair|recover|unreadable|corrupt|problem.*content|removed records|repaired records|修复|修復|恢复|復原|損毀|损坏|无法读取|無法讀取`)
var progressTitle = regexp.MustCompile(`(?i)^(publishing|exporting|printing|正在发布|正在导出|正在打印)([ .:：…]|$)`)
var cannotOpen = regexp.MustCompile(`(?i)cannot open|can't open|unable to open|could not be opened|无法打开|無法開啟|不能打开|不能開啟`)

func ClassifyDialog(d Dialog) string {
	if repairWords.MatchString(d.Title + "\n" + d.Text) {
		return "repair_required"
	}
	if cannotOpen.MatchString(d.Text) {
		return "open_failed"
	}
	if progressTitle.MatchString(d.Title) {
		return "progress"
	}
	return "blocked"
}

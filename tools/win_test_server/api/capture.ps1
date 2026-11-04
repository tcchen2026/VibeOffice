# Observe only; never press a dialog button or accept repairs.
param([long]$Handle, [string]$Directory)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.Drawing, Accessibility
# Microsoft's documented accessibility API, used for Office's legacy bosa_sdm dialogs:
# https://learn.microsoft.com/en-us/windows/win32/api/oleacc/nf-oleacc-accessibleobjectfromwindow
Add-Type -ReferencedAssemblies ([Accessibility.IAccessible].Assembly.Location) -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public static class DialogCapture {
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hwnd, IntPtr hdc, uint flags);
    [DllImport("oleacc.dll")] static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint id, ref Guid iid, [MarshalAs(UnmanagedType.Interface)] out object result);
    static void Read(Accessibility.IAccessible a, object child, int depth, List<string> text) {
        if (depth > 8 || text.Count > 1000) return;
        try { string name = a.get_accName(child); if (!String.IsNullOrEmpty(name)) text.Add(name); } catch { }
        if (!Object.Equals(child, 0)) return;
        int count = 0;
        try { count = Math.Min(a.accChildCount, 1000); } catch { }
        for (int i = 1; i <= count; i++) {
            Accessibility.IAccessible nested = null;
            try { nested = a.get_accChild(i) as Accessibility.IAccessible; } catch { }
            if (nested != null) Read(nested, 0, depth + 1, text);
            else Read(a, i, depth + 1, text);
        }
    }
    public static string[] Names(IntPtr hwnd) {
        Guid iid = typeof(Accessibility.IAccessible).GUID;
        object result;
        var text = new List<string>();
        if (AccessibleObjectFromWindow(hwnd, 0xFFFFFFFC, ref iid, out result) == 0)
            Read((Accessibility.IAccessible)result, 0, 0, text);
        return text.ToArray();
    }
}
'@
$element = [System.Windows.Automation.AutomationElement]::FromHandle([IntPtr]$Handle)
$texts = New-Object 'System.Collections.Generic.List[string]'
$texts.Add($element.Current.Name)
$all = $element.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)
foreach ($item in $all) { if ($item.Current.Name) { $texts.Add($item.Current.Name) } }
foreach ($name in [DialogCapture]::Names([IntPtr]$Handle)) { $texts.Add($name) }
[IO.File]::WriteAllText("$Directory\dialog.txt", (($texts | Select-Object -Unique) -join "`n"), (New-Object System.Text.UTF8Encoding($false)))
$rect = $element.Current.BoundingRectangle
if ($rect.Width -gt 0 -and $rect.Height -gt 0) {
    $bitmap = New-Object Drawing.Bitmap([int]$rect.Width, [int]$rect.Height)
    $graphics = [Drawing.Graphics]::FromImage($bitmap)
    try {
        # Capture this dialog even when another window covers it; never photograph unrelated desktop content.
        $dc = $graphics.GetHdc()
        try { $captured = [DialogCapture]::PrintWindow([IntPtr]$Handle, $dc, 2) }
        finally { $graphics.ReleaseHdc($dc) }
        if ($captured) { $bitmap.Save("$Directory\dialog.png", [Drawing.Imaging.ImageFormat]::Png) }
    } finally { $graphics.Dispose(); $bitmap.Dispose() }
}

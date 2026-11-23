// Run from the VibeOffice root with the system WKWebView. The native host fixture
// supplies a receiver so this check can run without a TV or a Local Network prompt.
import AppKit
import WebKit

func fail(_ message: String) -> Never {
    FileHandle.standardError.write(Data(("Mac cast check: " + message + "\n").utf8))
    exit(1)
}

let repo = URL(fileURLWithPath: FileManager.default.currentDirectoryPath, isDirectory: true)
let assets = repo.appendingPathComponent("public", isDirectory: true)
guard CommandLine.arguments.count == 2 else { fail("pass the output path for a captured H.264 keyframe") }
let output = URL(fileURLWithPath: CommandLine.arguments[1])

func script(_ name: String) -> String {
    do { return try String(contentsOf: repo.appendingPathComponent(name), encoding: .utf8) }
    catch { fail("\(name): \(error)") }
}

@MainActor final class SuiteFiles: NSObject, WKURLSchemeHandler {
    func webView(_ webView: WKWebView, start task: WKURLSchemeTask) {
        do {
            guard let url = task.request.url, url.host == "localhost" else { throw URLError(.badURL) }
            var file = assets.appendingPathComponent(url.path).standardizedFileURL
            if url.hasDirectoryPath { file.appendPathComponent("index.html") }
            guard file.path.hasPrefix(assets.path + "/") else { throw URLError(.badURL) }
            let data = try Data(contentsOf: file)
            let types = ["html": "text/html", "css": "text/css", "js": "text/javascript", "json": "application/json",
                         "svg": "image/svg+xml", "png": "image/png", "jpg": "image/jpeg", "woff2": "font/woff2"]
            let mime = types[file.pathExtension] ?? "application/octet-stream"
            let response = HTTPURLResponse(url: url, statusCode: 200, httpVersion: "HTTP/1.1",
                                           headerFields: ["Content-Type": mime, "Access-Control-Allow-Origin": "*"])!
            task.didReceive(response)
            task.didReceive(data)
            task.didFinish()
        } catch { task.didFailWithError(error) }
    }
    func webView(_ webView: WKWebView, stop task: WKURLSchemeTask) {}
}

@MainActor final class Runner: NSObject, WKNavigationDelegate {
    private var apps = ["quire", "ledger", "lectern"]
    private let exercise = script("tests/desktop/cast-exercise.js") + """

    try {
      if (!isSecureContext || VO.host.platform !== 'macos') throw new Error('Expected the secure macOS app origin');
      const app = await exerciseCast();
      const bytes = stubCast.frames.at(-1);
      let binary = '';
      for (let i = 0; i < bytes.length; i += 8192) binary += String.fromCharCode(...bytes.subarray(i, i + 8192));
      return { app, keyframe: btoa(binary) };
    } catch (e) { return { error: String(e) + ' ' + e.stack }; }
    """

    func next(_ view: WKWebView) {
        if apps.isEmpty { exit(0) }
        let app = apps.removeFirst()
        view.load(URLRequest(url: URL(string: "tauri://localhost/\(app)/?new&cast&appstore&lang=en")!))
    }
    func webView(_ view: WKWebView, didFinish navigation: WKNavigation!) {
        guard view.url?.host == "localhost" else { return }
        view.callAsyncJavaScript(exercise, arguments: [:], in: nil, in: .page) { result in
            do {
                let value = try result.get()
                guard let reply = value as? [String: Any] else { fail("unexpected WebKit result") }
                if let error = reply["error"] as? String { fail(error) }
                guard let app = reply["app"] as? String, let encoded = reply["keyframe"] as? String,
                      let bytes = Data(base64Encoded: encoded), !bytes.isEmpty else { fail("missing picture") }
                try bytes.write(to: output, options: .atomic)
                print("\(app) passed in macOS WKWebView (\(bytes.count) H.264 bytes)")
                self.next(view)
            } catch { fail(String(describing: error)) }
        }
    }
    func webView(_ view: WKWebView, didFailProvisionalNavigation navigation: WKNavigation!, withError error: Error) {
        fail(String(describing: error))
    }
}

let application = NSApplication.shared
application.setActivationPolicy(.accessory)
let config = WKWebViewConfiguration()
config.websiteDataStore = .nonPersistent()
config.setURLSchemeHandler(SuiteFiles(), forURLScheme: "tauri")
for name in ["tests/desktop/host-stub.js", "DLNAnything/web/canvas.js"] {
    config.userContentController.addUserScript(WKUserScript(source: script(name), injectionTime: .atDocumentStart, forMainFrameOnly: true))
}
let bounds = NSRect(x: 0, y: 0, width: 1280, height: 800)
let window = NSWindow(contentRect: bounds, styleMask: [.titled, .resizable], backing: .buffered, defer: false)
let view = WKWebView(frame: bounds, configuration: config)
let runner = Runner()
view.navigationDelegate = runner
window.contentView = view
window.makeKeyAndOrderFront(nil)
application.activate(ignoringOtherApps: true)
runner.next(view)
DispatchQueue.global().asyncAfter(deadline: .now() + 90) { fail("timed out waiting for WebKit") }
application.run()

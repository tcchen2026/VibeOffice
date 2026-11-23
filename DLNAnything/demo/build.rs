use std::{env, fs, path::PathBuf};

fn main() {
    let root = PathBuf::from(env::var_os("CARGO_MANIFEST_DIR").unwrap());
    let output = root.join("generated-ui");
    fs::create_dir_all(&output).unwrap();
    for (src, dst) in [
        ("../web/canvas.js", "canvas.js"),
        ("ui/index.html", "index.html"),
        ("ui/demo.js", "demo.js"),
        ("ui/demo.css", "demo.css"),
    ] {
        println!("cargo:rerun-if-changed={src}");
        let bytes = fs::read(root.join(src)).unwrap();
        let target = output.join(dst);
        if fs::read(&target).ok().as_ref() != Some(&bytes) {
            fs::write(target, bytes).unwrap();
        }
    }
    tauri_build::build()
}

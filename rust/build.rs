//! TaskDialogIndirect lives in Common Controls v6, which a console exe only gets
//! through an embedded manifest dependency. Done with MSVC linker flags, no extra crate.
fn main() {
    if std::env::var("CARGO_CFG_TARGET_ENV").as_deref() == Ok("msvc") {
        let dep = "/MANIFESTDEPENDENCY:type='win32' name='Microsoft.Windows.Common-Controls' \
                   version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' \
                   language='*'";
        // All targets: the lib's unit-test exe imports the dialog code too.
        println!("cargo:rustc-link-arg=/MANIFEST:EMBED");
        println!("cargo:rustc-link-arg={dep}");
    }
    println!("cargo:rerun-if-changed=build.rs");
}

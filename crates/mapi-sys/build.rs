fn main() -> Result<()> {
    mapi32_link::update_rustc_flags(mapi32_link::output_libs(
        mapi32_path::get_manifest_dir()?,
        mapi32_path::get_out_dir()?,
    )?)
}

#[derive(Debug, Error)]
pub enum Error {
    #[error(transparent)]
    Io(#[from] std::io::Error),
    #[error(transparent)]
    Var(#[from] std::env::VarError),
}

pub type Result<T> = std::result::Result<T, Error>;

#[macro_use]
extern crate thiserror;

mod mapi32_path {
    use std::{convert::From, env, path::PathBuf};

    pub fn get_out_dir() -> super::Result<PathBuf> {
        Ok(PathBuf::from(env::var("OUT_DIR")?))
    }

    pub fn get_manifest_dir() -> super::Result<PathBuf> {
        Ok(PathBuf::from(env::var("CARGO_MANIFEST_DIR")?))
    }
}

mod mapi32_link {
    use std::{env, path::PathBuf};

    pub fn output_libs(manifest_dir: PathBuf, out_dir: PathBuf) -> super::Result<PathBuf> {
        const MAPI32_LIBS: &[&str] = &["mapi32.lib", "olmapi32.lib"];
        const MAPI32_TARGETS: &[&str] = &["x64", "x86"];

        for target in MAPI32_TARGETS {
            for lib in MAPI32_LIBS {
                use std::fs;

                let mut lib_src = manifest_dir.clone();
                lib_src.push(target);
                lib_src.push(lib);

                let mut lib_dest = out_dir.clone();
                lib_dest.push(target);
                if !lib_dest.is_dir() {
                    fs::create_dir(lib_dest.as_path())?;
                }

                lib_dest.push(lib);
                eprintln!("Copy from {:?} -> {:?}", lib_src, lib_dest);
                fs::copy(lib_src.as_path(), lib_dest.as_path())?;
            }
        }

        Ok(out_dir)
    }

    pub fn update_rustc_flags(lib_path: PathBuf) -> super::Result<()> {
        let target_arch = match env::var("CARGO_CFG_TARGET_ARCH")?.as_str() {
            "x86_64" => "x64",
            "x86" => "x86",
            unimplemented => unimplemented!(
                "`{}` architecture set by `CARGO_CFG_TARGET_ARCH`",
                unimplemented
            ),
        };
        let mut lib_path = lib_path;
        lib_path.push(target_arch);

        match lib_path.to_str() {
            Some(path) if lib_path.exists() => println!("cargo:rustc-link-search=native={path}"),
            _ => unimplemented!("`{}` is not supported by MAPI", target_arch),
        };

        Ok(())
    }
}

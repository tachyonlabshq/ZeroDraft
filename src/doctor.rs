use anyhow::Result;
use serde::Serialize;
use std::env;

use crate::find_executable;

#[derive(Debug, Clone, Serialize)]
pub struct DoctorCheck {
    pub name: String,
    pub ok: bool,
    pub detail: String,
}

#[derive(Debug, Clone, Serialize)]
pub struct DoctorReport {
    pub status: String,
    pub current_exe: String,
    pub cwd: String,
    pub platform: String,
    pub soffice_path: Option<String>,
    pub checks: Vec<DoctorCheck>,
}

pub fn doctor_environment() -> Result<DoctorReport> {
    let current_exe = env::current_exe()?;
    let cwd = env::current_dir()?;
    let soffice = find_executable("soffice").or_else(|| find_executable("libreoffice"));

    let checks = vec![
        DoctorCheck {
            name: "current_executable".to_string(),
            ok: true,
            detail: current_exe.display().to_string(),
        },
        DoctorCheck {
            name: "working_directory".to_string(),
            ok: true,
            detail: cwd.display().to_string(),
        },
        DoctorCheck {
            name: "libreoffice_for_doc_conversion".to_string(),
            ok: soffice.is_some(),
            detail: soffice
                .as_ref()
                .map(|path| path.display().to_string())
                .unwrap_or_else(|| "not found on PATH".to_string()),
        },
    ];

    Ok(DoctorReport {
        status: "success".to_string(),
        current_exe: current_exe.display().to_string(),
        cwd: cwd.display().to_string(),
        platform: format!("{}-{}", env::consts::OS, env::consts::ARCH),
        soffice_path: soffice.map(|path| path.display().to_string()),
        checks,
    })
}

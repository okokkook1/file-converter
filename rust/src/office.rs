use std::{
    env, fs,
    os::windows::process::CommandExt,
    path::{Path, PathBuf},
    process::Command,
};

use anyhow::{Context, Result, ensure};
use url::Url;

use crate::{extension, write_output};

pub fn find_libreoffice() -> Option<PathBuf> {
    env::var_os("LIBREOFFICE_PATH")
        .map(PathBuf::from)
        .filter(|path| path.is_file())
        .or_else(|| {
            let path = env::var_os("PATH").unwrap_or_default();
            env::split_paths(&path)
                .chain([
                    PathBuf::from(r"C:\Program Files\LibreOffice\program"),
                    PathBuf::from(r"C:\Program Files (x86)\LibreOffice\program"),
                ])
                .flat_map(|dir| ["soffice.com", "soffice.exe"].map(|name| dir.join(name)))
                .find(|path| path.is_file())
        })
}

fn file_url(path: &Path) -> Result<String> {
    // URL conversion on Windows needs a drive path instead of a verbatim Win32 path.
    let absolute = std::path::absolute(path)?;
    Url::from_file_path(&absolute)
        .map(String::from)
        .map_err(|()| anyhow::anyhow!("無法轉為檔案網址：{}", path.display()))
}

fn command(profile: &Path) -> Result<Command> {
    let executable = find_libreoffice()
        .context("找不到 LibreOffice。請安裝 LibreOffice 或設定 LIBREOFFICE_PATH")?;
    let mut command = Command::new(executable);
    command
        .creation_flags(0x0800_0000)
        .arg(format!("-env:UserInstallation={}", file_url(profile)?))
        .args(["--headless", "--nologo", "--norestore", "--nodefault"]);
    Ok(command)
}

fn run(mut command: Command, produced: &Path) -> Result<()> {
    let result = command.output().context("無法啟動 LibreOffice")?;
    ensure!(
        result.status.success() && produced.metadata().is_ok_and(|metadata| metadata.len() > 0),
        "LibreOffice 未產生輸出\n{}\n{}",
        String::from_utf8_lossy(&result.stdout).trim(),
        String::from_utf8_lossy(&result.stderr).trim()
    );
    Ok(())
}

fn filter(extension: &str) -> Result<&'static str> {
    Ok(match extension {
        "doc" => "MS Word 97",
        "docx" => "Office Open XML Text",
        "odt" => "writer8",
        "pdf" => "writer_pdf_Export",
        _ => anyhow::bail!("不支援的 LibreOffice 輸出格式 .{extension}"),
    })
}

pub fn convert(input: &Path, output: &Path) -> Result<()> {
    ensure!(input.is_file(), "找不到輸入檔案：{}", input.display());
    let temp = tempfile::tempdir()?;
    let ext = extension(output);
    let mut command = command(&temp.path().join("profile"))?;
    if extension(input) == "pdf" {
        command.arg("--infilter=writer_pdf_import");
    }
    command
        .args(["--convert-to", &format!("{ext}:{}", filter(&ext)?)])
        .arg("--outdir")
        .arg(temp.path())
        .arg(std::path::absolute(input)?);
    let produced = temp
        .path()
        .join(input.file_name().context("缺少輸入檔名")?)
        .with_extension(&ext);
    run(command, &produced)?;
    write_output(output, |path| {
        fs::copy(produced, path)?;
        Ok(())
    })
}

pub fn merge_docx(inputs: &[PathBuf], output: &Path) -> Result<()> {
    let temp = tempfile::tempdir()?;
    let profile = temp.path().join("profile");
    let module = profile.join("user/basic/Standard/Module1.xba");
    let initialized = command(&profile)?.arg("--terminate_after_init").output()?;
    ensure!(
        initialized.status.success() && module.is_file(),
        "無法初始化 LibreOffice 暫存設定"
    );
    fs::write(module, include_str!("office/Merge.xba"))?;
    let produced = temp.path().join("merged.docx");
    let mut job = vec![file_url(&produced)?];
    for path in inputs {
        ensure!(path.is_file(), "找不到輸入檔案：{}", path.display());
        job.push(file_url(path)?);
    }
    let job_path = temp.path().join("job.txt");
    fs::write(&job_path, format!("{}\n", job.join("\n")))?;
    let error_path = temp.path().join("error.txt");
    let mut command = command(&profile)?;
    command
        .env("FILE_CONVERTER_JOB", file_url(&job_path)?)
        .env("FILE_CONVERTER_ERROR", file_url(&error_path)?)
        .arg("vnd.sun.star.script:Standard.Module1.RunMerge?language=Basic&location=application");
    run(command, &produced).with_context(|| {
        fs::read_to_string(&error_path).unwrap_or_else(|_| "文件合併失敗".into())
    })?;
    write_output(output, |path| {
        fs::copy(produced, path)?;
        Ok(())
    })
}

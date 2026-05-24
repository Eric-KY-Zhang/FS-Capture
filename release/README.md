# Release Artifacts

每次正式版本发布的 Windows zip 包放在这个目录。**zip 文件本身被 `.gitignore` 忽略**，只这份 README 被 track（让目录在 git 里持久存在）。

## 命名规范

```
FilingsAtlas-v{major}.{minor}.{patch}-windows.zip
```

例：
- `FilingsAtlas-v1.0.0-windows.zip`
- `FilingsAtlas-v1.1.0-windows.zip`

## 打包流程

每次新版本发布前的标准 3 步：

```powershell
cd "E:\Claude+CODEX Project\FS Capture\development"

# 1. 用 PyInstaller 重新打包源代码（产物：根目录的 Filings Atlas.exe + _internal/）
$empty = New-TemporaryFile
Start-Process -FilePath "$PWD\build.bat" `
    -WorkingDirectory $PWD `
    -Wait -NoNewWindow -RedirectStandardInput $empty.FullName
Remove-Item $empty.FullName

# 2. 把根目录的 EXE + _internal/ 压缩到 release/
cd ..
Compress-Archive `
    -Path ".\Filings Atlas.exe",".\_internal\*" `
    -DestinationPath ".\release\FilingsAtlas-v{X.Y.Z}-windows.zip" `
    -CompressionLevel Optimal -Force

# 3. push tag + 上传到 GitHub Release
git push origin v{X.Y.Z}
gh release create v{X.Y.Z} `
    ".\release\FilingsAtlas-v{X.Y.Z}-windows.zip" `
    --title "Filings Atlas v{X.Y.Z} — 全球披露图谱" `
    --notes-file CHANGELOG.md
```

## 已发布版本

| 版本 | zip 路径 | 大小 | 发布日期 | GitHub Release |
|---|---|---|---|---|
| v1.0.0 | `release/FilingsAtlas-v1.0.0-windows.zip` | ~187 MB | 2026-05-24 | https://github.com/Eric-KY-Zhang/FS-Capture/releases/tag/v1.0.0 |

## 历史版本（v0.x 内部迭代）

v0.6.1 / v0.7 / v0.8 / v0.9 是内部迭代，**未发布 GitHub Release**（按 `CLAUDE.md` § 10 发布策略），无对应 zip。

## 与 PyInstaller `dist/` 的关系

`development/dist/Filings Atlas/` 是 PyInstaller 的中间产物，**不入 git 也不直接分发**。这个目录的 zip 是从根目录 `Filings Atlas.exe` + `_internal/`（build.bat 拷贝过来的）二次打包而成。

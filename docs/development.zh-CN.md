# 开发与发布

English version: [development.md](development.md)

本文档汇总构建、发布和开发环境安装细节，这些内容不会放在根目录 `README.md` 中。

## 环境要求

- Windows
- .NET SDK 10 或更高版本

## 构建

```powershell
dotnet build .\OfficeCopyAsMarkdown.slnx
```

## 发布

发布推荐配置：

```powershell
.\scripts\publish.ps1
```

该命令会同时生成：

- framework-dependent 发布目录 `.\artifacts\publish\win-x64-framework-dependent`
- 当前用户安装包目录 `.\artifacts\installer`

如果没有安装 Inno Setup 6，可以先安装 `ISCC.exe`，或者跳过安装包步骤：

```powershell
.\scripts\publish.ps1 -SkipInstaller
```

其他发布方式：

```powershell
.\scripts\publish.ps1 -SelfContained
.\scripts\publish.ps1 -SelfContained -SingleFile
```

基于现有发布目录单独构建安装包：

```powershell
.\scripts\build-installer.ps1
```

## 开发环境安装

用于开发或手动部署时，可以使用辅助脚本安装发布产物：

```powershell
.\scripts\install.ps1
```

可选：创建开机自启动快捷方式：

```powershell
.\scripts\install.ps1 -Startup
```

脚本会把文件复制到：

```text
%LOCALAPPDATA%\Programs\OfficeCopyAsMarkdown
```

如果要显式安装 self-contained 配置：

```powershell
.\scripts\install.ps1 -Profile win-x64-self-contained
.\scripts\install.ps1 -Profile win-x64-self-contained-single-file
```

## 打包说明

- 默认推荐 framework-dependent 配置。
- self-contained 单文件版本更大，也更容易触发杀软启发式检测。

## Microsoft 参考链接

- [OneNote add-ins documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/onenote/)
- [Custom keyboard shortcuts in Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/design/keyboard-shortcuts)
- [HTML Clipboard Format](https://learn.microsoft.com/en-us/windows/win32/dataxchg/html-clipboard-format)

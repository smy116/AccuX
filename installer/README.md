# 安装脚本验证

`Build-Installer.ps1` 是本地和 GitHub Actions 共用的唯一安装包生成入口，使用 Inno Setup 6.7.3：

```powershell
pwsh -NoProfile -File .\installer\Build-Installer.ps1 `
  -Version 1.7.1 `
  -FileVersion 1.7.1.0

pwsh -NoProfile -File .\installer\Build-Installer.ps1 `
  -Version 1.7.1-ci.37.d202798 `
  -FileVersion 1.7.1.37
```

脚本会调用 `installer\AccuX.iss`，检查生成物的 MZ/PE 头和版本资源，并生成安装包、`.sha256` 校验文件和 `.manifest.json` 清单。清单记录显示/安装器/文件版本、提交 SHA、运行号、文件名、大小和哈希；`OutputDirectory` 和 `OutputBaseFilename` 可用于 CI 或隔离测试输出。

版本字段分工如下：

| 字段 | 正式版 | 预发布版 | 用途 |
| --- | --- | --- | --- |
| 显示版本 | `1.7.1` | `1.7.1-ci.37.d202798` | 安装向导、文件版本文本、清单 |
| 安装器数字版本 | `1.7.1` | `1.7.1` | Inno `AppVersion`，用于升级比较 |
| 文件版本 | `1.7.1.0` | `1.7.1.37` | PE 固定 `FileVersion` / `ProductVersion` |

预发布后缀不会进入 Windows 的固定数字版本字段；这避免部分系统无法双击启动预发布安装包。正式包名称仍为 `AccuXSetup-<版本>.exe`，测试包名称保留 `ci` 和短 SHA。

GitHub Actions 使用 `actions/upload-artifact@v7` 的 `archive: false` 分别上传 EXE、`.sha256` 和 manifest，因此 Actions Artifact 列表会直接显示可双击的 `.exe`，不再需要从 ZIP 中取出安装包。Release job 使用 `actions/download-artifact@v8` 下载并重新校验这些原始文件。

Windows 10/11 下载文件可能带有 Mark-of-the-Web；由于当前没有代码签名证书，附件管理器或 SmartScreen 仍可能显示阻止或安全提示。请按系统提示检查来源后运行。

## 前置条件

- 安装包仅支持通过安装向导完成的常规安装，不处理静默安装参数。
- 缺少 .NET Framework 4.8 时，安装向导会询问是否联网下载并启动常规安装。3010、1641 始终要求重启，即使 Release 注册表已经更新；同一次安装不能绕过此状态。
- 下载 .NET Framework 4.8 引导程序不做 SHA-256 校验：微软重发该引导程序会改变文件哈希，固定哈希会使安装硬失败。完整性依赖 HTTPS 与微软官方 fwlink；下载失败（网络错误等）仍会给出人工安装提示。
- 返回 0 后仍未检测到运行时，作为安装失败处理，不推断为成功或需要重启。
- 准备阶段失败返回 7，需要重启返回 8；COM 注册错误触发安装撤销并返回 4。这些返回值已在 Inno Setup 6.7.3 下验证。

安装默认启用日志。卸载可传 `/LOG="AccuX-uninstall.log"`，记录 RegAsm 输出与退出码。

## COM 失败恢复

在 `ssInstall` 中备份将被覆盖的旧文件。全部文件落盘后，第一个 `[Registry]` 条目的代码常量调用 RegAsm，调用前备份两个注册表视图中的 AccuX 专属键树。必须保留此入口：普通 BeforeInstall/AfterInstall 回调抛异常时，Inno 可能显示错误后继续安装。

注册或备份失败会中止安装；原生文件撤销完成后，`DeinitializeSetup` 恢复旧文件与注册表快照。首次安装删除本次新增的 COM 键；升级保留原来的 CodeBase、版本子键与其他旧值。失败的 RegAsm 本身也可能有部分写入，因此恢复不只覆盖已成功的位数。恢复失败会记录日志并提示修复。

此机制处理正常进程执行期间的失败，不是断电或强制终止进程后的持久事务。修改 `[Files]` 时需同步 `ssInstall` 的备份清单；修改公开 COM 类型时需同步 `BackupComView` 的根键清单。

## 隔离回归测试

```powershell
pwsh -NoProfile -File .\installer\Test-Installer.ps1
```

测试要求 Windows x64、Inno Setup 6.7.3 和已构建的 Release 程序集。它检查生产 Pascal 代码的 COM 事务入口与 `[Files]` / `BackupPackageFile` 清单一致，并在独立临时目录编译一份验证安装包；测试不会安装、注册或卸载真实 AccuX。

- 检查 COM 注册失败入口、失败恢复入口和提交标记仍然存在。
- 检查 `[Files]` 中每个 Release 输入文件都存在，并且都有对应的旧文件备份项。
- 编译验证安装包，确保版本参数和安装脚本在 CI 环境中可用。

产物与日志保存至 `artifacts/installer-tests/<随机编号>/`。真实 .NET 安装/重启、管理员权限下的实际 RegAsm、Excel/WPS 加载及卸载仍需在测试机验收。

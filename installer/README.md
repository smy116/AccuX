# 安装脚本验证

使用 Inno Setup 6.7.3 编译验证：

```powershell
& 'C:\Program Files (x86)\Inno Setup 6\ISCC.exe' installer\AccuX.iss
```

正式版本使用两段版本号，例如 `v1.4`。CI 会把 tag 转换为 `1.4`，并通过 Inno 预处理参数生成 `AccuXSetup-1.4.exe`；程序集和安装器文件版本使用 `1.4.0.0`。

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

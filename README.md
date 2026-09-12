# AccuX V1

面向财务人员的 Windows Excel / WPS 表格效率插件（COM Add-in）。

V1 现包含四个财务功能、一个工作簿目录功能、一个标记模块和一个区域对比模块：

1. **一键舍入** — 对选区金额统一四舍五入，每次点击弹出小数位窗口（默认 2 位）。
2. **金额折合** — 按除百 / 除千 / 除万折合选区金额，可选添加“万”字。
3. **选区求和** — 汇总选区内可见数字，四舍五入到 2 位，并通过对话框复制金额、万元金额或中文大写金额。
4. **金额大写** — 将选区金额转换为中文大写，原地覆盖选区。
5. **生成目录** — 在工作簿最前面生成可见工作表目录，并为名称创建内部超链接；若已存在“目录”工作表，先确认是否删除并重新生成。
6. **标记** — 将当前选区中的可见单元格底色标记为绿 / 红 / 黄 / 蓝。
7. **区域对比** — 选择两个区域执行严格的单元格存在对比，查看区域独有项与相同项，并支持标记、清除和导出。
8. **设置** — 统一管理版本信息、单元格处理阈值和 jsDelivr 升级检测。

核心目标不是功能数量，而是建立稳定、可扩展的基础架构：Excel/WPS 宿主差异层、明确的 Core/Host 单向依赖、统一 Range 操作管线、正确的数值/公式区分、模块内聚的业务算法。

## 环境要求

| 项 | 要求 |
| --- | --- |
| 操作系统 | Windows 10 / 11 |
| 运行时 | .NET Framework 4.8 |
| 宿主 | Microsoft Excel（已验证环境：Office 16 x64）或 WPS 表格 |
| 构建 | Visual Studio 2022 或 .NET SDK 9（含 MSBuild） |

## 解决方案结构

```
AccuX.sln
├─ Directory.Build.props        # net48 / C# 9 / 公共属性
├─ build.ps1                    # 还原 + 构建 + 测试 + 依赖边界校验
├─ src/
│  ├─ AccuX.Core                # 模块契约、命令分发、单元格分类、Range 管线、配置、日志
│  ├─ AccuX.Host                # 宿主适配：差异、批量读写、公式规范化、宿主状态；Features/ 为功能级 COM 实现
│  ├─ AccuX.Modules.BasicFinance# 基础功能 + 纯 C# 算法 + WPF 界面
│  ├─ AccuX.Modules.Mark        # 选区可见单元格底色标记
│  ├─ AccuX.Modules.Compare     # 两区域单元格存在对比
│  └─ AccuX.AddIn               # COM 入口、Ribbon、组合根、显式模块注册、图标
├─ tests/
│  ├─ AccuX.Core.Tests
│  ├─ AccuX.Modules.BasicFinance.Tests
│  └─ AccuX.Modules.Mark.Tests
├─ tools/                       # register / unregister / make-icons
├─ installer/AccuX.iss          # Inno Setup 安装工程
├─ .github/workflows/            # 推送构建与 tag 发布
└─ docs/CompatibilityMatrix.md  # 兼容性矩阵
```

### 依赖方向（编码强制规则）

```
AccuX.Core  ── 不引用 Host，不引用 Excel/WPS Interop
AccuX.Host  ── 引用 Core，实现 IRangeOperationHost / IHostStateScope 及功能窄接口（目录 / 批注 / 标记）
AccuX.Modules.*  ── 引用 Core，不直接解决宿主差异
AccuX.AddIn ── 引用 Host + Modules + Core，作为 Composition Root
```

`build.ps1` 会在每次构建后反射校验：`AccuX.Core` 不得引用任何 Interop / Office / WPS 程序集，也不得引用 `AccuX.Host`。

## 构建与测试

```powershell
# 一条命令完成还原、构建、测试、边界校验
powershell -ExecutionPolicy Bypass -File build.ps1

# Release 构建（默认版本为 1.4）
powershell -ExecutionPolicy Bypass -File build.ps1 -Configuration Release -Version 1.4

# 或手动
dotnet build AccuX.sln -c Debug
dotnet test AccuX.sln -c Debug
```

当前状态：解决方案编译 0 警告 0 错误；单元测试 **217 项全部通过**（Core 78 项，BasicFinance 112 项，Mark 6 项，Compare 6 项，AddIn 15 项）。

构建脚本会优先使用 Visual Studio Office15 PIA，也支持通过 `-OfficePiaPath` 显式指定目录；找不到该目录时兼容使用 GAC 中的 15.0.0.0 PIA。构建结束会校验 `Microsoft.Office.Interop.Excel.dll`、`office.dll` 和 `Microsoft.Vbe.Interop.dll` 均已复制到 Release 输出目录。

## 开发期注册与 F5 调试

AccuX 是 COM Add-in，必须先注册才能被 Excel 加载。

```powershell
# 注册（HKCU，无需管理员）
powershell -ExecutionPolicy Bypass -File tools\register.ps1

# 注销
powershell -ExecutionPolicy Bypass -File tools\unregister.ps1
```

注册脚本使用 32 位与 64 位 `RegAsm /regfile` 分别生成注册表脚本，改写为当前用户根键后导入（真正的 HKCU 注册，无需管理员）。Excel 写入 `LoadBehavior=3` 的 ProgId 子键；WPS 表格额外在 `AddinsWL` 根键写入 ProgId 白名单值：

```
HKCU\Software\Microsoft\Office\Excel\Addins\AccuX.AddIn.Connect
HKCU\Software\Kingsoft\Office\ET\AddinsWL
  "AccuX.AddIn.Connect"=""
```

已验证：注册后 `Type.GetTypeFromProgID('AccuX.AddIn.Connect')` 可解析并实例化，`GetCustomUI` 返回正确 Ribbon XML，Ribbon 图标均可转换为 `IPictureDisp`。

> `RegAsm` 会对未签名的程序集给出 `/codebase` 警告（RA0000）。开发期可忽略；正式发布建议为程序集添加强名称（强名称对 COM 注册与加载更稳妥）。

### Visual Studio F5 调试

`AccuX.AddIn` 是类库（COM Add-in 的正确输出类型），不能“直接启动”。启动目标通过
`src/AccuX.AddIn/Properties/launchSettings.json` 中的 **Excel** 配置文件指定：

```json
{
  "profiles": {
    "Excel": {
      "commandName": "Executable",
      "executablePath": "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE"
    }
  }
}
```

步骤：

1. 先执行一次 `tools\register.ps1`。
2. 在“解决方案资源管理器”中右键 **AccuX.AddIn → 设为启动项目**。
3. 在工具栏的调试目标下拉框中选择 **Excel** 配置文件（若下拉框为空，重新打开解决方案即可刷新）。
4. 按 **F5**，Visual Studio 启动 Excel 并自动附加调试器。

> 如果 Excel 不在上述路径，直接修改 `executablePath`（例如 `C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE`）。
> 也可以在 **项目属性 → 调试 → 打开调试启动配置文件 UI** 中修改，效果等价。

可命中以下断点：`Connect.OnConnection` → `AddInCompositionRoot.Start` → Ribbon callback → `CommandDispatcher.Execute` → `RangeOperationPipeline.Execute` → `ExcelRangeOperationHost.Read/Write` → 业务 Transform。

> Clean / Rebuild 后程序集路径不变，无需重新注册；若更换输出目录或改为 Release，重新运行注册脚本即可。

#### 常见问题：无法直接启动带有“类库输出类型”的项目

该错误表示 VS 找不到启动目标，不是项目配置错误。按上述步骤选择 **Excel** 启动配置文件即可。
SDK 风格项目的启动目标只由 `launchSettings.json` 决定（CPS 项目系统不支持旧的
`StartAction` / `StartProgram` 项目属性）。

### WPS 调试

WPS 与 Excel 共用同一套程序集与核心代码。本机未安装 WPS，因此：

- 若 WPS 可作为 Visual Studio 外部启动程序，按上述方式配置；
- 否则先启动 WPS，再使用 **调试 → 附加到进程** 附加到 WPS 表格进程。

WPS 验证项在 `docs/CompatibilityMatrix.md` 中标记为「未验证」，需在装有 WPS 的机器上补齐。

## 配置

配置文件路径：`%AppData%\AccuX\config.json`。文件不存在时使用默认值，不会报错。

```json
{
  "settings": {
    "largeSelectionWarning": 100000,
    "maxProcessCells": 500000,
    "autoCheckForUpdates": true,
    "lastUpdateCheckUtc": null
  },
  "basicFinance": {
    "roundDigits": 2
  },
  "compare": {
    "firstOnlyColor": "#FFFF66",
    "secondOnlyColor": "#FFFF66",
    "sameColor": "#CCFFCC"
  }
}
```

| 字段 | 说明 |
| --- | --- |
| `roundDigits` | 一键舍入窗口的预填小数位（默认 2，窗口内仍可修改） |
| `largeSelectionWarning` | 选区超过该单元格数时弹出确认提示 |
| `maxProcessCells` | 选区超过该单元格数时直接拒绝处理 |
| `autoCheckForUpdates` | 插件启动后是否自动检查 jsDelivr 更新清单（默认开启，每 24 小时最多一次） |
| `lastUpdateCheckUtc` | 最近一次自动或手动检测时间，由插件维护 |

`settings` 是统一设置节。警告阈值和最大阈值同时作用于基础财务功能与区域对比，且必须大于 0、警告阈值不能超过最大阈值。旧版 `basicFinance` / `compare` 中的阈值会在首次加载时迁移到 `settings`，旧字段保留以便回滚到旧版本；`compare` 中的颜色仍保留在原配置节。

样例文件位于 `src/AccuX.AddIn/config.sample.json`。

## 升级检查

设置窗口可显示当前版本、手动检查 jsDelivr 更新清单，并在校验安装包 SHA-256 后启动普通 Inno Setup 安装程序。自动检查默认开启，插件启动后后台执行，每 24 小时最多一次；网络失败、清单错误或附件缺失只写入日志，不弹窗。发现新版本后会提示进入设置，安装前请保存工作并关闭 Excel/WPS。升级使用固定附件名 `AccuXSetup-{version}.exe` 与 `AccuXSetup-{version}.exe.sha256`，仅接受两段式稳定版本号（例如 `1.4`）。

插件运行时只访问 jsDelivr：`https://cdn.jsdelivr.net/gh/smy116/AccuX@update-feed/latest.json`。发布工作流会把 GitHub Release 中的安装包、SHA-256 文件和更新说明同步到 `update-feed` 分支，再由 jsDelivr 提供访问；jsDelivr 不可用时升级检测静默失败，不影响插件正常使用。

更新清单包含 `version`、`tag`、`name`、`notes`、`releaseNotesUrl`、`installerUrl` 和 `sha256Url`；其中安装包与校验文件固定存放在 `update-feed/releases/{version}/`，清单只引用 jsDelivr 地址。

## 日志

默认目录：`%AppData%\AccuX\logs\accux-yyyyMMdd.log`。

按规格 §22 记录：AccuX 版本、宿主类型与版本、模块、命令、工作表、Range 地址、单元格数、耗时、结果、异常。

**不记录**真实财务金额、公式内容或工作表敏感数据。

## 数据安全策略

- **fail before write**：写回前完成全部读取、分类、业务计算与完整 WritePlan 生成；任一步失败即不写回。
- **数值进数值出，公式进公式出**：公式保持公式属性，不替换为固定值。
- 无法安全转换的公式（数组公式、动态数组等）默认跳过并计入结果统计。
- **隐藏行/列默认跳过**：Host 读取并标记隐藏状态，修改型功能由公共操作管线统一跳过，选区求和不计入隐藏单元格；结果统计会显示跳过数量。
- 一次 Command 只从当前 Selection 捕获一次 `RangeTarget`；后续读写与验证都针对同一 Target，不重新读取 Selection。
- 批量整块读写，禁止逐 Cell COM 操作。
- V1 不提供 Snapshot、AccuX Undo 或事务级回滚；宿主状态（`ScreenUpdating` 等）在异常时仍会恢复。

## 数据处理规则（V1 固定行为）

| 数据类型 | 一键舍入 | 金额折合 | 选区求和 | 金额大写 |
| --- | --- | --- | --- | --- |
| 普通数值 | 修改数值 | 修改数值 | 计入合计 | 输出大写 |
| 普通公式 | 修改公式 | 修改公式 | 计入合计 | 输出大写 |
| 文本 / 公式文本 | 跳过 | 跳过 | 跳过 | 跳过 |
| Date / FormulaDate | 跳过 | 跳过 | 跳过 | 跳过 |
| Boolean / FormulaBoolean | 跳过 | 跳过 | 跳过 | 跳过 |
| 空白 | 跳过 | 跳过 | 跳过 | 跳过 |
| Error / FormulaError | 跳过 | 跳过 | 跳过 | 跳过 |
| 复杂公式 | 修改公式 | 修改公式 | 计入数字结果 | 输出大写 |
| 不支持类型 | 跳过 | 跳过 | 跳过 | 跳过 |

> 隐藏行/列是与数据类型独立的公共规则：选区总单元格数仍按完整矩形计算，但隐藏行/列中的单元格不会进入四个财务功能的处理；选区求和只复制格式化后的合计数字。

## 安装包（正式发布）

使用 Inno Setup 6.7.3 编译 `installer/AccuX.iss`。先构建 Release 程序集，再执行安装器回归检查：

```
powershell -ExecutionPolicy Bypass -File build.ps1 -Configuration Release -Version 1.4
pwsh -NoProfile -File installer\Test-Installer.ps1
ISCC.exe /DAccuXVersion=1.4 /DAccuXFileVersion=1.4.0.0 installer\AccuX.iss
```

脚本负责 .NET Framework 4.8 前置检查、程序集部署、COM 注册、x64 适配、卸载与升级策略。默认安装包名为 `AccuXSetup-1.4.exe`，版本参数由 CI 传入时无需修改安装脚本。

## GitHub Actions 自动构建与发布

`.github/workflows/build-release.yml` 在所有分支推送和手动触发时构建 Windows 安装包，并在 Actions 的 Artifacts 中保留 30 天。工作流固定使用 .NET SDK 9、Inno Setup 6.7.3 和 Office 15 PIA。

只有两段版本 tag 才会创建正式 Release，并同步生成 jsDelivr 更新清单：

```powershell
git tag v1.4
git push origin v1.4
```

`v1.4` 会生成 `AccuXSetup-1.4.exe`、对应的 SHA-256 文件，发布名为 `AccuX v1.4` 的 Release，并推送到 jsDelivr 更新分支。`v1.4.0`、`v01.4` 和 `v1.4-beta` 会被工作流拒绝。普通分支构建的安装包名会追加 `ci.<运行号>.<短SHA>`，不会创建 Release。

## 人工验证清单

以下项目需要真实宿主，请按 `docs/CompatibilityMatrix.md` 逐项验证并回填结果：

1. 注册后打开 Excel，确认出现 **AccuX → 基础财务** 与 **AccuX → 标记** 两个 Ribbon 组及对应图标。
2. 选中普通数值区域，执行一键舍入 / 金额折合 / 选区求和 / 金额大写。
3. 选中普通区域，分别点击标绿 / 标红 / 标黄 / 标蓝，确认仅可见单元格变色，值、公式和数字格式不变。
4. 选中普通公式区域，确认公式被包裹（`ROUND` / `*÷`），**未变成固定值**。
5. 选中混合区域（数值 + 文本 + 日期 + 布尔 + 空白 + 错误），确认只处理金额且结果统计正确。
6. 选中区域后先点击按钮，在参数窗口打开期间切换工作表，确认仍只写回原 Target。
7. 在选区内隐藏行或列，确认修改型功能和标记功能均不写入隐藏单元格，选区求和不计入隐藏数字。
8. 打开“设置”，确认版本号、统一阈值和自动升级选项；修改并保存后，基础财务功能与区域对比均使用新阈值，取消则不生效。
9. 选区超过 `maxProcessCells` 时确认被拒绝；超过 `largeSelectionWarning` 时确认弹出确认框。
10. 手动检查升级，确认“已是最新”或新版本信息；新版本安装前确认 SHA-256 校验、保存提示和普通安装启动行为。
11. 选区求和后确认出现三行复制对话框；金额和万元金额为千分位、固定 2 位小数，大写金额可正常复制且窗口在点击按钮后关闭。
12. 在包含可见、隐藏和非常隐藏工作表的工作簿中执行“生成目录”，确认仅列出可见工作表、目录位于首位且名称可跳转；对同一工作簿再次执行“生成目录”，确认弹出替换确认框——选择“是”时旧“目录”表被删除并生成不含自身的新目录，选择“否”时不改动工作簿。
13. 点击“存在对比”，分别捕获两个工作簿或工作表的连续区域，确认区域1独有、区域2独有、相同项、标题排除、隐藏数据跳过、重复次数、底色标记、清除标记和四张导出表均符合预期。
14. 在装有 WPS 的机器上重复以上关键链路，回填 WPS 行。

## 已知限制（V1）

- 不实现 Snapshot / AccuX Undo / 事务回滚。
- 不实现超大 Range 分块边读边写；超过 `maxProcessCells` 直接拒绝。
- 不做动态插件发现、热加载或复杂依赖解析；`BasicFinance` 与 `Mark` 由 AddIn 显式注册。
- WPS 与 Inno Setup 尚未在本机验证。

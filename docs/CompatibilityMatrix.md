# AccuX 兼容性矩阵

> 规格 §25 要求：兼容记录不能只写笼统的“Excel / WPS 通过”，至少同时记录
> 宿主类型、宿主版本、x86/x64、功能/API、验证结果、备注/已知限制。
>
> **只有实际验证过的宿主 API 和部署链路才进入正式支持范围。**
> 下表中 `待验证` 行表示代码已实现但尚未在真实宿主上人工验证。

## 环境记录

| 项目 | 值 |
| --- | --- |
| AccuX 版本 | 1.1 |
| 开发机 Office | Microsoft Office 16（Excel x64，路径 `C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE`） |
| 开发机 WPS | WPS Office 12.1.0.28488（ET x64；本机 `Excel.Application` ProgID/CLSID 由 WPS 接管，COM `Excel.Application` 与 `Excel.Application.16` 均解析到 `et.exe`） |
| 目标框架 | .NET Framework 4.8 |
| 构建工具 | Visual Studio 2022 / MSBuild 17 |

## Excel 验证记录

| Capability | Host | Version | Arch | Result | Notes |
| --- | --- | --- | --- | --- | --- |
| COM Add-in Load | Excel | 16.0 | x64 | 通过 | HKCU 注册；`COMAddIns` 显示 Connect=True |
| Visual Studio F5 启动 | Excel | 16.0 | x64 | 待验证 | 需在 VS 中选择 Excel 启动配置文件 |
| Ribbon 加载 | Excel | 16.0 | x64 | 通过 | UI Automation 枚举到 AccuX 选项卡与「基础财务」组 |
| 标记 Ribbon 组 | Excel | 16.0 | x64 | 待验证 | 需确认「标记」独立组、两行两列布局及四个图标 |
| Ribbon 图标显示 | Excel | 16.0 | x64 | 待验证 | 按钮已渲染，图标视觉效果需人工确认 |
| Selection 获取 | Excel | 16.0 | x64 | 待验证 | 仅捕获一次 RangeTarget |
| Value 批量读取 | Excel | 16.0 | x64 | 待验证 | 整块 `Range.Value` |
| Formula 批量读取 | Excel | 16.0 | x64 | 待验证 | 整块 `Range.Formula` |
| Formula 写入 | Excel | 16.0 | x64 | 待验证 | 整块 `Range.Formula`，保持公式属性 |
| Range 批量写入 | Excel | 16.0 | x64 | 待验证 | 整块或按行连续区段写入 |
| 普通公式识别 | Excel | 16.0 | x64 | 待验证 | `Formula` 非空判定 |
| 数组公式识别 | Excel | 16.0 | x64 | 待验证 | `{}` 包裹判定，默认跳过 |
| 动态数组识别 | Excel | 16.0 | x64 | 待验证 | `_xlfn.` / `#` 溢出引用判定 |
| 合并单元格识别 | Excel | 16.0 | x64 | 待验证 | `Range.MergeCells` |
| NumberFormat 读取 | Excel | 16.0 | x64 | 待验证 | 用于日期识别 |
| 隐藏行/列识别 | Excel | 16.0 | x64 | 待验证 | Host 读取 `Rows.Hidden` / `Columns.Hidden`，公共 Pipeline 跳过隐藏单元格 |
| 选区求和 | Excel | 16.0 | x64 | 待验证 | 可见数字求和、两位舍入、三格式复制对话框 |
| 生成目录 | Excel | 16.0 | x64 | 通过 | 临时工作簿 COM 验证：仅列出可见工作表，目录插入首位，名称列使用内部超链接；存在“目录”表时 `replaceExisting=false` 拒绝、`true` 删除重建且新目录不含自身 |
| 可见单元格底色标记 | Excel | 16.0 | x64 | 待验证 | 需验证四种颜色、隐藏行列跳过以及值/公式/数字格式保持不变 |
| 剪切板复制 | Excel | 16.0 | x64 | 待验证 | 金额 / 万元金额 / 大写金额复制；失败时不显示 MsgBox |
| ScreenUpdating | Excel | 16.0 | x64 | 待验证 | HostStateScope 保存/恢复 |
| EnableEvents | Excel | 16.0 | x64 | 待验证 | HostStateScope 保存/恢复 |
| DisplayAlerts | Excel | 16.0 | x64 | 待验证 | HostStateScope 保存/恢复 |
| Calculation | Excel | 16.0 | x64 | 待验证 | V1 默认不切换计算模式 |
| WPF Owner | Excel | 16.0 | x64 | 待验证 | SetWindowLongPtr 关联主窗口 |
| 安装包加载 | Excel | 16.0 | x64 | 待验证 | Inno Setup 6.7.3 已完成 CI 编译，真实宿主加载仍需人工验证 |
| 卸载 / 升级 | Excel | 16.0 | x64 | 待验证 | Inno Setup 6.7.3 已完成 CI 编译，真实安装/卸载仍需人工验证 |

## WPS 验证记录

| Capability | Host | Version | Arch | Result | Notes |
| --- | --- | --- | --- | --- | --- |
| COM Add-in Load | WPS | - | - | 未验证 | 本机未安装 WPS |
| Ribbon 加载 | WPS | - | - | 未验证 | 本机未安装 WPS |
| 标记 Ribbon 组 | WPS | - | - | 未验证 | 本机未安装 WPS |
| Formula 读取 | WPS | - | - | 未验证 | 本机未安装 WPS |
| Formula 写回 | WPS | - | - | 未验证 | 本机未安装 WPS |
| Range Value | WPS | - | - | 未验证 | 本机未安装 WPS |
| 隐藏行/列识别 | WPS | - | - | 未验证 | 需验证 `Rows.Hidden` / `Columns.Hidden` 返回值 |
| 生成目录 | WPS | 12.1 | x64 | 通过 | 临时工作簿 COM 验证（ET 12.1.0.28488）：仅列出可见工作表、目录插入首位、名称列内部超链接，替换行为与 Excel 一致；Add-in 加载 / Ribbon / WPF 确认框仍待人工验证 |
| 可见单元格底色标记 | WPS | - | - | 未验证 | 需验证 `Interior.Pattern` / `Interior.Color` 与隐藏行列处理 |
| 选区求和 | WPS | - | - | 未验证 | 需验证可见数字求和、三格式复制对话框与窗口关闭行为 |
| 剪切板复制 | WPS | - | - | 未验证 | 本机未安装 WPS；需验证三种格式复制 |
| Ribbon callback | WPS | - | - | 未验证 | 本机未安装 WPS |
| WPF Window | WPS | - | - | 未验证 | 本机未安装 WPS |

## 已知差异与设计决策

| 项 | 说明 |
| --- | --- |
| `Formula` / `Formula2` | V1 统一通过 `Range.Formula` 读写；规范化边界位于 `AccuX.Host`。若 WPS 或新 Excel 出现差异，仅修改 Host。 |
| 公式参数分隔符 | V1 假设中文环境使用逗号；若发现分号区域设置差异，在 Host 规范化边界处理。 |
| COM 释放 | V1 不对临时 RCW 显式调用 `Marshal.ReleaseComObject`，以实际兼容性验证为准。 |
| Ribbon 回调可见性 | `Connect` 必须使用 `[ClassInterface(ClassInterfaceType.AutoDual)]`。Office 通过 IDispatch 按名称调用 `onLoad` / `onAction` / `getImage`；`ClassInterfaceType.None` 只暴露接口方法（`GetCustomUI` 可用），普通类方法无法被 IDispatch 解析，Office 会静默丢弃整个 Ribbon（选项卡不显示，且不报错）。实测：改回 AutoDual 后 `OnRibbonLoad` 触发，选项卡与五个按钮正常显示。 |
| Office PIA 版本 | 本机 GAC 仅有 Excel/office/Vbe.Interop PIA 的 **15.0.0.0** 版本；NuGet 的 16.x PIA 强依赖 office.dll 16.0，运行时 `FileNotFoundException`。因此统一引用并随包部署 GAC 15.0 PIA（`Private=true`）。15.0 PIA 的接口 IID 与 Excel 16 宿主一致，V1 使用的对象模型调用兼容。 |
| 分块边读边写 | V1 不实现；超过 `maxProcessCells` 直接拒绝，保证 fail before write。 |

## 如何更新本文件

1. 在真实宿主上执行 README 中的验证步骤。
2. 将 `待验证` 改为 `通过` / `失败`，并填写实际版本与位数。
3. 失败项必须补充备注，并把兼容处理收敛到 `AccuX.Host`（编码规则 7）。

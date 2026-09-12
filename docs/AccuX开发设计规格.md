# AccuX V1 开发设计规格

## 1. 项目概述

AccuX 是一款面向财务人员的 Windows Excel / WPS 表格效率插件。

产品目标是将财务工作中高频、重复、容易出错的 Excel 操作标准化和一键化，同时保持对 Microsoft Excel 与 WPS 表格的兼容。

V1 不追求“大而全”，只实现四个核心功能：

1. 一键舍入
2. 金额折合
3. 选区求和
4. 金额大写

V1 的重点不是功能数量，而是先建立一套稳定、可扩展的基础架构，使后续增加财务工具时不需要重新设计 Excel/WPS 兼容、Ribbon、Range 操作、日志、配置等公共能力。

---

# 2. V1 设计原则

AccuX V1 遵循以下原则：

### 2.1 Excel / WPS 双平台

同一套核心代码同时支持：

- Microsoft Excel
- WPS 表格

宿主差异必须统一封装。

业务模块不得直接依赖具体 Excel 或 WPS COM 对象。

---

### 2.2 财务数据安全优先

所有修改工作表数据的功能必须：

- 区分普通数值和公式；
- 默认不破坏用户公式；
- 跳过无法安全处理的单元格；
- 对大量数据采用批量操作；
- 在写回前完成 Selection、可写性、公式安全和待写入结果验证；
- V1 不承诺事务级原子写入、AccuX 自定义 Undo 或数据自动回滚。

V1 的数据安全策略是 **尽量在写入前失败（fail before write）**，而不是建立复杂的数据事务恢复系统。

---

### 2.3 基础架构可扩展

新增功能原则上应通过新增 Module 或 Command 完成，而不是修改 Add-in 主程序。

---

# 3. 技术选型

本次 V1 建议定稿为：

```text
开发环境：Visual Studio 2022
语言：C#
运行环境：.NET Framework 4.8
插件形式：COM Add-in
Ribbon：Office CustomUI XML
UI：WPF
配置：JSON
敏感配置：可选 DPAPI
调试：Visual Studio 直接启动 Excel 进入托管调试
安装包：Inno Setup（正式发布阶段）
```

V1 开发阶段应保证 `AccuX.AddIn` 可以在 Visual Studio 中按 F5 直接启动 Microsoft Excel 进行调试，使以下代码能够直接设置断点：

```text
COM Add-in 入口
Ribbon callback
CommandDispatcher
RangeOperationPipeline
AccuX.Host
业务 Module
```

WPS 调试继续使用同一套程序集和核心代码；根据具体 WPS 安装与启动方式，可配置 Visual Studio 外部启动程序或附加到 WPS 表格进程。

正式发布阶段不自行开发安装器框架，统一使用 **Inno Setup** 制作安装包，负责 .NET Framework 4.8 前置条件、COM Add-in 注册、位数适配及卸载/升级等部署工作。

# 4. 总体架构

V1 采用四个主要工程：

```text
AccuX.sln

src/
├─ AccuX.AddIn
├─ AccuX.Host
├─ AccuX.Core
└─ AccuX.Modules.BasicFinance
```

依赖方向必须固定为：

```text
AccuX.AddIn
   │
   ├──────────────→ AccuX.Host
   │                    │
   │                    ▼
   ├──────────────→ AccuX.Modules.BasicFinance
   │                    │
   │                    ▼
   └────────────────→ AccuX.Core
                        ▲
                        │
                   AccuX.Host
```

核心约束：

```text
AccuX.Core 不引用 AccuX.Host
AccuX.Core 不引用 Excel/WPS Interop
AccuX.Modules.* 不直接解决 Excel/WPS 差异
AccuX.Host 可以引用 AccuX.Core，并实现 Core 定义的窄宿主接口
AccuX.AddIn 是 Composition Root，负责创建并注入具体实现
```

`RangeOperationPipeline` 继续位于 `AccuX.Core`，但不得直接引用 `AccuX.Host`。Core 只定义 Pipeline 真正需要的宿主抽象，例如：

```text
IRangeOperationHost
IHostStateScope
```

或拆分为等价的更小接口。`AccuX.Host` 负责实现这些接口。

业务功能需要的宿主能力（目录、批注、标记等）遵循同一模式：Core 定义功能窄接口与纯 CLR 参数对象，
`AccuX.Host` 在 `Features/` 提供实现类，见 §5.2“能力实现类的组织”。

因此调用关系为：

```text
Excel / WPS
     │
     ▼
AccuX.Host implementation
     │ implements
     ▼
Core host abstractions
     │
     ▼
RangeOperationPipeline
     │
     ▼
Business transform
```

而不是：

```text
Core → Host → Core
```

这样既保留统一 Range 操作框架，又避免 Core 与 Host 的循环依赖。

四个工程职责如下：

```text
┌─────────────────────────────────────────────┐
│                Excel / WPS                  │
└──────────────────────┬──────────────────────┘
                       │
                       ▼
┌─────────────────────────────────────────────┐
│                AccuX.AddIn                  │
│ COM / Ribbon / 启动 / 生命周期             │
│ 静态模块注册 / Command 回调 / 依赖组合      │
└──────────────┬─────────────────┬────────────┘
               │                 │
               ▼                 ▼
┌───────────────────────┐   ┌──────────────────────────┐
│      AccuX.Host       │   │ AccuX.Modules.BasicFinance│
│ Excel/WPS 差异适配    │   │ 一键舍入 / 金额折合 / 求和 │
│ 批量读写 / Formula    │   │ 金额大写 / 业务逻辑       │
│ Host State            │   │ 纯 C# 财务算法            │
└───────────┬───────────┘   └────────────┬─────────────┘
            │                            │
            └─────────────┬──────────────┘
                          ▼
┌─────────────────────────────────────────────┐
│                  AccuX.Core                 │
│ Module 接口 / Command / Pipeline            │
│ Pipeline 所需宿主抽象 / Cell 分类           │
│ 配置接口 / 日志接口 / 公共模型               │
└─────────────────────────────────────────────┘
```

其中：

- `AccuX.AddIn` 负责启动和依赖组合，不包含具体财务算法；
- `AccuX.Host` 负责宿主差异、COM 边界和 Core 宿主接口的具体实现；
- `AccuX.Core` 只提供跨模块公共能力及必要的宿主抽象契约；
- `AccuX.Modules.BasicFinance` 持有四个 V1 功能及其业务算法；
- V1 保留模块接口，但不实现目录扫描、Assembly 动态发现、热加载等复杂插件框架。

# 5. 各工程职责

## 5.1 AccuX.AddIn

插件统一入口，也是整个解决方案的 **Composition Root**。

负责：

- COM Add-in 生命周期；
- Ribbon 加载；
- Ribbon callback；
- 初始化 Host、Config、Logger；
- 创建 `RangeOperationPipeline` 并注入 Host 接口实现；
- 显式注册 V1 已知模块；
- Command 注册与回调分发；
- WPF 窗口与宿主窗口关联。

V1 **不做复杂模块发现或动态加载**。

推荐直接由 AddIn 显式创建模块列表，例如：

```csharp
IReadOnlyList<IAccuXModule> modules = new IAccuXModule[]
{
    new BasicFinanceModule()
};
```

允许保留轻量的：

```text
ModuleRegistry
```

用于统一初始化、Command 收集和 Shutdown，但不得在 V1 实现：

```text
扫描 Modules 目录
Assembly.LoadFrom
反射发现任意第三方 DLL
运行时热加载 / 卸载
插件依赖解析
复杂版本隔离
```

未来确有独立模块动态部署需求时，再在保持 `IAccuXModule` 契约的前提下扩展加载机制。

单个模块初始化失败应被捕获并记录，不得导致整个 AccuX Add-in 崩溃。

`AccuX.AddIn` 不得包含具体财务算法。

## 5.2 AccuX.Host

`AccuX.Host` 负责处理 Microsoft Excel 与 WPS 表格之间的宿主差异，以及需要统一控制的宿主公共能力。

设计原则：

> **差异驱动封装，不过度抽象。**

对于 Excel 和 WPS 中接口名称、参数、返回值和实际行为一致，并已经通过兼容性验证的 COM API，可以直接使用，不再额外建立一套 AccuX 自有对象模型。

`AccuX.Host` 不负责完整复制 Excel/WPS Object Model，而只处理以下几类能力：

- Excel/WPS 存在实际差异的能力；
- 涉及公式安全处理的能力；
- 需要统一保存和恢复状态的能力；
- 需要进行批量读写和性能优化的 COM 操作；
- 多个模块都会重复使用且有必要统一实现的宿主能力。

### 宿主识别

统一提供当前宿主基本信息，包括：

- 当前宿主类型：Excel / WPS；
- 宿主版本；
- 主窗口句柄；
- 必要的宿主能力检测。

主要用于插件初始化、WPF 窗口 Owner 设置以及兼容性判断。

---

### Selection、RangeTarget 与 Range 数据访问

V1 对 Selection 的使用遵循：

> **一次操作只在开始阶段读取一次 Selection，并立即固化为 `RangeTarget`；后续读取、验证和写回都针对同一个 `RangeTarget`，不得再次依赖当前 Selection。**

原因是金额折合、自定义舍入等命令可能在执行过程中打开 WPF 参数窗口。用户在窗口打开期间可能切换工作表、Workbook 或 Selection。如果写回阶段重新读取当前 Selection，存在写入错误区域的风险。

建议由 Core 定义纯 CLR 的目标描述：

```csharp
public sealed class RangeTarget
{
    public string WorkbookKey { get; init; }
    public string WorksheetKey { get; init; }
    public string Address { get; init; }

    public int RowCount { get; init; }
    public int ColumnCount { get; init; }
    public long CellCount { get; init; }

    public bool IsMultiArea { get; init; }
}
```

其中：

- `WorkbookKey` / `WorksheetKey` 是由 Host 生成和解释的稳定身份信息；
- Core 和业务 Module 不解释这些 Key；
- `RangeTarget` 不得持有 `Workbook`、`Worksheet`、`Range` 等 COM 对象；
- Host 在后续读取或写回前重新解析并验证该 Target 是否仍然有效；
- 如果目标 Workbook / Worksheet 已关闭、目标地址失效或目标已不可安全写入，则终止操作，不自动改用新的 Selection。

推荐命令流程：

```text
用户点击 Ribbon
        ↓
CaptureTarget()  ← 唯一一次读取当前 Selection
        ↓
需要时打开 WPF 参数窗口
        ↓
Read(target)
        ↓
内存计算
        ↓
ValidateWrite(target)
        ↓
Write(target, writePlan)
```

V1 Host 重点提供：

- 从当前 Selection 创建 `RangeTarget`；
- 判断 Selection 是否为可处理的 Range；
- 判断是否为多区域 Selection；
- 判断是否包含合并单元格；
- 获取 Range 地址、行数、列数及单元格数量；
- 按 `RangeTarget` 批量读取 Value；
- 按 `RangeTarget` 批量读取 Formula；
- 读取选区覆盖的行、列隐藏状态并写入纯 CLR 单元格数据；
- 批量判断是否包含公式；
- 按 `RangeTarget` 批量写入 Value；
- 按 `RangeTarget` 批量写入 Formula。

隐藏行与隐藏列的处理规则：

- `RangeTarget` 的行数、列数和单元格数量仍按完整矩形选区计算；
- Host 在 `Read(target)` 阶段识别每个单元格所在行、列是否隐藏，并写入 `CellData`；
- `RangeOperationPipeline` 统一跳过隐藏行/列中的单元格，修改型业务 Module 不重复实现该规则；只读业务按同一 `CellData.IsHidden` 元数据过滤；
- `ValidateWrite(target, writePlan)` 再次检查待写单元格的隐藏状态，若读取后状态发生变化则终止写回；
- 结果统计区分选区是否包含隐藏行、隐藏列，并统计被跳过的隐藏单元格数量。

V1 在 `maxProcessCells` 硬限制以内采用整块批量读取和内存处理：

```text
RangeTarget
  ↓
批量读取 Value / Formula
  ↓
转换为 CLR 数据
  ↓
Core / Module 内存处理
  ↓
生成完整 WritePlan
  ↓
批量写回同一个 RangeTarget
```

禁止业务模块进行大量逐单元格 COM 读写。

对于以下经验证在 Excel/WPS 中行为一致的基础属性，可由 Host 边界内部直接使用，不需要额外建立影子对象模型：

- `Address`
- `Row`
- `Column`
- `Rows.Count`
- `Columns.Count`
- `Offset`
- `Resize`

如果后续发现具体兼容性差异，再将对应能力加入 Host Adapter。

这里的“可直接使用”只限 `AccuX.Host` 或 `AccuX.AddIn` 的宿主边界内部；业务 Module 不应因此直接持有或遍历 `Range` COM 对象。

---

### Formula 处理与规范化契约

公式处理是 V1 的重点宿主能力。

`AccuX.Host` 负责统一处理：

- 判断单元格是否为公式；
- Formula 批量读取；
- Formula 批量写入；
- 普通公式识别；
- 数组公式识别；
- 动态数组或其他特殊公式识别；
- 判断公式是否允许安全修改；
- Excel/WPS 在 `Formula`、`Formula2`、区域设置等公式 API 上的实际兼容性差异；
- 将宿主公式转换为业务层可使用的规范化公式信息。

Core 可以定义不依赖 Excel/WPS 的公式契约，例如：

```csharp
public sealed class FormulaInfo
{
    public string Expression { get; init; }
    public FormulaKind Kind { get; init; }
    public bool CanTransform { get; init; }
}

public enum FormulaKind
{
    Normal,
    Array,
    DynamicArray,
    Unsupported
}
```

其中：

- `Expression` 是 Host 提供给 Core / Module 的规范化公式表达式；
- Module 不区分 Excel `Formula`、`Formula2`、WPS Formula 或本地区域格式 API；
- Host 负责把规范化表达式安全写回当前宿主；
- 无法稳定规范化或安全写回的公式，`CanTransform = false`，默认跳过；
- 公式中的数字 literal、参数分隔符和区域设置差异由 Host / 规范化边界处理，业务 UI 不直接拼接本地化公式字符串。

业务模块不直接处理 Excel/WPS 的公式对象差异，只根据 `FormulaInfo` 和业务规则决定是否转换。

对于无法确认能够安全修改的公式，默认跳过，不强制处理。

---

### 数据类型归一化

当 Excel 与 WPS 对 COM 数据返回形式存在差异时，由 `AccuX.Host` 统一转换为 Core 可以直接处理的数据。

主要包括：

- 数值；
- 文本；
- 空值；
- Boolean；
- 日期；
- Error；
- Formula。

如果实际测试确认某种类型在 Excel/WPS 中行为完全一致，则直接使用原始返回值，不额外增加转换逻辑。

Host 层应避免将平台特有 COM 类型传播到 Core 和业务 Module。

---

### 合并单元格和特殊区域

V1 需要识别可能影响批量读写安全性的特殊区域，包括：

- 合并单元格；
- 多区域 Selection；
- 数组公式区域；
- 动态数组或其他无法安全修改的公式区域。

这些区域由 Host 层负责识别。

业务模块只需要根据统一结果决定：

```text
处理
跳过
提示用户
```

不得分别针对 Excel/WPS 编写判断逻辑。

---

### 宿主状态管理

V1 不建立数据 Snapshot、AccuX Undo、事务回滚等数据恢复机制。

宿主状态管理只处理 Excel/WPS Application 级运行状态，例如：

- `ScreenUpdating`
- `Calculation`
- `EnableEvents`
- `DisplayAlerts`
- `StatusBar`

建议将原来的 `HostOperationScope` 明确命名为：

```text
HostStateScope
```

并保留一个轻量接口，例如：

```csharp
public interface IHostStateScope : IDisposable
{
}
```

或者由 `IRangeOperationHost` 提供：

```csharp
IHostStateScope BeginStateScope(HostStateOptions options);
```

其职责只有：

```text
保存原始宿主状态
    ↓
按需修改本次操作真正需要修改的状态
    ↓
执行操作
    ↓
finally / Dispose
    ↓
恢复宿主原始状态
```

`HostStateScope` **不得承担**：

```text
Values / Formulas 备份
工作表数据恢复
AccuX Undo
事务 Commit / Rollback
```

如果某个 V1 命令完全不修改 Application 全局状态，则不要求为了形式统一而创建 Scope。

原则是：

> **修改了哪个宿主全局状态，就必须可靠恢复哪个状态；不为了架构完整而主动修改宿主状态。**

业务 Module 不直接修改这些 Application 状态。

### NumberFormat

对于金额处理需要使用的 NumberFormat 能力，可通过 Host 层提供统一访问。

主要包括：

- 获取 NumberFormat；
- 设置 NumberFormat。

如果 Excel/WPS 对具体格式字符串、区域设置或 `NumberFormatLocal` 存在差异，应在 Host 层解决。

如果行为一致，则直接使用宿主原始 API。

---

### COM Object Lifetime 与线程模型

V1 必须明确 COM 对象生命周期和线程边界。

核心规则：

```text
Excel / WPS COM 对象只允许存在于 AccuX.AddIn / AccuX.Host 宿主边界内部
Core / Modules 只接收 CLR DTO、RangeTarget、FormulaInfo 和业务数据
```

强制要求：

- Core 不保存 `Application`、`Workbook`、`Worksheet`、`Range` 等 COM 对象；
- Module 不保存或缓存任何宿主 COM 对象；
- `RangeTarget`、`OperationContext`、`CommandResult`、配置对象、日志对象中禁止携带 COM 引用；
- 禁止将当前 `Selection` 或 `Range` 长期缓存用于后续写回；
- 临时 COM 对象的获取、使用和释放策略统一由 Host 管理；
- 业务层不得自行调用 `Marshal.ReleaseComObject` / `Marshal.FinalReleaseComObject`；
- 是否对具体临时 RCW 显式释放，以 Excel/WPS 实际兼容性验证为准；
- 宿主拥有的 `Application` 等对象不得由业务代码擅自释放。

V1 线程模型采用：

> **所有 Excel/WPS COM 访问默认只在宿主 UI / STA 线程执行。**

因此禁止：

```csharp
Task.Run(() =>
{
    // 禁止在后台线程直接访问 Excel/WPS COM
});
```

V1 四个财务算法计算量很小，默认同步执行：

```text
STA 线程 COM Read
        ↓
纯 CLR 计算
        ↓
STA 线程 COM Write
```

后续如果出现真正 CPU 密集型算法，可以只把已经完全脱离 COM 的纯 CLR 数据复制到后台线程计算；任何 COM Read / Write 仍必须回到经过验证的宿主线程执行。

---

### 写入能力检查

执行修改型操作前，Host 层应提供必要的可写性判断，包括：

- 当前是否存在有效 Workbook；
- 当前是否存在有效 Worksheet；
- Selection 是否为有效 Range；
- 工作表是否受到保护；
- 当前 Range 是否允许修改；
- 是否包含无法安全写入的特殊区域。

业务层不依赖 COM Exception 来判断区域是否能够写入。

---

### Host Capability

对于实际发现存在 Excel/WPS 差异的功能，可以增加轻量的宿主能力检测。

例如：

```text
SupportsArrayFormula
SupportsDynamicArray
SupportsMultiAreaWrite
```

Capability 只针对已经确认存在兼容性差异的能力建立。

禁止为了未来可能存在的差异，提前定义大量无实际用途的 Capability。

Excel/WPS 的兼容性验证结果统一记录在：

```text
CompatibilityMatrix.md
```

---

### 与业务模块的边界

`AccuX.Modules.*` 不负责解决 Excel/WPS 差异。

业务模块不得：

- 根据 Excel/WPS 分别实现两套逻辑；
- 大量直接访问宿主 COM 对象；
- 自行处理 Formula 平台差异；
- 自行管理 ScreenUpdating、Calculation 等宿主状态；
- 大量逐 Cell 访问 COM。

禁止出现类似：

```text
if Excel
    执行 Excel 逻辑
else if WPS
    执行 WPS 逻辑
```

如果业务开发过程中发现 Excel/WPS 行为不同，应优先修改：

```text
AccuX.Host
```

而不是修改业务 Module。

---

### V1 实际实现范围

V1 不需要完整封装 Excel/WPS Object Model。

首版重点实现以下能力：

```text
宿主识别
主窗口句柄
Selection 获取与验证

Value 批量读取/写入
Formula 批量读取/写入
公式类型识别

合并单元格识别
特殊公式区域识别

NumberFormat

ScreenUpdating
Calculation
EnableEvents
DisplayAlerts
宿主状态保存与恢复
```

其他 Workbook、Worksheet、Range API，如果 Excel/WPS 行为一致，可直接调用。

只有当后续业务需求出现，或者实际发现 Excel/WPS 存在兼容性差异时，再增加对应 Host Adapter。

---

### 核心原则

判断某项宿主能力是否需要进入 `AccuX.Host` 时，遵循以下规则：

```text
Excel/WPS 存在实际差异
        → 封装

涉及公式安全
        → 封装

涉及批量 COM 性能
        → 封装

涉及宿主状态保存与恢复
        → 封装

多个模块存在重复且有统一价值
        → 视情况封装

Excel/WPS 行为一致且调用简单
        → 不封装
```

`AccuX.Host` 的目标是建立一个轻量、稳定、可扩展的 Excel/WPS 兼容层，而不是重新实现一套 Excel/WPS Object Model。

### 能力实现类的组织

`AccuX.Host` 内部按“一个 Core 窄接口一个实现类”组织，共享机制放在 `ExcelHostBase`：

```text
ExcelHostBase                          共享 COM 基础设施：Application、工作簿/工作表解析、安全属性读取、状态作用域
ExcelRangeOperationHost                IRangeOperationHost：批量读写与范围操作
Features/ExcelWorkbookDirectoryHost    IWorkbookDirectoryHost：目录生成的 COM 机制
Features/ExcelCellCommentHost          ICellCommentHost：批注读写
Features/ExcelCellMarkHost             ICellMarkHost：可见单元格底色标记
```

需要 COM 的业务功能实现放在 `Features/`，且只负责 COM 机制。文案、配色、列宽等业务决定
必须由模块通过参数对象（例如 `DirectoryOptions`）传入，Host 不内置功能专属常量；
这样“功能长什么样”留在可单测的模块层，Host 只保留“怎么落到 COM 上”。

窄接口与参数对象（DTO）都定义在 `AccuX.Core/Operations`，使用纯 CLR 类型，不携带 COM 引用。
新增需要 COM 的功能时按此结构增加实现类，不要继续堆进 `ExcelRangeOperationHost`。

## 5.3 AccuX.Core

`AccuX.Core` 只负责与具体业务模块无关、能够被多个模块复用的公共基础能力，以及 `RangeOperationPipeline` 所需的最小宿主抽象。

设计原则：

> **Core 是插件框架的公共核心，不是业务代码的公共存放区，也不认识具体 Excel/WPS COM 实现。**

V1 建议包括：

```text
Modules/
    IAccuXModule
    IAccuXContext
    ModuleContext

Commands/
    CommandDefinition
    CommandDispatcher
    CommandResult

Cells/
    CellValueClassifier
    CellValueType

Operations/
    RangeOperationPipeline
    OperationContext
    OperationResult
    RangeTarget
    RangeWritePlan
    FormulaInfo
    IRangeOperationHost
    IHostStateScope
    IHostContext
    IWorkbookDirectoryHost / DirectoryOptions
    ICellCommentHost / CellCommentTarget
    ICellMarkHost

Configuration/
    IConfigManager
    JsonConfigManager

Logging/
    ILogger
```

其中 `IRangeOperationHost` 是为了切断 `RangeOperationPipeline → AccuX.Host` 具体工程依赖而存在的窄接口。

它只暴露 Pipeline 真正需要的能力，例如概念上：

```csharp
public interface IRangeOperationHost
{
    RangeTarget CaptureTarget();
    RangeReadResult Read(RangeTarget target);
    WriteCheckResult ValidateWrite(RangeTarget target, ...);
    void Write(RangeTarget target, RangeWritePlan writePlan);
    IHostStateScope BeginStateScope(...);
}
```

最终方法名称和 DTO 可在实现时调整，但必须满足：

```text
Core 只依赖接口
Host 实现接口
AddIn 负责注入
Module 不拿具体 Host COM 对象
```

`RangeOperationPipeline` 因而只负责通用编排：

```text
验证
→ 批量读取抽象结果
→ 单元格分类
→ 调用业务 transform
→ 生成待写结果
→ 写入前最终检查
→ 通过接口批量写回
→ 结果统计与日志上下文
```

它不直接：

- 引用 `AccuX.Host`；
- 引用 Excel/WPS Interop；
- 持有 `Range` / `Worksheet` COM 对象；
- 判断 `HostKind` 后写平台分支。

以下业务能力 **不放入 Core**：

```text
RoundingService
AmountConversionService
ChineseAmountService
FormulaTransformService（当前阶段）
```

只有当某项业务能力已经被两个或以上独立模块真实复用，并且抽象边界稳定时，才考虑从 Module 提升到 Core。

禁止因为“以后可能会复用”而提前把业务 Service 放入 Core。

## 5.4 AccuX.Modules.BasicFinance

V1 唯一业务模块，负责“一键舍入、金额折合、选区求和、金额大写”四个功能及其全部业务逻辑。

主要职责：

- Ribbon Command 定义；
- 用户输入与参数校验；
- 四个功能自己的业务算法；
- 数值与公式的具体业务处理策略；
- 调用 Core 提供的公共能力；
- 调用统一 `RangeOperationPipeline`；
- WPF 界面与结果展示。

建议按功能聚合代码，而不是把所有 Command、Service、Model 分散到不同技术目录：

```text
AccuX.Modules.BasicFinance/
│
├─ Rounding/
│  ├─ RoundingCommand.cs
│  ├─ RoundingService.cs
│  └─ RoundingOptions.cs
│
├─ AmountConversion/
│  ├─ AmountConversionCommand.cs
│  ├─ AmountConversionService.cs
│  ├─ AmountConversionOptions.cs
│  ├─ AmountConversionView.xaml
│  └─ AmountConversionViewModel.cs
│
├─ ChineseAmount/
│  ├─ ChineseAmountCommand.cs
│  └─ ChineseAmountService.cs
│
├─ Common/
│  └─ FormulaTransformService.cs
│
└─ BasicFinanceModule.cs
```

`FormulaTransformService` 目前只服务于 BasicFinance 中的舍入和金额折合，因此也保留在模块内部；以后如果多个独立模块都需要统一公式转换能力，再考虑提升到 Core 或 Host。

模块内部的财务算法必须保持纯 C#，能够在不启动 Excel/WPS 的情况下进行单元测试。

业务模块不得自行解决 Excel/WPS 平台差异，也不得自行实现重复的批量 Range 操作框架。

---

# 6. 模块化设计

V1 只有一个业务模块，模块机制只保留最小边界，不建设动态插件系统。

统一接口精简为：

```csharp
public interface IAccuXModule
{
    string Id { get; }

    void Initialize(IAccuXContext context);

    IEnumerable<CommandDefinition> GetCommands();

    void Shutdown();
}
```

V1 不在 `IAccuXModule` 中放置：

```text
Name
Version
Order
GetRibbonGroup()
动态依赖信息
插件文件路径
```

这些信息当前没有真实运行时需求，不为未来假设场景提前扩展接口。

Ribbon 在 V1 中由 `AccuX.AddIn` 的静态 CustomUI XML 定义，按钮 ID 映射到 Command ID；业务模块只注册 Command，不动态贡献 Ribbon UI。

V1：

```text
AccuX.Modules.BasicFinance.dll
```

未来可以新增：

```text
AccuX.Modules.Reconciliation.dll
AccuX.Modules.Audit.dll
AccuX.Modules.DataCleaning.dll
```

但 V1 不因此实现一个通用第三方插件系统。

V1 模块注册采用：

> **最小接口 + 编译期项目引用 + AddIn 显式注册。**

即：

```text
AccuX.AddIn
    ↓
new BasicFinanceModule()
    ↓
ModuleRegistry.Register(...)
```

而不是：

```text
扫描 DLL
↓
反射发现
↓
动态加载
↓
版本解析
```

`ModuleRegistry` 只保存 AddIn 显式传入的模块实例和 Command 注册结果，不负责插件发现。

单个模块初始化或 Command 执行失败仍需隔离，不得导致整个 AccuX 无法启动。

# 7. Ribbon 设计

V1 只建立一个：

```text
AccuX
```

Tab。

其中一个功能组：

```text
AccuX
└─ 基础功能（2 行、3 列）
   第一行：一键舍入 | 金额折合 | 选区求和
   第二行：金额大写 | 生成目录 | 批注助手
```

不在 V1 堆放大量按钮。

Ribbon 各功能组的按钮布局高度不得超过 2 行。多按钮组使用纵向 `box` 包含最多两个横向行 `box`，显式控制行数，避免宿主默认排列成 3 行；新增按钮应横向扩展或拆分功能组，不得增加第三行。基础功能组固定为上述 2 行、3 列，按钮统一使用 `size="normal"`。此要求约束按钮布局行数，宿主 Ribbon 的实际像素高度由 Excel / WPS 控制；布局变更后须分别验证两个宿主中的显示效果。

Ribbon callback 统一进入 CommandDispatcher。

禁止每个模块自行暴露大量 COM callback。

图标资源统一由 AccuX.AddIn 管理，并通过 Office CustomUI XML 的 image / getImage 机制加载。

要求：

每个 Ribbon Button 必须有明确图标；
图标风格、尺寸和视觉语言保持统一；
图标资源不得散落在各业务 Module 中；
Module 只负责 Command，不负责直接加载 Ribbon 图标；
Excel 与 WPS 中必须分别验证图标显示是否正常；
图标加载失败不得影响 Command 本身执行。

推荐 Command ID：

```text
accux.basic.round
accux.basic.convert
accux.basic.uppercase
```

---

# 8. Host 使用方式

AccuX 不建立完整的 `IWorkbook / IWorksheet / IRange` 影子对象模型。

设计原则仍然是：

> **有差异才适配，没有差异不重复封装；跨 Pipeline 边界只暴露最小抽象。**

`AccuX.Host` 可以在其实现内部直接使用 Excel/WPS COM 对象，但不得把平台特有 COM 类型传播到 Core 或业务 Module。

Core 对 Host 的依赖仅通过 `IRangeOperationHost`、`IHostStateScope` 等窄接口表达。

`AccuX.Host` 对外还可以提供宿主上下文：

```csharp
public interface IHostContext
{
    HostKind HostKind { get; }

    string HostVersion { get; }

    IntPtr MainWindowHandle { get; }

    object Application { get; }
}
```

其中 `Application` 只供 AddIn/Host 边界内必要场景使用，不得作为通用对象传入业务 Module。

宿主类型：

```csharp
public enum HostKind
{
    Unknown,
    Excel,
    Wps
}
```

对于经 `CompatibilityMatrix.md` 验证、Excel/WPS 行为一致且调用简单的 Workbook、Worksheet、Range 基础 API，不额外建立影子模型。

对于以下类型的操作，必须由 `AccuX.Host` 实现并通过 Core 所需的窄接口或 Host 自身服务提供：

- Excel/WPS 存在实际差异的 API；
- Value / Formula 批量读取和写入；
- 普通公式、数组公式、动态数组等特殊公式识别；
- 合并单元格和特殊区域判断；
- 主窗口句柄；
- `ScreenUpdating`、`Calculation`、`EnableEvents`、`DisplayAlerts` 等宿主状态管理；
- 需要统一兼容行为或错误处理的高风险 COM 操作；
- 需要 COM 的业务功能实现（目录、批注、标记等，位于 `Features/`；文案与外观由模块通过
  Core 参数对象提供，Host 不内置功能专属常量）。

如果开发过程中发现原本认为一致的 API 在 Excel/WPS 中存在差异，应将能力收口到 `AccuX.Host`，并同步更新 `CompatibilityMatrix.md`。

# 9. 单元格分类

这是 V1 的核心公共能力。

选区读取后，统一识别为：

```csharp
public enum CellValueType
{
    Blank,

    ConstantNumber,
    FormulaNumber,

    Text,
    FormulaText,

    Date,
    FormulaDate,

    Boolean,
    FormulaBoolean,

    Error,
    FormulaError,

    Unsupported
}
```

增加 Date / Boolean 的原因是：财务操作不能把日期序列值或逻辑值误识别为普通金额数字。Host 负责根据宿主返回值、格式和公式结果等已验证信息进行归一化，`CellValueClassifier` 输出统一分类。

四个 V1 功能默认：

```text
Date / FormulaDate       → 跳过
Boolean / FormulaBoolean → 跳过
```

每个功能根据单元格类型决定处理策略。

不要让四个业务功能分别自行判断：

```text
是否为空
是否公式
是否数字
是否日期
是否 Boolean
```

统一由：

```text
CellValueClassifier
```

负责。

# 10. 统一 Range 操作管线

四个功能统一复用 `RangeOperationPipeline` 的选区捕获和批量读取能力，但必须遵守明确的依赖边界和固定目标规则；选区求和为只读命令，不执行写回阶段。

执行流程：

```text
用户点击 Ribbon
        ↓
CommandDispatcher
        ↓
通过 IRangeOperationHost CaptureTarget()
        ↓
生成 RangeTarget（之后不再读取当前 Selection）
        ↓
需要时打开 WPF 参数窗口
        ↓
RangeOperationPipeline
        ↓
Read(target) 批量读取 Value / Formula，返回 CLR 数据
        ↓
CellValueClassifier
        ↓
普通数值 / 公式 / Date / Boolean 等分类
        ↓
调用当前业务模块的纯 C# transform
        ↓
在内存中生成完整 RangeWritePlan
        ↓
ValidateWrite(target, writePlan)
        ↓
如有需要创建 HostStateScope
        ↓
Write(target, writePlan)
        ↓
恢复宿主状态（仅当实际修改过）
        ↓
记录日志 / 返回处理结果
```

关键依赖关系：

```text
RangeOperationPipeline
        ↓ depends on
IRangeOperationHost (Core interface)
        ↑ implemented by
AccuX.Host
```

固定目标约束：

```text
一次 Command = 一个 RangeTarget
```

禁止：

```text
写回阶段重新读取 Selection
参数窗口确认后重新读取 Selection
RangeOperationPipeline → AccuX.Host concrete class
Core → Excel/WPS Interop
业务 Module → Range COM object
```

如果 `RangeTarget` 对应的 Workbook / Worksheet / Range 在操作期间失效，命令直接失败并提示用户重新执行，不自动切换到新的 Selection。

业务功能只提供“怎么转换”的逻辑，不负责遍历 COM Cell，也不负责 Excel/WPS 差异。

V1 不在 Pipeline 中实现：

```text
Snapshot
AccuX Undo
Transaction
自动数据 Rollback
分块边读边写
```

# 11. 普通数值和公式的统一原则

V1 明确采用：

> 数值进，数值出；公式进，公式出。

对于能够安全转换的公式：

- 保持公式属性；
- 不将公式替换成固定值。

对于暂时无法安全转换的公式：

- 跳过；
- 不强制处理；
- 在处理结果中提示。

这是整个 V1 的核心数据安全规则。

---

# 12. 一键舍入

## 12.1 功能定义

对当前选区内金额进行统一四舍五入。

V1 支持：

```text
保留 2 位
自定义小数位
```

默认：

```text
2 位
```

---

## 12.2 数值处理

普通数字：

```text
原值
↓
decimal
↓
财务舍入
↓
写回数值
```

使用明确的：

```text
MidpointRounding.AwayFromZero
```

不得依赖 `Math.Round` 默认中点行为。

---

## 12.3 公式处理

对于普通单元格公式：

```text
原公式
↓
增加 ROUND 逻辑
↓
仍然写回 Formula
```

处理后：

```text
Formula → Formula
```

不得变为：

```text
Formula → Value
```

---

## 12.4 默认跳过

以下情况 V1 不强制处理：

- 文本；
- 空单元格；
- Error；
- 公式错误；
- 无法识别的数据；
- 特殊数组公式或类似复杂区域。

处理完成后显示：

```text
成功处理：xxx
跳过公式：xxx
跳过文本：xxx
跳过错误：xxx
```

---

# 13. 金额折合

## 13.1 V1 定义

V1 金额折合采用：

> 用户选择除百、除千、除万，对选区金额进行折合处理，可选是否添加“万”字。

暂不接在线汇率。

## 13.2 数值处理

普通数值：

```text
原金额
× 或 ÷
折合率
↓
按指定小数位舍入
↓
写回数值
```

核心计算使用：

```text
decimal
```

---

## 13.3 公式处理

公式单元格：

```text
原 Formula
↓
添加 × Rate 或 ÷ Rate
↓
必要时增加 ROUND
↓
写回 Formula
```

必须保持：

```text
Formula → Formula
```

不得直接写入当前计算结果。

---

## 13.4 除零检查

如果：

```text
除法模式
且
Rate = 0
```

必须阻止执行。

---

# 14. 金额大写

## 14.1 功能定义

将金额转换为中文金额大写。

核心算法：

```text
ChineseAmountService
```

输入：

```text
decimal
```

输出：

```text
string
```

---

## 14.2 普通数值和公式单元格

均直接转换为成中文大写文字。

## 14.3 输出方式

V1 默认：

```text
输出到当前选择区域
```

---

## 14.4 不使用 UDF

V1 不采用：

```text
=AccuXAmountUppercase(...)
```

这样的 Excel/WPS 自定义函数。

避免 Workbook 对 AccuX 插件产生永久依赖。

---

## 14.5 选区求和

选区求和为只读命令 `accux.basic.sum`：

- 只计入可见的普通数值和数字公式结果；
- 文本、日期、布尔值、空白和错误值不计入；
- 对全部可见数字求和后使用 `MidpointRounding.AwayFromZero` 舍入到 2 位；
- 计算完成后显示金额复制对话框，提供金额、万元金额和中文大写金额三种格式；
- 三种格式均使用只读输入框展示，用户点击对应按钮后复制到剪切板并关闭对话框；
- 金额与万元金额使用 `#,##0.00` 格式；无可求和数字时分别显示 `0.00`、`0.00` 和 `零元整`；
- 剪切板复制失败时不显示 MsgBox，仍关闭对话框并记录失败结果。

---

# 15. 四个功能的数据处理规则

V1 统一规则：

| 数据类型                 | 一键舍入 | 金额折合 | 选区求和 | 金额大写 |
| ------------------------ | -------- | -------- | -------- | -------- |
| 普通数值                 | 修改数值 | 修改数值 | 计入合计 | 输出大写 |
| 普通公式                 | 修改公式 | 修改公式 | 计入合计 | 输出大写 |
| 文本 / 公式文本          | 跳过     | 跳过     | 跳过     | 跳过     |
| Date / FormulaDate       | 跳过     | 跳过     | 跳过     | 跳过     |
| Boolean / FormulaBoolean | 跳过     | 跳过     | 跳过     | 跳过     |
| 空白                     | 跳过     | 跳过     | 跳过     | 跳过     |
| Error / FormulaError     | 跳过     | 跳过     | 跳过     | 跳过     |
| 复杂公式                 | 修改公式 | 修改公式 | 输出大写 |
| 不支持类型               | 跳过     | 跳过     | 跳过     |

这个规则应作为 V1 固定行为。

---

# 16. Range 性能与最大处理范围

禁止：

```text
逐 Cell COM Read
逐 Cell COM Write
```

V1 不实现“超大 Range 分块边读边写”。原因是 V1 同时采用 `fail before write` 且不提供事务级 Rollback；如果边读边写，后续 Chunk 失败时会天然产生部分写入状态，与当前安全策略冲突。

V1 采用两个阈值：

```text
largeSelectionWarning
maxProcessCells
```

行为：

```text
CellCount <= largeSelectionWarning
    → 正常执行

largeSelectionWarning < CellCount <= maxProcessCells
    → 明确提示用户，确认后执行

CellCount > maxProcessCells
    → V1 直接拒绝处理
```

阈值由统一设置节提供。默认值为警告 100000、最大 500000；正式发布前仍应通过 Excel/WPS 实测 benchmark 验证性能和兼容性。

在 `maxProcessCells` 以内，V1 原则上执行：

```text
RangeTarget
↓
批量读取全部 Value / Formula
↓
CLR 内存处理
↓
生成完整 RangeWritePlan
↓
完成写入前验证
↓
批量写回
```

这样可以兑现：

> **在开始修改工作表之前，完成当前操作所需的全部读取、分类、业务计算和待写结果生成。**

如果未来真实需求要求支持超过 V1 上限的超大区域，再单独设计 Chunk + Recovery / Partial Failure 策略，不在 V1 提前实现。

---

# 17. 公式修改

V1 的公式转换逻辑由 `AccuX.Modules.BasicFinance/Common/FormulaTransformService` 统一实现，供“一键舍入”和“金额折合”复用。

不要让：

```text
RoundingCommand
AmountConversionCommand
```

直接处理 Excel/WPS Formula API 或本地化公式差异。

## 17.1 Host 与 Module 的公式边界

公式读取流程：

```text
Excel / WPS Formula API
        ↓
AccuX.Host
        ↓
FormulaInfo（规范化表达式 + FormulaKind + CanTransform）
        ↓
BasicFinance
```

公式写回流程：

```text
BasicFinance 生成规范化公式表达式
        ↓
AccuX.Host
        ↓
转换为当前 Excel / WPS 可安全接受的 Formula API 表达
        ↓
写回同一个 RangeTarget
```

因此：

- Excel/WPS 的 `Formula` / `Formula2` 差异由 Host 处理；
- 区域设置导致的公式分隔符、数字 literal 等差异由 Host 的规范化边界处理；
- `FormulaTransformService` 不判断 Host 类型；
- 业务 Command 不直接调用 Excel/WPS Formula API。

## 17.2 FormulaTransformService 职责

`FormulaTransformService` 主要负责：

- 接收 `FormulaInfo.Expression`；
- 规范处理前导 `=`；
- 根据调用方明确给出的业务变换生成新的规范化公式表达式；
- 对原公式增加外层业务表达式；
- 生成语法结构明确的待写回表达式。

`FormulaTransformService` **不负责“防止重复包裹”或判断用户是不是已经执行过同一个业务操作。**

例如：

```text
=A1
↓ 用户执行一次金额折合
=(A1)*7
↓ 用户再次明确执行金额折合
=((A1)*7)*7
```

从公式转换 Service 的角度，这是合法的两次显式业务操作，不应自行去重。

## 17.3 重复操作规则属于具体 Command

是否允许重复执行、是否提示用户、是否需要识别已有业务包装，属于：

```text
RoundingCommand
AmountConversionCommand
```

自己的业务规则，而不是 FormulaTransformService 的通用职责。

V1 默认原则：

> **用户每次主动执行 Command 都视为一次新的明确操作。FormulaTransformService 不进行自动去重。**

如果未来某个具体 Command 需要幂等行为，只在该 Command / 对应业务 Service 中实现，并单独增加测试，不提升为通用公式规则。

该 Service 当前仅在 BasicFinance 内使用，因此不得提前放入 Core。以后如果多个独立模块真实需要同一套公式转换能力，再考虑提升为公共能力。

# 18. 写入前保护策略

V1 不实现以下机制：

```text
Values / Formulas Snapshot
撤销上一次 AccuX 操作
AccuX 自定义 Undo
事务级 Commit / Rollback
写入失败后的自动数据恢复
```

修改型功能：

```text
一键舍入
金额折合
```

必须遵循 **fail before write** 原则。

执行批量写回前至少完成：

```text
Workbook / Worksheet 有效性验证
RangeTarget 有效性验证
目标区域可写性验证
特殊区域检查
Value / Formula 批量读取
单元格分类
公式安全判断
全部业务计算
全部待写入结果生成
```

只有上述步骤全部成功后，才开始写回。

V1 明确不承诺 COM 批量写入具备数据库事务意义上的原子性。如果宿主在实际写回阶段发生异常：

```text
记录异常
恢复被 AccuX 修改过的宿主 Application 状态
向用户明确提示操作失败
```

但不尝试通过 AccuX 自建 Snapshot 自动恢复工作表数据。

该取舍用于控制 V1 复杂度；后续只有在真实用户需求和故障数据证明有必要时，才评估独立 Undo / Recovery 机制。

# 19. Selection / RangeTarget 保护

Selection 只用于创建一次 `RangeTarget`。

创建 Target 时统一验证：

- 是否存在 Workbook；
- 是否存在 Worksheet；
- Selection 是否为 Range；
- Selection 是否为空；
- 是否为不支持的多区域或特殊区域；
- CellCount 是否超过 `largeSelectionWarning`；
- CellCount 是否超过 `maxProcessCells`。

创建 Target 后：

```text
不得在本次 Command 中重新读取 Selection 作为写入目标
```

写入前针对同一个 `RangeTarget` 再次验证：

- Workbook / Worksheet 身份是否仍匹配；
- Range 地址是否仍可解析；
- 目标区域是否仍允许修改；
- 是否出现新的保护状态或其他阻止安全写入的条件。

超过 `largeSelectionWarning` 但未超过 `maxProcessCells` 时进行确认；超过 `maxProcessCells` 时 V1 直接拒绝处理。

阈值统一放入 `settings` 配置节，由 AddIn 设置窗口编辑；保存后更新共享 `HostOptions` 和基础功能提示，区域对比在下一次操作时读取最新值。

---

# 20. WPF

原项目继续采用 WPF。

WPF 主要用于：

```text
金额折合窗口
自定义舍入窗口
设置窗口
提示窗口
```

简单的一键操作无需每次弹出复杂界面。

窗口必须正确设置 Excel/WPS 主窗口为 Owner。

设置窗口同时显示当前 `AccuXVersion`、最近升级检测状态，并提供自动检测开关和手动检测按钮。自动检测在插件启动后后台执行，每 24 小时最多一次；仅当发现带有精确安装包与 SHA-256 附件的稳定 GitHub Release 时提示用户进入设置，下载校验通过后启动普通 Inno Setup 安装程序。

---

# 21. 配置

继续采用：

```text
JSON
+
可选 DPAPI
```

但 V1 配置保持简单。

例如：

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

`settings` 中的两个阈值同时作用于基础财务和区域对比，且必须大于 0、警告阈值不能超过最大阈值。旧版 `basicFinance` / `compare` 阈值首次加载时迁移到 `settings`，旧字段保留；`roundDigits` 和区域对比颜色继续保留在原配置节。`autoCheckForUpdates` 控制启动后的 GitHub Releases 检查，`lastUpdateCheckUtc` 由插件维护。

V1 没有 API Key，因此 DPAPI 暂时没有必须使用的场景。

接口仍然保留，为以后扩展准备。

---

# 22. 日志

V1 必须有基础日志。

记录：

```text
AccuX Version
Host
Host Version
Module
Command
Worksheet
Range Address
Cell Count
Duration
Result
Exception
```

默认禁止记录：

```text
真实财务金额
公式内容
工作表中的敏感数据
```

---

# 23. 异常处理

所有 Ribbon callback 必须：

```text
try
↓
Execute Command
↓
catch
↓
记录日志
↓
显示友好错误
```

任何异常都不得穿出 COM callback。

单个业务命令失败不得导致整个插件失效。

---

# 24. Excel / WPS 兼容策略

所有宿主差异只能出现在：

```text
AccuX.Host
```

禁止出现：

```csharp
if (isWps)
{
}
else
{
}
```

散落在业务模块。

如果发现某个 API Excel/WPS 行为不同：

```text
新增 Host Adapter 能力
```

而不是修改业务算法。

---

# 25. Compatibility Matrix

项目保留：

```text
CompatibilityMatrix.md
```

兼容记录不能只写笼统的“Excel / WPS 通过”，至少需要同时记录：

```text
宿主类型
宿主版本
x86 / x64
功能/API
验证结果
备注 / 已知限制
```

至少覆盖：

```text
Add-in 加载
Visual Studio 调试启动
Ribbon
Selection
Value
Formula
Formula 写入
Range 批量写入
特殊公式识别
ScreenUpdating / Calculation 等实际使用的 Host State
WPF Owner
安装包加载
卸载 / 升级
```

推荐记录形式：

| Capability      | Host  | Version  | Arch | Result    | Notes |
| --------------- | ----- | -------- | ---- | --------- | ----- |
| COM Add-in Load | Excel | 实测版本 | x64  | Pass/Fail |       |
| Formula Write   | WPS   | 实测版本 | x64  | Pass/Fail |       |

只有实际验证过的宿主 API 和部署链路才进入正式支持范围。

# 26. V1 测试

## Core Unit Test

只测试 Core 中真正的公共基础能力，例如：

```text
CellValueClassifier
RangeOperationPipeline 中可脱离宿主测试的逻辑
CommandDispatcher
配置与公共模型
```

Core Test 不包含一键舍入、金额折合、选区求和、金额大写等 BasicFinance 专属业务算法。

---

## BasicFinance Unit Test

必须覆盖：

```text
RoundingService
AmountConversionService
ChineseAmountService
FormulaTransformService
重复执行公式变换（验证 Service 不自动去重）
```

这些测试不依赖 Excel/WPS，直接验证模块内部纯 C# 业务逻辑。

---

## Excel Integration Test

验证：

```text
普通数字
普通公式
Date / FormulaDate
Boolean / FormulaBoolean
空白
文本
Error
混合区域
RangeTarget 固定目标
用户切换 Selection 后仍只操作原 Target
largeSelectionWarning
maxProcessCells 硬限制
```

---

## WPS Integration Test

使用与 Excel 相同的测试 Workbook。

重点检查：

```text
Formula 读取
Formula 写回
Range Value
Ribbon callback
WPF Window
```

---

## Visual Studio 调试约定

开发阶段必须把“可以从 Visual Studio 直接启动 Excel 调试 Add-in”作为标准开发体验，而不是依赖手工打开 Excel 后再附加进程。

`AccuX.AddIn` 项目建议配置 Visual Studio Debug 启动目标为本机 `EXCEL.EXE`。开发注册完成后：

```text
F5
↓
Visual Studio 启动 Excel
↓
Excel 加载 AccuX COM Add-in
↓
调试器自动附加
↓
COM Entry / Ribbon / Host / Pipeline / Module 断点可直接命中
```

开发期应提供简单、可重复的注册/注销方式，保证 Clean / Rebuild 后能够快速进入调试。具体注册脚本可以随实现确定，但不得要求开发者每次手工修改大量注册表项。

WPS 使用同一套核心程序集；如果 WPS 可稳定作为 Visual Studio 外部启动程序，则采用同样方式，否则使用启动 WPS 后 Attach to Process 的调试方式。

集成测试时应分别记录：

```text
Debug Host = Excel
Debug Host = WPS
Host Version
x86 / x64
```

## 安装与部署约定

V1 开发阶段优先保证本地注册和 Visual Studio 调试链路，不提前开发自定义安装器。

正式发布安装包统一使用：

```text
Inno Setup
```

安装工程在 Phase 5 建立，至少负责：

- 检查或声明 `.NET Framework 4.8` 前置条件；
- 安装 AccuX 程序集与资源；
- 注册 Excel / WPS 所需 COM Add-in 信息；
- 根据支持矩阵处理 x86 / x64；
- 支持正常卸载；
- 为后续版本升级保留稳定的 Product/Upgrade 策略。

安装注册方式必须和开发期注册方式保持同一逻辑来源，避免出现“VS 能调试、安装包无法加载”的两套配置。

# 27. V1 开发顺序

## Phase 0：兼容性验证

只实现：

```text
COM Add-in
Visual Studio F5 启动 Excel 调试
Ribbon
Selection
Value
Formula
WPF
```

要求 Excel/WPS 核心链路均通过，并开始维护 `CompatibilityMatrix.md`。

---

## Phase 1：基础架构

完成：

```text
AccuX.AddIn
AccuX.Host
AccuX.Core
IAccuXModule
ModuleRegistry（仅显式注册，不做动态发现）
CommandDispatcher
Config
Logger
IRangeOperationHost
IHostStateScope
RangeOperationPipeline
CellValueClassifier
```

同时确认工程依赖满足：

```text
Core !→ Host
Host → Core
Modules → Core
AddIn → Host + Core + Modules
```

暂不实现完整业务。

---

## Phase 2：一键舍入

完成：

```text
数值舍入
公式舍入
混合 Range
测试
```

---

## Phase 3：金额折合

完成：

```text
折合窗口
乘/除
数值处理
公式处理
小数位
测试
```

---

## Phase 4：金额大写

完成：

```text
金额大写算法
普通数值
公式结果
输出到原区域
测试
```

---

## Phase 5：稳定性与发布

完成：

```text
异常处理
大范围保护
Excel/WPS 回归测试
x86/x64 验证
Inno Setup 安装工程
安装 / 升级 / 卸载验证
```

V1 Phase 5 仍不包含 AccuX 自定义 Undo、Snapshot 或事务回滚。

# 28. 编码强制规则

Vibe Coding 必须遵循以下规则：

1. Core 不引用 Excel/WPS Interop。
2. Core 不引用 `AccuX.Host` 具体工程；`RangeOperationPipeline` 只依赖 Core 中定义的宿主抽象接口。
3. `AccuX.Host` 可以引用 Core，并负责实现 `IRangeOperationHost`、`IHostStateScope` 等宿主接口，以及 `Features/` 下的功能窄接口实现类。
4. Core 只存放跨模块公共能力和 Pipeline 必需的最小宿主抽象，不存放某个模块专属业务 Service。
5. 某项业务能力只有在两个或以上独立模块真实复用后，才考虑提升到 Core。
6. `RoundingService`、`AmountConversionService`、`ChineseAmountService` 属于 BasicFinance，不得放入 Core。
7. Excel/WPS 存在差异的宿主能力必须位于 `AccuX.Host`；经验证完全一致且调用简单的 API 不做无意义二次封装。需要 COM 的业务功能实现放 `AccuX.Host/Features`，文案与外观由模块通过 Core 参数对象提供。
8. Excel/WPS 差异不得进入业务模块。
9. 所有财务算法使用纯 C#，并可脱离 Excel/WPS 单元测试。
10. 财务金额计算优先使用 `decimal`。
11. 所有 Range 操作必须区分数值、公式、Date、Boolean、文本、空白和 Error。
12. 默认禁止 Formula → Value。
13. 无法安全处理的公式必须跳过。
14. 一次 Command 只能从当前 Selection 创建一次 `RangeTarget`；后续读取、验证和写回不得重新依赖当前 Selection。
15. `RangeTarget` 只能包含 CLR 身份信息和地址信息，不得携带 COM 对象。
16. V1 在 `maxProcessCells` 内完成整块批量读取、内存计算和完整 WritePlan 生成；超过上限直接拒绝，不实现分块边读边写。
17. Range 优先批量读取和写入，禁止逐 Cell 高频 COM 调用。
18. Formula API 差异和 Formula 规范化由 Host 处理；具体业务公式转换留在所属 Module。
19. `FormulaTransformService` 不实现“防止重复包裹”或通用去重；重复操作规则属于具体 Command。
20. 所有 Excel/WPS COM 对象只允许存在于 AddIn / Host 边界；Core / Module 不得保存 COM 引用。
21. 业务层不得自行调用 `Marshal.ReleaseComObject` 或 `Marshal.FinalReleaseComObject`。
22. 所有 Excel/WPS COM Read / Write 默认只在宿主 UI / STA 线程执行，禁止通过 `Task.Run` 直接访问 COM。
23. 所有 Ribbon callback 必须统一异常处理。
24. WPF 不实现财务计算逻辑。
25. Config 统一由 ConfigManager 管理。
26. `IAccuXModule` 只保留最小生命周期和 Command 注册能力；V1 不实现动态 Ribbon contribution。
27. V1 禁止实现不必要的目录扫描、反射插件发现、热加载和复杂依赖解析。
28. BasicFinance 由 AddIn 显式注册；单个模块失败不得导致 AccuX 整体失效。
29. `HostStateScope` 只恢复 AccuX 实际修改过的 Application 状态，不负责工作表数据恢复。
30. V1 不实现 Snapshot、AccuX Undo、Transaction 或自动数据 Rollback。
31. 修改型操作必须在写回前完成当前操作所需的全部读取、分类、业务计算和完整待写结果生成。
32. 开发版本必须支持从 Visual Studio 直接启动 Excel 进行 F5 调试。
33. 正式安装包使用 Inno Setup 制作。
34. 不得因为“以后可能复用”而提前公共化业务代码。
35. 所有 Ribbon 可点击按钮必须配置图标；图标资源统一由 AccuX.AddIn 管理，不得由业务 Module 自行加载 Ribbon 图标。

# 29. V1 Definition of Done

一个功能只有同时满足以下条件才算完成：

```text
[ ] 普通数字测试通过
[ ] 普通公式测试通过
[ ] Date / FormulaDate 跳过策略测试通过
[ ] Boolean / FormulaBoolean 跳过策略测试通过
[ ] 混合 Range 测试通过
[ ] 文本处理正确
[ ] 空白处理正确
[ ] Error 处理正确
[ ] 不破坏原公式
[ ] 写回前完成必要验证和完整结果计算
[ ] 一次 Command 只捕获一次 RangeTarget，写回不重新读取 Selection
[ ] 超过 maxProcessCells 时拒绝执行
[ ] Excel 测试通过
[ ] WPS 测试通过
[ ] 所属模块 Unit Test 通过；涉及 Core 公共能力时 Core Unit Test 通过
[ ] 大量 Range 不逐 Cell COM 操作
[ ] 异常能够正确记录
[ ] 如果修改宿主全局状态，异常后能够恢复该状态
[ ] 用户能够看到处理结果
```

基础架构进入可发布状态还必须满足：

```text
[ ] Core 不引用 Host / Excel / WPS Interop
[ ] RangeOperationPipeline 通过 Core host interface 工作
[ ] BasicFinance 使用显式模块注册，不依赖动态插件发现
[ ] IAccuXModule 使用精简接口，不包含 Ribbon 动态贡献能力
[ ] Core / Module 不持有 Excel/WPS COM 对象
[ ] COM Read / Write 遵守宿主 UI / STA 线程规则
[ ] Visual Studio F5 能直接启动 Excel 并命中 Add-in 断点
[ ] Inno Setup 安装包可完成安装、加载和卸载验证
```

V1 Definition of Done 不要求：

```text
AccuX Undo
Snapshot
Transaction Rollback
动态第三方插件加载
```

# 30. V1 最终结构

最终第一版应保持简单：

```text
AccuX
│
├─ AddIn
│   ├─ COM 入口
│   ├─ Ribbon
│   ├─ Composition Root
│   └─ 显式模块注册
│
├─ Host
│   ├─ ExcelHostBase（共享 COM 基础设施）
│   ├─ IRangeOperationHost 实现
│   ├─ IHostStateScope 实现
│   ├─ 宿主识别
│   ├─ Excel/WPS 差异适配
│   ├─ Value / Formula 批量读写
│   ├─ 宿主状态管理
│   └─ Features/：目录 / 批注 / 标记的功能级 COM 实现
│
├─ Core
│   ├─ Modules
│   │   └─ IAccuXModule
│   ├─ Commands
│   ├─ Cells（CellValueClassifier）
│   ├─ Operations
│   │   ├─ RangeOperationPipeline
│   │   ├─ IRangeOperationHost / IHostStateScope
│   │   └─ 功能窄接口与参数对象（IWorkbookDirectoryHost / DirectoryOptions 等）
│   ├─ Configuration
│   └─ Logging
│
├─ Modules.BasicFinance
│   ├─ Rounding
│   │   ├─ RoundingCommand
│   │   ├─ RoundingService
│   │   └─ RoundingOptions
│   │
│   ├─ AmountConversion
│   │   ├─ AmountConversionCommand
│   │   ├─ AmountConversionService
│   │   ├─ AmountConversionOptions
│   │   └─ WPF UI
│   │
│   ├─ ChineseAmount
│   │   ├─ ChineseAmountCommand
│   │   └─ ChineseAmountService
│   │
│   ├─ SelectionSum
│   │   ├─ SelectionSumCommand
│   │   └─ SelectionSumService
│   │
│   ├─ Directory
│   │   ├─ DirectoryCommand
│   │   └─ DirectoryTemplate（文案与外观参数）
│   │
│   ├─ Comment
│   │   ├─ CommentCommand
│   │   └─ CommentWindow
│   │
│   └─ Common
│       └─ FormulaTransformService
│
└─ Modules.Mark
    ├─ MarkCommand
    └─ MarkModule
```

发布工程：

```text
installer/
└─ Inno Setup Project
```

V1 的核心不是实现大量功能，而是完成：

```text
稳定 Excel/WPS 宿主差异层
        +
明确的 Core / Host 单向依赖
        +
统一 Range 操作框架
        +
正确区分数值和公式
        +
模块内聚的四个可靠财务功能
        +
顺畅的 VS → Excel 调试链路
        +
可重复的 Inno Setup 发布链路
```

新增功能时，代码默认先放在所属 Module 内部。

只有真正具有跨模块复用价值的能力，才提升到 Core。

V1 保留 `IAccuXModule` 契约，但模块采用编译期引用和显式注册，不为未来假设需求提前建设复杂插件系统。

因此未来增加一个新的财务工具时，理想情况下只需要：

```text
在所属 Module 中新增业务算法
+
新增 Command
+
新增 Ribbon 定义
+
在 AddIn 的模块注册表中显式注册新 Module（仅新增独立 Module 时）
```

不需要重新开发：

```text
Excel/WPS 差异适配
Range 批量读写
公共单元格分类
公共操作管线
异常处理
日志
配置
```

这就是 AccuX V1 需要保留的可扩展性：**接口稳定、依赖清晰、实现克制，先服务真实业务，再按真实复用需求扩展。**

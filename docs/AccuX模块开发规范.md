# AccuX Agent 新增模块开发规范

> 适用范围：AccuX 当前架构下新增或扩展业务 Module、Command、业务算法、WPF 参数界面及必要宿主能力。  
> 目标读者：负责在 AccuX 仓库中实施新增模块和功能的开发 Agent。  
> 本文是可独立执行的模块开发规范。Agent 应直接按本文进行设计、编码、测试和验收，不需要依赖其他设计文档才能完成新增模块。  
> 对本文未明确规定的实现细节，优先复用仓库现有接口和实现方式，并遵循“最小改动、不建立平行基础设施、不提前扩展架构”的原则。

---

## 0. 本规范的使用方式

Agent 收到“新增模块”“新增 Command”或“扩展现有模块”的任务后，应把本文作为直接执行规则。

执行时遵循以下优先级：

```text
当前任务的明确业务要求
        ↓
本文规定的架构与安全约束
        ↓
仓库中已经存在的公共接口和实现
        ↓
最小必要新增代码
```

如果需求没有明确要求修改公共架构，则不得主动：

```text
扩大 Core 职责
新增第二套配置 / 日志 / Range 框架
引入动态插件发现
建立 Excel/WPS 影子对象模型
增加 Snapshot / Undo / Transaction
改变 COM 线程模型
```

当本文没有规定某个具体方法名、构造参数或 DTO 字段时，Agent 应读取仓库现有代码并适配现有 API；不得为了让示例代码“对得上”而创建一套新的平行接口。

---

## 1. Agent 的首要目标

新增模块时，Agent 的目标不是“尽可能抽象”或“建立通用插件框架”，而是：

1. 在现有 AccuX 分层边界内实现业务能力；
2. 保持 Excel / WPS 双宿主兼容；
3. 不让业务模块接触 Excel/WPS COM 对象；
4. 对修改工作表的功能遵守数据安全与 fail-before-write 原则；
5. 复用 Core 已有公共能力，不重复建设 Range、日志、配置、宿主状态等基础设施；
6. 业务代码默认留在所属 Module，只有出现真实且稳定的跨模块复用后才提升到 Core；
7. 保持 V1 的轻量模块机制：编译期引用、AddIn 显式注册，不实现动态插件发现。

一句话原则：

> **新增模块应增加业务能力，而不是重新发明 AccuX 基础架构。**

---

## 2. 开工前必须先做的判断

Agent 在创建文件前，必须先判断需求属于以下哪一类。

### 2.1 只是现有模块的新功能

如果新能力与现有模块业务边界一致，应优先放入现有 Module，而不是创建新 Module。

例如：

- 与基础财务金额处理高度相关的功能，优先评估放入 `AccuX.Modules.BasicFinance`；
- 仅被 BasicFinance 内多个 Command 复用的 Service，继续放在 BasicFinance 内部；
- 不得因为“以后可能复用”而提前搬到 Core。

### 2.2 是独立业务域，适合新建 Module

只有业务边界相对独立时才创建新的 `AccuX.Modules.*` 工程，例如：

```text
AccuX.Modules.Reconciliation
AccuX.Modules.Audit
AccuX.Modules.DataCleaning
```

新 Module 仍然使用统一的：

- `IAccuXModule`；
- `CommandDefinition` / `CommandDispatcher`；
- `RangeOperationPipeline`；
- `CellValueClassifier`；
- `IConfigManager`；
- `ILogger`；
- Core 定义的宿主抽象。

### 2.3 需求其实是宿主差异

如果实现过程中发现问题本质是：

- Excel 与 WPS 行为不同；
- Formula / Formula2 / 本地化公式 API 不同；
- COM 批量读写行为不同；
- 特殊区域识别不同；
- NumberFormat 或 Application 状态处理不同；

则应修改或扩展 `AccuX.Host`，而不是在业务 Module 中加入：

```csharp
if (isWps) { ... }
else { ... }
```

同时，真实发现的兼容差异必须记录到项目的 `CompatibilityMatrix.md`，至少包含宿主类型、宿主版本、x86/x64、能力/API、验证结果和已知限制。

如果确认不是宿主差异，而是某个业务功能新增的 COM 操作（目录、批注、标记这类），按 §20.1 处理：Core 加功能窄接口与参数 DTO，`AccuX.Host/Features` 加实现类。

### 2.4 需求确实是跨模块公共能力

只有满足以下条件，才考虑修改 Core：

- 已有两个或以上独立 Module 真实使用；
- 抽象边界已经稳定；
- 与具体业务语义无关；
- 不依赖 Excel/WPS COM 实现；
- 放入 Core 后能减少真实重复，而不是为假设未来做准备。

否则，默认留在业务 Module。

---

## 3. 不可违反的架构边界

### 3.1 工程依赖方向

保持以下依赖关系：

```text
AccuX.AddIn
   ├──> AccuX.Host ───────────> AccuX.Core
   ├──> AccuX.Modules.* ──────> AccuX.Core
   └──────────────────────────> AccuX.Core
```

强制规则：

```text
AccuX.Core 不引用 AccuX.Host
AccuX.Core 不引用 Excel/WPS Interop
AccuX.Modules.* 不引用具体 Excel/WPS COM 实现
AccuX.Host 可以引用 AccuX.Core 并实现 Core 定义的宿主接口
AccuX.AddIn 是 Composition Root，负责实例创建、依赖组合和显式模块注册
```

### 3.2 Core 不是业务公共垃圾桶

禁止把下列类型的代码因为“看起来通用”就放入 Core：

- 某个财务规则；
- 某个对账算法；
- 某个审计判断；
- 某个模块专用 Formula 包装；
- 只被一个 Module 使用的 Service；
- 只是假设未来可能复用的 DTO / Helper。

### 3.3 不建设动态插件框架

V1 新增模块仍采用：

> **最小接口 + 编译期项目引用 + AddIn 显式注册。**

禁止为新增模块实现：

```text
扫描 Modules 目录
Assembly.LoadFrom
反射发现第三方模块
热加载 / 热卸载
复杂版本解析
插件依赖解析
动态 Ribbon Contribution
```

---

## 4. 新模块的标准目录结构

推荐按“业务功能聚合”，不要按纯技术类型把所有 Command、Service、Model 分散到全局目录。

示例：

```text
src/
└─ AccuX.Modules.Reconciliation/
   ├─ ReconciliationModule.cs
   │
   ├─ Match/
   │  ├─ MatchCommand.cs
   │  ├─ MatchService.cs
   │  ├─ MatchOptions.cs
   │  ├─ MatchResult.cs
   │  ├─ MatchView.xaml           # 仅在需要参数界面时
   │  └─ MatchViewModel.cs
   │
   ├─ DifferenceCheck/
   │  ├─ DifferenceCheckCommand.cs
   │  ├─ DifferenceCheckService.cs
   │  └─ DifferenceCheckOptions.cs
   │
   └─ Common/
      └─ ...                       # 仅限本模块内部真实复用
```

测试工程可按仓库现有约定建立，例如：

```text
tests/
└─ AccuX.Modules.Reconciliation.Tests/
   ├─ MatchServiceTests.cs
   └─ DifferenceCheckServiceTests.cs
```

Agent 不应为了新增一个功能重构整个解决方案目录。

---

## 5. `IAccuXModule` 实现规范

当前模块接口保持最小化：

```csharp
public interface IAccuXModule
{
    string Id { get; }

    void Initialize(IAccuXContext context);

    IEnumerable<CommandDefinition> GetCommands();

    void Shutdown();
}
```

### 5.1 模块类只做什么

模块实现主要负责：

- 保存模块级必要依赖；
- 初始化模块内部服务；
- 暴露本模块 Command 定义；
- 在 Shutdown 中释放模块自身管理的非 COM 资源；
- 不让单个模块初始化失败拖垮整个 Add-in。

### 5.2 模块类不要做什么

禁止把以下职责塞进 `IAccuXModule` 或模块类：

```text
Name / Version / Order 等无运行时必要性的元数据
Ribbon Group 动态生成
插件目录信息
动态依赖图
DLL 发现 / 加载
业务大段计算
Excel/WPS COM 缓存
```

### 5.3 模块骨架示例

> 以下代码仅表示职责结构。具体 `IAccuXContext`、`CommandDefinition` 构造方式以仓库实际接口为准，Agent 不得为了套用示例而创建平行 API。

```csharp
public sealed class ReconciliationModule : IAccuXModule
{
    private IAccuXContext _context;
    private IReadOnlyList<CommandDefinition> _commands;

    public string Id => "accux.reconciliation";

    public void Initialize(IAccuXContext context)
    {
        _context = context ?? throw new ArgumentNullException(nameof(context));

        // 从 context 使用现有公共能力：Config、Logger、Pipeline 等。
        // 不在这里获取或缓存 Excel/WPS Range/Application COM 对象。

        _commands = BuildCommands();
    }

    public IEnumerable<CommandDefinition> GetCommands()
        => _commands ?? Array.Empty<CommandDefinition>();

    public void Shutdown()
    {
        _commands = null;
        _context = null;
    }

    private IReadOnlyList<CommandDefinition> BuildCommands()
    {
        // 以仓库现有 CommandDefinition API 为准。
        throw new NotImplementedException();
    }
}
```

---

## 6. 新模块如何接入 AddIn

### 6.1 添加编译期引用

`AccuX.AddIn` 对新增业务 Module 增加明确的项目引用。

### 6.2 在 Composition Root 显式注册

示意：

```csharp
IReadOnlyList<IAccuXModule> modules = new IAccuXModule[]
{
    new BasicFinanceModule(),
    new ReconciliationModule()
};
```

如仓库已经存在 `ModuleRegistry`，通过现有 Registry 注册，不再创建第二套 Registry。

### 6.3 模块失败必须隔离

初始化模块时：

- 捕获单个模块初始化异常；
- 写入日志；
- 允许其他模块继续初始化；
- 不允许单个 Module 导致整个 AccuX Add-in 无法启动。

### 6.4 不通过模块动态生成 Ribbon

V1 Ribbon 由 `AccuX.AddIn` 的静态 CustomUI XML 定义。

新增按钮时：

1. 修改静态 Ribbon XML；
2. 为按钮绑定稳定 ID；
3. 映射到 Command ID；
4. Ribbon callback 统一进入 `CommandDispatcher`；
5. 不让 Module 暴露大量自己的 COM callback。
6. 所有 Ribbon 可点击按钮必须配置图标；图标资源统一由 AccuX.AddIn 管理，不得由业务 Module 自行加载 Ribbon 图标。

---

## 7. Command ID 与业务边界

推荐使用稳定、可读、带域前缀的 Command ID：

```text
accux.<module>.<command>
```

例如：

```text
accux.reconciliation.match
accux.reconciliation.diffcheck
accux.audit.formulacheck
```

Command 负责“业务用例编排”，Service 负责“纯业务计算”。

一个典型 Command 可以负责：

- 捕获一次 RangeTarget；
- 参数读取 / UI 调用；
- 参数校验；
- 构造业务 Transform；
- 调用 `RangeOperationPipeline`；
- 根据 `OperationResult` 形成用户可见结果；
- 定义该 Command 自己的重复执行规则。

Command 不应负责：

- 遍历 Excel/WPS COM Cell；
- 直接处理 Formula / Formula2 API；
- 自己管理 ScreenUpdating 等宿主状态；
- 自己实现重复的批量 Range 读写框架。

---

## 8. 修改工作表的 Command：强制标准流程

只要 Command 会修改工作表数据，必须优先采用统一 `RangeOperationPipeline`。

标准流程：

```text
用户点击 Ribbon
        ↓
CommandDispatcher
        ↓
IRangeOperationHost.CaptureTarget()
        ↓
生成 RangeTarget（本次 Command 唯一目标）
        ↓
需要时打开 WPF 参数窗口
        ↓
RangeOperationPipeline
        ↓
Read(target) 批量读取 Value / Formula，返回 CLR 数据
        ↓
CellValueClassifier
        ↓
业务 Transform（纯 C#）
        ↓
在内存中生成完整 RangeWritePlan
        ↓
ValidateWrite(target, writePlan)
        ↓
如确有需要，创建 HostStateScope
        ↓
Write(target, writePlan)
        ↓
恢复实际修改过的宿主状态
        ↓
日志 + OperationResult / CommandResult + 用户提示
```

### 8.1 一次 Command 只能捕获一次 RangeTarget

硬规则：

> **一次 Command = 一个固定 RangeTarget。**

捕获后禁止：

```text
参数窗口确认后再次读取 Selection
写回前重新读取 Selection
根据用户当前激活 Sheet 自动切换目标
RangeOperationPipeline 内部重新决定目标
```

如果用户在参数窗口打开期间切换了 Workbook / Worksheet / Selection：

- 仍然以最初的 `RangeTarget` 为目标；
- 写回前验证该 Target 是否仍有效；
- 如果失效，命令失败并提示用户重新执行；
- **绝不自动改写新的 Selection。**

### 8.2 RangeTarget 必须是纯 CLR DTO

`RangeTarget` 可包含：

- Workbook / Worksheet 的 Host 可解释身份 Key；
- Address；
- RowCount / ColumnCount / CellCount；
- IsMultiArea 等纯数据。

禁止携带：

```text
Application COM
Workbook COM
Worksheet COM
Range COM
Selection COM
```

---

## 9. fail-before-write 与最大处理范围

V1 数据安全采用：

> **尽量在第一次写入工作表之前发现问题。**

### 9.1 写回前至少完成

修改型操作开始写回前，至少完成：

```text
Workbook / Worksheet 有效性验证
RangeTarget 有效性验证
目标区域可写性验证
特殊区域检查
Value / Formula 批量读取
单元格分类
公式安全判断
全部业务计算
完整 RangeWritePlan 生成
写回前最终验证
```

### 9.2 不做分块边读边写

V1 不实现超大 Range 的 Chunk 边读边写。

原因：

- 当前策略强调 fail-before-write；
- V1 没有事务回滚；
- Chunk 1 已写入、Chunk 2 后续失败会造成天然部分写入。

因此 V1 在 `maxProcessCells` 内采用整块批量读取、内存计算、完整 WritePlan，然后批量写回。

### 9.3 两级范围阈值

配置中使用：

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
    → 直接拒绝处理
```

具体默认值必须依据 Excel/WPS benchmark，不得在新增模块里硬编码一个未经验证的“性能安全值”。

### 9.4 V1 不承诺事务原子性

如果实际 COM 批量 Write 阶段仍然异常：

- 记录异常；
- 恢复 AccuX 实际改动过的 Application 状态；
- 向用户明确提示失败；
- 不自行建立 Snapshot 回滚工作表。

Agent 不应私自为某个模块加入：

```text
AccuX Undo
Values / Formulas Snapshot
Transaction Commit / Rollback
写入失败后的自动数据恢复
```

除非当前任务明确要求引入此类能力，并同时补充相应的数据恢复策略、异常语义、测试和验收标准。

---

## 10. Range 与 COM 性能规则

禁止：

```text
逐 Cell COM Read
逐 Cell COM Write
业务 Module foreach Range.Cells
业务 Module 持有 Range 并长期缓存
```

应采用：

```text
Host 批量读取 Value / Formula
        ↓
转换为 CLR 数据
        ↓
Module / Core 内存处理
        ↓
生成 RangeWritePlan
        ↓
Host 批量写回
```

业务算法应尽可能在脱离 COM 后运行。

---

## 11. COM Object Lifetime 与线程模型

### 11.1 COM 对象只存在于宿主边界

Excel/WPS COM 对象只允许存在于：

```text
AccuX.AddIn
AccuX.Host
```

禁止出现在：

```text
AccuX.Core
AccuX.Modules.*
RangeTarget
OperationContext
CommandResult
配置对象
日志对象
业务 DTO
```

### 11.2 Module 不自行释放 COM

业务层禁止调用：

```csharp
Marshal.ReleaseComObject(...)
Marshal.FinalReleaseComObject(...)
```

临时 COM 对象的获取、使用、释放策略统一由 Host 管理。

### 11.3 COM Read / Write 默认在宿主 UI / STA 线程

禁止：

```csharp
Task.Run(() =>
{
    // 直接读取或写入 Excel/WPS COM
});
```

默认流程：

```text
STA：COM Read
    ↓
纯 CLR 计算
    ↓
STA：COM Write
```

未来若出现 CPU 密集算法，可将已经完全脱离 COM 的 CLR 数据复制到后台线程计算；COM Read / Write 仍回到经过验证的宿主线程。

---

## 12. 单元格分类规则

模块不得各自判断：

```text
是否为空
是否公式
是否数字
是否日期
是否 Boolean
是否 Error
```

统一使用 `CellValueClassifier`，分类至少覆盖：

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

### 12.1 Date / Boolean 不得误当金额

对财务类数值操作，默认：

```text
Date / FormulaDate       → 跳过
Boolean / FormulaBoolean → 跳过
```

新增模块如需处理 Date / Boolean，必须由该 Command 明确声明业务语义并新增对应测试，不能通过“它底层也是数字”来处理。

### 12.2 Formula 与 Value 的安全原则

对于修改型财务操作，默认：

> **数值进，数值出；公式进，公式出。**

禁止默认：

```text
Formula → 当前计算 Value
```

无法安全转换的公式应跳过并在结果中统计/提示。

---

## 13. Formula 处理规范

### 13.1 Host 负责公式 API 与规范化边界

公式读取：

```text
Excel / WPS Formula API
        ↓
AccuX.Host
        ↓
FormulaInfo
  - Expression
  - FormulaKind
  - CanTransform
        ↓
Module
```

公式写回：

```text
Module 生成规范化公式表达式
        ↓
AccuX.Host
        ↓
转换为当前宿主可安全接受的 Formula API 表达
        ↓
写回固定 RangeTarget
```

Module 不处理：

- Excel `Formula` 与 `Formula2` 差异；
- WPS Formula API 差异；
- 区域设置导致的参数分隔符差异；
- 本地化数字 literal；
- HostKind 分支。

### 13.2 FormulaTransformService 只做公式表达式变换

若某 Module 内存在类似 `FormulaTransformService` 的公式变换 Service，它只应：

- 接收规范化 `FormulaInfo.Expression`；
- 处理前导 `=` 等规范形式；
- 按调用方明确给出的业务变换构造新的规范化表达式；
- 返回待写回表达式。

它不应：

```text
判断用户是否执行过同一 Command
自动检测“已经包裹过”并去重
为了幂等性偷偷忽略第二次业务操作
```

### 13.3 重复操作规则属于具体 Command

默认规则：

> **用户每次主动执行 Command，都视为一次新的明确业务操作。**

例如：

```text
=A1
↓ 第一次金额折合
=(A1)*7
↓ 用户再次主动执行金额折合
=((A1)*7)*7
```

Formula 变换 Service 不自动去重。

如果某个具体 Command 需要幂等、重复执行提示、已有业务包装识别：

- 只在该 Command / 对应业务 Service 中实现；
- 明确业务规则；
- 单独增加测试；
- 不提升为 FormulaTransformService 的通用职责。

---

## 14. 业务算法与 Service 设计

业务算法应保持纯 C#，可以在不启动 Excel/WPS 的情况下单元测试。

推荐结构：

```text
Command
  ├─ 参数与用户交互
  ├─ RangeTarget / Pipeline 编排
  └─ 调用纯业务 Service

Service
  ├─ 输入 CLR DTO / primitive
  ├─ 执行业务计算
  └─ 输出 CLR 结果
```

不要让 WPF ViewModel 或 Host 实现业务计算。

财务金额计算优先使用：

```csharp
decimal
```

如果存在舍入，必须显式确定中点规则，不依赖隐式默认行为。

---

## 15. 配置读写规范

AccuX 配置统一由 Core 的配置能力管理：

```text
IConfigManager
JsonConfigManager
JSON
可选 DPAPI
```

### 15.1 Module 不直接读写配置文件

禁止在业务模块里自行：

```text
File.ReadAllText(config.json)
File.WriteAllText(config.json)
创建 module-specific config manager
创建第二套 JSON 配置路径规则
直接调用 DPAPI 并绕过统一配置接口
```

应通过 `IAccuXContext` 或现有依赖注入路径取得统一 `IConfigManager`。

### 15.2 非敏感配置

普通功能设置使用 JSON 明文保存，例如：

```json
{
  "reconciliation": {
    "tolerance": 0.01,
    "largeSelectionWarning": 100000,
    "maxProcessCells": 500000
  }
}
```

上述数值仅示意配置结构；Range 阈值实际默认值应使用项目统一配置并依据 benchmark 确定。

### 15.3 敏感配置

设计允许使用 DPAPI 保存敏感配置。

V1 当前没有必须加密的 API Key，但新增模块若未来引入 Token、Secret 等敏感信息：

- 仍通过统一 ConfigManager 能力读写；
- 使用现有的可选 DPAPI 支持；
- 不在业务 Module 另建加密框架；
- 不把 Secret 写日志；
- 不把 Secret 放入代码、默认 JSON、测试 fixture 或异常消息。

### 15.4 配置接口签名尚未固定时的 Agent 行为

如果仓库实际 `IConfigManager` 方法签名与本文示意不同：

1. 先复用仓库现有接口；
2. 若确实缺少“模块化键空间 / 加密读写”等能力，优先最小扩展现有配置契约；
3. 不创建平行 `ReconciliationConfigManager`、`SecureConfigServiceV2` 等第二套基础设施；
4. 修改 Core 公共配置契约时必须补 Core Unit Test。

---

## 16. 日志规范

新增 Module / Command 必须使用统一 `ILogger`。

操作日志建议携带：

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
完整公式内容
工作表中的敏感业务数据
API Token / Secret
```

如果排错确实需要更多信息，必须优先记录结构化元数据，而不是用户业务内容。

---

## 17. 异常处理规范

### 17.1 Ribbon callback 不允许异常外泄

所有 Ribbon callback 必须遵循：

```text
try
  ↓
CommandDispatcher / Execute Command
  ↓
catch
  ↓
记录日志
  ↓
显示友好错误
```

任何异常不得穿出 COM callback。

### 17.2 Command 失败不能拖垮插件

- 单个 Command 失败，只影响当前操作；
- 单个 Module 初始化失败，不影响其他 Module；
- Host 或 Pipeline 的可预期校验失败应转为清晰结果，而不是依赖 COM Exception 作为业务判断流程。

### 17.3 用户提示应可执行

例如不要只显示：

```text
发生错误。
```

应尽量转换为：

```text
目标区域已失效，请重新选择区域后再次执行。
当前选区超过允许处理的最大单元格数量。
工作表受保护，目标区域不可写入。
部分公式无法安全转换，已跳过。
```

---

## 18. HostStateScope 使用规则

只有当当前操作确实需要修改 Application 全局状态时才创建 HostStateScope。

可管理的状态例如：

```text
ScreenUpdating
Calculation
EnableEvents
DisplayAlerts
StatusBar
```

规则：

> **修改了哪个状态，就可靠恢复哪个状态；没有必要修改，就不要为了形式统一而修改。**

`HostStateScope` 只负责宿主状态，不负责：

```text
工作表数据 Snapshot
Values / Formulas 恢复
AccuX Undo
Transaction Rollback
```

业务 Module 不直接修改这些 Application 状态。

---

## 19. WPF 使用规范

WPF 只负责：

- 参数输入；
- 参数校验的 UI 呈现；
- 用户确认；
- 结果展示；
- 设置界面。

禁止把核心业务算法写在：

```text
View code-behind
ViewModel
ValueConverter
```

窗口必须正确设置 Excel/WPS 主窗口为 Owner。

对于一键操作，不要为了“统一体验”强制弹复杂参数窗口。

最重要的是：

> 打开 WPF 前已经捕获的 RangeTarget 不得因为用户在窗口期间切换 Selection 而改变。

---

## 20. 什么时候允许扩展 Host / Core 宿主抽象

只有 Pipeline 或多个业务模块真正需要某项宿主能力，而且该能力属于以下类别之一时，才扩展 Host / Core 宿主抽象：

```text
Excel/WPS 存在实际差异
涉及公式安全
涉及批量 COM 性能
涉及宿主状态保存与恢复
多个模块重复且确有统一价值
```

扩展时遵守：

1. 接口应窄，只暴露调用方真正需要的能力；
2. DTO 使用纯 CLR 类型；
3. Core 只定义抽象，不引用 Host；
4. Host 实现接口；
5. AddIn 负责注入；
6. 不因为新增一个业务功能就建立完整 `IWorkbook/IWorksheet/IRange` 影子对象模型。

如果 Excel/WPS API 行为一致、调用简单，并且只在 Host/AddIn 边界内部使用，则无需额外抽象。

### 20.1 单个业务功能需要 COM 时

目录生成、批注、标记这类“属于某个业务功能的 COM 操作”不扩展 `IRangeOperationHost`，按以下流程落地：

1. 在 `AccuX.Core/Operations` 定义**功能窄接口**（如 `IWorkbookDirectoryHost` / `ICellCommentHost` / `ICellMarkHost`）；
2. 参数与结果使用 Core 的纯 CLR DTO（如 `DirectoryOptions` / `CellCommentTarget`），文案、配色、尺寸等业务决定由模块提供，Host 不内置功能专属常量；
3. 在 `AccuX.Host/Features` 增加实现类，从 `ExcelHostBase` 继承共享 COM 机制（工作簿/工作表解析、安全属性读取、状态作用域）；
4. AddIn 组合根为每个窄接口创建并注入独立实例。

`IRangeOperationHost` 只保留 Pipeline 真正需要的范围读写能力，新功能不要往上加方法。

---

## 21. 新增模块的推荐实施顺序

Agent 应按以下顺序工作，减少错误架构先行：

### 第 1 步：确认业务边界

输出并确认：

```text
Module Id
新增 Module 还是现有 Module 新功能
Command 列表
是否修改工作表
是否涉及公式
是否需要 WPF
是否需要新配置
是否发现新的 Excel/WPS 差异
```

### 第 2 步：检查已有公共能力

优先检查并复用：

```text
IAccuXModule / IAccuXContext
CommandDefinition / CommandDispatcher
RangeOperationPipeline
IRangeOperationHost
IWorkbookDirectoryHost / ICellCommentHost / ICellMarkHost
DirectoryOptions / CellCommentTarget
CellValueClassifier
RangeTarget / RangeWritePlan / FormulaInfo
IConfigManager
ILogger
IHostStateScope
```

发现已有能力时，不建立同功能平行实现。

### 第 3 步：先实现纯业务 Service 与测试

- Service 输入输出使用 CLR 类型；
- 覆盖正常、边界、错误场景；
- 不依赖 Excel/WPS 启动。

### 第 4 步：实现 Command 编排

- 一次捕获 RangeTarget；
- 参数 UI；
- Pipeline；
- 结果统计；
- 具体 Command 的重复执行规则。

### 第 5 步：只在确有需要时扩 Host

如果业务实现发现宿主差异：

- 收口到 Host；
- 补 Host 兼容实现；
- 更新 CompatibilityMatrix；
- 不把平台分支留在 Module。

如果只是本功能需要的新 COM 操作，按 §20.1 处理：Core 加功能窄接口 + 参数 DTO，
`AccuX.Host/Features` 加实现类，不扩展 `IRangeOperationHost`。

### 第 6 步：注册 Module 与 Ribbon

- AddIn 项目引用；
- 显式注册；
- 静态 Ribbon XML；
- Command ID 映射。

### 第 7 步：补配置、日志、异常与用户提示

这些不是“最后可选优化”，而是完成标准的一部分。

### 第 8 步：执行测试与双宿主验证

先 Unit Test，再 Excel，再 WPS。

---

## 22. 测试要求

### 22.1 Module Unit Test

所属 Module 的纯业务算法必须有 Unit Test。

至少覆盖与业务相关的：

- 正常输入；
- 边界值；
- 非法参数；
- 重复执行规则；
- 公式变换规则（如果模块涉及公式）；
- 不支持类型的处理策略。

### 22.2 修改型 Range 功能的公共测试矩阵

只要功能处理 Range，至少覆盖：

```text
[ ] 普通数字
[ ] 普通公式
[ ] Date / FormulaDate
[ ] Boolean / FormulaBoolean
[ ] 文本 / FormulaText
[ ] 空白
[ ] Error / FormulaError
[ ] 混合 Range
[ ] 不支持 / 特殊公式
[ ] 不破坏原公式
[ ] 大选区提示
[ ] 超过 maxProcessCells 拒绝
[ ] RangeTarget 在操作期间失效
[ ] 写回前可写性状态发生变化
```

### 22.3 Core 修改需要 Core Unit Test

如果新增模块导致修改：

```text
CellValueClassifier
RangeOperationPipeline
IConfigManager
公共 DTO
公共宿主抽象
Command 基础设施
```

必须增加或更新 Core Unit Test。

### 22.4 Excel / WPS 双宿主测试

一个涉及宿主交互的功能，不能只在 Excel 通过就算完成。

必须：

```text
[ ] Excel 测试通过
[ ] WPS 测试通过
```

若出现差异，应由 Host 解决并记录兼容矩阵。

### 22.5 F5 调试链路不得破坏

开发版本应继续支持 Visual Studio 直接启动 Excel 并命中：

```text
COM Add-in 入口
Ribbon callback
CommandDispatcher
RangeOperationPipeline
AccuX.Host
业务 Module
```

---

## 23. 新增模块 Definition of Done

Agent 在宣布模块完成前，逐项检查：

### 23.1 架构

```text
[ ] 新业务代码默认位于所属 Module，而非随意放入 Core
[ ] Core 没有新增对 Host / Excel / WPS Interop 的引用
[ ] Module 没有引用或缓存 Excel/WPS COM 对象
[ ] Module 没有 Excel/WPS 平台 if/else 分支
[ ] 没有新增动态插件扫描 / 反射加载 / 热加载框架
[ ] 新 Module 已由 AddIn 编译期引用并显式注册
[ ] Ribbon 使用静态 XML + Command ID 映射
```

### 23.2 Range 与数据安全

```text
[ ] 一次 Command 只捕获一次 RangeTarget
[ ] 参数窗口后没有重新读取 Selection
[ ] RangeTarget 不携带 COM 对象
[ ] 使用批量 Read / Write，而非逐 Cell COM 高频操作
[ ] 写回前完成全部读取、分类、业务计算和完整 WritePlan
[ ] 写回前再次验证同一个 RangeTarget
[ ] 超过 maxProcessCells 直接拒绝
[ ] 没有实现与 fail-before-write 冲突的分块边读边写
```

### 23.3 数据类型与公式

```text
[ ] 使用 CellValueClassifier
[ ] Date / Boolean 不被误识别为金额
[ ] Formula 不默认转换为 Value
[ ] 无法安全转换的 Formula 被跳过并统计
[ ] Formula API / 本地化差异由 Host 处理
[ ] Formula 规范化表达式在 Module 内按业务转换
[ ] FormulaTransformService 没有自动“防止重复包裹”
[ ] 重复执行规则在具体 Command / 业务 Service 中测试
```

### 23.4 COM 与线程

```text
[ ] Core / Module 不保存 Application / Workbook / Worksheet / Range
[ ] Module 不调用 Marshal.ReleaseComObject / FinalReleaseComObject
[ ] COM Read / Write 在宿主 UI / STA 线程
[ ] 没有 Task.Run 直接访问 Excel/WPS COM
```

### 23.5 配置、日志、异常

```text
[ ] 配置统一通过 IConfigManager / ConfigManager
[ ] Module 没有自行维护第二套 JSON 文件
[ ] 敏感配置如有需要走统一 DPAPI 能力
[ ] 日志不记录真实金额、公式内容或 Secret
[ ] Ribbon callback 异常不会穿出 COM 边界
[ ] 单个 Command / Module 失败不会导致整个插件失效
[ ] 用户能看到明确的成功 / 跳过 / 失败结果
```

### 23.6 测试

```text
[ ] 所属 Module Unit Test 通过
[ ] 如修改 Core，Core Unit Test 通过
[ ] 普通数字 / 普通公式测试通过
[ ] Date / Boolean 跳过或业务策略测试通过
[ ] 文本 / 空白 / Error 测试通过
[ ] 混合 Range 测试通过
[ ] 大量 Range 不逐 Cell COM 操作
[ ] Excel 测试通过
[ ] WPS 测试通过
[ ] 如修改宿主全局状态，异常后能恢复实际修改的状态
```

---

## 24. Agent 禁止清单

以下行为默认禁止。只有当前任务明确要求改变对应架构，并且同时完成必要的公共契约、兼容性、测试和迁移设计时，才允许例外：

```text
禁止业务 Module 直接拿 Excel.Range / WPS Range
禁止业务 Module 缓存 Selection 用于稍后写回
禁止参数窗口关闭后重新决定写入 Selection
禁止 Core 引用 AccuX.Host
禁止 Core 引用 Excel/WPS Interop
禁止为每个业务功能各写一套 Range 遍历框架
禁止大量逐 Cell COM Read / Write
禁止 Formula → Value 作为默认处理
禁止把日期序列值当普通金额
禁止把 Boolean 当 0/1 金额处理
禁止 FormulaTransformService 自动去重或“防止重复包裹”
禁止业务代码自行处理 Formula / Formula2 / 本地化公式差异
禁止业务层自行 ReleaseComObject
禁止 Task.Run 内直接访问宿主 COM
禁止 Module 直接修改 ScreenUpdating / Calculation / EnableEvents 等状态
禁止 Module 自己读写 JSON 配置文件
禁止另建一套日志框架
禁止动态 DLL 扫描 / 反射插件发现 / 热加载
禁止为了“未来可能复用”提前把业务 Service 提升到 Core
禁止实现未经设计批准的 Snapshot / AccuX Undo / Transaction Rollback
```

---

## 25. Agent 需要提交的开发说明

每次新增 Module 或大型 Command，Agent 最终应附带一份简短变更说明，至少包含：

```text
1. 模块 / Command
   - Module Id：
   - Command Id：
   - 业务目的：

2. 代码位置
   - 新增工程 / 目录：
   - 主要文件：

3. 公共能力复用
   - Pipeline：是 / 否 / 不适用
   - CellValueClassifier：是 / 否 / 不适用
   - ConfigManager：是 / 否
   - Logger：是 / 否

4. 架构变更
   - 是否修改 Core：
   - 为什么必须修改：
   - 是否修改 Host：
   - 新增/修改的宿主窄接口或 Features 实现类：
   - 对应 Excel/WPS 差异：
   - 是否更新 CompatibilityMatrix：

5. 数据安全
   - RangeTarget 捕获时点：
   - 是否写回工作表：
   - fail-before-write 实现方式：
   - maxProcessCells 行为：
   - Formula 处理规则：
   - 重复执行规则：

6. 配置
   - 新增配置键：
   - 是否敏感：
   - 明文 JSON / DPAPI：

7. 测试
   - Module Unit Test：
   - Core Unit Test：
   - Excel：
   - WPS：
   - 特殊 / 边界场景：
```

该说明用于帮助后续 Agent 快速判断本次开发是否破坏现有边界。

---

## 26. Agent 的默认决策规则

当需求不完整、但无需向用户追问即可安全推进时，使用以下默认值：

| 问题                                     | 默认决策                                                 |
| ---------------------------------------- | -------------------------------------------------------- |
| 新功能放哪里？                           | 先放所属 Module                                          |
| 是否提升 Core？                          | 否，除非已有真实跨模块复用                               |
| 是否新增 Host Adapter？                  | 仅当发现真实宿主差异 / 公式安全 / COM 性能需求           |
| 新功能需要 COM 怎么办？                  | Core 加功能窄接口 + 参数 DTO，Host/Features 加实现类，不扩展 IRangeOperationHost |
| 是否新增模块动态发现？                   | 否                                                       |
| 是否动态生成 Ribbon？                    | 否，使用静态 Ribbon XML                                  |
| Range 目标怎么确定？                     | Command 开始时 CaptureTarget 一次                        |
| 参数窗口后是否重新读 Selection？         | 否                                                       |
| 大 Range 怎么处理？                      | 阈值内整块批量；超过 maxProcessCells 拒绝                |
| 是否 Chunk 边读边写？                    | 否                                                       |
| 是否逐 Cell COM？                        | 否                                                       |
| 公式怎么处理？                           | Host 规范化，Module 做业务变换，安全时 Formula → Formula |
| 重复操作怎么处理？                       | 默认每次主动 Command 都是新操作；具体规则归 Command      |
| Date / Boolean 怎么处理？                | 财务数值类功能默认跳过                                   |
| COM 能否进 Module/Core？                 | 不能                                                     |
| COM 能否在 Task.Run 使用？               | 不能                                                     |
| 配置怎么存？                             | 统一 ConfigManager；非敏感 JSON，敏感可选 DPAPI          |
| 日志能否记录用户金额/公式？              | 默认不能                                                 |
| 是否实现 Undo / Snapshot / Transaction？ | V1 不实现                                                |

---

## 27. 最小示例：新增一个修改型业务 Command 时的思维模型

假设在 `AccuX.Modules.Reconciliation` 中新增“统一差异容差处理”命令，Agent 应按以下方式拆分：

```text
Ribbon XML
  ↓ Command ID: accux.reconciliation.normalize-diff
CommandDispatcher
  ↓
NormalizeDifferenceCommand
  ├─ CaptureTarget() 一次
  ├─ 打开参数窗口（如需要）
  ├─ 读取 ConfigManager 默认容差
  ├─ 调用 RangeOperationPipeline
  └─ 显示 OperationResult
         ↓
RangeOperationPipeline
  ├─ Host.Read(target)
  ├─ CellValueClassifier
  ├─ 调用 NormalizeDifferenceService（纯 C#）
  ├─ 完整生成 RangeWritePlan
  ├─ Host.ValidateWrite(target, writePlan)
  └─ Host.Write(target, writePlan)
         ↓
Excel / WPS Host implementation
```

不应变成：

```text
NormalizeDifferenceCommand
  ↓
直接获取 Excel.Application.Selection
  ↓
foreach (Excel.Range cell in selection.Cells)
  ↓
Task.Run 中计算并写回
  ↓
if (isWps) 特殊处理
```

---

## 28. 最后原则

后续 Agent 新增模块时，优先守住以下六条：

1. **Module 只做业务，Host 只做宿主差异，Core 只做真实公共能力，AddIn 只做组合与入口。**
2. **一次 Command 只捕获一次 RangeTarget，永远不在写回阶段跟随新的 Selection。**
3. **在 `maxProcessCells` 内批量读、内存算、生成完整 WritePlan、验证后再写；V1 不边读边写。**
4. **COM 不进入 Core / Module，COM Read / Write 留在宿主 UI / STA 线程。**
5. **公式平台差异由 Host 规范化；具体业务公式变换与重复执行规则留在所属 Command / Module。**
6. **不为假设未来做过度抽象：没有真实跨模块复用，就不要把业务能力提升到 Core，也不要建设动态插件系统。**

如果一个实现同时满足这六条，并通过本指南的 Definition of Done，通常就是符合 AccuX 当前架构方向的实现。

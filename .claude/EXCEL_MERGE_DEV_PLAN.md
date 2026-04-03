# ExcelMerge 二次开发需求文档

> 本文档供 Claude Code CLI 在开发过程中参考，用于理解项目背景、代码结构、功能差距和开发计划。

## 1. 项目背景

### 1.1 问题

团队约 150 人，使用 SVN (TortoiseSVN) 管理项目，项目是 Unity 手游，资产中包含大量 Excel 文件（策划表等）。SVN 内置的 diff 工具无法有效比较 Excel 二进制文件，导致无法快速查看版本间差异。

### 1.2 目标

基于开源项目 **ExcelMerge**（MIT License）进行二次开发，对标商业软件 **xlCompare** 的核心功能，打造一款适用于 SVN 工作流的 Excel Diff & Merge 工具。

### 1.3 基础仓库

- **原始仓库地址**: https://github.com/skanmera/ExcelMerge
- **备注：当前开发仓库地址**: https://github.com/CGWang/ExcelMerge
- **语言**: C# / WPF
- **协议**: MIT
- **Stars**: ~824
- **最后 Release**: v1.3.4 (2018-01-04)
- **提交数**: 171

## 2. 现有代码结构

```
ExcelMerge.sln
├── ExcelMerge/                  # 核心 diff 引擎库
│   ├── ExcelSheet.cs            # Excel 文件读取 + Sheet 级 diff 算法
│   ├── ExcelRow.cs              # 行数据模型
│   ├── ExcelCell.cs             # 单元格数据模型
│   ├── ExcelColumn.cs           # 列数据模型
│   ├── ExcelSheetDiff.cs        # Sheet diff 结果数据结构
│   ├── ExcelSheetDiffConfig.cs  # diff 配置（header 行、列等）
│   ├── ExcelCellStatus.cs       # 枚举：None/Modified/Added/Removed
│   ├── ExcelColumnStatus.cs     # 列状态枚举
│   └── ExcelRowStatus.cs        # 行状态枚举
│
├── ExcelMerge.GUI/              # WPF 图形界面
│   ├── MainWindow.xaml(.cs)     # 主窗口
│   ├── Views/                   # diff 视图组件
│   ├── ViewModels/              # MVVM ViewModel
│   ├── Settings/                # 颜色、外部命令等配置
│   └── Converters/              # WPF 值转换器
│
├── FastWpfGrid/                 # 自研高性能虚拟化网格控件（核心资产）
│   ├── FastGridControl.cs       # 网格主控件，支持大量单元格虚拟化渲染
│   └── ...                      # 行/列模型、滚动、选择等
│
├── NetDiff/                     # 通用 diff 算法库
│   ├── DiffUtil.cs              # 基于 LCS 的 diff 算法实现
│   ├── DiffResult.cs            # diff 结果（Modified/Added/Removed）
│   └── ...
│
├── ExcelMerge.ShellExtension/   # Windows Explorer 右键菜单集成
├── ExcelMerge.Installer/        # WiX MSI 安装包项目
├── .nuget/                      # NuGet 配置
└── README.md
```

### 2.1 关键依赖

- **NPOI** 或 **ClosedXML**：Excel 文件读写（需确认具体用哪个，检查 .csproj 的 NuGet 引用）
- **FastWpfGrid**：项目自带的高性能 WPF DataGrid 替代品，支持虚拟化渲染大量单元格
- **NetDiff**：项目自带的 LCS diff 算法库

### 2.2 现有功能

- ✅ 2-way side-by-side diff 显示
- ✅ 支持 xls, xlsx, csv, tsv
- ✅ 单元格级别差异高亮（修改/新增/删除用不同颜色）
- ✅ 行级别 diff（基于 LCS 算法对齐）
- ✅ 列级别 diff（有已知 bug）
- ✅ 多 Sheet 切换
- ✅ 指定 header 行/列进行 diff
- ✅ 命令行接口（`ExcelMerge.GUI diff -s <src> -d <dst>`）
- ✅ Git / Mercurial 外部 diff tool 集成
- ✅ Windows 右键菜单集成
- ✅ 差异导航快捷键（Ctrl+方向键跳转下一个差异）
- ✅ diff log 输出
- ✅ 自定义颜色配置
- ✅ 外部命令注册（fallback 到其他 diff 工具）

### 2.3 已知问题

- 列插入/删除时显示位置可能不正确（README 中已标注 Known Problem）
- 项目基于 .NET Framework 4.x，依赖较旧
- 无 merge 功能（原作者标注为目标但从未实现）
- 不支持公式、注释、VBA 的比较

## 3. 功能差距分析（对标 xlCompare）

### 3.1 优先级定义

- **P0**: 必须实现，直接解决当前痛点
- **P1**: 重要功能，显著提升体验
- **P2**: 锦上添花，可后续迭代

### 3.2 功能清单

| 优先级 | 功能 | 当前状态 | 说明 |
|--------|------|----------|------|
| **P0** | SVN (TortoiseSVN) 命令行集成 | ❌ 缺失 | 需支持 `%base %mine %bname %yname` 参数格式 |
| **P0** | 修复行对齐算法 | ⚠️ 有 bug | 插入/删除行时 LCS 对齐不准确 |
| **P0** | 修复列对齐算法 | ⚠️ 有 bug | 列增删时显示位置错误（README Known Problem） |
| **P0** | 构建现代化 | ⚠️ 过时 | 确认能在 VS2022 正常构建，更新 NuGet 依赖 |
| **P1** | 公式 diff 模式 | ❌ 缺失 | 可选择比较"值"或"公式"，底层库应已支持读取公式 |
| **P1** | 单元格注释/批注 diff | ❌ 缺失 | 显示注释差异 |
| **P1** | diff 报告导出 | ⚠️ 仅 log | 导出为 xlsx 差异报告（新工作簿，标记所有差异） |
| **P1** | 单元格内字符级 diff 高亮 | ❌ 缺失 | 长文本单元格中精确显示哪些字符变了 |
| **P1** | 忽略规则 | ❌ 缺失 | 忽略空白差异、数值精度差异（如 0.001 以内视为相同） |
| **P2** | 3-way diff 显示 | ❌ 缺失 | 显示 BASE / LOCAL / REMOTE 三方差异 |
| **P2** | Merge 操作 | ❌ 缺失 | 从 diff 视图中选择性合并变更 |
| **P2** | VBA 宏代码 diff | ❌ 缺失 | 提取 VBA 源码做文本 diff |
| **P2** | 单元格格式/样式 diff | ❌ 缺失 | 字体、背景色、边框等格式变更 |
| **P2** | 命名范围 diff | ❌ 缺失 | 比较 Named Ranges 定义差异 |

## 4. 分阶段开发计划

### Phase 1: 实用化改造（P0 全部 + 部分 P1）

**目标**: 团队能在 TortoiseSVN 中一键调起，准确查看 Excel 文件版本差异。

#### 4.1.1 构建现代化

- [ ] 用 VS2022 打开解决方案，确认构建状态
- [ ] 更新所有 NuGet 包到兼容版本
- [ ] 修复可能的编译错误/警告
- [ ] 确认 MSI 安装包能正常生成
- [ ] **决策点**: 是否升级到 .NET 6/8（如保持 .NET Framework 4.x 则用户无需安装运行时；如升级则用 Self-Contained 发布）

#### 4.1.2 SVN 命令行集成

TortoiseSVN 外部 diff 工具的调用方式：

```
"<tool_path>" %base %mine
```

带标题的完整格式：

```
"<tool_path>" %base %mine /title1=%bname /title2=%yname /leftreadonly
```

**需要做的改动**:

1. 扩展命令行参数解析，增加 `--base-name` 和 `--mine-name` 参数用于窗口标题显示
2. 增加 `--readonly-left` 参数（base 文件应该只读）
3. 增加 `--quit-on-close` 参数（关闭 diff 窗口时退出进程）
4. 支持位置参数（`ExcelMerge.GUI diff file1 file2`，不必写 -s -d）
5. 文档中附 TortoiseSVN 配置步骤

TortoiseSVN 配置路径：Settings → Diff Viewer → Advanced → 添加 `.xlsx` / `.xls` / `.csv` 扩展名，指向：

```
"C:\Program Files\ExcelMerge\ExcelMerge.GUI.exe" diff -s %base -d %mine --quit-on-close
```

#### 4.1.3 修复行对齐算法

**问题**: NetDiff 模块使用 LCS 算法做行匹配，当存在行插入/删除时，对齐结果不准确。

**排查方向**:

1. 检查 `ExcelSheet.Diff()` 方法中行比较的逻辑
2. 检查 `RowComparer` 的相等性判断——可能需要支持"模糊匹配"（大部分列相同即认为是同一行）
3. 检查 header 行提取是否正确影响了 diff 起始位置
4. 增加基于"key column"的行匹配模式（用户指定某一列作为主键来对齐行，如 ID 列）

**测试方法**: 准备测试 xlsx 文件对：
- `test_row_insert.xlsx` — 在中间插入 3 行
- `test_row_delete.xlsx` — 删除中间 3 行
- `test_row_move.xlsx` — 行顺序调换
- 验证 diff 结果中 Added/Removed 行标记正确

#### 4.1.4 修复列对齐算法

与行对齐类似，检查 `ExcelSheet.Diff()` 中列处理逻辑。参考 `ExcelColumnStatus` 和 `columnStatusMap` 的生成过程。

#### 4.1.5 公式 diff 模式（P1，建议 Phase 1 一并完成）

- 在 `ExcelCell` 中增加 `Formula` 属性（读取时从 NPOI/ClosedXML 获取）
- 在 UI 中增加一个 toggle："比较值" / "比较公式"
- `ExcelSheetDiffConfig` 增加 `CompareFormula` 配置项
- 比较逻辑中根据配置取 `Value` 或 `Formula` 进行比较

### Phase 2: 体验增强（P1 剩余）

#### 4.2.1 单元格注释 diff

- 读取 Excel 注释（NPOI: `cell.CellComment`）
- 在 diff 视图中标记有注释差异的单元格（如角标图标）
- 点击可查看注释内容对比

#### 4.2.2 diff 报告导出为 xlsx

- 创建新工作簿，复制原始数据
- 用背景色标记差异（绿色=新增，红色=删除，黄色=修改）
- 增加一个 Summary sheet 列出所有差异的坐标和内容

#### 4.2.3 字符级 diff 高亮

- 对 Modified 状态的单元格，做文本级 diff（可复用 NetDiff 的 LCS）
- 在单元格详情面板中高亮显示具体哪些字符发生了变化

#### 4.2.4 忽略规则

- 忽略空白差异（trim 后比较）
- 数值精度容差（如 `|a - b| < 0.001` 视为相同）
- 忽略指定 sheet / 行范围 / 列范围

### Phase 3: 高级功能（P2）

#### 4.3.1 3-way Diff

- UI 扩展为三面板（BASE / LOCAL / REMOTE）
- NetDiff 扩展为 3-way diff 算法（以 BASE 为基准，分别计算 BASE→LOCAL 和 BASE→REMOTE 的变更，检测冲突）
- 冲突定义：同一单元格在 LOCAL 和 REMOTE 中都相对 BASE 发生了不同变更

SVN merge 调用格式：

```
"<tool_path>" /base:%base /mine:%mine /theirs:%theirs /merged:%merged
```

#### 4.3.2 Merge 操作

- 在 diff 视图中，用户可以选择"采用左侧"或"采用右侧"
- 支持逐单元格、逐行、逐 sheet 的合并操作
- 合并结果保存到指定文件

#### 4.3.3 VBA 代码 diff

- 解压 xlsm 文件读取 vbaProject.bin
- 提取 VBA 模块源码
- 用标准文本 diff 显示 VBA 代码差异
- 可参考开源项目 `xltrail/git-xl` 的 VBA 提取逻辑

## 5. 技术约束与决策

### 5.1 .NET 版本选择

**决定: 升级到 .NET 8（LTS）**

理由：
- .NET Framework 4.5.2 已于 2022 年 EOL，不再接收安全更新
- 主要依赖（NPOI、Prism 等）的新版本已迁移到 .NET 6+，留在 Framework 会导致依赖锁死在旧版本
- 新 C# 语言特性（模式匹配、nullable reference types、record 等）提升代码质量和开发效率
- 性能更优，尤其大文件 diff 场景受益明显
- WPF 在 .NET 8 上完全支持，API 兼容性极高，迁移成本可控

部署方案：
- 使用 Self-Contained 发布（`--self-contained true -p:PublishSingleFile=true`）
- 用户无需安装 .NET 运行时，单个 exe 或 MSI 分发
- 安装包约 80-150MB，对桌面工具完全可接受
- 团队 150 人统一 Windows 环境，推送到共享目录即可

迁移要点：
- 项目文件从传统 .csproj 转为 SDK-style 格式
- packages.config 迁移为 PackageReference
- NuGet 依赖全部升级到 .NET 8 兼容版本
- .vdproj 安装项目需替换为其他方案（如 Inno Setup 或 dotnet publish 直接分发）

### 5.2 Excel 读写库

检查当前项目使用的 Excel 库（大概率是 NPOI 或 ClosedXML），确认：
- 是否支持读取公式（不只是计算后的值）
- 是否支持读取批注/注释
- 是否支持 xlsm（含宏的工作簿）
- xls（老格式）的兼容性

### 5.3 部署方式

目标：生成 MSI 安装包 或 单个 exe，可通过内部共享分发，用户无需手动安装 .NET 运行时。

### 5.4 不依赖 Excel

ExcelMerge 当前不依赖 Microsoft Excel 安装，二次开发中必须保持这一特性。使用 NPOI / ClosedXML 等纯库方式读写 Excel。

## 6. SVN 集成配置指南（最终交付物之一）

开发完成后需提供团队配置文档，内容包括：

### TortoiseSVN Diff Viewer 配置

1. 右键 → TortoiseSVN → Settings
2. Diff Viewer → Advanced
3. 添加扩展名 `.xlsx`，命令行设置为：

```
"<安装路径>\ExcelMerge.GUI.exe" diff -s %base -d %mine --quit-on-close
```

4. 添加扩展名 `.xls`，同上
5. 添加扩展名 `.csv`，同上

### TortoiseSVN Merge Tool 配置（Phase 3 完成后）

```
"<安装路径>\ExcelMerge.GUI.exe" merge --base %base --mine %mine --theirs %theirs --output %merged
```

## 7. 测试策略

### 7.1 单元测试（Claude Code 可独立完成）

针对核心 diff 引擎编写单元测试，覆盖：
- 空文件 diff
- 相同文件 diff（应无差异）
- 单个单元格修改
- 行插入/删除/移动
- 列插入/删除
- 多 Sheet diff
- 公式 vs 值比较
- 大文件性能（1000+ 行）

测试文件放在 `TestData/` 目录下，使用 xUnit 或 NUnit 框架。

### 7.2 GUI 验证（需人工操作）

Claude Code 无法看到 WPF 界面，以下场景需要开发者手动验证并反馈：
- diff 视图颜色是否正确
- 行/列对齐是否与预期一致
- 快捷键导航是否正常
- 大文件滚动是否流畅
- SVN 调起是否正常工作

反馈格式建议：

```
[场景] 打开 test_row_insert.xlsx vs test_base.xlsx
[期望] 第 5-7 行应显示为绿色（新增）
[实际] 第 5 行绿色正确，第 6-7 行偏移到了第 8-9 行
[截图] （如有）
```

## 8. 开发注意事项

### 8.1 代码风格

- 保持与现有代码风格一致（C# 命名规范、MVVM 模式）
- 新增文件使用与项目一致的命名空间
- 每个功能模块提交前确保编译通过

### 8.2 向后兼容

- 保持现有命令行参数兼容（-s, -d, -c 等不变）
- 新增参数使用 `--long-option` 风格
- 保持右键菜单集成功能可用

### 8.3 Git 工作流

- 每个 Phase 一个分支（`feature/phase1-svn-integration` 等）
- 每个子功能一个 commit，message 清晰
- Phase 完成后合并到 main

### 8.4 FastWpfGrid 是核心资产

`FastWpfGrid` 是这个项目最有价值的组件——它是专门为大量单元格 diff 显示优化的虚拟化 WPF 网格控件。修改它时要格外小心：
- 不要破坏虚拟化渲染逻辑（它只渲染可见区域）
- 单元格颜色渲染逻辑在这里，diff 颜色的改动需要同步修改
- 滚动同步（左右面板同步滚动）也在这里实现

## 9. 参考资料

- **xlCompare 官网**（对标功能参考）: https://xlcompare.com
- **xlCompare SVN 集成文档**: https://xlcompare.com/svn-integration.html
- **TortoiseSVN 外部 Diff 工具配置**: https://tortoisesvn.net/docs/release/TortoiseSVN_en/tsvn-dug-diff.html
- **ExcelMerge 原始仓库**: https://github.com/skanmera/ExcelMerge
- **Git XL (VBA diff 参考)**: https://github.com/xltrail/git-xl
- **xlsxDiff (Python diff 算法参考)**: https://github.com/rafal-dot/xlsxDiff

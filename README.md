# 库存批次匹配工具 (OrderMatchTool)

一款面向库存/分单场景的桌面小工具：从 Excel 导入"批次号 + 货物型号 + 库存"数据后，按指定型号 + 目标库存数量，自动用**最少的批次条数**凑出该数量，并支持一键把结果整列复制到 Excel。

> 例：仓库里某型号有若干条不同批次的库存（10 / 15 / 5 / 20…），客户要 30 件 —— 工具自动选出"用最少几条批次记录刚好凑出 30"的最优组合。

---

## 功能特性

- **Excel 导入**：支持 `.xlsx`（openpyxl）和 `.xls`（xlrd）两种格式，自动识别表头。
- **数据存储**：使用 SQLite 数据库，数据持久化在 exe 同目录的 `data.db` 文件中。
- **可切换数据库**：支持指定任意 `.db` 文件作为当前数据源（多人协作时可指向共享文件夹中的同一个数据库）。
- **型号筛选**：左上侧可按货物型号过滤数据列表，便于核对。
- **最优批次匹配**：基于 0/1 背包动态规划，**用最少的批次号数量**精确凑出目标库存（每条批次记录最多使用一次）。
- **结果一键复制**：
  - 复制批次号列
  - 复制库存列
  - 复制"批次号 + 库存"两列
  - 复制并删除 / 仅删除 已匹配的批次记录
- **打包为单文件 EXE**：通过 PyInstaller 一键打包成 `OrderMatchTool.exe`，无需 Python 环境即可运行。

---

## Excel 文件格式

工具会自动识别表头（第 1 行）中的关键字，找到对应的三列。**列名包含以下任一关键字**即可：

| 必需列 | 可识别的关键字（包含即可） |
| --- | --- |
| 批次号 | `批次号` / `批次` / `分单号` / `单号` / `订单号` / `编号` |
| 货物型号 | `货物型号` / `型号` / `品类` / `类别` / `分类` / `类型` / `品名` |
| 库存 | `库存` / `件数` / `数量` / `数目` / `件` |

示例表头（第一行）：

| 批次号 | 货物型号 | 库存 |
| --- | --- | --- |
| B20240101-001 | A型 | 10 |
| B20240101-002 | A型 | 15 |
| B20240101-003 | B型 | 8 |

无效行（缺关键列、库存非正数、空值）会被自动跳过并在导入完成后显示跳过条数。

---

## 使用方式（终端用户）

### 方式一：直接下载 EXE（推荐）

1. 进入仓库 [Releases](../../releases) 或最新一次绿色 [Actions Run](../../actions) 的 Artifacts，下载 `OrderMatchTool-windows-exe`。
2. 解压得到 `OrderMatchTool.exe`，放到任意目录。
3. 双击运行。首次运行会在 exe 同目录自动创建 `data.db`。

### 方式二：从源码运行

需要 Python 3.8+：

```bash
pip install openpyxl xlrd
python app.py
```

> 仅处理 `.xlsx` 文件可不安装 `xlrd`；只需处理 `.xls` 文件可不安装 `openpyxl`，但通常建议两个都装上。

---

## 操作流程

1. **导入 Excel**：点左上角"导入 Excel"，选择文件，确认导入结果。
2. **（可选）筛选型号**：用上方下拉框筛选要核对的货物型号。
3. **匹配查询**：在下半部分选择"货物型号" + 输入"目标库存"，点"开始匹配"。
4. **复制结果**：选用合适的复制按钮，把批次号 / 库存列直接粘贴到目标 Excel 对应列即可。
5. **删除已用记录**（可选）：点"复制并删除记录"或"仅删除记录"，把已被领用的批次从数据库里清掉。

> 可点右上角"切换数据库"指向团队共享的 `.db` 文件，配合网络共享盘实现多人协作。所选路径会被记录在 exe 同目录的 `db_config.json` 中。

---

## 匹配算法说明

- 目标：在某型号的全部批次记录中，**找出库存数之和恰好等于目标值、且使用记录条数最少**的子集。
- 实现：标准 0/1 背包 DP，状态 `dp_count[s]` = "凑出和 s 所需的最少批次条数"，转移时同时保存到达每个状态的完整 path（每条批次记录最多用一次）。
- 上限：单次匹配最多考虑前 300 条候选记录。
- 若所有批次库存之和小于目标值、或无法精确凑出目标值，会提示"未找到等于该数量的组合"。

---

## 开发 / 打包

### 本地开发

```bash
pip install openpyxl xlrd pyinstaller
python app.py
```

### 本地打包成 EXE（Windows）

仓库已附带打包脚本：

```cmd
build.bat
```

或手动执行：

```bash
pyinstaller --onefile --noconsole --name OrderMatchTool --clean app.py
```

输出文件：`dist/OrderMatchTool.exe`。

### CI 自动打包

仓库 `.github/workflows/` 已配置：每次 push 或 PR 改动 `app.py` 时，GitHub Actions 会在 Windows runner 上自动执行 PyInstaller 构建，并把 `OrderMatchTool.exe` 上传为 artifact，可在对应 run 页面直接下载。

---

## 文件说明

| 文件 | 用途 |
| --- | --- |
| `app.py` | 主程序（GUI、数据库、Excel 解析、匹配算法） |
| `OrderMatchTool.spec` | PyInstaller 打包配置 |
| `build.bat` | Windows 本地一键打包脚本 |
| `.github/workflows/` | CI（在 Windows 上自动打 EXE） |
| `data.db` | 运行时自动创建的 SQLite 数据库（已 gitignore） |
| `db_config.json` | 记录最近一次选择的数据库路径（已 gitignore） |

---

## 依赖

- Python 3.8+
- [openpyxl](https://pypi.org/project/openpyxl/)（读 `.xlsx`）
- [xlrd](https://pypi.org/project/xlrd/)（读 `.xls`，**xlrd 2.0+ 不再支持 .xlsx**，本工具仅用其处理 .xls，符合限制）
- [PyInstaller](https://pyinstaller.org/)（打包阶段使用）
- Tkinter（Python 标准库，Windows 自带）

---

## License

仅供内部使用，未声明开源协议。

# Excel 公式转数值工具

把一份 Excel 工作簿里**所有工作表的公式**一次性转成数值，图表、图片、合并单元格、条件格式、
行高列宽全部原样保留；**原始文件全程只读，不会被修改**。

---

## 一、目录内容

| 文件 | 作用 |
|---|---|
| `xlsx_deformula.py` | 主程序（图形界面 + 命令行两种用法，纯 Python 标准库，零第三方依赖） |
| `一键打包exe.bat` | 在 Windows 上双击，自动装 PyInstaller 并生成单文件 exe |
| `selftest.py` | 自检（造样例→转换→校验公式/数值/图片单元格/源文件不变），CI 与本地都能跑 |
| `make_ico.py` | 由 `icon.png` 生成纯 BMP 条目的 `icon.ico`（PIL 默认写 PNG 条目，PyInstaller 嵌入时易报错） |
| `icon.ico` / `icon.png` | 程序图标 |
| `.github/workflows/build-exe.yml` | 没有 Windows 电脑时，用 GitHub Actions 云端打包 exe |
| `README.md` | 本说明 |

---

## 二、直接运行（已装 Python 的电脑）

```bat
python xlsx_deformula.py
```

命令行模式（批量、可写进批处理）：

```bat
python xlsx_deformula.py --cli "D:\日报\2026.09.03基坑监测日报.xlsx"
python xlsx_deformula.py --cli "源文件.xlsx" "D:\输出\结果.xlsx"
```

不写第二个参数时，输出文件自动放在源文件同目录，文件名加 `_无公式`；
若同名文件已存在，会自动加 `(2)`、`(3)`，不会覆盖任何文件。

---

## 三、打包成 Windows 单文件 exe（内含运行库，目标机无需 Python）

### 方式 A：自己的 Windows 电脑上打包（推荐，1 分钟）

1. 装 Python（只需这台**打包机**装，最终用户不用装）：
   <https://www.python.org/downloads/windows/> ，安装时勾选 *Add python.exe to PATH*。
   > 官方安装包自带 `tkinter`，本项目不需要任何第三方库。
2. 把整个文件夹拷到 Windows 电脑。
3. **双击 `一键打包exe.bat`**。脚本会：找 Python → 建 `.buildenv` 虚拟环境 →
   装 PyInstaller → 打包 → 自动打开 `dist` 文件夹。
4. 产物：`dist\Excel公式转数值.exe`（约 10–12 MB）。
   拷到任何 64 位 Windows 双击即用，**目标电脑无需安装 Python 或任何库**。

手动等价命令（想自己敲/T 换参数时）：

```bat
python -m pip install pyinstaller
pyinstaller --noconfirm --clean --onefile --windowed ^
  --name "Excel公式转数值" --icon icon.ico ^
  --distpath dist --workpath build --specpath build ^
  xlsx_deformula.py
```

- `--onefile` 打成单个 exe（启动会先解压到临时目录，约 1–2 秒）；
- 想启动更快、且不介意一个文件夹，把 `--onefile` 换成 `--onedir`；
- `--windowed` 表示不弹黑色控制台窗口；需要命令行模式时保留控制台就用 `--console`。

### 方式 B：没有 Windows 电脑 —— GitHub Actions 云端打包

1. 新建一个 GitHub 仓库，把本文件夹（含 `.github/` 目录）推上去；
2. 仓库 → **Actions** → 选 `build-windows-exe` → **Run workflow**；
3. 几分钟后在该次运行页面的 **Artifacts** 下载 `Excel公式转数值-windows.zip`，
   解压即得 exe（Windows 官方运行器打包，自带校验和冒烟测试步骤）。

> 说明：PyInstaller **不能跨平台交叉编译**——Linux 上只能产出 Linux 可执行文件，
> Windows exe 必须在 Windows 上打（本机或 GitHub 的 Windows 运行器）。

---

## 四、界面用法

1. **① 选择文件**：`添加文件…` 可多选 `.xlsx` / `.xlsm`；
2. **② 输出位置**：三种方式任选
   - 与原文件同目录，自动加 `_无公式`（默认）
   - 输出到指定文件夹（自动保留原文件名 + 后缀）
   - 保存为指定文件（仅单个源文件时可用）
3. 点 **开始转换**，下方日志会逐文件显示：工作表数、公式单元格数、结果路径。
4. `打开输出文件夹` 直接定位结果。

安全约束：程序拒绝把结果写到源文件路径上；源文件只以只读方式打开。

---

## 五、它是怎么做到"不丢图表"的

不用 openpyxl 重新保存工作簿（那样会丢掉全部图表、图片、部分格式），而是：

1. 直接按 zip 解包，在 `xl/worksheets/*.xml` 层面把每个单元格的 `<f>…</f>`（公式）删掉，
   保留紧随其后的 `<v>`（Excel/WPS 上次算出的缓存结果）；
2. 公式结果是**文本**的单元格（`t="str"`）改写成 `inlineStr`，否则去掉公式后这种类型不合法；
3. 删除 `calcChain.xml` 及其登记项、去掉 `fullCalcOnLoad`，避免 Excel 打开时报"需要修复"；
4. **内嵌图片的单元格整格原样保留**：WPS 特有的 `=DISPIMG("ID_…",1)`（常见于"图纸"页）
   不是计算型公式而是图片载体，程序识别后**跳过该格**——公式、样式、图片全部不动，
   日志里会单独报告"保留图片单元格 N 个"；
5. 其余 zip 成员（图表、图片、样式、主题）原样拷贝，最后重打包并自检。

---

## 六、已知限制（务必看）

1. **本工具不做计算**，只搬运公式已有的缓存结果。若源文件从未被
   Excel/WPS 保存过（例如程序直接生成、从未打开），公式没有缓存值，这些格处理后会是空白。
   界面会在开始前提示"其中 N 个公式没有缓存结果"，看到这个提示请先用 Excel/WPS 打开源文件
   保存一次，再运行本工具。
2. 除内嵌图片单元格外，其余单元格公式全部清除；**条件格式里的公式规则会保留**
   （它只影响颜色显示，不影响数值）。
3. 处理后的文件是静态快照：之后修改数据不会再联动重算。
4. 处理前建议先留一份原文件（本工具不会改源文件，但值得养习惯）。
5. 首次运行未签名的 exe，Windows SmartScreen 可能提示"未知发布者"，
   点"更多信息 → 仍要运行"；单文件模式偶尔被杀软误报，可改用 `--onedir`。


---

## 七、CI（GitHub Actions）打包失败排查记录

2026-10 首次云端构建：**第 6 步"打包单文件 exe"失败**（后续步骤被跳过），
并有 **Node.js 20 弃用警告**。已修复，要点如下：

| 现象 | 原因 | 修法 |
|---|---|---|
| PyInstaller 步骤退出码 1 | Windows 运行器的 stdout 是**管道**，默认按 cp1252 编码；PyInstaller 日志里一旦出现中文路径（`--name Excel公式转数值`）就 `UnicodeEncodeError` 直接崩 | 工作流加 `PYTHONUTF8=1`、`PYTHONIOENCODING=utf-8`；内部包名改 ASCII（`XlsxFormulaToValue`），打包后再 `Copy-Item` 成中文名 |
| 同上（另一处风险） | 用 PowerShell 反引号续写多行命令，参数易被截断，pyinstaller 收不到脚本名 | 改为 PowerShell 数组 `$common = @(...)` 单次调用，不用续行符 |
| 图标嵌入失败风险 | PIL 生成的 ico 每个尺寸都是 **PNG 压缩条目**，PyInstaller 在 Windows 嵌入资源时容易报错 | 改用 `make_ico.py` 生成**纯 BMP 条目**的 ico；工作流仍保留"带图标失败→无图标重试"的自动回退 |
| Node.js 20 弃用警告 | `checkout@v4`、`setup-python@v5`、`upload-artifact@v4` 都跑在 Node 20 上 | 升级到 `checkout@v5`、`setup-python@v6`、`upload-artifact@v5`（Node 24） |

现在工作流共 11 步，失败点会明确停在某一步；新增两步质量闸门：

- **引擎自检**：`python selftest.py`，不通过就不打包；
- **冒烟测试**：用打包好的 exe 真跑一次 `--cli 样例.xlsx 输出.xlsx`，再用
  `python selftest.py --verify` 校验产物——不仅证明"能打包"，还证明"打包出来的程序真能干活"。

若仍失败：Actions 页面点开失败的那一步，把日志最后 30 行发出来即可；若怀疑图标，
删掉工作流里 `--icon icon.ico` 那行也会自动走无图标路径。

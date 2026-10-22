# Excel 公式转数值工具

把一份 Excel 工作簿里**所有工作表的公式**一次性转成数值，图表、图片、合并单元格、条件格式、
行高列宽全部原样保留；**原始文件全程只读，不会被修改**。

界面要点（v1.1）：
- **拖拽导入**：把 `.xlsx / .xlsm` 文件直接拖进窗口即可，支持一次拖多个、也支持整文件夹拖入；
  缺少拖拽组件时自动退回「添加文件…」按钮，不会崩。
- **输出文件夹每次打开都要重新选择**（启动时不预填任何路径，没选会被拦下并提示）。
- **历史输出文件夹**会被记住（最近优先、去重、最多 10 个），下拉即可直接选，不用逐层点目录。

---

## 一、目录内容

| 文件 | 作用 |
|---|---|
| `xlsx_deformula.py` | 主程序（图形界面 + 命令行两种用法，纯 Python 标准库，零第三方依赖） |
| `一键打包exe.bat` | 在 Windows 上双击，自动装 PyInstaller 并生成单文件 exe |
| `selftest.py` | 自检（造样例→转换→校验公式/数值/图片单元格/源文件不变，以及拖拽解析、历史记忆），CI 与本地都能跑 |
| `make_ico.py` | 由 `icon.png` 生成纯 BMP 条目的 `icon.ico`（PIL 默认写 PNG 条目，PyInstaller 嵌入时易报错） |
| `icon.ico` / `icon.png` | 程序图标 |
| `.github/workflows/build-exe.yml` | 云端打包**单文件 exe**（一个 exe，启动时自解压） |
| `.github/workflows/build-exe-folder.yml` | 云端打包**文件夹产物**（exe + `_internal` 运行库，数百个文件） |
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
python -m pip install pyinstaller tkinterdnd2
pyinstaller --noconfirm --clean --onefile --windowed ^
  --name "Excel公式转数值" --icon icon.ico ^
  --collect-all tkinterdnd2 ^
  --distpath dist --workpath build --specpath build ^
  xlsx_deformula.py
```

> `--collect-all tkinterdnd2` 必须带上，否则 exe 里没有 tkdnd 组件，**拖拽功能失效**
> （程序仍能运行，只能用按钮选文件）。打完可用 `Excel公式转数值.exe --check-dnd` 验证：
> 退出码 0 = 拖拽可用，3 = 不可用。

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

1. **① 导入文件**（两种方式都行）
   - **拖拽**：从资源管理器把 `.xlsx` / `.xlsm` 拖到窗口上方蓝色区域或文件列表里 ——
     可一次拖多个，也可把**整个文件夹**拖进来（自动收集其中的表格，跳过 `~$` 临时文件和非 Excel 文件）；
     拖入非表格文件会弹窗提示并忽略该条目。
   - **按钮**：`添加文件…` 多选文件；`移除选中` / `清空列表` 调整列表。
2. **② 输出位置**（三种方式任选）
   - 与原文件同目录，自动加 `_无公式`（默认）
   - 输出到指定文件夹：**每次打开程序都要重新选择**；点 `浏览…` 选新目录，
     或从**下拉历史记录**里直接挑一个（选过即记住，最近优先、最多 10 个）。
     注意：下拉只是"候选"，需要你本次点选/选择后才会生效；没选就点开始会提示并中止。
   - 保存为指定文件（仅单个源文件时可用）
3. 点 **开始转换**，下方日志逐文件显示：工作表数、公式单元格数、内嵌图片单元格数、结果路径。
4. `打开输出文件夹` 直接定位结果。

配置文件：`%APPDATA%\XlsxDeformula\config.json`（非 Windows 为 `~/.config/XlsxDeformula/config.json`），
只保存历史输出目录，不含任何路径以外的信息。

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
| **冒烟测试**报"未生成输出文件"（exe 其实正常） | `--windowed` 产物是 **GUI 子系统程序**，PowerShell 直接 `& $exe` **不会等待**，进程还没写完文件就检查了，`$LASTEXITCODE` 也是陈旧值 | 改用 `Start-Process -FilePath $exe -ArgumentList ... -Wait -PassThru` 取 `$p.ExitCode`，并对输出文件做最多 20 秒轮询；失败时临时构建一个 `--console` 版把真实 stderr 抓进摘要 |
| Node.js 20 弃用警告 | `checkout@v4`、`setup-python@v5`、`upload-artifact@v4/v5` 仍跑在 Node 20 上 | 升级到当前 Node 24 版本：`checkout@v7`、`setup-python@v7`、`upload-artifact@v7`（输入参数不变，已核对 `action.yml`） |
| 第二次仍失败：PyInstaller 退出码 1，自检已通过，报错落在日志靠后的 "Building EXE from EXE-00.toc" 阶段 | Windows 运行器的 **Defender 实时防护锁住正在写入的 exe**（`[WinError 5] Access is denied`）；单文件模式要把整个包追加进 exe 尾部，最容易撞上 | ① 构建前把工作目录加入 Defender 排除项；② 打包改**多套单文件配置依次尝试**（onefile → onefile 重试 → 带图标 → `--noupx`），**只用退出码 0 且产物为文件的那一套**；③ 每套日志写进运行页面 Summary 并作为 artifact 留存；④ PyInstaller 固定用已验证版本 6.22.3，失败再退回最新版 |
| 看不到失败原因 | 日志下载接口需要仓库 admin 权限（403） | 每套尝试的最后 80 行直接写进 `$GITHUB_STEP_SUMMARY`，运行页面 Summary 标签页公开可读，不需要下载日志 |

现在工作流共 11 步，失败点会明确停在某一步；新增两步质量闸门：

- **引擎自检**：`python selftest.py`，不通过就不打包；
- **冒烟测试**：用打包好的 exe 真跑一次 `--cli 样例.xlsx 输出.xlsx`，再用
  `python selftest.py --verify` 校验产物——不仅证明"能打包"，还证明"打包出来的程序真能干活"。

若仍失败：Actions 页面点开失败的那一步，把日志最后 30 行发出来即可；若怀疑图标，
删掉工作流里 `--icon icon.ico` 那行也会自动走无图标路径。


---

## 八、云端打包的硬性约定与任务摘要

工作流（`.github/workflows/build-exe.yml`）现在只接受**单文件产物**：

1. 依次尝试 4 套 `--onefile` 配置，逐套记录 `pyinstaller` 的**退出码**；
2. 判定成功的条件同时满足三条：**退出码 0** + `dist/XlsxFormulaToValue.exe` 是**文件** +
   `dist/XlsxFormulaToValue/` **目录不存在**（一旦出现目录即判失败，符合"只输出单个文件"的要求）；
3. 命中后只保留一个文件 `dist/Excel公式转数值.exe`（ASCII 中间产物会被删掉），
   Artifact 里也只有这一个 exe；
4. 全部配置都失败时，用 `::error::` 输出工作流命令错误注解并让步骤失败。

任务摘要按 GitHub 官方 workflow commands 文档，在**运行时**向 `$GITHUB_STEP_SUMMARY`
追加 Markdown 生成（[文档](https://docs.github.com/actions/using-workflows/workflow-commands-for-github-actions#adding-a-job-summary)），内容包括：

- 运行器 / 提交 SHA / Python 版本 / stdout 编码 / 文件系统编码 / C 盘可用空间；
- Defender 排除项是否设置成功；
- **每套配置的退出码、是否有单文件产物、是否误建目录、结论**（Markdown 表格）；
- 失败配置的日志尾部 80 行（折叠在 `<details>` 里）；
- 冒烟测试结果（单文件 exe 实跑 `--cli` 并校验产物）；
- 收尾的运行结果与产物要求。

因此即使失败，也不需要下载日志：打开该次运行的 **Summary** 标签页即可看到全部退出码与报错。


---

## 九、冒烟测试（CI）的实现要点

`--windowed` 打包出来的是 **GUI 子系统程序**，这一点会导致两个坑，工作流里都已处理：

1. **不能用 `& $exe` 调用**：PowerShell 对 GUI 程序不阻塞等待，会立刻返回，
   于是 `$LASTEXITCODE` 是上一命令的陈旧值、输出文件也可能还没落盘，误判成"没生成输出文件"。
   正确写法：

   ```powershell
   $p = Start-Process -FilePath $exe -ArgumentList @('--cli','smoke.xlsx','smoke_out.xlsx') `
                      -WorkingDirectory (Get-Location).Path -Wait -PassThru
   $code = $p.ExitCode
   ```

2. **文件落盘与进程退出有毫秒级错位**：拿到退出码后再对 `smoke_out.xlsx` 轮询最多 20 秒。

冒烟流程：造样例 → 单文件 exe 实跑 `--cli` → 校验退出码 → 校验输出文件 → `selftest.py --verify`
校验产物（公式已转数值 / 图片单元格保留 / 源文件未改动）。任一步失败会自动**临时再构建一个
`--console` 版本**跑同一命令，把 stdout/stderr 写进任务摘要 —— 因为 windowed 程序本身没有控制台输出，
不这么做失败原因是不可见的。诊断版不进产物、只用于定位。


---

## 十、拖拽功能的实现与排查

- 拖拽依赖 **tkinterdnd2**（Tcl 扩展 tkdnd 的 Python 包装），tkinter 本身不支持文件拖放。
  打包时必须 `--collect-all tkinterdnd2`，把各平台 tkdnd 二进制一起塞进 exe。
- 程序有三层兜底，任何一层失败都不会崩：
  1. `import tkinterdnd2` 失败 → 直接用普通 `tk.Tk()`，界面照常，只能按钮选文件；
  2. 拖拽注册失败（`drop_target_register` 抛异常）→ 忽略，按钮照常；
  3. 想彻底关掉拖拽排障：设环境变量 `XLSX_DEFORMULA_NO_DND=1` 再启动。
- 自检命令（windowed exe 也能返回退出码）：
  ```bat
  Excel公式转数值.exe --check-dnd
  ```
- CI 里已加「校验拖拽组件已打进 exe」这一步，退出码不是 0 直接判失败，
  避免发布一个"看起来正常、其实拖不进去"的包。
- 已知限制：拖拽文件在**提权运行的资源管理器**（管理员窗口）里拖进普通权限程序会被 Windows 拦，
  这是系统 UIPI 限制；以相同权限运行即可。


---

## 十一、两套云端打包工作流：单文件 vs 文件夹

仓库里有两个工作流，产物形态不同，按需选一个跑（都在 Actions 页面手动 `Run workflow`）：

| | `build-exe.yml` | `build-exe-folder.yml` |
| --- | --- | --- |
| PyInstaller 模式 | `--onefile` | `--onedir` |
| 产物形态 | **1 个 exe**（约 11–12 MB） | **1 个文件夹**（主程序 + `_internal/`，约 450+ 文件、40–60 MB） |
| Artifact 名称 | `Excel公式转数值-windows` | `Excel公式转数值-windows-文件夹`（另有 `-zip`） |
| 启动速度 | 首次启动要自解压，慢 1–2 秒 | 直接启动，最快 |
| 分发要求 | 单文件，随便拷 | **必须整个文件夹一起拷**，只拿 exe 无法运行 |
| 杀软误报 | 相对容易触发（自解压特征） | 相对少 |

### 文件夹模式工作流的判定与产物

- 只用 `--onedir`，依次尝试 4 套配置（`onedir` → 重试 → 带图标 → `--noupx`）；
- 成功条件三合一：`pyinstaller` 退出码 0 ＋ `dist/<名字>/` 是目录 ＋ 目录里有主程序 ＋
  `_internal/` 内文件数 ≥ 10（证明是完整的多文件目录产物，不是空壳）；
- 产物整理：把 `dist/XlsxFormulaToValue/` 改名为 **`dist/Excel公式转数值/`**，主程序改名为
  `Excel公式转数值.exe`（**`_internal` 目录名不能动**，改了程序就找不到运行库）；
- 校验：目录内文件数 ≥ 10、拖拽组件 `--check-dnd` 必须返回 0、目录产物实跑一次 `--cli` 并校验产物；
- 上传两个 Artifact：`Excel公式转数值-windows-文件夹`（解压即得整个文件夹）和
  `Excel公式转数值-windows-zip`（`Compress-Archive` 打好的 zip，根目录只有一个文件夹，拿去直接分发）；
- 任务摘要同样运行时生成，包含每套配置的退出码、目录内文件数、总大小、冒烟测试结果。

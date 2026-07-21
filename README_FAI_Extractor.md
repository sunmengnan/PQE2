# FAI / IPQC Excel 提取工具

## 功能

- 输入一个 `.xlsx` / `.xlsm` 检验 Excel。
- 自动遍历所有 Sheet，识别包含 `尺寸序号`、`检验日期(yymmdd)`、`检验时间(hh:mm)` 的巡检表。
- 提取每个尺寸/外观行对应日期、时间下的所有工位数据，输出到 `Sample 1`、`Sample 2` ...。
- 自动跳过无实际测量数据的行。
- 自动跳过结果为 `OK` / `NG` / `/` 等判定或占位内容的行。
- 如果原 Excel 的 `G4` 包含 `工序`，提取 `H4` 到 `Sampling process`。
- 如果原 Excel 的 `I4` 包含 `机台`，提取 `J4` 到 `Sampling line#/Machine#`。
- 如果原 Excel 中存在 `炉号` 行，会按每个日期/时间采样组提取对应炉号到 `Furnace No.`。
- 输出列：`Date `、`Sampling process`、`Sampling time`、`Sampling line#/Machine#`、`Furnace No.`、`FAI`、`Sample 1` ...。
- 导出文件会额外生成 `CPK summary` 工作表：按 Source/工序/机台/FAI 汇总样本，计算 CPK；当 CPK < 1.33 时，自动按双边公差方式给出 `Proposed Lower Tol = -Proposed Upper Tol`，并用对称上下限重新计算 CPK。

## 运行环境

按现有项目要求，使用 conda `yolov5` 环境执行。

## 命令行运行

在本目录执行：

```bash
/home/nordbo/anaconda3/envs/yolov5/bin/python fai_excel_extractor.py input.xlsx -o input_extracted.xlsx
```

多个 Excel 合并提取：

```bash
/home/nordbo/anaconda3/envs/yolov5/bin/python fai_excel_extractor.py input1.xlsx input2.xlsx -o merged_extracted.xlsx --include-source-sheet
```

可选参数：

- `--prefix-fai`：把 `4` 输出为 `FAI4`。
- `--include-source-sheet`：额外输出 `Source sheet` 追溯列。
- 多文件导出时，`--include-source-sheet` 会同时输出 `Source file` 和 `Source sheet`。

## 图形界面运行

```bash
conda run -n yolov5 streamlit run fai_excel_extractor_ui.py
```

然后在浏览器页面上传 Excel 或填写本地路径，点击“开始提取”，即可预览并下载结果。

## 桌面程序双击运行

本目录已生成桌面入口：`FAI Excel提取工具.desktop`。

如果系统提示不信任启动器，请右键该文件，选择“允许启动”或“信任此启动器”，之后即可双击打开桌面程序。

也可以双击或运行：

```bash
./open_fai_extractor_desktop.sh
```

桌面程序支持：

- 点击 `Browse` 可一次选择多个 Excel 文件并合并提取。
- 按日期范围筛选，例如 `2026-06-01`、`2026/06/01`、`260601`。
- 按时间范围筛选，例如 `08:00`。
- 按尺寸 / FAI 包含条件筛选，例如 `4`、`5.1`、`FAI4`。
- 提取后预览数据，并可点击“导出当前筛选”单独导出筛选结果。
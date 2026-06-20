# Log to Excel 工具集

工具集版本：v1.2.0

说明：v1.2.0 中 `log_to_excel.py` 和 `log_alarm_pipeline.py` 的日志明细均输出为 CSV；汇总分析默认输出为 Excel。

这组脚本用于处理设备导出的 logfile 原始日志包，并生成便于筛选、统计和汇总的 Excel 表格。

## 文件说明

| 文件 | 作用 |
| --- | --- |
| `log_to_excel.py` | 将设备日志包、`.log` 或 `.log.gz` 转换为结构化 CSV |
| `multi_alarm_excel_summary.py` | 对一个或多个结构化日志 Excel 做告警标识符统计 |
| `log_alarm_pipeline.py` | 一次性完成日志包转 CSV 明细和告警汇总分析 |

## 环境要求

- macOS 本地运行
- Python 3
- `openpyxl`

安装依赖：

```bash
python3 -m pip install openpyxl
```

查看版本：

```bash
python3 log_to_excel.py --version
python3 multi_alarm_excel_summary.py --version
python3 log_alarm_pipeline.py --version
```

## 1. 日志包转 CSV

脚本：

```bash
python3 log_to_excel.py <输入文件>
```

支持输入：

- `.tar.gz` / `.tgz`
- `.log.gz`
- `.log`

示例：

```bash
python3 log_to_excel.py "/path/to/CPLEAF_CN-260616-logfile.tar.gz"
```

默认输出目录固定为脚本所在目录下的 `output/`：

```text
output/CPLEAF_CN-260616-logfile.csv
```

指定输出文件名：

```bash
python3 log_to_excel.py "/path/to/CPLEAF_CN-260616-logfile.tar.gz" -o CPLEAF_CN.csv
```

注意：`-o` 只指定文件名，文件仍然保存到 `output/`。

完成后打开输出文件：

```bash
python3 log_to_excel.py "/path/to/CPLEAF_CN-260616-logfile.tar.gz" --open
```

### 输出文件

生成的 CSV 字段：

```text
日志文件
原始行号
日志序号
日期
时间
设备名
模块名称
等级
助记符
模块名称/等级/助记符
日志内容
```

多行日志会自动合并到上一条日志的 `日志内容` 单元格中。

## 2. 告警标识符汇总

脚本：

```bash
python3 multi_alarm_excel_summary.py <日志Excel1> <日志Excel2> ...
```

示例：

```bash
python3 multi_alarm_excel_summary.py output/CPLEAF_CN-260616-logfile.xlsx output/CPLEAF_EH-260616-logfile.xlsx
```

默认输出目录：

```text
alarm_summary_output/
```

默认输出格式为 `.xlsx`，列宽和格式已按当前示例固定。

如需输出 CSV：

```bash
python3 multi_alarm_excel_summary.py --csv output/CPLEAF_CN-260616-logfile.xlsx output/CPLEAF_EH-260616-logfile.xlsx
```

### 汇总输出字段

```text
告警标识符
总次数
备注
示例来源文件
示例日志序号
示例时间
示例告警内容
各输入文件对应次数
```

默认会将常见无需关注的运维类日志放在末尾，并在 `备注` 列标记为：

```text
无需关注
```

当前内置无需关注标识符包括：

```text
SHELL/6/SHELL_CMD
PING/6/PING_STATISTICS
SSHS/6/SSHS_LOG
SSHS/6/SSHS_DISCONNECT
SSHS/6/SSHS_CONNECT
SHELL/5/SHELL_LOGIN
SHELL/5/SHELL_LOGOUT
NETCONF/6/SSH_XML_LOGOUT
NETCONF/6/SSH_XML_LOGIN
```

## 推荐流程

大批量日志推荐直接使用一键脚本 `log_alarm_pipeline.py`，它会同时生成 CSV 明细和 Excel 汇总：

```bash
python3 log_alarm_pipeline.py "/path/to/CPLEAF_CN-260616-logfile.tar.gz" "/path/to/CPLEAF_EH-260616-logfile.tar.gz"
```

如果需要分步处理，也可以先将设备日志包转换成结构化 CSV：

```bash
python3 log_to_excel.py "/path/to/CPLEAF_CN-260616-logfile.tar.gz"
python3 log_to_excel.py "/path/to/CPLEAF_EH-260616-logfile.tar.gz"
```

`multi_alarm_excel_summary.py` 保留对 Excel 明细表的汇总能力，适用于已有结构化 Excel 明细表：

```bash
python3 multi_alarm_excel_summary.py output/CPLEAF_CN-260616-logfile.xlsx output/CPLEAF_EH-260616-logfile.xlsx
```

在 `alarm_summary_output/` 查看汇总结果。

## 3. 一键生成日志明细和汇总分析

脚本：

```bash
python3 log_alarm_pipeline.py <原始日志包1> <原始日志包2> ...
```

示例：

```bash
python3 log_alarm_pipeline.py "/path/to/CPLEAF_CN-260616-logfile.tar.gz" "/path/to/CPLEAF_EH-260616-logfile.tar.gz"
```

执行后会同时生成：

- 每个输入日志包对应的一份日志明细 CSV
- 一份告警标识符汇总 Excel

默认都保存到 `alarm_summary_output/`。如果指定 `--summary-output-dir`，日志明细 CSV 和汇总文件都会保存到该目录。

如需同时额外生成 CSV 汇总：

```bash
python3 log_alarm_pipeline.py --csv "/path/to/CPLEAF_CN-260616-logfile.tar.gz" "/path/to/CPLEAF_EH-260616-logfile.tar.gz"
```

处理单个输入文件时，可以指定日志明细文件名：

```bash
python3 log_alarm_pipeline.py "/path/to/CPLEAF_CN-260616-logfile.tar.gz" -o CPLEAF_CN.csv
```

指定输出目录：

```bash
python3 log_alarm_pipeline.py --summary-output-dir alarm_summary_output/my_batch "/path/to/CPLEAF_CN-260616-logfile.tar.gz"
```

## v1.0.0 范围

- 支持直接读取设备导出的整个 logfile 压缩包
- 支持 `.log`、`.log.gz`、`.tar.gz/.tgz`
- 支持日志序号提取
- 支持多行日志续行合并
- 支持日志明细 CSV 输出
- 支持多个日志 Excel 的告警标识符汇总
- 支持汇总结果输出为 Excel 或 CSV

## v1.1.0 新增

- 新增 `log_alarm_pipeline.py`
- 支持一次命令完成“原始日志包 -> 日志明细 Excel -> 告警汇总分析”
- 保留 v1.0.0 的两个独立脚本，便于分步骤处理

## v1.2.0 更新

- `log_alarm_pipeline.py` 的日志明细默认输出为 CSV
- 汇总分析默认继续输出为 Excel
- 日志明细 CSV 和汇总分析文件统一保存到 `--summary-output-dir` 指定目录
- 避免百万级日志明细触发 Excel 单 sheet 行数上限和大文件保存性能问题

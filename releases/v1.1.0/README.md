# Log to Excel 工具集

工具集版本：v1.1.0

说明：`log_to_excel.py` 和 `multi_alarm_excel_summary.py` 保留 v1.0.0 的分步能力；v1.1.0 新增 `log_alarm_pipeline.py` 合并流程脚本。

这组脚本用于处理设备导出的 logfile 原始日志包，并生成便于筛选、统计和汇总的 Excel 表格。

## 文件说明

| 文件 | 作用 |
| --- | --- |
| `log_to_excel.py` | 将设备日志包、`.log` 或 `.log.gz` 转换为结构化 Excel |
| `multi_alarm_excel_summary.py` | 对一个或多个结构化日志 Excel 做告警标识符统计 |
| `log_alarm_pipeline.py` | 一次性完成日志包转 Excel 和告警汇总分析 |

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

## 1. 日志包转 Excel

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
output/CPLEAF_CN-260616-logfile.xlsx
```

指定输出文件名：

```bash
python3 log_to_excel.py "/path/to/CPLEAF_CN-260616-logfile.tar.gz" -o CPLEAF_CN.xlsx
```

注意：`-o` 只指定文件名，文件仍然保存到 `output/`。

完成后打开输出文件：

```bash
python3 log_to_excel.py "/path/to/CPLEAF_CN-260616-logfile.tar.gz" --open
```

### 输出表格

生成的 Excel 包含 3 个工作表：

| 工作表 | 内容 |
| --- | --- |
| `logs` | 所有成功解析的日志明细 |
| `summary` | 行数、续行数、模块/等级/助记符统计 |
| `unparsed` | 无法归属到任何日志记录的异常行 |

`logs` 工作表字段：

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

1. 将设备日志包转换成结构化 Excel：

```bash
python3 log_to_excel.py "/path/to/CPLEAF_CN-260616-logfile.tar.gz"
python3 log_to_excel.py "/path/to/CPLEAF_EH-260616-logfile.tar.gz"
```

2. 对多个设备的输出 Excel 做告警汇总：

```bash
python3 multi_alarm_excel_summary.py output/CPLEAF_CN-260616-logfile.xlsx output/CPLEAF_EH-260616-logfile.xlsx
```

3. 在 `alarm_summary_output/` 查看汇总结果。

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

- 每个输入日志包对应的一份日志明细 Excel，保存到 `output/`
- 一份告警标识符汇总 Excel，保存到 `alarm_summary_output/`

如需同时额外生成 CSV 汇总：

```bash
python3 log_alarm_pipeline.py --csv "/path/to/CPLEAF_CN-260616-logfile.tar.gz" "/path/to/CPLEAF_EH-260616-logfile.tar.gz"
```

处理单个输入文件时，可以指定日志明细文件名：

```bash
python3 log_alarm_pipeline.py "/path/to/CPLEAF_CN-260616-logfile.tar.gz" -o CPLEAF_CN.xlsx
```

指定汇总输出目录：

```bash
python3 log_alarm_pipeline.py --summary-output-dir alarm_summary_output/my_batch "/path/to/CPLEAF_CN-260616-logfile.tar.gz"
```

## v1.0.0 范围

- 支持直接读取设备导出的整个 logfile 压缩包
- 支持 `.log`、`.log.gz`、`.tar.gz/.tgz`
- 支持日志序号提取
- 支持多行日志续行合并
- 支持日志明细 Excel 输出
- 支持多个日志 Excel 的告警标识符汇总
- 支持汇总结果输出为 Excel 或 CSV

## v1.1.0 新增

- 新增 `log_alarm_pipeline.py`
- 支持一次命令完成“原始日志包 -> 日志明细 Excel -> 告警汇总分析”
- 保留 v1.0.0 的两个独立脚本，便于分步骤处理

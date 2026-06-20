#!/usr/bin/env python3
"""
Convert raw device logfile archives to CSV details and generate alarm summaries in one run.
"""

from __future__ import annotations

import argparse
import csv
import sys
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from typing import Optional

import log_to_excel
import multi_alarm_excel_summary as alarm_summary

__version__ = "1.2.0"


def log(message: str) -> None:
    print(message, flush=True)


def output_base_name() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def detail_output_path(input_path: Path, output_dir: Path, output_name: Optional[str] = None) -> Path:
    if output_name:
        output_path = output_dir / Path(output_name).name
        return output_path.with_suffix(".csv")

    name = input_path.name
    for suffix in (".tar.gz", ".tgz", ".log.gz", ".log"):
        if name.endswith(suffix):
            name = name[: -len(suffix)]
            break
    return output_dir / f"{name}.csv"


def record_to_csv_row(record: log_to_excel.LogRecord) -> list[object]:
    return [
        record.source_log_file,
        record.line_number,
        record.sequence,
        record.date.strftime("%Y-%m-%d"),
        record.time.strftime("%H:%M:%S.%f")[:-3],
        record.device,
        record.module,
        record.level,
        record.mnemonic,
        record.key,
        record.content,
    ]


def add_record_to_aggregate(
    record: log_to_excel.LogRecord,
    detail_file_name: str,
    aggregate: dict[str, dict],
) -> None:
    summary_record = aggregate[record.key]
    summary_record["total"] += 1
    summary_record["files"][detail_file_name] += 1

    if not summary_record["example_detail"]:
        summary_record["example_detail"] = record.content
        summary_record["example_file"] = detail_file_name
        summary_record["example_sequence"] = record.sequence
        summary_record["example_datetime"] = f"{record.date:%Y-%m-%d} {record.time:%H:%M:%S}"


def parse_logfile_to_detail_csv(
    input_path: Path,
    output_path: Path,
    aggregate: dict[str, dict],
) -> dict[str, int | str]:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    total_lines = 0
    log_records = 0
    continuation_lines = 0
    unparsed_lines = 0
    current_record: Optional[log_to_excel.LogRecord] = None

    with output_path.open("w", newline="", encoding="utf-8-sig") as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(log_to_excel.LOG_HEADERS)

        def flush_current_record() -> None:
            nonlocal current_record, log_records
            if current_record is None:
                return

            writer.writerow(
                [
                    log_to_excel.remove_illegal_characters(value)
                    for value in record_to_csv_row(current_record)
                ]
            )
            add_record_to_aggregate(current_record, output_path.name, aggregate)
            log_records += 1
            current_record = None

        for source_log_file, line_number, line in log_to_excel.iter_input_lines(input_path):
            total_lines += 1
            if total_lines % 100000 == 0:
                log(f"已读取 {total_lines:,} 行，已解析 {log_records:,} 条日志")

            try:
                next_record = log_to_excel.parse_log_line(line, source_log_file, line_number)
            except ValueError:
                if current_record is not None:
                    current_record.content = f"{current_record.content}\n{line}"
                    continuation_lines += 1
                else:
                    unparsed_lines += 1
                continue

            flush_current_record()
            current_record = next_record

        flush_current_record()

    return {
        "input": str(input_path),
        "output": str(output_path),
        "total_lines": total_lines,
        "log_records": log_records,
        "continuation_lines": continuation_lines,
        "unparsed_lines": unparsed_lines,
    }


def write_summary_outputs(
    aggregate: dict[str, dict],
    detail_files: list[Path],
    file_total_rows: dict[str, int],
    output_dir: Path,
    include_csv: bool,
) -> tuple[Path, Optional[Path]]:
    output_dir.mkdir(parents=True, exist_ok=True)
    base_name = output_base_name()

    summary_xlsx = output_dir / f"{base_name}_告警标识符汇总_含示例.xlsx"
    alarm_summary.write_summary(summary_xlsx, aggregate, detail_files, file_total_rows)

    summary_csv = None
    if include_csv:
        summary_csv = output_dir / f"{base_name}_告警标识符汇总_含示例.csv"
        alarm_summary.write_summary_csv(summary_csv, aggregate, detail_files, file_total_rows)

    return summary_xlsx, summary_csv


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="一次性完成原始日志包转 CSV 明细，并生成告警标识符汇总分析。"
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    parser.add_argument("input_files", nargs="+", help="输入文件，支持 .tar.gz/.tgz、.log.gz 或 .log")
    parser.add_argument(
        "-o",
        "--output",
        help="仅处理单个输入文件时可指定日志明细 CSV 文件名，保存到 --summary-output-dir 目录",
    )
    parser.add_argument(
        "--summary-output-dir",
        default="alarm_summary_output",
        help="日志明细和汇总分析输出目录，默认 alarm_summary_output",
    )
    parser.add_argument(
        "--csv",
        action="store_true",
        help="除默认汇总 Excel 外，额外生成一份汇总 CSV",
    )
    parser.add_argument("--open", action="store_true", help="完成后在 macOS 中打开汇总 Excel")
    return parser


def main(argv: Optional[list[str]] = None) -> int:
    parser = build_arg_parser()
    args = parser.parse_args(argv)

    input_paths = [Path(name).expanduser().resolve() for name in args.input_files]
    missing = [str(path) for path in input_paths if not path.exists()]
    if missing:
        print("以下文件不存在：", file=sys.stderr)
        for path in missing:
            print(f"  {path}", file=sys.stderr)
        return 1

    if args.output and len(input_paths) != 1:
        parser.error("-o/--output 只能在处理单个输入文件时使用")

    output_dir = Path(args.summary_output_dir).expanduser().resolve()

    aggregate = defaultdict(
        lambda: {
            "total": 0,
            "files": Counter(),
            "example_detail": "",
            "example_file": "",
            "example_sequence": "",
            "example_datetime": "",
        }
    )
    detail_files: list[Path] = []
    file_total_rows: dict[str, int] = {}

    for index, input_path in enumerate(input_paths, start=1):
        log(f"[{index}/{len(input_paths)}] 生成日志明细：{input_path}")
        output_path = detail_output_path(input_path, output_dir, args.output)
        result = parse_logfile_to_detail_csv(input_path, output_path, aggregate)
        detail_files.append(output_path)
        log(
            "    明细完成："
            f"日志记录 {result['log_records']:,}，"
            f"续行 {result['continuation_lines']:,}，"
            f"无法归属 {result['unparsed_lines']:,}"
        )

        file_total_rows[output_path.name] = int(result["log_records"])
        log(f"    告警行数：{int(result['log_records']):,}")

    summary_xlsx, summary_csv = write_summary_outputs(
        aggregate=aggregate,
        detail_files=detail_files,
        file_total_rows=file_total_rows,
        output_dir=output_dir,
        include_csv=args.csv,
    )

    ignored_total = sum(
        data["total"] for alert_id, data in aggregate.items() if alert_id in alarm_summary.IGNORED_ALERT_IDS
    )
    log("处理完成")
    log(f"输入文件数：{len(input_paths)}")
    log(f"总告警行数：{sum(file_total_rows.values()):,}")
    log(f"告警标识符种类数：{len(aggregate)}")
    log(f"无需关注告警总行数：{ignored_total:,}")
    log(f"日志明细文件：{', '.join(str(path) for path in detail_files)}")
    log(f"汇总 Excel：{summary_xlsx.resolve()}")
    if summary_csv:
        log(f"汇总 CSV：{summary_csv.resolve()}")

    if args.open:
        import subprocess

        subprocess.run(["open", str(summary_xlsx)], check=False)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())

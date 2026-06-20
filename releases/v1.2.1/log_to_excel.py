#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import gzip
import io
import re
import subprocess
import tarfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterator, Optional, TextIO

__version__ = "1.2.1"

LOG_PATTERN = re.compile(
    r"^%@(?P<sequence>\d+)%"
    r"(?P<timestamp>[A-Za-z]{3}\s+\d+\s+\d{2}:\d{2}:\d{2}:\d{3}\s+\d{4})"
    r"\s+(?P<device>\S+)"
    r"\s+(?P<key>[^\s:]+):\s?"
    r"(?P<content>.*)$"
)

MONTH_MAP = {
    "Jan": 1,
    "Feb": 2,
    "Mar": 3,
    "Apr": 4,
    "May": 5,
    "Jun": 6,
    "Jul": 7,
    "Aug": 8,
    "Sep": 9,
    "Oct": 10,
    "Nov": 11,
    "Dec": 12,
}

LOG_HEADERS = [
    "日志文件",
    "原始行号",
    "日志序号",
    "日期",
    "时间",
    "设备名",
    "模块名称",
    "等级",
    "助记符",
    "模块名称/等级/助记符",
    "日志内容",
]


@dataclass
class LogRecord:
    source_log_file: str
    line_number: int
    sequence: str
    date: datetime
    time: object
    device: str
    module: str
    level: str
    mnemonic: str
    key: str
    content: str


def log(message: str) -> None:
    print(message, flush=True)


def remove_illegal_characters(value):
    if isinstance(value, str):
        return re.sub(r"[\x00-\x09\x0B\x0C\x0E-\x1F\x7F]", "", value)
    return value


def parse_timestamp(timestamp: str) -> tuple[datetime, object]:
    parts = timestamp.split()
    if len(parts) != 4:
        raise ValueError(f"invalid timestamp: {timestamp}")

    month_name, day, time_text, year = parts
    if month_name not in MONTH_MAP:
        raise ValueError(f"invalid month: {month_name}")

    hour, minute, second, millisecond = time_text.split(":")
    parsed_date = datetime(int(year), MONTH_MAP[month_name], int(day))
    parsed_time = datetime.strptime(
        f"{hour}:{minute}:{second}.{millisecond}", "%H:%M:%S.%f"
    ).time()
    return parsed_date, parsed_time


def parse_log_line(line: str, source_log_file: str, line_number: int) -> LogRecord:
    match = LOG_PATTERN.match(line)
    if not match:
        raise ValueError("not a logfile record")

    key = match.group("key")
    key_parts = key.split("/")
    if len(key_parts) != 3:
        raise ValueError(f"invalid module/level/mnemonic: {key}")

    parsed_date, parsed_time = parse_timestamp(match.group("timestamp"))
    return LogRecord(
        source_log_file=source_log_file,
        line_number=line_number,
        sequence=match.group("sequence"),
        date=parsed_date,
        time=parsed_time,
        device=match.group("device"),
        module=key_parts[0],
        level=key_parts[1],
        mnemonic=key_parts[2],
        key=key,
        content=match.group("content"),
    )


def record_to_row(record: LogRecord) -> list[object]:
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


def normalize_log_file_name(source_name: str) -> str:
    name = Path(source_name).name
    for suffix in (".log.gz", ".log"):
        if name.endswith(suffix):
            return name[: -len(suffix)]
    return Path(name).stem


def iter_text_lines(file: TextIO, source_log_file: str) -> Iterator[tuple[str, int, str]]:
    for line_number, line in enumerate(file, start=1):
        yield source_log_file, line_number, line.rstrip("\r\n")


def iter_plain_log(path: Path) -> Iterator[tuple[str, int, str]]:
    with path.open("r", encoding="utf-8", errors="replace", newline="") as file:
        yield from iter_text_lines(file, normalize_log_file_name(path.name))


def iter_gzip_log(path: Path) -> Iterator[tuple[str, int, str]]:
    with gzip.open(path, "rt", encoding="utf-8", errors="replace", newline="") as file:
        yield from iter_text_lines(file, normalize_log_file_name(path.name))


def iter_tar_logs(path: Path) -> Iterator[tuple[str, int, str]]:
    with tarfile.open(path, "r:gz") as tar:
        for member in tar.getmembers():
            if not member.isfile():
                continue
            if not (member.name.endswith(".log") or member.name.endswith(".log.gz")):
                continue

            extracted = tar.extractfile(member)
            if extracted is None:
                continue

            binary_file = gzip.GzipFile(fileobj=extracted) if member.name.endswith(".gz") else extracted
            with io.TextIOWrapper(binary_file, encoding="utf-8", errors="replace", newline="") as text_file:
                yield from iter_text_lines(text_file, normalize_log_file_name(member.name))


def iter_input_lines(input_path: Path) -> Iterator[tuple[str, int, str]]:
    if input_path.name.endswith((".tar.gz", ".tgz")):
        yield from iter_tar_logs(input_path)
    elif input_path.name.endswith(".log.gz"):
        yield from iter_gzip_log(input_path)
    elif input_path.name.endswith(".log"):
        yield from iter_plain_log(input_path)
    else:
        raise ValueError("输入文件仅支持 .tar.gz/.tgz、.log.gz 或 .log")


def output_directory() -> Path:
    return Path(__file__).resolve().parent / "output"


def default_output_path(input_path: Path) -> Path:
    name = input_path.name
    for suffix in (".tar.gz", ".tgz", ".log.gz", ".log"):
        if name.endswith(suffix):
            name = name[: -len(suffix)]
            break
    return output_directory() / f"{name}.csv"


def requested_output_path(output_name: str) -> Path:
    return (output_directory() / Path(output_name).name).with_suffix(".csv")


def parse_logfile_to_csv(input_path: Path, output_path: Path) -> dict[str, int | str]:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    total_lines = 0
    log_records = 0
    continuation_lines = 0
    unparsed_lines = 0
    current_record: Optional[LogRecord] = None

    with output_path.open("w", newline="", encoding="utf-8-sig") as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(LOG_HEADERS)

        def flush_current_record() -> None:
            nonlocal current_record, log_records
            if current_record is None:
                return

            writer.writerow(
                [remove_illegal_characters(value) for value in record_to_row(current_record)]
            )
            log_records += 1
            current_record = None

        for source_log_file, line_number, line in iter_input_lines(input_path):
            total_lines += 1
            if total_lines % 100000 == 0:
                log(f"已读取 {total_lines:,} 行，已解析 {log_records:,} 条日志")

            try:
                next_record = parse_log_line(line, source_log_file, line_number)
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


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="将设备 logfile 原始文件或日志压缩包转换为 CSV 明细。"
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    parser.add_argument("input", help="输入文件，支持 .tar.gz/.tgz、.log.gz 或 .log")
    parser.add_argument("-o", "--output", help="输出 CSV 文件名，固定保存到脚本目录的 output 文件夹")
    parser.add_argument("--open", action="store_true", help="完成后在 macOS 中打开输出文件")
    return parser


def main(argv: Optional[list[str]] = None) -> int:
    parser = build_arg_parser()
    args = parser.parse_args(argv)

    input_path = Path(args.input).expanduser().resolve()
    if not input_path.exists():
        parser.error(f"输入文件不存在：{input_path}")

    output_path = requested_output_path(args.output) if args.output else default_output_path(input_path)

    log(f"开始处理：{input_path}")
    result = parse_logfile_to_csv(input_path, output_path)
    log(
        "处理完成："
        f"物理行 {result['total_lines']:,}，"
        f"日志记录 {result['log_records']:,}，"
        f"续行 {result['continuation_lines']:,}，"
        f"无法归属 {result['unparsed_lines']:,}"
    )
    log(f"已生成：{output_path}")

    if args.open:
        subprocess.run(["open", str(output_path)], check=False)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())

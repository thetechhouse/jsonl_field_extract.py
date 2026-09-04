#!/usr/bin/env python3
"""Extract field-matching records from JSONL, NDJSON, and JSON files.

JSONL uses ripgrep as an optional high-speed prefilter. Conventional JSON is
parsed record-by-record with ijson, so top-level arrays need not fit in memory.
By default, every discovered field path containing ``email`` (case-insensitive)
is searched. All matching records are combined into one JSONL output file.
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, BinaryIO, Iterable, Iterator, Sequence

try:
    import orjson  # type: ignore
except ImportError:
    orjson = None

try:
    import ijson  # type: ignore
except ImportError:
    ijson = None


DEFAULT_GLOBS = ("*.jsonl", "*.ndjson", "*.json")
PathTokens = tuple[str, ...]


@dataclass
class FieldInfo:
    occurrences: int = 0
    examples: list[str] = field(default_factory=list)


@dataclass
class Stats:
    jsonl_candidates: int = 0
    json_records: int = 0
    matches: int = 0
    invalid_json: int = 0
    files_scanned: int = 0
    bytes_scanned: int = 0

    def add(self, other: "Stats") -> None:
        self.jsonl_candidates += other.jsonl_candidates
        self.json_records += other.json_records
        self.matches += other.matches
        self.invalid_json += other.invalid_json
        self.files_scanned += other.files_scanned
        self.bytes_scanned += other.bytes_scanned


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Discover fields in JSONL/NDJSON/JSON files and extract records "
            "containing partial values in selected fields."
        )
    )
    parser.add_argument("source", nargs="?", help="Source file or directory")
    parser.add_argument("-o", "--output", help="Destination JSONL file")
    parser.add_argument(
        "-g",
        "--glob",
        action="append",
        dest="globs",
        help=(
            "Filename glob; repeatable. Defaults: *.jsonl, *.ndjson, *.json. "
            "A .json file is treated as conventional JSON; others as JSONL."
        ),
    )
    parser.add_argument(
        "-f",
        "--field",
        action="append",
        dest="fields",
        help=(
            "Explicit field path, e.g. sender.email or recipients[].email; "
            "repeatable. Overrides automatic email-field selection."
        ),
    )
    parser.add_argument(
        "--field-name-contains",
        default="email",
        help=(
            "Automatically select discovered field paths containing this text, "
            "case-insensitively (default: email)"
        ),
    )
    parser.add_argument(
        "-t", "--term", action="append", dest="terms", help="Partial value; repeatable"
    )
    parser.add_argument(
        "--term-mode",
        choices=("any", "all"),
        default="any",
        help="Require any or all terms in the selected fields (default: any)",
    )
    parser.add_argument(
        "--case-sensitive", action="store_true", help="Use case-sensitive matching"
    )
    parser.add_argument(
        "--engine",
        choices=("auto", "rg", "python"),
        default="auto",
        help=(
            "JSONL engine (default: auto, preferring ripgrep). Conventional "
            "JSON always uses streaming parsing."
        ),
    )
    parser.add_argument(
        "--json-record-path",
        help=(
            "Dot-separated path to a record array inside top-level JSON objects, "
            "e.g. data.records. Top-level arrays need no value."
        ),
    )
    parser.add_argument(
        "--sample-per-file",
        type=int,
        default=50,
        help="Records sampled from each file during discovery (default: 50)",
    )
    parser.add_argument(
        "--max-fields",
        type=int,
        default=10_000,
        help="Safety cap for discovered field paths (default: 10000)",
    )
    parser.add_argument(
        "--overwrite", action="store_true", help="Replace an existing output file"
    )
    parser.add_argument(
        "--no-discovery",
        action="store_true",
        help="Skip field discovery (requires --field)",
    )
    return parser.parse_args()


def json_loads(raw: bytes) -> Any:
    return orjson.loads(raw) if orjson is not None else json.loads(raw)


def json_dumps_line(record: Any) -> bytes:
    if orjson is not None:
        return orjson.dumps(record, option=orjson.OPT_APPEND_NEWLINE)
    text = json.dumps(record, ensure_ascii=False, separators=(",", ":"))
    return (text + "\n").encode("utf-8")


def human_size(size: int) -> str:
    units = ("B", "KiB", "MiB", "GiB", "TiB", "PiB")
    value = float(size)
    for unit in units:
        if value < 1024 or unit == units[-1]:
            return f"{value:.1f} {unit}"
        value /= 1024
    return f"{value:.1f} PiB"


def shorten(value: Any, limit: int = 70) -> str:
    text = value if isinstance(value, str) else json.dumps(
        value, ensure_ascii=False, separators=(",", ":")
    )
    text = text.replace("\n", "\\n").replace("\r", "\\r")
    return text if len(text) <= limit else text[: limit - 1] + "…"


def display_path(tokens: PathTokens) -> str:
    return ".".join(tokens)


def parse_field_path(path: str) -> PathTokens:
    tokens = tuple(part.strip() for part in path.split(".") if part.strip())
    if not tokens:
        raise ValueError(f"Invalid field path: {path!r}")
    return tokens


def collect_source_files(
    source: Path, patterns: Sequence[str], excluded: Sequence[Path] = ()
) -> list[Path]:
    excluded_resolved = set()
    for path in excluded:
        try:
            excluded_resolved.add(path.resolve())
        except OSError:
            pass

    if source.is_file():
        candidates: Iterable[Path] = (source,)
    else:
        candidates = (
            path
            for path in source.rglob("*")
            if path.is_file() and any(path.match(pattern) for pattern in patterns)
        )

    files: list[Path] = []
    seen: set[Path] = set()
    for path in candidates:
        try:
            resolved = path.resolve()
        except OSError:
            resolved = path.absolute()
        if resolved in excluded_resolved or resolved in seen:
            continue
        seen.add(resolved)
        files.append(path)
    return files


def is_conventional_json(path: Path) -> bool:
    return path.suffix.casefold() == ".json"


def source_summary(files: Sequence[Path]) -> tuple[int, int]:
    total = 0
    for path in files:
        try:
            total += path.stat().st_size
        except OSError:
            pass
    return len(files), total


def first_non_whitespace(path: Path) -> bytes:
    with path.open("rb") as handle:
        while chunk := handle.read(64 * 1024):
            stripped = chunk.lstrip()
            if stripped:
                return stripped[:1]
    return b""


def ijson_prefix(path: Path, record_path: str | None) -> str:
    if record_path:
        cleaned = record_path.strip().removeprefix("$").strip(".").replace("[]", "")
        return "item" if not cleaned else f"{cleaned}.item"

    opening = first_non_whitespace(path)
    if opening == b"[":
        return "item"
    if opening == b"{":
        if path.stat().st_size > 256 * 1024 * 1024:
            raise RuntimeError(
                f"{path} is a large top-level JSON object. Specify the array "
                "containing its records with --json-record-path, for example "
                "--json-record-path records."
            )
        return ""
    raise ValueError(f"{path} is empty or does not begin with a JSON array/object")


def iter_json_array_stdlib(path: Path, chunk_size: int = 4 * 1024 * 1024) -> Iterator[Any]:
    """Incrementally decode a top-level JSON array using only the standard library."""
    decoder = json.JSONDecoder()
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        buffer = ""
        position = 0
        eof = False

        def refill(compact: bool = True) -> None:
            nonlocal buffer, position, eof
            if compact and position:
                buffer = buffer[position:]
                position = 0
            chunk = handle.read(chunk_size)
            if chunk:
                buffer += chunk
            else:
                eof = True

        def skip_space() -> None:
            nonlocal position
            while True:
                while position < len(buffer) and buffer[position].isspace():
                    position += 1
                if position < len(buffer) or eof:
                    return
                refill()

        refill()
        skip_space()
        if position >= len(buffer) or buffer[position] != "[":
            raise ValueError(f"{path} is not a top-level JSON array")
        position += 1

        while True:
            skip_space()
            if position >= len(buffer):
                raise ValueError(f"Unexpected end of JSON array in {path}")
            if buffer[position] == "]":
                return

            value_start = position
            while True:
                try:
                    record, end = decoder.raw_decode(buffer, position)
                    break
                except json.JSONDecodeError:
                    if eof:
                        raise
                    if value_start:
                        buffer = buffer[value_start:]
                        position = 0
                        value_start = 0
                    refill(compact=False)

            yield record
            position = end
            skip_space()
            if position >= len(buffer):
                raise ValueError(f"Unexpected end of JSON array in {path}")
            delimiter = buffer[position]
            position += 1
            if delimiter == ",":
                continue
            if delimiter == "]":
                return
            raise ValueError(
                f"Expected ',' or ']' after a JSON record in {path}; got {delimiter!r}"
            )


def iter_json_records(path: Path, record_path: str | None) -> Iterator[Any]:
    opening = first_non_whitespace(path)
    if ijson is not None:
        prefix = ijson_prefix(path, record_path)
        with path.open("rb") as handle:
            yield from ijson.items(handle, prefix, use_float=True)
        return

    if record_path:
        raise RuntimeError(
            "--json-record-path requires ijson. Install it with: "
            "python3 -m pip install ijson"
        )
    if opening == b"[":
        yield from iter_json_array_stdlib(path)
        return
    if opening == b"{":
        ijson_prefix(path, None)  # Enforce the large-object memory safeguard.
        with path.open("r", encoding="utf-8-sig") as handle:
            yield json.load(handle)
        return
    raise ValueError(f"{path} is empty or is not a JSON array/object")


def walk_scalars(
    value: Any,
    path: PathTokens,
    fields: dict[PathTokens, FieldInfo],
    max_fields: int,
) -> None:
    if isinstance(value, dict):
        for key, child in value.items():
            child_path = path + (str(key),)
            if len(fields) >= max_fields and child_path not in fields:
                continue
            walk_scalars(child, child_path, fields, max_fields)
    elif isinstance(value, list):
        if not path:
            return
        array_path = path[:-1] + (path[-1] + "[]",)
        for child in value:
            walk_scalars(child, array_path, fields, max_fields)
    elif path:
        info = fields.setdefault(path, FieldInfo())
        info.occurrences += 1
        if len(info.examples) < 2:
            example = shorten(value)
            if example not in info.examples:
                info.examples.append(example)


def sample_jsonl_file(
    path: Path,
    sample_per_file: int,
    fields: dict[PathTokens, FieldInfo],
    max_fields: int,
) -> tuple[int, int]:
    valid = invalid = 0
    with path.open("rb") as handle:
        for raw in handle:
            if not raw.strip():
                continue
            try:
                record = json_loads(raw)
            except (ValueError, json.JSONDecodeError):
                invalid += 1
                continue
            valid += 1
            walk_scalars(record, (), fields, max_fields)
            if valid >= sample_per_file:
                break
    return valid, invalid


def sample_json_file(
    path: Path,
    record_path: str | None,
    sample_per_file: int,
    fields: dict[PathTokens, FieldInfo],
    max_fields: int,
) -> tuple[int, int]:
    valid = 0
    try:
        for record in iter_json_records(path, record_path):
            valid += 1
            walk_scalars(record, (), fields, max_fields)
            if valid >= sample_per_file:
                break
    except Exception as exc:
        print(f"Warning: cannot sample structured JSON {path}: {exc}", file=sys.stderr)
        return valid, 1
    return valid, 0


def discover_fields(
    files: Sequence[Path],
    record_path: str | None,
    sample_per_file: int,
    max_fields: int,
) -> tuple[dict[PathTokens, FieldInfo], int, int]:
    fields: dict[PathTokens, FieldInfo] = {}
    valid = invalid = 0
    for path in files:
        try:
            if is_conventional_json(path):
                good, bad = sample_json_file(
                    path, record_path, sample_per_file, fields, max_fields
                )
            else:
                good, bad = sample_jsonl_file(path, sample_per_file, fields, max_fields)
        except OSError as exc:
            print(f"Warning: cannot sample {path}: {exc}", file=sys.stderr)
            good, bad = 0, 1
        valid += good
        invalid += bad
    return fields, valid, invalid


def select_fields_by_name(
    fields: dict[PathTokens, FieldInfo], name_fragment: str
) -> list[PathTokens]:
    needle = name_fragment.casefold()
    return sorted(
        (path for path in fields if needle in display_path(path).casefold()),
        key=lambda item: display_path(item).casefold(),
    )


def values_at_path(record: Any, path: PathTokens) -> Iterator[Any]:
    current = [record]
    for token in path:
        is_array = token.endswith("[]")
        key = token[:-2] if is_array else token
        following: list[Any] = []
        for node in current:
            if not isinstance(node, dict) or key not in node:
                continue
            value = node[key]
            if is_array:
                if isinstance(value, list):
                    following.extend(value)
            else:
                following.append(value)
        current = following
        if not current:
            break
    yield from current


def searchable_text(value: Any) -> str:
    if isinstance(value, str):
        return value
    if value is None:
        return "null"
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (int, float)):
        return str(value)
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


def record_matches(
    record: Any,
    fields: Sequence[PathTokens],
    terms: Sequence[str],
    case_sensitive: bool,
    term_mode: str,
) -> bool:
    values = [searchable_text(value) for path in fields for value in values_at_path(record, path)]
    if not values:
        return False
    if not case_sensitive:
        values = [value.casefold() for value in values]
        checked_terms = [term.casefold() for term in terms]
    else:
        checked_terms = list(terms)
    found = (any(term in value for value in values) for term in checked_terms)
    return all(found) if term_mode == "all" else any(found)


def verify_jsonl_candidate(
    raw: bytes,
    output: BinaryIO,
    fields: Sequence[PathTokens],
    terms: Sequence[str],
    case_sensitive: bool,
    term_mode: str,
    stats: Stats,
) -> None:
    stats.jsonl_candidates += 1
    try:
        record = json_loads(raw)
    except (ValueError, json.JSONDecodeError):
        stats.invalid_json += 1
        return
    if record_matches(record, fields, terms, case_sensitive, term_mode):
        output.write(raw)
        if not raw.endswith(b"\n"):
            output.write(b"\n")
        stats.matches += 1


def path_batches(paths: Sequence[Path], max_items: int = 256) -> Iterator[list[Path]]:
    for start in range(0, len(paths), max_items):
        yield list(paths[start : start + max_items])


def rg_command(files: Sequence[Path], terms: Sequence[str], case_sensitive: bool) -> list[str]:
    command = [
        "rg",
        "--fixed-strings",
        "--no-filename",
        "--no-heading",
        "--no-line-number",
        "--color=never",
        "--text",
        "--no-ignore",
        "--hidden",
    ]
    if not case_sensitive:
        command.append("--ignore-case")
    for term in terms:
        command.extend(("--regexp", term))
    command.append("--")
    command.extend(str(path) for path in files)
    return command


def scan_jsonl_with_rg(
    files: Sequence[Path],
    output: BinaryIO,
    fields: Sequence[PathTokens],
    terms: Sequence[str],
    case_sensitive: bool,
    term_mode: str,
) -> Stats:
    stats = Stats(files_scanned=len(files))
    for batch in path_batches(files):
        process = subprocess.Popen(rg_command(batch, terms, case_sensitive), stdout=subprocess.PIPE)
        assert process.stdout is not None
        try:
            for raw in process.stdout:
                verify_jsonl_candidate(
                    raw, output, fields, terms, case_sensitive, term_mode, stats
                )
        except KeyboardInterrupt:
            process.terminate()
            process.wait()
            raise
        return_code = process.wait()
        if return_code not in (0, 1):
            raise RuntimeError(f"ripgrep exited with status {return_code}")
    return stats


def scan_jsonl_with_python(
    files: Sequence[Path],
    output: BinaryIO,
    fields: Sequence[PathTokens],
    terms: Sequence[str],
    case_sensitive: bool,
    term_mode: str,
) -> Stats:
    stats = Stats()
    encoded_terms = [term.encode("utf-8") for term in terms]
    if not case_sensitive:
        encoded_terms = [term.lower() for term in encoded_terms]
    for path in files:
        try:
            with path.open("rb") as handle:
                for raw in handle:
                    stats.bytes_scanned += len(raw)
                    candidate = raw if case_sensitive else raw.lower()
                    if any(term in candidate for term in encoded_terms):
                        verify_jsonl_candidate(
                            raw, output, fields, terms, case_sensitive, term_mode, stats
                        )
            stats.files_scanned += 1
        except OSError as exc:
            print(f"Warning: cannot read {path}: {exc}", file=sys.stderr)
    return stats


def scan_conventional_json(
    files: Sequence[Path],
    record_path: str | None,
    output: BinaryIO,
    fields: Sequence[PathTokens],
    terms: Sequence[str],
    case_sensitive: bool,
    term_mode: str,
) -> Stats:
    stats = Stats()
    for path in files:
        try:
            for record in iter_json_records(path, record_path):
                stats.json_records += 1
                if record_matches(record, fields, terms, case_sensitive, term_mode):
                    output.write(json_dumps_line(record))
                    stats.matches += 1
            stats.files_scanned += 1
            try:
                stats.bytes_scanned += path.stat().st_size
            except OSError:
                pass
        except Exception as exc:
            raise RuntimeError(f"Could not fully process structured JSON {path}: {exc}") from exc
    return stats


def prompt_nonempty(prompt: str) -> str:
    while True:
        value = input(prompt).strip()
        if value:
            return value


def main() -> int:
    args = parse_args()
    source = Path(
        args.source or prompt_nonempty("Source JSONL/JSON file or directory: ")
    ).expanduser()
    if not source.exists():
        print(f"Error: source does not exist: {source}", file=sys.stderr)
        return 2

    output = Path(args.output or prompt_nonempty("Output JSONL file: ")).expanduser()
    if output.exists() and not args.overwrite:
        print(f"Error: output exists; use --overwrite to replace it: {output}", file=sys.stderr)
        return 2
    try:
        if output.resolve() == source.resolve():
            print("Error: source and output cannot be the same file.", file=sys.stderr)
            return 2
    except OSError:
        pass

    patterns = tuple(args.globs or DEFAULT_GLOBS)
    output.parent.mkdir(parents=True, exist_ok=True)
    files = collect_source_files(source, patterns, (output,))
    file_count, total_size = source_summary(files)
    if file_count == 0:
        print(f"Error: no files matching {', '.join(patterns)} were found.", file=sys.stderr)
        return 2

    json_files = [path for path in files if is_conventional_json(path)]
    jsonl_files = [path for path in files if not is_conventional_json(path)]
    if json_files and args.json_record_path and ijson is None:
        print(
            "Error: --json-record-path requires ijson. Install it with:\n"
            "  python3 -m pip install ijson",
            file=sys.stderr,
        )
        return 2

    print(
        f"Found {file_count:,} source file(s), totalling {human_size(total_size)}: "
        f"{len(jsonl_files):,} JSONL/NDJSON and {len(json_files):,} JSON."
    )

    if args.fields:
        try:
            selected_fields = [parse_field_path(item) for item in args.fields]
        except ValueError as exc:
            print(f"Error: {exc}", file=sys.stderr)
            return 2
    elif args.no_discovery:
        print("Error: --no-discovery requires at least one --field.", file=sys.stderr)
        return 2
    else:
        print(f"Sampling up to {args.sample_per_file} record(s) per file…")
        discovered, valid, invalid = discover_fields(
            files, args.json_record_path, args.sample_per_file, args.max_fields
        )
        print(f"Sampled {valid:,} record(s); encountered {invalid:,} invalid item(s).")
        if len(discovered) >= args.max_fields:
            print(f"Warning: field discovery reached --max-fields ({args.max_fields:,}).")
        selected_fields = select_fields_by_name(discovered, args.field_name_contains)
        if not selected_fields:
            print(
                f"Error: no discovered field path contains "
                f"{args.field_name_contains!r} (case-insensitive). Increase "
                "--sample-per-file or specify one or more --field values.",
                file=sys.stderr,
            )
            return 2
        print(
            "Automatically selected field(s): "
            + ", ".join(display_path(path) for path in selected_fields)
        )

    terms = [term for term in (args.terms or []) if term]
    if not terms:
        print("\nEnter partial search values. Submit an empty value when finished.")
        while True:
            value = input(f"Search term {len(terms) + 1}: ")
            if not value:
                break
            terms.append(value)
    if not terms:
        print("Error: at least one search term is required.", file=sys.stderr)
        return 2

    engine = args.engine
    if engine == "auto":
        engine = "rg" if shutil.which("rg") else "python"
    if jsonl_files and engine == "rg" and shutil.which("rg") is None:
        print("Error: --engine rg requested but ripgrep is not installed.", file=sys.stderr)
        return 2

    print("\nSelected field(s): " + ", ".join(display_path(path) for path in selected_fields))
    print("Search term(s): " + ", ".join(repr(term) for term in terms))
    print(
        f"JSONL engine: {engine}; JSON engine: streaming; "
        f"term mode: {args.term_mode}; case-sensitive: {args.case_sensitive}"
    )
    if args.json_record_path:
        print(f"Nested JSON record array: {args.json_record_path}")
    temporary_path: Path | None = None
    started = time.monotonic()
    combined = Stats()
    try:
        with tempfile.NamedTemporaryFile(
            mode="w+b",
            prefix=".json-record-extract-",
            suffix=".tmp",
            dir=output.parent,
            delete=False,
        ) as destination:
            temporary_path = Path(destination.name)
            if jsonl_files:
                scanner = scan_jsonl_with_rg if engine == "rg" else scan_jsonl_with_python
                combined.add(
                    scanner(
                        jsonl_files,
                        destination,
                        selected_fields,
                        terms,
                        args.case_sensitive,
                        args.term_mode,
                    )
                )
            if json_files:
                combined.add(
                    scan_conventional_json(
                        json_files,
                        args.json_record_path,
                        destination,
                        selected_fields,
                        terms,
                        args.case_sensitive,
                        args.term_mode,
                    )
                )
            destination.flush()
            os.fsync(destination.fileno())
        os.replace(temporary_path, output)
        temporary_path = None
    except KeyboardInterrupt:
        print("\nCancelled; partial output was not retained.", file=sys.stderr)
        return 130
    except Exception as exc:
        print(f"\nError: {exc}", file=sys.stderr)
        return 1
    finally:
        if temporary_path is not None:
            try:
                temporary_path.unlink()
            except OSError:
                pass

    elapsed = time.monotonic() - started
    output_size = output.stat().st_size
    print(
        f"\nComplete: wrote {combined.matches:,} matching record(s) "
        f"({human_size(output_size)}) to {output}"
    )
    print(
        f"Checked {combined.jsonl_candidates:,} JSONL candidate line(s) and "
        f"{combined.json_records:,} structured JSON record(s); "
        f"{combined.invalid_json:,} invalid/incomplete input item(s); "
        f"elapsed {elapsed:,.1f}s."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

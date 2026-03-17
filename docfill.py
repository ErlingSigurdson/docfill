#!/usr/bin/env python3
from __future__ import annotations

import argparse
import bisect
import io
import json
import re
import shutil
import subprocess
import sys
import tempfile
import os
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Mapping, Sequence
from zipfile import ZipFile

import xml.etree.ElementTree as ET


DOCX_NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}

ODT_NS = {
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
}

XML_NS = "http://www.w3.org/XML/1998/namespace"
XML_SPACE_ATTR = f"{{{XML_NS}}}space"

for prefix, uri in DOCX_NS.items():
    ET.register_namespace(prefix, uri)

for prefix, uri in ODT_NS.items():
    ET.register_namespace(prefix, uri)


@dataclass
class TextSlot:
    owner: ET.Element
    attr: str  # "text" or "tail"

    def get(self) -> str:
        value = getattr(self.owner, self.attr)
        return "" if value is None else value

    def set(self, value: str) -> None:
        setattr(self.owner, self.attr, value)
        if self.attr == "text":
            if value and (value[:1].isspace() or value[-1:].isspace()):
                self.owner.set(XML_SPACE_ATTR, "preserve")
            elif self.owner.get(XML_SPACE_ATTR) == "preserve":
                del self.owner.attrib[XML_SPACE_ATTR]


@dataclass
class MatchInfo:
    matched_text: str
    canonical_placeholder: str
    replacement: str
    start: int
    end: int


@dataclass
class ReplacementSpec:
    pattern: re.Pattern[str] | None
    replacement_by_lookup: Dict[str, str]
    canonical_by_lookup: Dict[str, str]
    ignore_case: bool
    all_placeholders: List[str]

    def resolve_placeholder(self, matched_text: str) -> tuple[str, str]:
        lookup_key = matched_text.casefold() if self.ignore_case else matched_text
        return self.canonical_by_lookup[lookup_key], self.replacement_by_lookup[lookup_key]


@dataclass
class ProcessReport:
    output_path: Path | None
    matches_total: int
    records: List[MatchInfo]
    found_counts: Counter[str]
    missing_placeholders: List[str]


class RussianHelpFormatter(argparse.RawTextHelpFormatter):
    def add_usage(self, usage, actions, groups, prefix=None):
        if prefix is None:
            prefix = "Использование: "
        return super().add_usage(usage, actions, groups, prefix)

    def start_section(self, heading):
        translations = {
            "positional arguments": "Позиционные аргументы",
            "options": "Ключи",
            "optional arguments": "Ключи",
        }
        super().start_section(translations.get(heading, heading))


class RussianArgumentParser(argparse.ArgumentParser):
    def error(self, message):
        translated = self._translate_error_message(message)
        self.print_usage(sys.stderr)
        self.exit(2, f"Ошибка: {translated}\n")

    @staticmethod
    def _translate_error_message(message: str) -> str:
        translations = [
            (r"^unrecognized arguments: (.+)$", r"нераспознанные аргументы: \1"),
            (r"^the following arguments are required: (.+)$", r"отсутствуют обязательные аргументы: \1"),
            (r"^argument ([^:]+): expected one argument$", r"аргумент \1: ожидается одно значение"),
            (r"^argument ([^:]+): invalid choice: (.+) \(choose from (.+)\)$", r"аргумент \1: недопустимое значение: \2 (допустимые варианты: \3)"),
            (r"^argument ([^:]+): invalid int value: (.+)$", r"аргумент \1: недопустимое целое число: \2"),
            (r"^argument ([^:]+): invalid float value: (.+)$", r"аргумент \1: недопустимое число: \2"),
            (r"^argument ([^:]+): invalid (.+) value: (.+)$", r"аргумент \1: недопустимое значение: \3"),
            (r"^argument ([^:]+): not allowed with argument (.+)$", r"аргумент \1: нельзя использовать вместе с аргументом \2"),
            (r"^argument ([^:]+): ignored explicit argument (.+)$", r"аргумент \1: явное значение будет проигнорировано: \2"),
        ]

        for pattern, replacement in translations:
            translated = re.sub(pattern, replacement, message)
            if translated != message:
                return translated

        return message


def scalar_to_string(value) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (int, float)):
        return str(value)
    return str(value)


def flatten_json(value, prefix: str = "") -> Dict[str, str]:
    result: Dict[str, str] = {}

    if isinstance(value, dict):
        for key, item in value.items():
            key_str = str(key)
            new_prefix = key_str if not prefix else f"{prefix}.{key_str}"
            if isinstance(item, (dict, list)):
                result.update(flatten_json(item, new_prefix))
            else:
                result[new_prefix] = scalar_to_string(item)
        return result

    if isinstance(value, list):
        for index, item in enumerate(value):
            new_prefix = str(index) if not prefix else f"{prefix}.{index}"
            if isinstance(item, (dict, list)):
                result.update(flatten_json(item, new_prefix))
            else:
                result[new_prefix] = scalar_to_string(item)
        return result

    if prefix:
        result[prefix] = scalar_to_string(value)

    return result


def build_placeholder_spec(replacements: Mapping[str, str], ignore_case: bool) -> ReplacementSpec:
    placeholders = [placeholder for placeholder in replacements.keys() if placeholder]
    if not placeholders:
        return ReplacementSpec(None, {}, {}, ignore_case, [])

    replacement_by_lookup: Dict[str, str] = {}
    canonical_by_lookup: Dict[str, str] = {}

    for placeholder in placeholders:
        lookup_key = placeholder.casefold() if ignore_case else placeholder
        existing = canonical_by_lookup.get(lookup_key)
        if existing is not None and existing != placeholder:
            mode = "без учёта регистра" if ignore_case else "с учётом регистра"
            raise ValueError(
                "Обнаружены неоднозначные плейсхолдеры для поиска "
                f"{mode}: {existing!r} и {placeholder!r}."
            )
        canonical_by_lookup[lookup_key] = placeholder
        replacement_by_lookup[lookup_key] = replacements[placeholder]

    escaped = [re.escape(placeholder) for placeholder in sorted(placeholders, key=len, reverse=True)]
    flags = re.IGNORECASE if ignore_case else 0
    pattern = re.compile("|".join(escaped), flags)

    return ReplacementSpec(
        pattern=pattern,
        replacement_by_lookup=replacement_by_lookup,
        canonical_by_lookup=canonical_by_lookup,
        ignore_case=ignore_case,
        all_placeholders=sorted(placeholders),
    )


def iter_namespace_declarations(xml_bytes: bytes) -> Iterable[tuple[str, str]]:
    seen: set[tuple[str, str]] = set()
    for _, ns in ET.iterparse(io.BytesIO(xml_bytes), events=("start-ns",)):
        if ns not in seen:
            seen.add(ns)
            yield ns


def register_namespaces_from_xml(xml_bytes: bytes) -> None:
    for prefix, uri in iter_namespace_declarations(xml_bytes):
        ET.register_namespace(prefix or "", uri)


def ensure_root_namespace_declarations(serialized_xml: bytes, original_xml: bytes) -> bytes:
    original_ns = list(iter_namespace_declarations(original_xml))
    if not original_ns:
        return serialized_xml

    start = serialized_xml.find(b"<")
    if start < 0:
        return serialized_xml

    if serialized_xml[start:start + 2] == b"<?":
        decl_end = serialized_xml.find(b"?>", start)
        if decl_end < 0:
            return serialized_xml
        start = serialized_xml.find(b"<", decl_end + 2)
        if start < 0:
            return serialized_xml

    end = serialized_xml.find(b">", start)
    if end < 0:
        return serialized_xml

    root_open = serialized_xml[start:end]
    additions: List[bytes] = []

    for prefix, uri in original_ns:
        if prefix:
            decl = f'xmlns:{prefix}="{uri}"'.encode("utf-8")
        else:
            decl = f'xmlns="{uri}"'.encode("utf-8")
        if decl not in root_open:
            additions.append(decl)

    if not additions:
        return serialized_xml

    new_root_open = root_open + b" " + b" ".join(additions)
    return serialized_xml[:start] + new_root_open + serialized_xml[end:]


def collect_matches(full_text: str, spec: ReplacementSpec) -> List[MatchInfo]:
    if not full_text or spec.pattern is None:
        return []

    matches: List[MatchInfo] = []
    for match in spec.pattern.finditer(full_text):
        matched_text = match.group(0)
        canonical_placeholder, replacement = spec.resolve_placeholder(matched_text)
        matches.append(
            MatchInfo(
                matched_text=matched_text,
                canonical_placeholder=canonical_placeholder,
                replacement=replacement,
                start=match.start(),
                end=match.end(),
            )
        )
    return matches


def replace_in_slots(slots: Sequence[TextSlot], spec: ReplacementSpec) -> tuple[bool, List[MatchInfo]]:
    if not slots or spec.pattern is None:
        return False, []

    original_texts = [slot.get() for slot in slots]
    full_text = "".join(original_texts)
    matches = collect_matches(full_text, spec)
    if not matches:
        return False, []

    starts: List[int] = []
    pos = 0
    for text in original_texts:
        starts.append(pos)
        pos += len(text)

    current_texts = original_texts[:]

    for match in reversed(matches):
        start, end = match.start, match.end
        replacement = match.replacement

        first_idx = bisect.bisect_right(starts, start) - 1
        last_idx = bisect.bisect_right(starts, end - 1) - 1

        first_start = starts[first_idx]
        last_start = starts[last_idx]

        start_off = start - first_start
        end_off = end - last_start

        if first_idx == last_idx:
            current = current_texts[first_idx]
            current_texts[first_idx] = current[:start_off] + replacement + current[end_off:]
            continue

        first_text = current_texts[first_idx]
        last_text = current_texts[last_idx]
        current_texts[first_idx] = first_text[:start_off] + replacement + last_text[end_off:]

        for idx in range(first_idx + 1, last_idx + 1):
            current_texts[idx] = ""

    changed = current_texts != original_texts
    if changed:
        for slot, text in zip(slots, current_texts):
            slot.set(text)

    return changed, matches


def find_matches_in_slots(slots: Sequence[TextSlot], spec: ReplacementSpec) -> List[MatchInfo]:
    if not slots:
        return []
    full_text = "".join(slot.get() for slot in slots)
    return collect_matches(full_text, spec)


def collect_docx_slots(paragraph: ET.Element) -> List[TextSlot]:
    slots: List[TextSlot] = []
    w_t = f"{{{DOCX_NS['w']}}}t"

    for node in paragraph.iter(w_t):
        if node.text:
            slots.append(TextSlot(node, "text"))

    return slots


def collect_odt_slots(node: ET.Element) -> List[TextSlot]:
    slots: List[TextSlot] = []

    def visit(elem: ET.Element) -> None:
        if elem.text:
            slots.append(TextSlot(elem, "text"))
        for child in elem:
            visit(child)
            if child.tail:
                slots.append(TextSlot(child, "tail"))

    visit(node)
    return slots


def transform_docx_xml(xml_bytes: bytes, spec: ReplacementSpec) -> tuple[bytes, List[MatchInfo]]:
    register_namespaces_from_xml(xml_bytes)
    root = ET.fromstring(xml_bytes)
    paragraph_tag = f"{{{DOCX_NS['w']}}}p"
    changed = False
    records: List[MatchInfo] = []

    for paragraph in root.iter(paragraph_tag):
        slots = collect_docx_slots(paragraph)
        paragraph_changed, paragraph_matches = replace_in_slots(slots, spec)
        changed |= paragraph_changed
        records.extend(paragraph_matches)

    if not changed:
        return xml_bytes, records

    serialized = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    serialized = ensure_root_namespace_declarations(serialized, xml_bytes)
    return serialized, records


def scan_docx_xml(xml_bytes: bytes, spec: ReplacementSpec) -> List[MatchInfo]:
    root = ET.fromstring(xml_bytes)
    paragraph_tag = f"{{{DOCX_NS['w']}}}p"
    records: List[MatchInfo] = []

    for paragraph in root.iter(paragraph_tag):
        slots = collect_docx_slots(paragraph)
        records.extend(find_matches_in_slots(slots, spec))

    return records


def transform_odt_xml(xml_bytes: bytes, spec: ReplacementSpec) -> tuple[bytes, List[MatchInfo]]:
    register_namespaces_from_xml(xml_bytes)
    root = ET.fromstring(xml_bytes)
    paragraph_tags = {
        f"{{{ODT_NS['text']}}}p",
        f"{{{ODT_NS['text']}}}h",
    }
    changed = False
    records: List[MatchInfo] = []

    for elem in root.iter():
        if elem.tag in paragraph_tags:
            slots = collect_odt_slots(elem)
            elem_changed, elem_matches = replace_in_slots(slots, spec)
            changed |= elem_changed
            records.extend(elem_matches)

    if not changed:
        return xml_bytes, records

    serialized = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    serialized = ensure_root_namespace_declarations(serialized, xml_bytes)
    return serialized, records


def scan_odt_xml(xml_bytes: bytes, spec: ReplacementSpec) -> List[MatchInfo]:
    root = ET.fromstring(xml_bytes)
    paragraph_tags = {
        f"{{{ODT_NS['text']}}}p",
        f"{{{ODT_NS['text']}}}h",
    }
    records: List[MatchInfo] = []

    for elem in root.iter():
        if elem.tag in paragraph_tags:
            slots = collect_odt_slots(elem)
            records.extend(find_matches_in_slots(slots, spec))

    return records


def rewrite_zip_package(
    src_path: Path,
    dst_path: Path,
    spec: ReplacementSpec,
    package_kind: str,
) -> List[MatchInfo]:
    with tempfile.NamedTemporaryFile(delete=False, suffix=dst_path.suffix) as tmp_file:
        temp_path = Path(tmp_file.name)

    all_records: List[MatchInfo] = []

    try:
        with ZipFile(src_path, "r") as zin, ZipFile(temp_path, "w") as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)

                if package_kind == "docx":
                    if info.filename.startswith("word/") and info.filename.endswith(".xml"):
                        data, records = transform_docx_xml(data, spec)
                        all_records.extend(records)
                elif package_kind == "odt":
                    if info.filename.endswith(".xml"):
                        data, records = transform_odt_xml(data, spec)
                        all_records.extend(records)
                else:
                    raise ValueError(f"Unsupported package kind: {package_kind}")

                zout.writestr(info, data)

        dst_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(temp_path), str(dst_path))
        return all_records
    finally:
        if temp_path.exists():
            temp_path.unlink(missing_ok=True)


def scan_zip_package(src_path: Path, spec: ReplacementSpec, package_kind: str) -> List[MatchInfo]:
    all_records: List[MatchInfo] = []

    with ZipFile(src_path, "r") as zin:
        for info in zin.infolist():
            data = zin.read(info.filename)

            if package_kind == "docx":
                if info.filename.startswith("word/") and info.filename.endswith(".xml"):
                    all_records.extend(scan_docx_xml(data, spec))
            elif package_kind == "odt":
                if info.filename.endswith(".xml"):
                    all_records.extend(scan_odt_xml(data, spec))
            else:
                raise ValueError(f"Unsupported package kind: {package_kind}")

    return all_records


def convert_doc_to_docx(src_path: Path, libre_office_exec: str) -> Path:
    with tempfile.TemporaryDirectory() as tmp_dir_name:
        tmp_dir = Path(tmp_dir_name)

        completed = subprocess.run(
            [
                libre_office_exec,
                "--headless",
                "--convert-to",
                "docx",
                "--outdir",
                str(tmp_dir),
                str(src_path),
            ],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            check=False,
        )

        if completed.returncode != 0:
            raise RuntimeError(
                "Не удалось конвертировать .doc через LibreOffice.\n"
                f"Команда завершилась с кодом {completed.returncode}.\n"
                f"stdout:\n{completed.stdout}\n"
                f"stderr:\n{completed.stderr}"
            )

        converted = tmp_dir / f"{src_path.stem}.docx"
        if not converted.exists():
            raise RuntimeError(
                "LibreOffice завершился без ошибки, но результирующий .docx не найден."
            )

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_file:
            persistent_path = Path(tmp_file.name)

        shutil.copyfile(converted, persistent_path)
        return persistent_path


def load_replacements(json_path: Path) -> Dict[str, str]:
    with json_path.open("r", encoding="utf-8") as file:
        data = json.load(file)

    replacements = flatten_json(data)
    if not replacements:
        raise ValueError("В JSON не найдено скалярных значений для подстановки.")

    return replacements


def guess_libre_office_exec() -> str | None:
    candidates: List[str] = []

    for name in ("soffice", "libreoffice"):
        resolved = shutil.which(name)
        if resolved:
            candidates.append(resolved)

    if sys.platform.startswith("win"):
        platform_candidates = [
            Path("C:/Program Files/LibreOffice/program/soffice.exe"),
            Path("C:/Program Files (x86)/LibreOffice/program/soffice.exe"),
        ]
    elif sys.platform == "darwin":
        platform_candidates = [
            Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
        ]
    else:
        platform_candidates = [
            Path("/usr/bin/soffice"),
            Path("/usr/local/bin/soffice"),
            Path("/snap/bin/libreoffice"),
        ]

    for candidate in platform_candidates:
        if candidate.is_file():
            candidates.append(str(candidate))

    seen: set[str] = set()
    for candidate in candidates:
        if candidate not in seen:
            seen.add(candidate)
            return candidate

    return None


def resolve_libre_office_exec(user_value: str | None) -> str:
    if user_value:
        return user_value

    detected = guess_libre_office_exec()
    if detected:
        return detected

    raise RuntimeError(
        "Для обработки файлов .doc требуется LibreOffice.\n"
        "Программа не смогла автоматически найти запускаемый файл LibreOffice, "
        "необходимый для конвертации из .doc в .docx.\n"
        "Укажите путь явно через ключ --libre-office-exec."
    )


def build_output_path(src_path: Path, suffix: str, in_place: bool, force_ext: str | None = None) -> Path:
    if in_place:
        if force_ext and src_path.suffix.lower() != force_ext.lower():
            raise ValueError("Нельзя безопасно перезаписать файл с изменением расширения.")
        return src_path

    if force_ext is not None:
        return src_path.with_name(f"{src_path.stem}{suffix}{force_ext}")

    return src_path.with_name(f"{src_path.stem}{suffix}{src_path.suffix}")


def build_process_report(output_path: Path | None, records: List[MatchInfo], spec: ReplacementSpec) -> ProcessReport:
    found_counts: Counter[str] = Counter(match.canonical_placeholder for match in records)
    missing_placeholders = [placeholder for placeholder in spec.all_placeholders if placeholder not in found_counts]
    return ProcessReport(
        output_path=output_path,
        matches_total=len(records),
        records=records,
        found_counts=found_counts,
        missing_placeholders=missing_placeholders,
    )


def check_one_file(src_path: Path, spec: ReplacementSpec, libre_office_exec: str | None) -> ProcessReport:
    ext = src_path.suffix.lower()

    if ext == ".docx":
        records = scan_zip_package(src_path, spec, "docx")
        return build_process_report(None, records, spec)

    if ext == ".odt":
        records = scan_zip_package(src_path, spec, "odt")
        return build_process_report(None, records, spec)

    if ext == ".doc":
        temp_docx = convert_doc_to_docx(src_path, resolve_libre_office_exec(libre_office_exec))
        try:
            records = scan_zip_package(temp_docx, spec, "docx")
            return build_process_report(None, records, spec)
        finally:
            temp_docx.unlink(missing_ok=True)

    raise ValueError(
        f"Неподдерживаемый формат: {src_path.name}. Поддерживаются только .doc, .docx и .odt."
    )


def process_one_file(
    src_path: Path,
    spec: ReplacementSpec,
    suffix: str,
    in_place: bool,
    libre_office_exec: str | None,
) -> ProcessReport:
    ext = src_path.suffix.lower()

    if ext == ".docx":
        dst_path = build_output_path(src_path, suffix, in_place)
        records = rewrite_zip_package(src_path, dst_path, spec, "docx")
        return build_process_report(dst_path, records, spec)

    if ext == ".odt":
        dst_path = build_output_path(src_path, suffix, in_place)
        records = rewrite_zip_package(src_path, dst_path, spec, "odt")
        return build_process_report(dst_path, records, spec)

    if ext == ".doc":
        if in_place:
            raise ValueError(
                "Режим --in-place не поддерживается для .doc: старый бинарный формат "
                "конвертируется в .docx."
            )

        temp_docx = convert_doc_to_docx(src_path, resolve_libre_office_exec(libre_office_exec))
        try:
            dst_path = build_output_path(src_path, suffix, False, ".docx")
            records = rewrite_zip_package(temp_docx, dst_path, spec, "docx")
            return build_process_report(dst_path, records, spec)
        finally:
            temp_docx.unlink(missing_ok=True)

    raise ValueError(
        f"Неподдерживаемый формат: {src_path.name}. Поддерживаются только .doc, .docx и .odt."
    )


def print_verbose_replacements(records: Sequence[MatchInfo]) -> None:
    if not records:
        print("  Список замен: нет совпадений.")
        return

    print("  Список замен:")
    for record in records:
        print(f"    {record.matched_text} -> {record.replacement}")


def print_check_report(src_path: Path, report: ProcessReport, verbose: bool) -> None:
    print(
        f"[CHECK] {src_path}: совпадений={report.matches_total}, "
        f"уникальных_найдено={len(report.found_counts)}, "
        f"уникальных_не_найдено={len(report.missing_placeholders)}"
    )

    if verbose:
        print_verbose_replacements(report.records)
        print("  Не найдены:")
        if report.missing_placeholders:
            for placeholder in report.missing_placeholders:
                print(f"    {placeholder}")
        else:
            print("    нет")


def print_process_report(src_path: Path, report: ProcessReport, verbose: bool, in_place: bool) -> None:
    if report.output_path is None:
        raise ValueError("Отсутствует путь к выходному файлу в отчёте обработки.")

    if in_place:
        print(f"[OK] {src_path} (заменено плейсхолдеров: {report.matches_total})")
    else:
        print(f"[OK] {src_path} -> {report.output_path} (заменено плейсхолдеров: {report.matches_total})")

    if verbose:
        print_verbose_replacements(report.records)


def make_parser() -> argparse.ArgumentParser:
    prog_name = Path(os.path.basename(sys.argv[0] if sys.argv else "docfill.py")).name or "docfill.py"

    parser = RussianArgumentParser(
        prog=prog_name,
        add_help=False,
        formatter_class=RussianHelpFormatter,
        usage=(
            "%(prog)s [--ignore-case] [--check] [-v|--verbose] [--suffix SUFFIX] "
            "[--in-place] [--libre-office-exec LIBRE_OFFICE_EXEC] placeholders_json document [document ...]"
        ),
        description=(
            "Подставляет значения из JSON в документы .docx и .odt, а старые .doc\n"
            "обрабатывает через промежуточную конвертацию в .docx.\n\n"
            "Плейсхолдеры строятся из иерархии JSON-ключей через точку, например:\n"
            "ООО_Ромашка.Руководитель.Должность.Именительный\n\n"
            "Плейсхолдер — это заменяемый текст в документе."
        ),
        epilog=(
            "Примеры:\n"
            f"  {prog_name} data.json contract.docx letter.odt\n"
            f"  {prog_name} --ignore-case data.json contract.docx\n"
            f"  {prog_name} --check data.json contract.docx letter.odt\n"
            f"  {prog_name} -v data.json contract.docx\n"
            f"  {prog_name} --suffix .filled data.json a.docx b.odt\n"
            f"  {prog_name} --libre-office-exec \"C:\\Program Files\\LibreOffice\\program\\soffice.exe\" data.json legacy.doc\n"
        ),
    )
    parser.add_argument(
        "-h",
        "--help",
        action="help",
        help="Показать эту справку и выйти.",
    )
    parser.add_argument(
        "json_file",
        help="Путь к JSON-файлу со значениями плейсхолдеров.",
    )
    parser.add_argument(
        "documents",
        nargs="+",
        metavar="document",
        help="Один или несколько файлов .doc, .docx или .odt.",
    )
    parser.add_argument(
        "--ignore-case",
        action="store_true",
        help=(
            "Искать плейсхолдеры без учёта регистра. Если в JSON есть плейсхолдеры,\n"
            "отличающиеся только регистром, программа выдаст ошибку и завершится."
        ),
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help=(
            "Ничего не записывать; только проверить документы и показать,\n"
            "какие плейсхолдеры в них найдены."
        ),
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help=(
            "Показать подробный вывод. Для обычной обработки это полный\n"
            "список замен вида 'было -> стало'."
        ),
    )
    parser.add_argument(
        "--suffix",
        default=".rendered",
        help="Суффикс выходного файла (по умолчанию: %(default)s).",
    )
    parser.add_argument(
        "--in-place",
        action="store_true",
        help=(
            "Перезаписывать исходные .docx и .odt на месте. Для .doc\n"
            "не поддерживается."
        ),
    )
    parser.add_argument(
        "--libre-office-exec",
        dest="libre_office_exec",
        help=(
            "Команда или путь к запускаемому файлу LibreOffice. Нужен только\n"
            "для файлов формата .doc и только если программе не удалось\n"
            "самостоятельно найти LibreOffice, необходимый для конвертации\n"
            "из .doc в .docx."
        ),
    )
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    parser = make_parser()
    args = parser.parse_args(argv)

    json_path = Path(args.json_file)
    if not json_path.is_file():
        print(f"JSON-файл не найден: {json_path}", file=sys.stderr)
        return 1

    try:
        replacements = load_replacements(json_path)
        spec = build_placeholder_spec(replacements, args.ignore_case)
    except Exception as exc:
        print(f"Ошибка подготовки плейсхолдеров: {exc}", file=sys.stderr)
        return 1

    failures = 0

    for doc_name in args.documents:
        src_path = Path(doc_name)

        if not src_path.is_file():
            print(f"[ERROR] Файл не найден: {src_path}", file=sys.stderr)
            failures += 1
            continue

        try:
            if args.check:
                report = check_one_file(src_path, spec, args.libre_office_exec)
                print_check_report(src_path, report, args.verbose)
            else:
                report = process_one_file(
                    src_path=src_path,
                    spec=spec,
                    suffix=args.suffix,
                    in_place=args.in_place,
                    libre_office_exec=args.libre_office_exec,
                )
                print_process_report(src_path, report, args.verbose, args.in_place)
        except Exception as exc:
            print(f"[ERROR] {src_path}: {exc}", file=sys.stderr)
            failures += 1

    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())

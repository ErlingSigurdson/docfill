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


def build_replacement_spec(replacements: Mapping[str, str], ignore_case: bool) -> ReplacementSpec:
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


def convert_doc_to_docx(src_path: Path, libreoffice_exec: str) -> Path:
    with tempfile.TemporaryDirectory() as tmp_dir_name:
        tmp_dir = Path(tmp_dir_name)

        completed = subprocess.run(
            [
                libreoffice_exec,
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
                "Не удалось конвертировать .doc с помощью LibreOffice.\n"
                f"Команда завершилась с кодом {completed.returncode}.\n"
                f"stdout:\n{completed.stdout}\n"
                f"stderr:\n{completed.stderr}"
            )

        converted = tmp_dir / f"{src_path.stem}.docx"
        if not converted.exists():
            raise RuntimeError(
                "LibreOffice завершил работу без ошибки, но итоговый файл в формате DOCX не найден."
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
        raise ValueError(
            f"В JSON-файле {json_path} не найдено скалярных значений для подстановки."
        )

    return replacements


def load_and_merge_replacements(json_paths: Sequence[Path]) -> Dict[str, str]:
    merged: Dict[str, str] = {}
    source_by_placeholder: Dict[str, Path] = {}

    for json_path in json_paths:
        replacements = load_replacements(json_path)

        for placeholder, value in replacements.items():
            existing_value = merged.get(placeholder)
            if existing_value is None:
                merged[placeholder] = value
                source_by_placeholder[placeholder] = json_path
                continue

            if existing_value != value:
                first_source = source_by_placeholder[placeholder]
                raise ValueError(
                    "Обнаружен конфликт плейсхолдеров между JSON-файлами: "
                    f"{placeholder!r} имеет разные значения в {first_source} и {json_path}."
                )

    if not merged:
        raise ValueError("Не найдено плейсхолдеров для подстановки ни в одном JSON-файле.")

    return merged


def parse_alias_pairs(alias_values: Sequence[str]) -> List[tuple[str, str]]:
    pairs: List[tuple[str, str]] = []

    for raw_value in alias_values:
        for item in raw_value.split(","):
            item = item.strip()
            if not item:
                continue

            if "=" not in item:
                raise ValueError(
                    "Некорректное значение опции --alias: "
                    f"{raw_value!r}. Ожидается формат ИСХОДНОЕ=НОВОЕ[,ИСХОДНОЕ=НОВОЕ ...]."
                )

            source_prefix, alias_prefix = item.split("=", 1)
            source_prefix = source_prefix.strip()
            alias_prefix = alias_prefix.strip()

            if not source_prefix or not alias_prefix:
                raise ValueError(
                    "Некорректное значение опции --alias: "
                    f"{raw_value!r}. И исходное, и новое имя должны быть непустыми."
                )

            pairs.append((source_prefix, alias_prefix))

    return pairs


def apply_aliases(replacements: Mapping[str, str], alias_pairs: Sequence[tuple[str, str]]) -> Dict[str, str]:
    if not alias_pairs:
        return dict(replacements)

    result: Dict[str, str] = dict(replacements)
    source_of_placeholder: Dict[str, str] = {
        placeholder: "JSON-файл"
        for placeholder in result
    }

    for source_prefix, alias_prefix in alias_pairs:
        matched_any = False

        for placeholder, value in replacements.items():
            if placeholder == source_prefix:
                aliased_placeholder = alias_prefix
            elif placeholder.startswith(source_prefix + "."):
                aliased_placeholder = alias_prefix + placeholder[len(source_prefix):]
            else:
                continue

            matched_any = True
            existing_value = result.get(aliased_placeholder)
            if existing_value is None:
                result[aliased_placeholder] = value
                source_of_placeholder[aliased_placeholder] = (
                    f"alias {source_prefix!r} -> {alias_prefix!r}"
                )
                continue

            if existing_value != value:
                existing_source = source_of_placeholder.get(aliased_placeholder, "другой источник")
                raise ValueError(
                    "Обнаружен конфликт при применении опции --alias: "
                    f"{aliased_placeholder!r} уже существует ({existing_source}) "
                    f"и получает другое значение при alias {source_prefix!r} -> {alias_prefix!r}."
                )

        if not matched_any:
            raise ValueError(
                "Опция --alias ссылается на несуществующий префикс плейсхолдеров: "
                f"{source_prefix!r}."
            )

    return result

def guess_libreoffice_exec() -> str | None:
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


def resolve_libreoffice_exec(user_value: str | None) -> str:
    if user_value:
        return user_value

    detected = guess_libreoffice_exec()
    if detected:
        return detected

    raise RuntimeError(
        "Для обработки файлов в формате DOC (как правило, имеют расширение '.doc') требуется LibreOffice.\n"
        "Программа не смогла автоматически найти запускаемый файл LibreOffice, "
        "необходимый для конвертации из DOC в DOCX.\n"
        "Укажите путь к исполняемому файлу LibreOffice или его имя напрямую с помощью ключа --libreoffice-exec."
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


def check_one_file(src_path: Path, spec: ReplacementSpec, libreoffice_exec: str | None) -> ProcessReport:
    ext = src_path.suffix.lower()

    if ext == ".docx":
        records = scan_zip_package(src_path, spec, "docx")
        return build_process_report(None, records, spec)

    if ext == ".odt":
        records = scan_zip_package(src_path, spec, "odt")
        return build_process_report(None, records, spec)

    if ext == ".doc":
        temp_docx = convert_doc_to_docx(src_path, resolve_libreoffice_exec(libreoffice_exec))
        try:
            records = scan_zip_package(temp_docx, spec, "docx")
            return build_process_report(None, records, spec)
        finally:
            temp_docx.unlink(missing_ok=True)

    raise ValueError(
        f"Неподдерживаемый формат: {src_path.name}. Поддерживаются только DOCX, ODT и DOC."
    )


def process_one_file(
    src_path: Path,
    spec: ReplacementSpec,
    suffix: str,
    in_place: bool,
    libreoffice_exec: str | None,
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
                "Опция --in-place не поддерживается для DOC."
            )

        temp_docx = convert_doc_to_docx(src_path, resolve_libreoffice_exec(libreoffice_exec))
        try:
            dst_path = build_output_path(src_path, suffix, False, ".docx")
            records = rewrite_zip_package(temp_docx, dst_path, spec, "docx")
            return build_process_report(dst_path, records, spec)
        finally:
            temp_docx.unlink(missing_ok=True)

    raise ValueError(
        f"Неподдерживаемый формат: {src_path.name}. Поддерживаются только DOCX, ODT и DOC."
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
            "[--in-place] [--alias ALIAS] [--libreoffice-exec LIBREOFFICE_EXEC] "
            "input_file [input_file ...]"
        ),
        description=(
            "Подставляет значения из одного или нескольких JSON-файлов в один или несколько документов\n"
            "в формате DOCX, ODT или DOC. JSON-файлы распознаются по расширению .json.\n"
        ),
    )
    parser.add_argument(
        "-h",
        "--help",
        action="help",
        help="Показать эту справку и выйти.",
    )
    parser.add_argument(
        "input_files",
        nargs="+",
        metavar="input_file",
        help=(
            "Один или несколько целевых файлов. JSON-файлы используются как источники данных,\n"
            "а целевые файлы .doc, .docx и .odt — как документы, в которых выполняется подстановка."
        ),
    )
    parser.add_argument(
        "--ignore-case",
        action="store_true",
        help=(
            "Искать плейсхолдеры без учёта регистра. Если в JSON есть плейсхолдеры,\n"
            "отличающиеся только регистром, программа завершится с ошибкой."
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
            "Выполнять подстановку непосредственно в целевых файлах .docx и .odt, а не создавать для этого их копии. Для .doc\n"
            "не поддерживается."
        ),
    )

    parser.add_argument(
        "--alias",
        action="append",
        default=[],
        metavar="SOURCE=TARGET[,SOURCE=TARGET ...]",
        help=(
            "Установить соответствия между полями JSON на время текущего запуска."
        ),
    )
    parser.add_argument(
        "--libreoffice-exec",
        dest="libreoffice_exec",
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

    input_paths = [Path(item) for item in args.input_files]
    json_paths = [path for path in input_paths if path.suffix.lower() == ".json"]
    document_paths = [path for path in input_paths if path.suffix.lower() in {".doc", ".docx", ".odt"}]
    unsupported_paths = [
        path for path in input_paths
        if path.suffix.lower() not in {".json", ".doc", ".docx", ".odt"}
    ]

    if unsupported_paths:
        for path in unsupported_paths:
            print(
                f"Неподдерживаемый входной файл: {path}. "
                "Поддерживаются только .json, .doc, .docx и .odt.",
                file=sys.stderr,
            )
        return 1

    if not json_paths:
        print(
            "Не указан ни один JSON-файл со значениями плейсхолдеров.",
            file=sys.stderr,
        )
        return 1

    if not document_paths:
        print(
            "Не указан ни один документ для обработки. "
            "Передайте хотя бы один файл .doc, .docx или .odt.",
            file=sys.stderr,
        )
        return 1

    missing_inputs = [path for path in input_paths if not path.is_file()]
    if missing_inputs:
        for path in missing_inputs:
            print(f"[ERROR] Файл не найден: {path}", file=sys.stderr)
        return 1

    try:
        replacements = load_and_merge_replacements(json_paths)
        alias_pairs = parse_alias_pairs(args.alias)
        replacements = apply_aliases(replacements, alias_pairs)
        spec = build_replacement_spec(replacements, args.ignore_case)
    except Exception as exc:
        print(f"Ошибка подготовки плейсхолдеров: {exc}", file=sys.stderr)
        return 1

    failures = 0

    for src_path in document_paths:
        try:
            if args.check:
                report = check_one_file(src_path, spec, args.libreoffice_exec)
                print_check_report(src_path, report, args.verbose)
            else:
                report = process_one_file(
                    src_path=src_path,
                    spec=spec,
                    suffix=args.suffix,
                    in_place=args.in_place,
                    libreoffice_exec=args.libreoffice_exec,
                )
                print_process_report(src_path, report, args.verbose, args.in_place)
        except Exception as exc:
            print(f"[ERROR] {src_path}: {exc}", file=sys.stderr)
            failures += 1

    return 1 if failures else 0

if __name__ == "__main__":
    raise SystemExit(main())

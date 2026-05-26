from __future__ import annotations

import argparse
import sys
import tempfile
from pathlib import Path

import config as cfg
from app import (
    cleanup_word_story_text,
    make_pairs,
    word_table_column_count,
    word_table_row_count,
)


def parse_args() -> argparse.Namespace:
    """Read options for the parsing inspection report."""
    parser = argparse.ArgumentParser(
        description=(
            "Inspect how Word COM reads paired Kyrgyz/Russian documents, including "
            "headers, footers, body paragraphs, list labels, and tables."
        )
    )
    parser.add_argument("--kyrgyz-dir", default=cfg.DEFAULT_KYRGYZ_DIR, help="Folder with Kyrgyz documents.")
    parser.add_argument("--russian-dir", default=cfg.DEFAULT_RUSSIAN_DIR, help="Folder with Russian documents.")
    parser.add_argument("--pair", help="Only inspect one filename stem, for example 'test' or 'legacy'.")
    parser.add_argument("--output", help="Optional text file path for the report.")
    parser.add_argument(
        "--no-doc-conversion-check",
        action="store_true",
        help="Do not compare .doc files against temporary Word-exported .docx copies.",
    )
    return parser.parse_args()


class WordParsingInspector:
    """Build a readable report of what Word COM sees in source documents."""

    def __init__(self) -> None:
        self.word = None

    def __enter__(self) -> "WordParsingInspector":
        try:
            import win32com.client  # type: ignore
        except ImportError as exc:
            raise RuntimeError("inspect_parsing.py requires pywin32 on Windows.") from exc

        self.word = win32com.client.DispatchEx(cfg.WORD_APPLICATION_NAME)
        self.word.Visible = False
        self.word.DisplayAlerts = cfg.WORD_DISPLAY_ALERTS_NONE
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        if self.word is not None:
            self.word.Quit()

    def inspect_pair(self, pair, check_doc_conversion: bool) -> list[str]:
        """Inspect both sides of one matched document pair."""
        lines = [
            "=" * 100,
            f"PAIR: {pair.name}",
            f"KYRGYZ:  {pair.kyrgyz}",
            f"RUSSIAN: {pair.russian}",
            "",
        ]
        lines.extend(self.inspect_document(pair.kyrgyz, "KYRGYZ ORIGINAL"))
        lines.extend(self.inspect_document(pair.russian, "RUSSIAN ORIGINAL"))

        if check_doc_conversion:
            with tempfile.TemporaryDirectory(prefix=cfg.PREPARED_DOCX_TEMP_PREFIX) as temp_dir:
                temp_root = Path(temp_dir)
                for source, label in (
                    (pair.kyrgyz, "KYRGYZ AFTER .DOC -> .DOCX"),
                    (pair.russian, "RUSSIAN AFTER .DOC -> .DOCX"),
                ):
                    if source.suffix.casefold() == cfg.LEGACY_DOC_EXTENSION:
                        converted = temp_root / f"{source.stem}_{label.split()[0].lower()}{cfg.DOCX_EXTENSION}"
                        self.export_doc_to_docx(source, converted)
                        lines.extend(self.inspect_document(converted, label))

        return lines

    def export_doc_to_docx(self, source_path: Path, target_path: Path) -> None:
        """Export a legacy .doc to .docx through Word COM for comparison."""
        target_path.parent.mkdir(parents=True, exist_ok=True)
        doc = self.open_document(source_path)
        try:
            doc.SaveAs2(str(target_path.resolve()), FileFormat=cfg.WORD_FORMAT_XML_DOCUMENT)
        finally:
            doc.Close(SaveChanges=cfg.WORD_DO_NOT_SAVE_CHANGES)

    def inspect_document(self, path: Path, label: str) -> list[str]:
        """Inspect one Word document path."""
        doc = self.open_document(path)
        try:
            lines = [
                "-" * 100,
                f"{label}: {path}",
                f"HEADER: {self.story_text(doc, cfg.HEADER_STORY)}",
                f"FOOTER: {self.story_text(doc, cfg.FOOTER_STORY)}",
                "BODY:",
            ]
            lines.extend(self.body_lines(doc))
            lines.append("")
            return lines
        finally:
            doc.Close(SaveChanges=cfg.WORD_DO_NOT_SAVE_CHANGES)

    def open_document(self, path: Path):
        """Open a document read-only without adding it to Word recent files."""
        return self.word.Documents.Open(
            FileName=str(path.resolve()),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            OpenAndRepair=True,
        )

    def story_text(self, doc, story: str) -> str:
        """Read visible header/footer text from all common section variants."""
        pieces: list[str] = []
        for section in doc.Sections:
            collection = section.Headers if story == cfg.HEADER_STORY else section.Footers
            for index in cfg.WORD_HEADER_FOOTER_INDEXES:
                try:
                    text = cleanup_word_story_text(collection(index).Range.Text)
                except Exception:
                    continue
                if text:
                    pieces.append(text)
        return cfg.MULTI_SECTION_TEXT_SEPARATOR.join(dict.fromkeys(pieces))

    def body_lines(self, doc) -> list[str]:
        """Read top-level body paragraphs and tables in document order."""
        blocks = []
        for paragraph in doc.Paragraphs:
            try:
                if paragraph.Range.Information(cfg.WORD_INFO_WITHIN_TABLE):
                    continue
            except Exception:
                continue
            blocks.append(("paragraph", paragraph.Range.Start, paragraph))

        for table in doc.Tables:
            blocks.append(("table", table.Range.Start, table))

        lines: list[str] = []
        paragraph_count = 0
        table_count = 0
        for kind, _, item in sorted(blocks, key=lambda block: block[1]):
            if kind == "paragraph":
                paragraph_count += 1
                lines.extend(self.paragraph_lines(item, paragraph_count))
            else:
                table_count += 1
                lines.extend(self.table_lines(item, table_count))

        if not lines:
            lines.append("  <empty>")
        return lines

    def paragraph_lines(self, paragraph, index: int) -> list[str]:
        """Format one paragraph, including Word's visible list label."""
        text = cleanup_word_story_text(paragraph.Range.Text)
        if not text:
            return []

        list_label = word_list_label(paragraph.Range)
        style = word_style_name(paragraph)
        prefix = f"  P{index:03d}"
        if list_label:
            prefix += f" [{list_label}]"
        if style:
            prefix += f" <{style}>"
        return [f"{prefix}: {text}"]

    def table_lines(self, table, index: int) -> list[str]:
        """Format one table, including cell text and cell paragraph list labels."""
        rows = word_table_row_count(table)
        columns = word_table_column_count(table)
        lines = [f"  TABLE {index}: rows={rows}, columns={columns}"]
        for row_index in range(1, rows + 1):
            row_values = []
            for column_index in range(1, columns + 1):
                try:
                    cell = table.Cell(row_index, column_index)
                except Exception:
                    row_values.append("<missing>")
                    continue
                row_values.append(cell_text_with_list_labels(cell))
            lines.append(f"    R{row_index}: {row_values}")
        return lines


def word_list_label(word_range) -> str:
    """Return Word's visible list label for a range, such as 'I)' or '1.'."""
    try:
        return str(word_range.ListFormat.ListString or "").strip()
    except Exception:
        return ""


def word_style_name(paragraph) -> str:
    """Return the paragraph style name, if Word exposes it."""
    try:
        return str(paragraph.Style.NameLocal)
    except Exception:
        return ""


def cell_text_with_list_labels(cell) -> str:
    """Read table cell text while preserving visible list labels in the report."""
    parts: list[str] = []
    for paragraph in cell.Range.Paragraphs:
        text = cleanup_word_story_text(paragraph.Range.Text)
        if not text:
            continue
        label = word_list_label(paragraph.Range)
        parts.append(f"{label} {text}".strip() if label else text)
    return " | ".join(parts)


def build_report(args: argparse.Namespace) -> str:
    """Create the full parsing report for all selected pairs."""
    pairs, warnings = make_pairs(Path(args.kyrgyz_dir), Path(args.russian_dir))
    if args.pair:
        pairs = [pair for pair in pairs if pair.name.casefold() == args.pair.casefold()]

    lines: list[str] = []
    for warning in warnings:
        lines.append(f"WARNING: {warning}")
    if not pairs:
        lines.append("No matching document pairs found.")
        return "\n".join(lines)

    with WordParsingInspector() as inspector:
        for pair in pairs:
            lines.extend(inspector.inspect_pair(pair, not args.no_doc_conversion_check))

    return "\n".join(lines)


def main() -> int:
    """Script entry point."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    args = parse_args()
    report = build_report(args)
    if args.output:
        Path(args.output).write_text(report, encoding="utf-8")
        print(f"Wrote parsing report to {args.output}")
    else:
        print(report)
    return 0


if __name__ == "__main__":
    sys.exit(main())

from __future__ import annotations

import html
import re
import textwrap
from dataclasses import dataclass
from html.parser import HTMLParser
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
HTML_PATH = BASE_DIR / "trial_balance_builder_guide.html"
PDF_PATH = BASE_DIR / "trial_balance_builder_guide.pdf"

PAGE_WIDTH = 612.0
PAGE_HEIGHT = 792.0
MARGIN = 54.0
CONTENT_WIDTH = PAGE_WIDTH - (2 * MARGIN)
BODY_COLOR = (0.12, 0.17, 0.14)
HEADING_COLOR = (0.11, 0.29, 0.21)
BOX_FILL = (0.97, 0.98, 0.97)
BOX_STROKE = (0.81, 0.85, 0.82)


@dataclass(frozen=True)
class Block:
    kind: str
    text: str
    box: bool = False


class GuideHTMLParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.blocks: list[Block] = []
        self._current_tag: str | None = None
        self._current_parts: list[str] = []
        self._current_small = False
        self._list_stack: list[dict[str, int | str]] = []
        self._box_depth = 0

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attr_map = {key: value or "" for key, value in attrs}
        if tag == "div" and "box" in attr_map.get("class", "").split():
            self._box_depth += 1
            return
        if tag in {"ul", "ol"}:
            self._list_stack.append({"kind": tag, "counter": 0})
            return
        if tag in {"h1", "h2", "p", "li"}:
            self._current_tag = tag
            self._current_parts = []
            self._current_small = "small" in attr_map.get("class", "").split()
            return
        if tag == "br" and self._current_tag:
            self._current_parts.append("\n")

    def handle_endtag(self, tag: str) -> None:
        if tag == "div" and self._box_depth:
            self._box_depth -= 1
            return
        if tag in {"ul", "ol"}:
            if self._list_stack:
                self._list_stack.pop()
            return
        if tag in {"h1", "h2", "p", "li"} and self._current_tag == tag:
            self._flush_current()

    def handle_data(self, data: str) -> None:
        if self._current_tag:
            self._current_parts.append(data)

    def _flush_current(self) -> None:
        text = html.unescape("".join(self._current_parts))
        text = re.sub(r"\s+", " ", text).strip()
        if not text:
            self._current_tag = None
            self._current_parts = []
            self._current_small = False
            return

        kind = "small" if self._current_small else str(self._current_tag)
        if self._current_tag == "li" and self._list_stack:
            current_list = self._list_stack[-1]
            if current_list["kind"] == "ol":
                current_list["counter"] = int(current_list["counter"]) + 1
                prefix = f"{current_list['counter']}. "
            else:
                prefix = "- "
            text = f"{prefix}{text}"

        self.blocks.append(Block(kind=kind, text=text, box=self._box_depth > 0))
        self._current_tag = None
        self._current_parts = []
        self._current_small = False


def _font_for(kind: str) -> tuple[str, float]:
    if kind == "h1":
        return "F2", 22.0
    if kind == "h2":
        return "F2", 15.0
    if kind == "small":
        return "F1", 10.0
    return "F1", 11.0


def _leading_for(size: float) -> float:
    return round(size * 1.35, 2)


def _spacing_after(kind: str) -> float:
    if kind == "h1":
        return 14.0
    if kind == "h2":
        return 10.0
    if kind == "small":
        return 8.0
    return 8.0


def _wrap_text(text: str, max_width: float, font_size: float, *, bullet: bool = False) -> list[str]:
    width_factor = 0.58 if font_size >= 15 else 0.53
    chars = max(24, int(max_width / (font_size * width_factor)))
    if bullet and chars > 2:
        chars -= 2
    wrapped = textwrap.wrap(
        text,
        width=chars,
        break_long_words=False,
        break_on_hyphens=False,
    )
    return wrapped or [text]


def _measure_block(block: Block, width: float) -> tuple[list[str], float]:
    font_name, font_size = _font_for(block.kind)
    bullet = block.kind == "li"
    lines = _wrap_text(block.text, width, font_size, bullet=bullet)
    height = (len(lines) * _leading_for(font_size)) + _spacing_after(block.kind)
    return lines, height


def _escape_pdf_text(value: str) -> str:
    return value.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


class SimplePDF:
    def __init__(self) -> None:
        self.pages: list[str] = []
        self._commands: list[str] = []
        self._y = PAGE_HEIGHT - MARGIN

    def _new_page(self) -> None:
        if self._commands:
            self.pages.append("\n".join(self._commands))
        self._commands = []
        self._y = PAGE_HEIGHT - MARGIN

    def _ensure_page_started(self) -> None:
        if not self._commands and not self.pages:
            self._new_page()

    def _ensure_space(self, height: float) -> None:
        self._ensure_page_started()
        if self._y - height < MARGIN:
            self._new_page()

    def _set_fill_color(self, rgb: tuple[float, float, float]) -> None:
        self._commands.append(f"{rgb[0]:.3f} {rgb[1]:.3f} {rgb[2]:.3f} rg")

    def _set_stroke_color(self, rgb: tuple[float, float, float]) -> None:
        self._commands.append(f"{rgb[0]:.3f} {rgb[1]:.3f} {rgb[2]:.3f} RG")

    def _draw_text(self, x: float, y: float, text: str, font_name: str, font_size: float, color: tuple[float, float, float]) -> None:
        escaped = _escape_pdf_text(text)
        self._commands.append("BT")
        self._commands.append(f"/{font_name} {font_size:.2f} Tf")
        self._commands.append(f"{color[0]:.3f} {color[1]:.3f} {color[2]:.3f} rg")
        self._commands.append(f"1 0 0 1 {x:.2f} {y:.2f} Tm")
        self._commands.append(f"({escaped}) Tj")
        self._commands.append("ET")

    def _draw_box(self, x: float, y: float, width: float, height: float) -> None:
        self._set_fill_color(BOX_FILL)
        self._commands.append(f"{x:.2f} {y:.2f} {width:.2f} {height:.2f} re f")
        self._set_stroke_color(BOX_STROKE)
        self._commands.append("1 w")
        self._commands.append(f"{x:.2f} {y:.2f} {width:.2f} {height:.2f} re S")

    def add_block(self, block: Block) -> None:
        text_width = CONTENT_WIDTH
        text_x = MARGIN
        if block.kind == "li":
            text_x += 12.0
            text_width -= 12.0

        font_name, font_size = _font_for(block.kind)
        lines, height = _measure_block(block, text_width)
        self._ensure_space(height)

        color = HEADING_COLOR if block.kind in {"h1", "h2"} else BODY_COLOR
        baseline_y = self._y - font_size
        for line in lines:
            self._draw_text(text_x, baseline_y, line, font_name, font_size, color)
            baseline_y -= _leading_for(font_size)
        self._y -= height

    def add_box_group(self, blocks: list[Block]) -> None:
        inner_padding = 12.0
        text_width = CONTENT_WIDTH - (2 * inner_padding)
        prepared: list[tuple[Block, list[str], float]] = []
        content_height = 0.0
        for block in blocks:
            lines, block_height = _measure_block(block, text_width)
            prepared.append((block, lines, block_height))
            content_height += block_height
        total_height = content_height + (2 * inner_padding)
        self._ensure_space(total_height)

        rect_bottom = self._y - total_height
        self._draw_box(MARGIN, rect_bottom, CONTENT_WIDTH, total_height)

        text_y = self._y - inner_padding
        for block, lines, block_height in prepared:
            font_name, font_size = _font_for(block.kind)
            color = HEADING_COLOR if block.kind in {"h1", "h2"} else BODY_COLOR
            baseline_y = text_y - font_size
            for line in lines:
                self._draw_text(MARGIN + inner_padding, baseline_y, line, font_name, font_size, color)
                baseline_y -= _leading_for(font_size)
            text_y -= block_height
        self._y -= total_height + 4.0

    def finish(self) -> list[str]:
        if self._commands:
            self.pages.append("\n".join(self._commands))
            self._commands = []
        return self.pages


def _parse_blocks(html_text: str) -> list[Block]:
    parser = GuideHTMLParser()
    parser.feed(html_text)
    return parser.blocks


def _build_page_streams(blocks: list[Block]) -> list[str]:
    pdf = SimplePDF()
    index = 0
    while index < len(blocks):
        block = blocks[index]
        if block.box:
            group: list[Block] = []
            while index < len(blocks) and blocks[index].box:
                group.append(blocks[index])
                index += 1
            pdf.add_box_group(group)
            continue
        pdf.add_block(block)
        index += 1
    return pdf.finish()


def _pdf_bytes_from_pages(page_streams: list[str]) -> bytes:
    if not page_streams:
        page_streams = [""]

    objects: dict[int, bytes] = {}
    objects[1] = b"<< /Type /Catalog /Pages 2 0 R >>"

    kids: list[str] = []
    page_count = len(page_streams)
    for page_index, stream in enumerate(page_streams):
        page_obj = 5 + (page_index * 2)
        content_obj = page_obj + 1
        kids.append(f"{page_obj} 0 R")
        stream_bytes = stream.encode("latin-1", errors="replace")
        objects[content_obj] = (
            f"<< /Length {len(stream_bytes)} >>\nstream\n".encode("latin-1")
            + stream_bytes
            + b"\nendstream"
        )
        objects[page_obj] = (
            f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {PAGE_WIDTH:.0f} {PAGE_HEIGHT:.0f}] "
            f"/Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> /Contents {content_obj} 0 R >>"
        ).encode("latin-1")

    objects[2] = f"<< /Type /Pages /Count {page_count} /Kids [{' '.join(kids)}] >>".encode("latin-1")
    objects[3] = b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"
    objects[4] = b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>"

    output = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offsets: dict[int, int] = {}
    for object_number in range(1, max(objects) + 1):
        offsets[object_number] = len(output)
        output.extend(f"{object_number} 0 obj\n".encode("latin-1"))
        output.extend(objects[object_number])
        output.extend(b"\nendobj\n")

    xref_offset = len(output)
    total_objects = max(objects) + 1
    output.extend(f"xref\n0 {total_objects}\n".encode("latin-1"))
    output.extend(b"0000000000 65535 f \n")
    for object_number in range(1, total_objects):
        output.extend(f"{offsets[object_number]:010d} 00000 n \n".encode("latin-1"))
    output.extend(
        (
            f"trailer\n<< /Size {total_objects} /Root 1 0 R >>\n"
            f"startxref\n{xref_offset}\n%%EOF\n"
        ).encode("latin-1")
    )
    return bytes(output)


def build_pdf() -> Path:
    html_text = HTML_PATH.read_text(encoding="utf-8")
    blocks = _parse_blocks(html_text)
    page_streams = _build_page_streams(blocks)
    PDF_PATH.write_bytes(_pdf_bytes_from_pages(page_streams))
    return PDF_PATH


if __name__ == "__main__":
    output_path = build_pdf()
    print(output_path)

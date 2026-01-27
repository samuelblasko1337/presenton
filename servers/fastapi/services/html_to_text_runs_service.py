from html.parser import HTMLParser
from typing import List, Optional
import re

from models.pptx_models import PptxFontModel, PptxTextRunModel


# --- NEW: helpers to keep PPTX export robust and support Markdown emphasis ---

_TAG_RE = re.compile(r"</?[a-zA-Z][^>]*>")  # detect real HTML tags
_MD_BOLD_RE = re.compile(r"(\*\*|__)(.+?)\1")
# italic: avoid matching bullet markers like "* " at line start
_MD_ITALIC_RE = re.compile(r"(?<!\*)\*(?!\s)(.+?)(?<!\s)\*(?!\*)|(?<!_)_(?!\s)(.+?)(?<!\s)_(?!_)")

def _escape_angles_outside_tags(text: str) -> str:
    """
    Escape < and > so that strings like '<1%' or 'MTTR <4h' won't be treated as tags.
    We only do this when the string does NOT already contain real HTML tags.
    """
    if _TAG_RE.search(text):
        return text
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def _markdown_emphasis_to_html(text: str) -> str:
    """
    Convert Markdown emphasis to HTML tags.
    IMPORTANT: Also works when the string already contains HTML tags by only converting
    text segments outside of tags.
    """
    def _convert_segment(seg: str) -> str:
        # Convert bold first
        seg = _MD_BOLD_RE.sub(r"<strong>\2</strong>", seg)

        # Convert italic; group(1) matches *...*, group(2) matches _..._
        def _italics_sub(m: re.Match) -> str:
            inner = m.group(1) if m.group(1) is not None else m.group(2)
            return f"<em>{inner}</em>"

        seg = _MD_ITALIC_RE.sub(_italics_sub, seg)
        return seg

    # Split into tags and non-tags, convert only non-tag parts
    parts = re.split(r"(<[^>]+>)", text)
    for i, p in enumerate(parts):
        if p.startswith("<") and p.endswith(">"):
            continue
        parts[i] = _convert_segment(p)
    return "".join(parts)



class InlineHTMLToRunsParser(HTMLParser):
    def __init__(self, base_font: PptxFontModel):
        super().__init__(convert_charrefs=True)
        self.base_font = base_font
        self.tag_stack: List[str] = []
        self.text_runs: List[PptxTextRunModel] = []

    def _current_font(self) -> PptxFontModel:
        font_json = self.base_font.model_dump()
        is_bold = any(tag in ("strong", "b") for tag in self.tag_stack)
        is_italic = any(tag in ("em", "i") for tag in self.tag_stack)
        is_underline = any(tag == "u" for tag in self.tag_stack)
        is_strike = any(tag in ("s", "strike", "del") for tag in self.tag_stack)
        is_code = any(tag == "code" for tag in self.tag_stack)

        if is_bold:
            font_json["font_weight"] = 700
        if is_italic:
            font_json["italic"] = True
        if is_underline:
            font_json["underline"] = True
        if is_strike:
            font_json["strike"] = True
        if is_code:
            font_json["name"] = "Courier New"

        return PptxFontModel(**font_json)

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        if tag == "br":
            self.text_runs.append(PptxTextRunModel(text="\n"))
            return
        self.tag_stack.append(tag)

    def handle_endtag(self, tag):
        tag = tag.lower()
        for i in range(len(self.tag_stack) - 1, -1, -1):
            if self.tag_stack[i] == tag:
                del self.tag_stack[i]
                break

    def handle_data(self, data):
        if data == "":
            return
        self.text_runs.append(PptxTextRunModel(text=data, font=self._current_font()))


def parse_html_text_to_text_runs(
    text: str, base_font: Optional[PptxFontModel] = None
) -> List[PptxTextRunModel]:
    # Normalize line breaks
    normalized_text = text.replace("\r\n", "\n").replace("\r", "\n")

    # NEW: if no HTML tags exist, safely escape < and > and convert markdown emphasis
    normalized_text = _escape_angles_outside_tags(normalized_text)
    normalized_text = _markdown_emphasis_to_html(normalized_text)

    # Keep newline behavior consistent
    normalized_text = normalized_text.replace("\n", "<br>")

    parser = InlineHTMLToRunsParser(base_font if base_font else PptxFontModel())
    parser.feed(normalized_text)
    return parser.text_runs



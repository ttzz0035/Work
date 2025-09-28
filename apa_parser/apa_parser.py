# Apps/APAParser/apa_parser.py
from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path
from typing import List, Dict, Any
import xml.etree.ElementTree as ET
import re


class APAFormatError(Exception):
    pass


@dataclass
class APAParseResult:
    source_path: Path
    xml_text: str      # XML全文（宣言～ルート閉じタグまで）
    xml_encoding: str  # 使用したエンコーディング推定
    root_tag: str
    summary: Dict[str, Any]


class APAParser:
    """
    APA専用パーサ：
      - バイナリから不要文字（NUL/制御）を削除
      - <?xml ～ </Root> までのXML全文を抽出・保持
    """

    def parse_many(self, paths: List[Path]) -> List[APAParseResult]:
        return [self.parse_file(p) for p in paths]

    def parse_file(self, path: Path) -> APAParseResult:
        raw = path.read_bytes()

        # 1) 不要文字（NULや制御）を削除
        cleaned = self._sanitize(raw)

        # 2) XML宣言開始位置
        start = cleaned.find(b'<?xml')
        if start == -1:
            # UTF-16LE痕跡が混在している場合の救済
            cleaned2 = cleaned.replace(b'\x00', b'')
            start = cleaned2.find(b'<?xml')
            if start == -1:
                raise APAFormatError("XML宣言が見つかりません")
            xml_bytes = cleaned2[start:]
        else:
            xml_bytes = cleaned[start:]

        # 3) デコード（UTF-16LE優先 → UTF-8）
        xml_text_full = self._decode(xml_bytes)

        # 4) ルート閉じタグまで切り出し（全文）
        xml_text = self._slice_to_root_close(xml_text_full)

        # 5) XMLパース確認
        try:
            root = ET.fromstring(xml_text)
        except ET.ParseError as e:
            raise APAFormatError(f"XML解析エラー: {e}")

        # 6) サマリ生成（最低限）
        summary = self._build_summary(root)

        return APAParseResult(
            source_path=path,
            xml_text=xml_text,
            xml_encoding="utf-16le" if "\x00" in xml_text_full[:200] else "utf-8",
            root_tag=root.tag,
            summary=summary,
        )

    # ---------------- 内部 ----------------
    def _sanitize(self, data: bytes) -> bytes:
        # NUL削除 + 制御文字削除（\t,\n,\r は許容）
        return bytes(
            b for b in data
            if b != 0x00 and not (b < 0x20 and b not in (0x09, 0x0A, 0x0D)) and b != 0x7F
        )

    def _decode(self, data: bytes) -> str:
        try:
            return data.decode("utf-16le")
        except Exception:
            return data.decode("utf-8", errors="replace")

    def _slice_to_root_close(self, text: str) -> str:
        m = re.search(r'<\?xml[^>]*\?>\s*<([A-Za-z0-9:_\-]+)', text)
        if not m:
            return text
        root = m.group(1)
        close = f"</{root}>"
        j = text.rfind(close)
        return text[: j + len(close)] if j != -1 else text

    def _build_summary(self, root: ET.Element) -> Dict[str, Any]:
        def attr(e, name): return e.attrib.get(name, "") if e is not None else ""
        hdr = root.find(".//Header")
        eth = root.find(".//ETH-IP")
        return {
            "Version": attr(hdr, "wVersion"),
            "IPAddress": attr(eth, "szIPAddress"),
            "Subnet": attr(eth, "szSubnetMask"),
            "Gateway": attr(eth, "szGateway"),
            "DeviceCount": len(root.findall(".//DeviceList/*")),
            "LocalSlaveCount": len(root.findall(".//LocalSlaveList/*")),
        }

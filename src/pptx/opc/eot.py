"""Convert TTF/OTF font bytes to Embedded OpenType (EOT) format.

PowerPoint stores embedded fonts as .fntdata files in EOT format (W3C EOT spec,
version 0x00020001). This module builds the EOT header from font metadata and
appends the raw font data.

Reference: https://www.w3.org/Submission/EOT/
"""

from __future__ import annotations

import struct
from io import BytesIO

from fontTools.ttLib import TTFont


def ttf_to_eot(font_bytes: bytes) -> bytes:
    """Convert raw TTF/OTF *font_bytes* to EOT format.

    Returns the complete EOT file as bytes (header + font data).
    """
    font = TTFont(BytesIO(font_bytes))

    os2 = font["OS/2"]
    head = font["head"]
    name_table = font["name"]

    # Extract PANOSE bytes from OS/2 table
    panose_values = os2.panose
    panose_bytes = bytes([
        panose_values.bFamilyType,
        panose_values.bSerifStyle,
        panose_values.bWeight,
        panose_values.bProportion,
        panose_values.bContrast,
        panose_values.bStrokeVariation,
        panose_values.bArmStyle,
        panose_values.bLetterForm,
        panose_values.bMidline,
        panose_values.bXHeight,
    ])

    # Italic flag: bit 0 of fsSelection
    italic = 0x01 if (os2.fsSelection & 0x01) else 0x00

    # Extract name strings as UTF-16LE bytes
    family_name = _get_name_bytes(name_table, 1)
    style_name = _get_name_bytes(name_table, 2)
    version_name = _get_name_bytes(name_table, 5)
    full_name = _get_name_bytes(name_table, 4)

    font.close()

    font_data_size = len(font_bytes)

    # Build the header (everything before FontData)
    header = bytearray()

    # EOTSize placeholder (4 bytes) — will be filled at the end
    header += struct.pack("<I", 0)
    # FontDataSize
    header += struct.pack("<I", font_data_size)
    # Version 0x00020001
    header += struct.pack("<I", 0x00020001)
    # Flags: 0 (no compression, no XOR)
    header += struct.pack("<I", 0x00000000)
    # FontPANOSE (10 bytes)
    header += panose_bytes
    # Charset: DEFAULT_CHARSET = 0x01
    header += struct.pack("B", 0x01)
    # Italic
    header += struct.pack("B", italic)
    # Weight
    header += struct.pack("<I", os2.usWeightClass)
    # fsType
    header += struct.pack("<H", os2.fsType)
    # MagicNumber
    header += struct.pack("<H", 0x504C)
    # UnicodeRange1-4
    header += struct.pack("<I", os2.ulUnicodeRange1)
    header += struct.pack("<I", os2.ulUnicodeRange2)
    header += struct.pack("<I", os2.ulUnicodeRange3)
    header += struct.pack("<I", os2.ulUnicodeRange4)
    # CodePageRange1-2
    header += struct.pack("<I", os2.ulCodePageRange1)
    header += struct.pack("<I", os2.ulCodePageRange2)
    # CheckSumAdjustment
    header += struct.pack("<I", head.checkSumAdjustment & 0xFFFFFFFF)
    # Reserved1-4
    header += struct.pack("<IIII", 0, 0, 0, 0)

    # Name strings with padding
    for name_bytes in (family_name, style_name, version_name, full_name):
        header += struct.pack("<H", 0x0000)  # Padding
        header += struct.pack("<H", len(name_bytes))  # Size
        header += name_bytes  # UTF-16LE string

    # Version 0x00020001 adds RootString
    header += struct.pack("<H", 0x0000)  # Padding5
    header += struct.pack("<H", 0)  # RootStringSize (empty = no URL restriction)

    # Now set EOTSize = header size + font data size
    eot_size = len(header) + font_data_size
    struct.pack_into("<I", header, 0, eot_size)

    return bytes(header) + font_bytes


def _get_name_bytes(name_table, name_id: int) -> bytes:
    """Get name record *name_id* as UTF-16LE bytes.

    Tries platform 3 (Windows), encoding 1 (Unicode BMP), lang 0x0409 (English)
    first, then falls back to any available record.
    """
    record = name_table.getName(name_id, 3, 1, 0x0409)
    if record is None:
        # Fallback: try any platform
        for r in name_table.names:
            if r.nameID == name_id:
                record = r
                break
    if record is None:
        return b""
    text = str(record)
    return text.encode("utf-16-le")

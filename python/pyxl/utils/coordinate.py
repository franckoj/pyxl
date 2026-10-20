"""Coordinate utilities for openpyxl compatibility."""

def get_column_letter(col_idx: int) -> str:
    """Convert a 1-based column index to an Excel column letter."""
    if not (1 <= col_idx <= 16384):
        raise ValueError(f"Column index {col_idx} is out of bounds (allowed: 1-16384)")
    col = col_idx
    letters = []
    while col > 0:
        rem = (col - 1) % 26
        letters.append(chr(ord('A') + rem))
        col = (col - 1) // 26
    return "".join(reversed(letters))


def column_index_from_string(col_str: str) -> int:
    """Convert an Excel column letter to a 1-based column index."""
    clean = col_str.strip().lstrip("$").upper()
    if not clean or not clean.isalpha():
        raise ValueError(f"Invalid column string: {col_str}")
    col = 0
    for char in clean:
        col = col * 26 + (ord(char) - ord('A') + 1)
    if not (1 <= col <= 16384):
        raise ValueError(f"Column index for {col_str} is out of bounds (allowed: 1-16384)")
    return col


def coordinate_to_tuple(coord: str) -> tuple[int, int]:
    """Convert a coordinate string (e.g. 'A1', '$B$10') to (row, col) (1-based)."""
    clean = coord.strip().replace("$", "")
    letters = "".join(c for c in clean if c.isalpha())
    digits = "".join(c for c in clean if c.isdigit())
    if not letters or not digits or len(letters) + len(digits) != len(clean):
        raise ValueError(f"Invalid coordinate format: {coord}")
    row = int(digits)
    col = column_index_from_string(letters)
    if not (1 <= row <= 1048576):
        raise ValueError(f"Row {row} is out of bounds (allowed: 1-1048576)")
    return (row, col)

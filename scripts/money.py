"""Money handling for Construction CFO.

Rule: money is stored as INTEGER minor units (cents) in SQLite, and all
arithmetic is done in Decimal. Binary float is never used for money.

    to_minor("1234.56", scale=2) -> 123456
    from_minor(123456, scale=2)  -> Decimal("1234.56")
    fmt_minor(123456, "AUD", 2)  -> "AUD 1,234.56"
"""

from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP


def to_minor(value, scale: int = 2) -> int:
    """Convert a money value (str/Decimal/int) to integer minor units.

    Floats are rejected — pass a string or Decimal so we never inherit
    binary-float rounding error.
    """
    if isinstance(value, float):
        raise TypeError(
            "refusing to convert a float to money; pass a string or Decimal"
        )
    d = Decimal(str(value)) if not isinstance(value, Decimal) else value
    factor = Decimal(10) ** scale
    return int((d * factor).quantize(Decimal(1), rounding=ROUND_HALF_UP))


def from_minor(minor: int, scale: int = 2) -> Decimal:
    """Convert integer minor units back to a Decimal money value."""
    factor = Decimal(10) ** scale
    return (Decimal(minor) / factor).quantize(
        Decimal(1).scaleb(-scale), rounding=ROUND_HALF_UP
    )


def fmt_minor(minor: int, currency: str = "AUD", scale: int = 2) -> str:
    """Human-readable money string with thousands separators."""
    d = from_minor(minor, scale)
    sign = "-" if d < 0 else ""
    whole, _, frac = f"{abs(d):.{scale}f}".partition(".")
    whole = f"{int(whole):,}"
    body = f"{whole}.{frac}" if scale else whole
    return f"{currency} {sign}{body}"

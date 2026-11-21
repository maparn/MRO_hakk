#!/usr/bin/env python3
"""Official ECB exchange-rate lookup with a local last-known-good cache."""

from __future__ import annotations

import json
from decimal import Decimal, InvalidOperation
from pathlib import Path
from tempfile import NamedTemporaryFile
from urllib.request import Request, urlopen
from xml.etree import ElementTree


def parse_ecb_rate(payload: bytes, currency: str = "USD") -> tuple[Decimal, str]:
    root = ElementTree.fromstring(payload)
    rate_date = None
    wanted = currency.upper()
    for element in root.iter():
        if element.attrib.get("time"):
            rate_date = element.attrib["time"]
        if element.attrib.get("currency", "").upper() == wanted:
            try:
                rate = Decimal(element.attrib["rate"])
            except (KeyError, InvalidOperation) as error:
                raise ValueError(f"ECB returned an invalid {wanted} rate") from error
            if rate <= 0 or not rate_date:
                raise ValueError(f"ECB returned an incomplete {wanted} rate")
            return rate, rate_date
    raise ValueError(f"ECB rate for {wanted} was not found")


def _write_cache(path: Path, record: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with NamedTemporaryFile(
        mode="w", encoding="utf-8", dir=path.parent,
        prefix=".mro-rate-", suffix=".tmp", delete=False,
    ) as handle:
        temporary = Path(handle.name)
        json.dump(record, handle, indent=2)
        handle.write("\n")
    try:
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)


def get_ecb_rate(url: str, cache_path: Path, *, currency: str = "USD",
                 timeout: float = 10) -> dict:
    """Return live ECB rate, or the cached official rate if ECB is unavailable."""
    try:
        request = Request(url, headers={"User-Agent": "Magnetic-MRO/1.0"})
        with urlopen(request, timeout=timeout) as response:
            rate, rate_date = parse_ecb_rate(response.read(), currency)
        record = {
            "provider": "ECB", "base_currency": "EUR",
            "target_currency": currency.upper(), "rate": str(rate),
            "date": rate_date, "source_url": url,
        }
        _write_cache(cache_path, record)
        return {**record, "status": "LIVE"}
    except (OSError, ValueError, ElementTree.ParseError) as live_error:
        try:
            cached = json.loads(cache_path.read_text(encoding="utf-8"))
            rate = Decimal(str(cached["rate"]))
            if (cached.get("provider") != "ECB" or rate <= 0
                    or cached.get("target_currency") != currency.upper()):
                raise ValueError("cache metadata mismatch")
            return {**cached, "rate": str(rate), "status": "CACHED",
                    "warning": str(live_error)}
        except (OSError, KeyError, ValueError, InvalidOperation, json.JSONDecodeError) as cache_error:
            raise ValueError(
                f"ECB rate unavailable and no valid cache exists: {live_error}; {cache_error}"
            ) from live_error

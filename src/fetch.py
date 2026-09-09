"""Download helpers with an ETag-based incremental cache.

The ONS bucket is a plain public S3 bucket, so we can list it and read each
object's ETag.  We keep the ETag of every month we have already processed in
`state.json`; on the daily run only the months whose ETag changed are
downloaded again (in practice: the current month and, for a few days, the
previous one).
"""

from __future__ import annotations

import io
import json
import os
import time
import xml.etree.ElementTree as ET
from pathlib import Path

import pandas as pd
import requests

from .sources import ONS_S3

S3_NS = "{http://s3.amazonaws.com/doc/2006-03-01/}"
# O WAF da CCEE devolve 403 para clientes que não parecem navegador (o runner do
# GitHub Actions cai nisso). Os dados são públicos e CC-BY; o que muda aqui são
# só os cabeçalhos, para o firewall não classificar a requisição como bot.
UA = {
    "User-Agent": ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36"),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "pt-BR,pt;q=0.9,en;q=0.8",
}
TIMEOUT = 180
# Erros 4xx não adianta repetir — exceto 429, que é pedido para esperar.
NAO_REPETIR = {400, 401, 403, 404, 405, 410, 451}


def _session() -> requests.Session:
    s = requests.Session()
    s.headers.update(UA)
    return s


SESSION = _session()


def get(url: str, session: requests.Session | None = None, **kw) -> requests.Response:
    """GET com algumas tentativas — o bucket do ONS às vezes devolve 503."""
    s = session or SESSION
    last = None
    for attempt in range(5):
        try:
            r = s.get(url, timeout=TIMEOUT, **kw)
            if r.status_code in NAO_REPETIR:
                raise RuntimeError(f"HTTP {r.status_code} em {url}")
            if r.status_code < 500:
                r.raise_for_status()
                return r
            last = RuntimeError(f"{r.status_code} for {url}")
        except RuntimeError:
            raise
        except Exception as exc:  # noqa: BLE001
            last = exc
        time.sleep(2 * (attempt + 1))
    raise RuntimeError(f"failed to GET {url}: {last}")


def list_prefix(prefix: str) -> dict[str, dict]:
    """List one S3 prefix -> {filename: {"etag":..., "size":..., "modified":...}}."""
    out: dict[str, dict] = {}
    token = None
    while True:
        url = f"{ONS_S3}/?list-type=2&prefix={prefix}/&max-keys=1000"
        if token:
            url += f"&continuation-token={requests.utils.quote(token, safe='')}"
        root = ET.fromstring(get(url).content)
        for c in root.findall(f"{S3_NS}Contents"):
            key = c.findtext(f"{S3_NS}Key", "")
            name = key.rsplit("/", 1)[-1]
            if not name:
                continue
            out[name] = {
                "etag": (c.findtext(f"{S3_NS}ETag", "") or "").strip('"'),
                "size": int(c.findtext(f"{S3_NS}Size", "0") or 0),
                "modified": c.findtext(f"{S3_NS}LastModified", ""),
            }
        if root.findtext(f"{S3_NS}IsTruncated", "false") != "true":
            break
        token = root.findtext(f"{S3_NS}NextContinuationToken")
        if not token:
            break
    return out


def read_csv(url: str, **kw) -> pd.DataFrame:
    """Read an ONS open-data CSV (semicolon separated, UTF-8, dot decimals)."""
    raw = get(url).content
    defaults = dict(sep=";", encoding="utf-8", low_memory=False)
    defaults.update(kw)
    try:
        return pd.read_csv(io.BytesIO(raw), **defaults)
    except UnicodeDecodeError:
        defaults["encoding"] = "latin-1"
        return pd.read_csv(io.BytesIO(raw), **defaults)


# --------------------------------------------------------------------------
# state file
# --------------------------------------------------------------------------
class State:
    def __init__(self, path: str | os.PathLike):
        self.path = Path(path)
        self.data: dict = {}
        if self.path.exists():
            try:
                self.data = json.loads(self.path.read_text())
            except json.JSONDecodeError:
                self.data = {}

    def etag(self, dataset: str, filename: str) -> str | None:
        return self.data.get("etags", {}).get(dataset, {}).get(filename)

    def set_etag(self, dataset: str, filename: str, etag: str) -> None:
        self.data.setdefault("etags", {}).setdefault(dataset, {})[filename] = etag

    def save(self, **extra) -> None:
        self.data.update(extra)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(json.dumps(self.data, indent=1, ensure_ascii=False))

"""
spellcheck.py — Hunspell spell-check helpers for the QA Check tab.

Reintroduced in Task #33 as an OPT-IN check, later promoted to opt-OUT and
expanded from 5 to 38 supported languages: the 5 most common (en/es/fr/de/it)
are pre-bundled at build time so the .exe works fully offline for the majority
of jobs, and the remaining 33 are auto-downloaded the first time their
language is detected in a file (cached under ``~/.cache/anonymizer/``). Uses
``spylls`` (pure-Python Hunspell port, no C extension required).

Public API
----------
- DICTIONARY_SOURCES         : explicit URL map for all 38 supported languages
- BUNDLED_LANGUAGES          : 5-tuple of codes pre-bundled by build_exe (the
                               other 33 entries in DICTIONARY_SOURCES are
                               downloaded on demand at runtime)
- bundled_dir() / cache_dir(): on-disk locations for .aff/.dic pairs
- normalize_lang_code(code)  : ISO normalisation ("es-MX" -> "es_MX")
- get_dictionary(lang)       : returns spylls Dictionary | None (lazy load+cache)
- spell_check_text(text, dictionary, ignore) -> list[str] of misspelled tokens
- ensure_bundled_dictionaries(target_dir, *, strict) : used by build_exe.py
"""

from __future__ import annotations

import os
import re
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

# -----------------------------------------------------------------------------
# Source-of-truth dictionary map. The LibreOffice repo layout is irregular
# (German lives under `de/` with the `_frami` suffix, French under `fr_FR/`
# with `fr.aff`/`fr.dic`, Norwegian under `no/`, Romanian under `ro/`,
# Lithuanian uses `lt.{aff,dic}` inside `lt_LT/`, etc.) — never assume
# `<lang>/<lang>.aff`. Every URL below was verified live (HTTP 200) against
# https://github.com/LibreOffice/dictionaries during the task that promoted
# spell-check from "5 bundled languages" to "any detected language is
# auto-downloaded". When adding a new language, run a quick fetch against
# https://raw.githubusercontent.com/LibreOffice/dictionaries/master/<folder>/
# <file> and confirm a 200 before inserting it here.
# -----------------------------------------------------------------------------
DICTIONARY_SOURCES: Dict[str, Dict[str, str]] = {
    # -- English variants
    "en_US": {"folder": "en",     "aff": "en_US.aff",       "dic": "en_US.dic"},
    "en_GB": {"folder": "en",     "aff": "en_GB.aff",       "dic": "en_GB.dic"},
    "en_AU": {"folder": "en",     "aff": "en_AU.aff",       "dic": "en_AU.dic"},
    "en_CA": {"folder": "en",     "aff": "en_CA.aff",       "dic": "en_CA.dic"},
    # -- Spanish variants
    "es_ES": {"folder": "es",     "aff": "es_ES.aff",       "dic": "es_ES.dic"},
    "es_MX": {"folder": "es",     "aff": "es_MX.aff",       "dic": "es_MX.dic"},
    "es_AR": {"folder": "es",     "aff": "es_AR.aff",       "dic": "es_AR.dic"},
    # -- French
    "fr_FR": {"folder": "fr_FR",  "aff": "fr.aff",          "dic": "fr.dic"},
    # -- German variants (frami flavour shipped by LibreOffice)
    "de_DE": {"folder": "de",     "aff": "de_DE_frami.aff", "dic": "de_DE_frami.dic"},
    "de_AT": {"folder": "de",     "aff": "de_AT_frami.aff", "dic": "de_AT_frami.dic"},
    "de_CH": {"folder": "de",     "aff": "de_CH_frami.aff", "dic": "de_CH_frami.dic"},
    # -- Italian
    "it_IT": {"folder": "it_IT",  "aff": "it_IT.aff",       "dic": "it_IT.dic"},
    # -- Portuguese (BR and PT are linguistically distinct, both kept)
    "pt_BR": {"folder": "pt_BR",  "aff": "pt_BR.aff",       "dic": "pt_BR.dic"},
    "pt_PT": {"folder": "pt_PT",  "aff": "pt_PT.aff",       "dic": "pt_PT.dic"},
    # -- Dutch
    "nl_NL": {"folder": "nl_NL",  "aff": "nl_NL.aff",       "dic": "nl_NL.dic"},
    # -- Slavic
    "pl_PL": {"folder": "pl_PL",  "aff": "pl_PL.aff",       "dic": "pl_PL.dic"},
    "ru_RU": {"folder": "ru_RU",  "aff": "ru_RU.aff",       "dic": "ru_RU.dic"},
    "cs_CZ": {"folder": "cs_CZ",  "aff": "cs_CZ.aff",       "dic": "cs_CZ.dic"},
    "sk_SK": {"folder": "sk_SK",  "aff": "sk_SK.aff",       "dic": "sk_SK.dic"},
    "uk_UA": {"folder": "uk_UA",  "aff": "uk_UA.aff",       "dic": "uk_UA.dic"},
    "bg_BG": {"folder": "bg_BG",  "aff": "bg_BG.aff",       "dic": "bg_BG.dic"},
    "hr_HR": {"folder": "hr_HR",  "aff": "hr_HR.aff",       "dic": "hr_HR.dic"},
    "sl_SI": {"folder": "sl_SI",  "aff": "sl_SI.aff",       "dic": "sl_SI.dic"},
    "sr":    {"folder": "sr",     "aff": "sr.aff",          "dic": "sr.dic"},
    "sr_Latn":{"folder": "sr",    "aff": "sr-Latn.aff",     "dic": "sr-Latn.dic"},
    # -- Nordic
    "sv_SE": {"folder": "sv_SE",  "aff": "sv_SE.aff",       "dic": "sv_SE.dic"},
    "da_DK": {"folder": "da_DK",  "aff": "da_DK.aff",       "dic": "da_DK.dic"},
    "nb_NO": {"folder": "no",     "aff": "nb_NO.aff",       "dic": "nb_NO.dic"},
    "nn_NO": {"folder": "no",     "aff": "nn_NO.aff",       "dic": "nn_NO.dic"},
    # -- Other European
    "hu_HU": {"folder": "hu_HU",  "aff": "hu_HU.aff",       "dic": "hu_HU.dic"},
    "el_GR": {"folder": "el_GR",  "aff": "el_GR.aff",       "dic": "el_GR.dic"},
    "ro_RO": {"folder": "ro",     "aff": "ro_RO.aff",       "dic": "ro_RO.dic"},
    "tr_TR": {"folder": "tr_TR",  "aff": "tr_TR.aff",       "dic": "tr_TR.dic"},
    "lt_LT": {"folder": "lt_LT",  "aff": "lt.aff",          "dic": "lt.dic"},
    "lv_LV": {"folder": "lv_LV",  "aff": "lv_LV.aff",       "dic": "lv_LV.dic"},
    "et_EE": {"folder": "et_EE",  "aff": "et_EE.aff",       "dic": "et_EE.dic"},
    # -- Semitic
    "he_IL": {"folder": "he_IL",  "aff": "he_IL.aff",       "dic": "he_IL.dic"},
    "ar":    {"folder": "ar",     "aff": "ar.aff",          "dic": "ar.dic"},
}

# Languages physically copied into the offline .exe at build time. We keep this
# small (5 languages) so the bundle stays under ~10 MB. Every other entry in
# DICTIONARY_SOURCES is downloaded on demand when its language is detected in
# the file the user is checking — this is the runtime behaviour the user
# explicitly asked for ("siempre que detecte el idioma lo descargue").
BUNDLED_LANGUAGES: Tuple[str, ...] = ("en_US", "es_ES", "fr_FR", "de_DE", "it_IT")

_BASE_URL = "https://raw.githubusercontent.com/LibreOffice/dictionaries/master"

# -----------------------------------------------------------------------------
# Map common ISO/IETF codes to our cache keys. Order matters: more specific
# keys (with region) are looked up first; bare language codes fall through.
# Region-specific entries route to the closest available LibreOffice variant.
# Anything not listed here AND not a top-level key in DICTIONARY_SOURCES is
# treated as unsupported (spell-check is silently skipped, the rest of the QA
# run continues).
# -----------------------------------------------------------------------------
_LANG_NORMALISATION: Dict[str, str] = {
    # English: route every English region to the closest LibreOffice variant.
    "en":    "en_US",
    "en-us": "en_US", "en_us": "en_US",
    "en-gb": "en_GB", "en_gb": "en_GB",
    "en-uk": "en_GB",
    "en-au": "en_AU", "en_au": "en_AU",
    "en-ca": "en_CA", "en_ca": "en_CA",
    "en-nz": "en_AU", "en-ie": "en_GB", "en-za": "en_GB",
    # Spanish: Latin-American variants route to es_MX when available.
    "es":    "es_ES",
    "es-es": "es_ES", "es_es": "es_ES",
    "es-mx": "es_MX", "es_mx": "es_MX",
    "es-ar": "es_AR", "es_ar": "es_AR",
    "es-co": "es_MX", "es-cl": "es_MX", "es-pe": "es_MX",
    "es-uy": "es_AR", "es-ve": "es_MX", "es-419": "es_MX",
    # French: Belgian/Swiss/Canadian fall back to fr_FR (single LO dict).
    "fr":    "fr_FR",
    "fr-fr": "fr_FR", "fr_fr": "fr_FR",
    "fr-ca": "fr_FR", "fr-be": "fr_FR", "fr-ch": "fr_FR", "fr-lu": "fr_FR",
    # German: Austria + Switzerland have their own frami dictionaries.
    "de":    "de_DE",
    "de-de": "de_DE", "de_de": "de_DE",
    "de-at": "de_AT", "de_at": "de_AT",
    "de-ch": "de_CH", "de_ch": "de_CH", "de-li": "de_CH",
    # Italian
    "it":    "it_IT", "it-it": "it_IT", "it_it": "it_IT", "it-ch": "it_IT",
    # Portuguese: keep BR vs PT distinct (very different orthographies).
    "pt":    "pt_PT",
    "pt-pt": "pt_PT", "pt_pt": "pt_PT",
    "pt-br": "pt_BR", "pt_br": "pt_BR",
    "pt-ao": "pt_PT", "pt-mz": "pt_PT",
    # Dutch (Belgium falls back to NL since LO has no nl_BE dictionary).
    "nl":    "nl_NL", "nl-nl": "nl_NL", "nl_nl": "nl_NL", "nl-be": "nl_NL",
    # Slavic
    "pl":    "pl_PL", "pl-pl": "pl_PL", "pl_pl": "pl_PL",
    "ru":    "ru_RU", "ru-ru": "ru_RU", "ru_ru": "ru_RU",
    "cs":    "cs_CZ", "cs-cz": "cs_CZ", "cs_cz": "cs_CZ",
    "sk":    "sk_SK", "sk-sk": "sk_SK", "sk_sk": "sk_SK",
    "uk":    "uk_UA", "uk-ua": "uk_UA", "uk_ua": "uk_UA",
    "bg":    "bg_BG", "bg-bg": "bg_BG", "bg_bg": "bg_BG",
    "hr":    "hr_HR", "hr-hr": "hr_HR", "hr_hr": "hr_HR",
    "sl":    "sl_SI", "sl-si": "sl_SI", "sl_si": "sl_SI",
    "sr":    "sr",    "sr-rs": "sr",    "sr_rs": "sr",
    "sr-cyrl": "sr",  "sr-latn": "sr_Latn",
    # Nordic — Norwegian Bokmål is the default for `no`.
    "sv":    "sv_SE", "sv-se": "sv_SE", "sv_se": "sv_SE", "sv-fi": "sv_SE",
    "da":    "da_DK", "da-dk": "da_DK", "da_dk": "da_DK",
    "no":    "nb_NO",
    "nb":    "nb_NO", "nb-no": "nb_NO", "nb_no": "nb_NO",
    "nn":    "nn_NO", "nn-no": "nn_NO", "nn_no": "nn_NO",
    # Other European
    "hu":    "hu_HU", "hu-hu": "hu_HU", "hu_hu": "hu_HU",
    "el":    "el_GR", "el-gr": "el_GR", "el_gr": "el_GR", "gr": "el_GR",
    "ro":    "ro_RO", "ro-ro": "ro_RO", "ro_ro": "ro_RO", "ro-md": "ro_RO",
    "tr":    "tr_TR", "tr-tr": "tr_TR", "tr_tr": "tr_TR",
    "lt":    "lt_LT", "lt-lt": "lt_LT", "lt_lt": "lt_LT",
    "lv":    "lv_LV", "lv-lv": "lv_LV", "lv_lv": "lv_LV",
    "et":    "et_EE", "et-ee": "et_EE", "et_ee": "et_EE",
    # Semitic
    "he":    "he_IL", "he-il": "he_IL", "he_il": "he_IL", "iw": "he_IL",
    "ar":    "ar",    "ar-sa": "ar", "ar-eg": "ar", "ar-ae": "ar",
    "ar-ma": "ar", "ar-tn": "ar", "ar-jo": "ar", "ar-lb": "ar",
}


def normalize_lang_code(code: Optional[str]) -> Optional[str]:
    """Map an ISO/IETF code to a key in :data:`DICTIONARY_SOURCES`, or None.

    Resolution order (most specific first):
      1. Exact match in :data:`_LANG_NORMALISATION` (e.g. ``es-MX`` -> ``es_MX``).
      2. Exact match against an existing :data:`DICTIONARY_SOURCES` key
         (e.g. the file already labels its target as ``de_AT`` -> kept).
      3. Bare language head in :data:`_LANG_NORMALISATION`
         (e.g. ``zh-Hans`` -> ``zh`` -> not in map -> None).
    Returns ``None`` for genuinely unsupported languages so the caller can
    surface a "spell-check skipped" notice without crashing the QA run.
    """
    if not code:
        return None
    raw = code.strip().lower().replace("_", "-")
    if raw in _LANG_NORMALISATION:
        return _LANG_NORMALISATION[raw]
    # Allow files that already use a DICTIONARY_SOURCES key verbatim
    # (case-insensitive match against the canonical keys).
    raw_underscore = raw.replace("-", "_")
    for canonical in DICTIONARY_SOURCES:
        if canonical.lower() == raw_underscore:
            return canonical
    head = raw.split("-", 1)[0]
    return _LANG_NORMALISATION.get(head)


# -----------------------------------------------------------------------------
# On-disk locations
# -----------------------------------------------------------------------------

def bundled_dir() -> Path:
    """Directory where build_exe pre-downloads the 5 default dictionaries.

    Resolution depends on whether the app is running frozen (PyInstaller
    one-file .exe) or as plain source:

    * **Frozen** — PyInstaller's ``--add-data dictionaries;dictionaries``
      extracts the folder under ``sys._MEIPASS`` at startup. The previous
      ``Path(__file__).resolve().parent`` resolved to PyInstaller's
      bootloader temp dir for the *script* — NOT ``_MEIPASS`` — so the
      bundled dictionaries were never found and every language fell
      through to the lazy-download path (including IT, which was supposed
      to ship offline). We now prefer ``sys._MEIPASS`` (where ``--add-data``
      actually lands) and fall back to ``sys.executable.parent`` so users
      who keep a manual ``dictionaries/`` folder beside the .exe also
      work.
    * **Source / dev** — same as before: the ``dictionaries/`` folder
      sitting next to ``spellcheck.py``.
    """
    if getattr(sys, "frozen", False):
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            candidate = Path(meipass) / "dictionaries"
            if candidate.exists():
                return candidate
        # Fallback: dictionaries/ folder next to the .exe (user-managed).
        exe_dir = Path(sys.executable).resolve().parent / "dictionaries"
        if exe_dir.exists():
            return exe_dir
        # Last resort — return the _MEIPASS path even if missing so the
        # downstream candidate-check logs a sensible "not found" path.
        if meipass:
            return Path(meipass) / "dictionaries"
    return Path(__file__).resolve().parent / "dictionaries"


# Memoized cache-dir choice. Computed once per process by ``cache_dir()`` so
# the resolution cost (and the side-effect of creating the directory) only
# happens at the first call. ``reset_caches()`` clears it for tests.
# Sentinel ``_UNRESOLVED_CACHE_DIR`` distinguishes "not yet probed" from the
# legitimate "probed, every candidate failed" outcome (stored as ``None``)
# so we don't re-walk the filesystem on every QA segment.
class _Unset:
    pass


_UNRESOLVED_CACHE_DIR = _Unset()
_RESOLVED_CACHE_DIR: Any = _UNRESOLVED_CACHE_DIR  # Path | None | _Unset


def _cache_dir_candidates() -> List[Path]:
    """Ordered list of locations to try as the lazy-download cache.

    1. ``~/.cache/anonymizer/dictionaries/`` — standard XDG-ish location,
       writable by default for the logged-in user.
    2. ``<exe-parent>/cache/dictionaries/`` — useful for locked-down
       corporate machines where ``%USERPROFILE%\\.cache`` is read-only,
       and for portable installs where the user expects everything to
       live next to the .exe.
    3. ``%TEMP%/anonymizer-cache/dictionaries/`` — last resort. Always
       writable on a sane Windows session; data may be wiped between
       runs but at least spell-check works once.
    """
    import tempfile

    candidates: List[Path] = [Path.home() / ".cache" / "anonymizer" / "dictionaries"]
    if getattr(sys, "frozen", False):
        try:
            exe_parent = Path(sys.executable).resolve().parent
            candidates.append(exe_parent / "cache" / "dictionaries")
        except (OSError, ValueError):
            pass
    candidates.append(Path(tempfile.gettempdir()) / "anonymizer-cache" / "dictionaries")
    return candidates


def _is_writable(path: Path) -> bool:
    """Return True iff we can create ``path`` and write a temp file inside it."""
    try:
        path.mkdir(parents=True, exist_ok=True)
        probe = path / ".write_probe.tmp"
        with open(probe, "wb") as fh:
            fh.write(b"x")
        probe.unlink()
        return True
    except OSError:
        return False


def cache_dir() -> Optional[Path]:
    """Return the first writable cache directory from :func:`_cache_dir_candidates`.

    Returns ``None`` if every candidate is non-writable (e.g. very locked-down
    machine with no usable home, no exe-sibling write permission AND no
    accessible temp dir). Callers must handle the ``None`` case so the QA
    run keeps going even when there is nowhere to cache downloads.

    The choice is memoized for the lifetime of the process so we don't keep
    probing every directory on every QA segment.
    """
    global _RESOLVED_CACHE_DIR
    if not isinstance(_RESOLVED_CACHE_DIR, _Unset):
        return _RESOLVED_CACHE_DIR
    for candidate in _cache_dir_candidates():
        if _is_writable(candidate):
            _RESOLVED_CACHE_DIR = candidate
            return candidate
    _RESOLVED_CACHE_DIR = None
    return None


def _candidate_files(lang: str) -> Optional[Tuple[Path, Path, Path]]:
    """Return (lang_dir, aff_path, dic_path) for a known language, or None."""
    info = DICTIONARY_SOURCES.get(lang)
    if not info:
        return None
    bundled = bundled_dir() / lang
    cache_root = cache_dir()
    # Prefer cache (writable, may be newer); fall back to bundled (read-only-ish).
    bases: List[Path] = []
    if cache_root is not None:
        bases.append(cache_root / lang)
    bases.append(bundled)
    for base in bases:
        aff = base / info["aff"]
        dic = base / info["dic"]
        if aff.exists() and dic.exists():
            return base, aff, dic
    return None


# -----------------------------------------------------------------------------
# Network helpers
# -----------------------------------------------------------------------------

class DownloadError(Exception):
    pass


def _http_get(url: str, *, timeout: float = 8.0, retries: int = 2) -> bytes:
    """GET ``url`` with explicit timeout and bounded retries.

    Backoff is 1s then 2s. Raises :class:`DownloadError` if every attempt
    fails — never blocks indefinitely.
    """
    last_exc: Optional[Exception] = None
    for attempt in range(retries + 1):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "anonymizer-qa/1.0"})
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                if resp.status != 200:
                    raise DownloadError(f"HTTP {resp.status} for {url}")
                return resp.read()
        except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError, OSError) as exc:
            last_exc = exc
            if attempt < retries:
                time.sleep(2 ** attempt)  # 1s, 2s
    raise DownloadError(f"All {retries + 1} attempts failed for {url}: {last_exc}")


# -----------------------------------------------------------------------------
# Dictionary file validation
# -----------------------------------------------------------------------------

def _validate_dic(path: Path) -> bool:
    """Hunspell .dic must start (after optional BOM) with an integer word-count."""
    try:
        if path.stat().st_size == 0:
            return False
        with open(path, "rb") as fh:
            chunk = fh.read(64)
        if chunk.startswith(b"\xef\xbb\xbf"):
            chunk = chunk[3:]
        # Keep only the first non-empty line.
        for line in chunk.splitlines():
            line = line.strip()
            if not line:
                continue
            try:
                n = int(line.decode("latin-1", errors="replace"))
                return n > 0
            except ValueError:
                return False
        return False
    except OSError:
        return False


def _validate_aff(path: Path) -> bool:
    """Hunspell .aff must contain at least one SET or FLAG line in its first 2 KB."""
    try:
        if path.stat().st_size == 0:
            return False
        with open(path, "rb") as fh:
            head = fh.read(2048)
        return b"\nSET " in (b"\n" + head) or b"\nFLAG " in (b"\n" + head) or b"\nTRY " in (b"\n" + head)
    except OSError:
        return False


def _delete_pair_silently(*paths: Path) -> None:
    for p in paths:
        try:
            if p.exists():
                p.unlink()
        except OSError:
            pass


# -----------------------------------------------------------------------------
# Atomic download + validation
# -----------------------------------------------------------------------------

def _atomic_download(url: str, dest: Path, *, timeout: float = 8.0) -> None:
    """Download ``url`` to ``<dest>.tmp`` then atomically rename to ``dest``."""
    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp = dest.with_suffix(dest.suffix + ".tmp")
    data = _http_get(url, timeout=timeout)
    with open(tmp, "wb") as fh:
        fh.write(data)
    os.replace(str(tmp), str(dest))


def _download_pair(lang: str, target_dir: Path, *, timeout: float = 8.0) -> Tuple[Path, Path]:
    """Download both .aff and .dic for ``lang`` into ``target_dir``.

    On any failure: deletes both files (no half-pair) and re-raises DownloadError.
    """
    info = DICTIONARY_SOURCES[lang]
    aff_dest = target_dir / info["aff"]
    dic_dest = target_dir / info["dic"]
    try:
        _atomic_download(f"{_BASE_URL}/{info['folder']}/{info['aff']}", aff_dest, timeout=timeout)
        _atomic_download(f"{_BASE_URL}/{info['folder']}/{info['dic']}", dic_dest, timeout=timeout)
    except Exception:
        _delete_pair_silently(aff_dest, dic_dest,
                              aff_dest.with_suffix(aff_dest.suffix + ".tmp"),
                              dic_dest.with_suffix(dic_dest.suffix + ".tmp"))
        raise

    if not (_validate_aff(aff_dest) and _validate_dic(dic_dest)):
        _delete_pair_silently(aff_dest, dic_dest)
        raise DownloadError(
            f"Dictionary for '{lang}' downloaded but failed validation "
            f"(partial transfer?). Cleaned up; will retry on next QA run."
        )
    return aff_dest, dic_dest


# -----------------------------------------------------------------------------
# Build-time helper: pre-bundle the 5 defaults (called from build_exe.py)
# -----------------------------------------------------------------------------

def ensure_bundled_dictionaries(target_dir: Path, *, strict: bool = True,
                                 logger=print) -> Dict[str, str]:
    """Pre-download the 5 default dictionaries into ``target_dir``.

    Returns ``{lang: 'ok'|'cached'|'error: …'}`` per language. With
    ``strict=True`` (used by build_exe), raises :class:`DownloadError` on
    the first failure so a broken release never ships.
    """
    target_dir = Path(target_dir)
    target_dir.mkdir(parents=True, exist_ok=True)
    report: Dict[str, str] = {}
    for lang in BUNDLED_LANGUAGES:
        info = DICTIONARY_SOURCES[lang]
        lang_dir = target_dir / lang
        aff = lang_dir / info["aff"]
        dic = lang_dir / info["dic"]
        if aff.exists() and dic.exists() and _validate_aff(aff) and _validate_dic(dic):
            report[lang] = "cached"
            logger(f"  [+] {lang}: already present ({info['aff']}, {info['dic']})")
            continue
        try:
            _download_pair(lang, lang_dir, timeout=8.0)
            report[lang] = "ok"
            logger(f"  [+] {lang}: downloaded ({info['aff']}, {info['dic']})")
        except DownloadError as exc:
            report[lang] = f"error: {exc}"
            logger(f"  [x] {lang}: {exc}")
            if strict:
                raise
    return report


# -----------------------------------------------------------------------------
# Runtime: lazy load + cache spylls Dictionary objects
# -----------------------------------------------------------------------------

_DICT_CACHE: Dict[str, Any] = {}
_NEGATIVE_CACHE: Dict[str, str] = {}  # lang -> reason it's unavailable


def _try_import_spylls():
    try:
        from spylls.hunspell import Dictionary as _Dictionary  # type: ignore
        return _Dictionary
    except Exception as exc:  # pragma: no cover - defensive
        return None


def _load_dictionary_from_disk(lang: str):
    """Build a spylls Dictionary from the cache or bundled folder."""
    Dictionary = _try_import_spylls()
    if Dictionary is None:
        _NEGATIVE_CACHE[lang] = "spylls library not installed"
        return None
    candidate = _candidate_files(lang)
    if candidate is None:
        return None
    base, aff, dic = candidate
    # Validate files we found (cache may have been corrupted between runs).
    if not (_validate_aff(aff) and _validate_dic(dic)):
        # Surface the corruption explicitly: log a WARNING and stash the
        # reason in the negative cache so the user-visible "skipped" notice
        # mentions corruption if a re-download also fails.
        import logging as _logging
        _logging.getLogger(__name__).warning(
            "Corrupted spell-check dictionary for %r at %s — deleting and "
            "will attempt re-download (aff=%s, dic=%s)",
            lang, base, aff.name, dic.name,
        )
        _NEGATIVE_CACHE[lang] = (
            f"corrupted dictionary files removed from cache "
            f"({aff.name}, {dic.name})"
        )
        _delete_pair_silently(aff, dic)
        # Clear the negative cache entry so a fresh download is still
        # attempted; the entry only persists if that download also fails.
        _NEGATIVE_CACHE.pop(lang, None)
        return None
    # spylls.hunspell.Dictionary.from_files takes the path WITHOUT extension
    # of the .dic / .aff pair (which must share the same stem). Our cache
    # preserves the original LibreOffice filenames, which DO share a stem.
    info = DICTIONARY_SOURCES[lang]
    stem = info["aff"].rsplit(".", 1)[0]
    assert stem == info["dic"].rsplit(".", 1)[0], (
        f"DICTIONARY_SOURCES['{lang}']: aff/dic stems must match "
        f"({info['aff']!r} vs {info['dic']!r})"
    )
    try:
        return Dictionary.from_files(str(base / stem))
    except Exception as exc:  # pragma: no cover - defensive
        _NEGATIVE_CACHE[lang] = f"spylls failed to load: {exc}"
        return None


def get_dictionary(lang: Optional[str]) -> Optional[Any]:
    """Return a loaded spylls Dictionary for ``lang`` or None.

    Resolution order:
      1. In-memory cache.
      2. Negative cache (don't retry within the same session if it failed).
      3. On-disk cache (~/.cache/anonymizer/dictionaries/<lang>/).
      4. Bundled dir (offline_package/dictionaries/<lang>/, shipped with .exe).
      5. Lazy download from LibreOffice into the cache.
    """
    if not lang:
        return None
    if lang in _DICT_CACHE:
        return _DICT_CACHE[lang]
    if lang in _NEGATIVE_CACHE:
        return None
    if lang not in DICTIONARY_SOURCES:
        _NEGATIVE_CACHE[lang] = f"unsupported language '{lang}'"
        return None

    d = _load_dictionary_from_disk(lang)
    if d is not None:
        _DICT_CACHE[lang] = d
        return d

    # If the bundled load failed with a DEFINITIVE reason (spylls
    # library missing, spylls rejected the bundled .aff/.dic, etc.)
    # don't bother attempting a lazy download — it would just hit the
    # same parse failure AND overwrite the real reason in the notice.
    # The corruption-recovery branch in _load_dictionary_from_disk
    # explicitly pops _NEGATIVE_CACHE first so it does fall through
    # here and re-download as recovery. Any other case where the disk
    # load failed cleanly (returning None without setting a reason —
    # i.e. the bundled files don't exist) also falls through, since
    # no entry is present.
    if lang in _NEGATIVE_CACHE:
        return None

    # Lazy download into the user cache. If no cache dir is writable
    # (locked-down corporate machine, read-only home AND no temp access),
    # skip the download cleanly instead of crashing the QA run.
    cache_root = cache_dir()
    if cache_root is None:
        _NEGATIVE_CACHE[lang] = (
            "no writable cache directory (checked ~/.cache, exe-sibling "
            "cache/, and the system temp dir)"
        )
        return None
    try:
        _download_pair(lang, cache_root / lang, timeout=8.0)
    except DownloadError as exc:
        _NEGATIVE_CACHE[lang] = str(exc)
        return None
    except OSError as exc:
        # PermissionError, disk-full, antivirus-quarantine etc. surface here
        # because _atomic_download writes to disk. Capture the real reason
        # so the QA-tab notice tells the user what went wrong.
        _NEGATIVE_CACHE[lang] = f"cache write failed: {exc}"
        return None

    d = _load_dictionary_from_disk(lang)
    if d is None:
        # Preserve the specific spylls exception captured by
        # _load_dictionary_from_disk (e.g. "spylls failed to load: …")
        # instead of overwriting it with a generic placeholder. The
        # user-visible "skipped" notice and any Windows-only bug report
        # then carries the real cause (encoding, parse error, etc.).
        if lang not in _NEGATIVE_CACHE:
            _NEGATIVE_CACHE[lang] = "loaded but spylls rejected"
        return None
    _DICT_CACHE[lang] = d
    return d


def get_negative_reason(lang: str) -> Optional[str]:
    return _NEGATIVE_CACHE.get(lang)


def reset_caches() -> None:
    """Test helper: clear in-memory dictionary + negative caches."""
    global _RESOLVED_CACHE_DIR
    _DICT_CACHE.clear()
    _NEGATIVE_CACHE.clear()
    _RESOLVED_CACHE_DIR = _UNRESOLVED_CACHE_DIR


# -----------------------------------------------------------------------------
# Tokenisation + spell-check
# -----------------------------------------------------------------------------

# Letter runs (with internal apostrophe / hyphen) — matches "don't", "well-known",
# "señor", "côté". Excludes pure numbers, single letters, punctuation.
_TOKEN_RE = re.compile(
    r"\b[^\W\d_]+(?:['\u2019\-][^\W\d_]+)*\b",
    re.UNICODE,
)


def _tokenize_for_spellcheck(text: str) -> List[str]:
    if not text:
        return []
    return [t for t in _TOKEN_RE.findall(text) if len(t) > 1]


def spell_check_text(text: str, dictionary: Any,
                     ignore_words: Optional[set] = None) -> List[str]:
    """Return a deduplicated, ordered list of misspelled tokens in ``text``.

    Uses the dictionary's case-insensitive lookup heuristic: a token is
    considered correct if either the original or its lowercased form passes.
    Single-character tokens, pure numbers, and tokens in ``ignore_words``
    (case-insensitive) are skipped.
    """
    if not text or dictionary is None:
        return []
    ignore = {w.lower() for w in (ignore_words or set())}
    seen: set = set()
    out: List[str] = []
    for tok in _tokenize_for_spellcheck(text):
        key = tok.lower()
        if key in seen or key in ignore:
            continue
        seen.add(key)
        try:
            if dictionary.lookup(tok) or dictionary.lookup(tok.lower()):
                continue
        except Exception:
            # Defensive: spylls can raise on weird inputs; treat as correct.
            continue
        out.append(tok)
    return out

#!/usr/bin/env python3
"""
Excel CIQ -> TL1 generator

- Reads an Excel workbook containing one or more sheets/sections.
- Finds sections by cells that contain: "CLI Command: xxxx".
- Builds TL1 lines by matching Parameters -> CLI Name -> Target Value.
- Supports horizontal target tables and vertical target tables.
- Supports "==> Refer to <sheet> sheet" when target values are in another sheet.
- Lets you process the whole workbook, one sheet, or selected sections.
- Prompts for input workbook, output folder, and output .tl1 filename.

Required:
    pip install openpyxl
"""

from __future__ import annotations

import os
import re
import sys
from datetime import datetime
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

try:
    from openpyxl import load_workbook
except ImportError:
    print("ERROR: openpyxl is required. Install it with: pip install openpyxl")
    sys.exit(1)


# -----------------------------
# Configuration
# -----------------------------
SCRIPT_VERSION = "2.108.0"
VERSION_TAG = "v" + SCRIPT_VERSION.replace(".", "_")
MAX_TL1_LINES_PER_FILE = 10000
LARGE_SHEET_ROW_THRESHOLD = 100000
DEFAULT_CLI_SCAN_MAX_ROWS = 20000
DEFAULT_TARGET_MARKER_SCAN_MAX_ROWS = 2000
MAX_BLANK_ROWS_IN_TARGET_TABLE = 25

SKIP_VALUES = {
    "[tbd]", "[st-input]", "[telus-input]", "[enc]", "tbd", "skip", "n/a"
}
STATUS_VALUES = {"✓", "◌", "🗙", "↞", "⯑"}
TARGET_RE = re.compile(r"\btarget\s+values?\b", re.I)
CURRENT_RE = re.compile(r"\bcurrent\s+values?\b", re.I)
CLI_RE = re.compile(r"cli\s*command\s*:\s*(.+)", re.I)
REFER_RE = re.compile(r"refer\s+to\s+(.+?)\s+sheet", re.I)


@dataclass
class Section:
    sheet_name: str
    title: str
    cli_row: int
    cli_col: int
    raw_cli: str
    command: str
    cli_order: List[str]
    param_to_cli: Dict[str, str]


# -----------------------------
# General helpers
# -----------------------------
def as_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).replace("\xa0", " ").strip()


def norm_key(value: Any) -> str:
    return re.sub(r"\s+", " ", as_text(value)).strip().lower()


def clean_cli_name(value: Any) -> str:
    s = as_text(value)
    s = s.replace("[", "").replace("]", "")
    s = s.strip().strip(",;:")
    return s.upper()


def clean_param(value: Any) -> str:
    """Normalize YANG/CIQ parameter paths for reliable matching.

    CIQ templates sometimes wrap long child-list paths across lines, for example:
        amfmme-ne-grp-elem-entry/
        amfmme-ne-grp-ne-id
    while the Target Values header has the same path on one line.  Treat both as
    the same key so child parameters such as NEID/CAPA are not skipped.
    """
    s = as_text(value).strip()
    s = re.sub(r"\s*/\s*", "/", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip()



def is_placeholder_param(value: Any) -> bool:
    """Return True when the Parameter/YANG cell is only a placeholder.

    CIQ sections sometimes use "-", "--", "N/A", or "None" instead of a real
    parameter name. Such values must not be used as mapping keys because several
    rows can share the same placeholder and overwrite each other.
    """
    text = clean_param(value).strip()
    normalized = text.lower()
    if not text:
        return True
    if re.fullmatch(r"-+", text):
        return True
    return normalized in {"n/a", "na", "none", "null", "[none]"}


def infer_cli_from_param(param: str, cli_order: Sequence[str]) -> Optional[str]:
    """Infer CLI name from parameter/header text when it matches the CLI list.

    This protects against CIQ rows where CLI Name is shifted/swapped and target
    headers use direct CLI names or YANG-style names.
    """
    if not param or not cli_order:
        return None

    cli_set = {x.upper() for x in cli_order}
    p_clean = clean_param(param)
    p_cli = clean_cli_name(p_clean)
    p_low = p_clean.lower()

    # Direct CLI header such as [RTALIDX] or RTALIDX.
    if p_cli in cli_set:
        return p_cli

    explicit_aliases = {
        "call-prof-index": "INDEX",
        "call-prof-vrfsp": "VRFSP",
        "call-prof-eplmn": "EPLMN",
        "call-prof-forbta": "FORBTA",
        "call-prof-forbla": "FORBLA",
        "call-prof-rqos": "RQOS",
        "call-prof-ims": "IMS",
        "call-prof-emg": "EMG",
        "call-prof-dcn_mmeg": "DCN_MMEG",
        "call-prof-nepsidx": "NEPSIDX",
        "call-prof-dcnr": "DCNR",
        "call-prof-srur": "SRUR",
        "call-prof-override": "OVERRIDE",
        "call-prof-usage": "USAGE",
        "call-prof-vplmn-addr": "VPLMN_ADDR",
        "call-prof-feat2": "FEAT2",
        "rtladix": "RTALIDX",
        "[rtalidx]": "RTALIDX",
        "call-prof-sgwrestoration": "SGWRST",
        "call-prof-nrrestriction": "NRRSTR",
        "call-prof-defaultapnindex": "DAPNIDX",
        "call-prof-ratrestirctioncause": "RATRSTRC",
        "call-prof-ratrestrictioncause": "RATRSTRC",
        "call-prof-rat-rstrc": "RATRSTRC",
        "call-prof-ratrstrc": "RATRSTRC",
        "call-prof-primarylaplmn": "PRILAPLMN",
        "call-prof-secondarylaplmn": "SNDLAPLMN",
        "call-prof-prilaplmn": "PRILAPLMN",
        "call-prof-sndlaplmn": "SNDLAPLMN",
        "call-prof-ltem": "LTEM",
    }
    if p_low in explicit_aliases and explicit_aliases[p_low] in cli_set:
        return explicit_aliases[p_low]

    suffix = p_clean.split("/")[-1].split("-")[-1].replace("-", "_").upper()
    if suffix in cli_set:
        return suffix

    p_tail = p_low.split("/")[-1]
    suffix_aliases = {
        "mtc-t3412-home": "MTC_T3412_HOME",
        "mtc-t3412-roam": "MTC_T3412_ROAM",
        "esr-ps": "ESR_PS",
        "psm-supp": "PSM_SUPP",
        "psm-t3324-pri": "PSM_T3324_PRI",
        "psm-t3412-pri": "PSM_T3412_PRI",
        "edrx-supp": "EDRX_SUPP",
        "edrx-ptw": "EDRX_PTW",
        "edrx-cycle": "EDRX_CYCLE",
        "edrx-ptw-nb": "EDRX_PTW_NB",
        "edrx-cycle-nb": "EDRX_CYCLE_NB",
        "edrx-ptw-pri": "EDRX_PTW_PRI",
        "edrx-cycle-pri": "EDRX_CYCLE_PRI",
        "dl-buff-use": "DL_BUFF_USE",
        "dl-infinite-buff": "DL_INFINITE_BUFF",
        "dl-buff-dur-offset": "DL_BUFF_DUR_OFFSET",
        "dl-buff-cnt": "DL_BUFF_CNT",
        "dl-buff-indir-use": "DL_BUFF_INDIR_USE",
        "ciot-cp": "CIOT_CP",
        "ciot-up": "CIOT_UP",
        "ciot-wopdn": "CIOT_WOPDN",
        "ciot-release-act": "CIOT_RELEASE_ACT",
        "ciot-release-time": "CIOT_RELEASE_TIME",
        "ciot-multi-tdr-rrt": "CIOT_MULTI_TDR_RRT",
        "ciot-rds-supp": "CIOT_RDS_SUPP",
        "sgwrestoration": "SGWRST",
        "nrrestriction": "NRRSTR",
        "defaultapnindex": "DAPNIDX",
        "ratrestirctioncause": "RATRSTRC",
        "ratrestrictioncause": "RATRSTRC",
        "primarylaplmn": "PRILAPLMN",
        "secondarylaplmn": "SNDLAPLMN",
        "vplmn-addr": "VPLMN_ADDR",
        "vplmn_addr": "VPLMN_ADDR",
        "dcn_mmeg": "DCN_MMEG",
        "nepsidx": "NEPSIDX",
        "ratrstrc": "RATRSTRC",
        "prilaplmn": "PRILAPLMN",
        "sndlaplmn": "SNDLAPLMN",
    }
    for key, cli in suffix_aliases.items():
        if p_tail.endswith(key) and cli in cli_set:
            return cli

    return None


def is_blank(value: Any) -> bool:
    return value is None or as_text(value) == ""


def is_skip_value(value: Any) -> bool:
    s = norm_key(value)
    return s in SKIP_VALUES or s.startswith("skip")


def is_status(value: Any) -> bool:
    return as_text(value) in STATUS_VALUES


def is_arrow_value(value: Any) -> bool:
    """Return True for CIQ ditto/carry-forward markers."""
    text = as_text(value).strip()
    return text in {"↪", "⇢", "→", "-->", "--->", "->", "=>"}


def normalize_ne_type(value: str) -> str:
    """Normalize NE-GRP NE type aliases required by TL1 output."""
    text = (value or "").upper()
    if text == "EIR-MME":
        return "DEIR"
    if text == "GMLC-MME":
        return "DGMLC"
    return text


def quote_if_needed(value: str) -> str:
    """Wrap TL1 values containing whitespace in double quotes.

    Example: NAME=Lab TOR UDM -> NAME="Lab TOR UDM".
    Existing quotes are preserved. Blank values are not quoted.
    """
    if value == "":
        return value
    if len(value) >= 2 and value[0] == '"' and value[-1] == '"':
        return value
    if any(ch.isspace() for ch in value):
        return f'"{value}"'
    return value


def normalize_rtalidx_value(value: str) -> str:
    """Normalize RTALIDX list values for MMECALL-PROF and similar commands.

    Examples:
        1 | 2       -> 1&2
        1|2         -> 1&2
        1 / 2       -> 1&2
        1, 2        -> 1&2
        [1, 2]      -> 1&2
        1 and 2     -> 1&2

    Single values are returned unchanged apart from whitespace cleanup.
    """
    text = (value or "").strip()
    if not text:
        return ""

    # Remove optional wrappers while preserving the content.
    text = text.strip().strip("[]{}()").strip()

    # Normalize common human/table separators to ampersand.
    text = re.sub(r"\s+(?:and|or)\s+", "&", text, flags=re.I)
    text = re.sub(r"\s*[|,/;]\s*", "&", text)
    text = re.sub(r"\s*&\s*", "&", text)

    # Remove extra whitespace around pure numeric tokens; keep non-numeric
    # tokens intact except for outer spacing.
    parts = [part.strip() for part in text.split("&") if part.strip()]
    if len(parts) <= 1:
        single = parts[0] if parts else ""
        if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_-]*", single or ""):
            return single.upper()
        return single

    return "&".join(parts)



def looks_like_formula_text(value: Any) -> bool:
    return isinstance(value, str) and value.startswith("=")


def clean_cell_value(value: Any) -> Optional[str]:
    """Return the base TL1 value text, or None when the value should be skipped/omitted."""
    if value is None:
        return ""   # keep blank as CLI_NAME=

    if is_status(value):
        return None

    if isinstance(value, float) and value.is_integer():
        text = str(int(value))
    else:
        text = as_text(value)

    if not text:
        return ""

    # Convert [default = none] -> none, [default = nspt] -> nspt first.
    m = re.search(r"\[\s*default\s*=\s*([^\]]+)\]", text, re.I)
    if m:
        text = m.group(1).strip()

    if is_skip_value(text):
        return None

    # Remove surrounding quotes only; field-specific formatting is applied later.
    text = text.strip().strip('"').strip("'")
    return text


def format_tl1_value(cli_name: str, value: str, command: str = "") -> str:
    """Apply TL1 output case rules for a specific CLI parameter.

    Most TL1 values are uppercased, but values assigned to POD must be kept
    lowercase because Kubernetes/OpenShift pod/service interface names are
    case-sensitive in the generated command.
    """
    if value == "":
        return ""

    cli = (cli_name or "").strip().upper()
    cmd = (command or "").strip().upper()

    # CHG-FD-LMT value normalization.
    # The service/container name "istio-proxy" is case-sensitive and must remain
    # lowercase instead of being converted to ISTIO-PROXY.
    if cmd == "CHG-FD-LMT" and value.strip().lower() == "istio-proxy":
        return "istio-proxy"

    # CHG-DIAG-INF MSG normalization.
    # These source values are YANG-style strings but TL1 expects short codes:
    #   status-inhibit -> INH
    #   status-allow   -> ALW
    if cmd == "CHG-DIAG-INF" and cli == "MSG":
        normalized = value.strip().lower()
        if normalized == "status-inhibit":
            return "INH"
        if normalized == "status-allow":
            return "ALW"

    # RTALIDX can be entered as "1 | 2", "1|2", "1 / 2", "1, 2", etc.
    # TL1 expects ampersand-separated values, especially for CHG-MMECALL-PROF:
    #   RTALIDX=1&2
    if cli == "RTALIDX":
        return quote_if_needed(normalize_rtalidx_value(value))

    # Keep POD values lowercase globally.
    # Example: POD=SVC-MME-ANINTF-SCTP-0 -> POD=svc-mme-anintf-sctp-0
    if cli == "POD":
        return quote_if_needed(value.lower())

    # Keep prior important case-sensitive fields stable when this simplified
    # script is used directly.
    preserve_case = {"NFID", "FQDN", "SVCID", "INSTID", "PGWFQDN", "DFTDNN", "APN"}
    lower_case = {"SVC", "HOST", "REALM"}
    if cli in preserve_case:
        return quote_if_needed(value)
    if cli in lower_case:
        return quote_if_needed(value.lower())
    if cmd == "CRTE-SVC-VER" and cli == "NF":
        return quote_if_needed(value.lower())

    text = value.upper()
    if cli == "TYPE" and text == "MEMORY":
        text = "MEM"
    return quote_if_needed(text)


def iter_used_cells(ws) -> Iterable[Tuple[int, int, Any]]:
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                yield cell.row, cell.column, cell.value


def get_effective_value(ws_values, row: int, col: int) -> Any:
    return ws_values.cell(row=row, column=col).value


def row_values(ws, row: int, start_col: int = 1, end_col: Optional[int] = None) -> List[Any]:
    if end_col is None:
        end_col = ws.max_column
    return [ws.cell(row=row, column=c).value for c in range(start_col, end_col + 1)]


# -----------------------------
# CLI and section parsing
# -----------------------------

def extract_fixed_cli_values(raw_cli_text: str) -> Dict[str, str]:
    """Extract fixed KEY=VALUE assignments from a CLI Command template."""
    m = CLI_RE.search(raw_cli_text)
    cli = m.group(1).strip() if m else raw_cli_text.strip()

    fixed: Dict[str, str] = {}
    command_part, sep, suffix = cli.partition(";")

    if ":" in command_part:
        command_name, rest = command_part.split(":", 1)
        for part in rest.split(","):
            token = part.strip().strip("[]")
            if "=" not in token:
                continue
            key, value = token.split("=", 1)
            key = clean_cli_name(key)
            value = as_text(value).strip().strip('"').strip("'")
            if key and value:
                fixed[key] = format_tl1_value(key, value, command_name.strip())

    if sep and suffix:
        command_name = command_part.split(":", 1)[0].strip()
        for key, value in re.findall(
            r"\b([A-Za-z][A-Za-z0-9_]*)\s*=\s*([A-Za-z0-9_.\-']+)",
            suffix,
        ):
            cli_key = clean_cli_name(key)
            if cli_key:
                fixed[cli_key] = format_tl1_value(cli_key, value, command_name)

    return fixed


def parse_cli(raw_cli_text: str) -> Tuple[str, List[str]]:
    """Parse a CLI template into command name and ordered CLI field names.

    Fixed assignments such as TYPE=PAGE are represented in cli_order as TYPE.
    """
    m = CLI_RE.search(raw_cli_text)
    cli = m.group(1).strip() if m else raw_cli_text.strip()
    cli = cli.rstrip()

    command_part = cli.split(";", 1)[0].strip()
    if ":" not in command_part:
        return command_part.rstrip(";").upper(), []

    command, rest = command_part.split(":", 1)
    parts: List[str] = []
    for part in rest.split(","):
        token = part.strip().strip("[]")
        if not token:
            continue
        if "=" in token:
            token = token.split("=", 1)[0].strip()
        name = clean_cli_name(token)
        if name and name not in parts:
            parts.append(name)

    return command.strip().upper(), parts

def find_section_title(ws_values, cli_row: int) -> str:
    """Find nearest non-empty title above CLI command. Preferred title is a cell with '(...)'."""
    best = ""
    for r in range(cli_row - 1, max(0, cli_row - 12), -1):
        for c in range(1, ws_values.max_column + 1):
            text = as_text(ws_values.cell(r, c).value)
            if not text:
                continue
            low = text.lower()
            if any(x in low for x in ["yang path", "review status", "config status"]):
                continue
            if "(" in text and ")" in text:
                return text
            if not best and len(text) > 2:
                best = text
    return best or ws_values.title



def is_out_of_scope_section(ws_values, section: Section, next_cli_row: int) -> bool:
    """Return True when THIS CLI section is marked Out of Scope.

    Review Status / Config Status are located above the section's CLI Command.
    Scan only the current section header area ending at section.cli_row.  Do not
    scan forward toward the next CLI command because the next section's header
    may start several rows before its CLI row and could otherwise cause a false
    Out-of-Scope match for the current section.
    """
    start_row = max(1, section.cli_row - 15)
    end_row = section.cli_row

    for r in range(start_row, end_row + 1):
        row_values_text = [
            as_text(ws_values.cell(r, c).value)
            for c in range(1, ws_values.max_column + 1)
        ]
        row_norm = [re.sub(r"\s+", " ", v).strip().lower() for v in row_values_text if v]

        if not row_norm:
            continue

        has_status_label = any(
            v in {"review status", "config status"} or
            v.startswith("review status") or
            v.startswith("config status")
            for v in row_norm
        )
        has_out_of_scope = any(
            v == "out of scope" or "out of scope" in v
            for v in row_norm
        )

        if has_status_label and has_out_of_scope:
            return True

    return False

def find_parameter_mapping(ws_values, cli_row: int, cli_col: int, next_cli_row: int) -> Dict[str, str]:
    """Find Parameters table after CLI row and return parameter leaf -> CLI name."""
    max_row = min(ws_values.max_row, next_cli_row - 1 if next_cli_row else ws_values.max_row)
    header_row = None

    for r in range(cli_row, max_row + 1):
        texts = [norm_key(ws_values.cell(r, c).value) for c in range(1, ws_values.max_column + 1)]
        if any(t in {"parameters", "paramaters"} for t in texts):
            header_row = r
            break

    if header_row is None:
        return {}

    # Locate PARAMATERS and CLI Name columns.
    param_col = None
    cli_name_col = None
    for c in range(1, ws_values.max_column + 1):
        text = norm_key(ws_values.cell(header_row, c).value)
        if text in {"parameters", "paramaters"}:
            param_col = c
        elif text == "cli name":
            cli_name_col = c

    if param_col is None or cli_name_col is None:
        return {}

    mapping: Dict[str, str] = {}
    for r in range(header_row + 1, max_row + 1):
        # Stop at first Target Values block, next section title, or next CLI command.
        row_text = " | ".join(as_text(ws_values.cell(r, c).value) for c in range(1, ws_values.max_column + 1))
        if TARGET_RE.search(row_text) or CLI_RE.search(row_text):
            break

        param = clean_param(ws_values.cell(r, param_col).value)
        cli_name = clean_cli_name(ws_values.cell(r, cli_name_col).value)

        # Placeholder Parameter values such as "-" are not real parameter
        # names. Do not use them as mapping keys or later rows will overwrite
        # earlier rows that use the same placeholder.
        if is_placeholder_param(param):
            param = ""

        # Normal mapping: YANG/CIQ parameter path -> CLI Name.
        if param and cli_name:
            mapping[param] = cli_name

        # Fallback mapping: CLI Name -> CLI Name.
        # Some sheets, including MMERTAL-INF, keep the Parameter/YANG path cell
        # blank and use CLI headers directly in the Target Values table:
        #   RTALIDX | CAUSE | MCC | MNC | TAC
        # Without this alias the parser cannot match the target headers and
        # emits "No TL1 rows generated".
        if cli_name:
            mapping.setdefault(cli_name, cli_name)

    return mapping


def discover_sections(wb_values, sheets_to_scan: Sequence[str]) -> List[Section]:
    sections: List[Section] = []
    for sheet_name in sheets_to_scan:
        ws = wb_values[sheet_name]
        cli_cells = []

        # Large data sheets can contain hundreds of thousands of rows. CLI Command
        # sections are normally in the CIQ section/header area, so scan only the
        # first N rows by default to avoid a full 400k+ row pass on data-only sheets.
        # Set TL1_CLI_SCAN_MAX_ROWS=0 to force a full-sheet CLI scan.
        scan_limit_env = os.environ.get("TL1_CLI_SCAN_MAX_ROWS", str(DEFAULT_CLI_SCAN_MAX_ROWS)).strip()
        try:
            scan_limit = int(scan_limit_env)
        except ValueError:
            scan_limit = DEFAULT_CLI_SCAN_MAX_ROWS
        max_scan_row = ws.max_row if scan_limit <= 0 else min(ws.max_row, scan_limit)
        if ws.max_row > max_scan_row:
            print(f"Scanning first {max_scan_row:,} rows for CLI sections on large sheet '{sheet_name}' "
                  f"({ws.max_row:,} total rows). Set TL1_CLI_SCAN_MAX_ROWS=0 to scan all rows.")

        for row in ws.iter_rows(min_row=1, max_row=max_scan_row):
            for cell in row:
                if cell.value is None:
                    continue
                text = as_text(cell.value)
                if CLI_RE.search(text):
                    cli_cells.append((cell.row, cell.column, text))
        cli_cells.sort()

        for i, (r, c, raw) in enumerate(cli_cells):
            next_cli_row = cli_cells[i + 1][0] if i + 1 < len(cli_cells) else ws.max_row + 1
            command, cli_order = parse_cli(raw)
            title = find_section_title(ws, r)
            mapping = find_parameter_mapping(ws, r, c, next_cli_row)
            # Keep the explicit CLI Name column as the source of truth.
            # Earlier versions inferred CLI names from the YANG leaf suffix for every row,
            # but that broke valid child-list mappings such as:
            #   smf-inf-smf-svc-entry/smf-inf-smf-svc-name -> SVC
            # because the suffix "name" was incorrectly remapped to NAME.
            # Only infer when the explicit CLI name is missing or not part of the command.
            for _param, _cli in list(mapping.items()):
                if _cli and _cli in cli_order:
                    continue
                _inferred = infer_cli_from_param(_param, cli_order)
                if _inferred:
                    mapping[_param] = _inferred

            # Add direct aliases for every CLI placeholder. This makes Target
            # Values headers like INDEX, [RTALIDX], or SGWRST parse even when
            # the Parameter/YANG-path column is blank or shifted.
            for _cli in cli_order:
                mapping.setdefault(_cli.upper(), _cli.upper())
                mapping.setdefault(f"[{_cli.upper()}]", _cli.upper())
            sections.append(
                Section(
                    sheet_name=sheet_name,
                    title=title,
                    cli_row=r,
                    cli_col=c,
                    raw_cli=raw,
                    command=command,
                    cli_order=cli_order,
                    param_to_cli=mapping,
                )
            )
    return sections




def find_parameter_table_end(ws_values, section: Section, next_cli_row: int) -> int:
    """Return the last row in the Parameters table for this section.

    The Parameters header row may itself contain a "Target Value" helper column,
    so we must first find the PARAMETERS/PARAMATERS row, then stop at the next
    true Target Values block below the parameter rows.
    """
    known_params = set(section.param_to_cli.keys())
    if not known_params:
        return section.cli_row

    max_row = min(ws_values.max_row, next_cli_row - 1)
    header_row = None
    for r in range(section.cli_row + 1, max_row + 1):
        row_texts = [norm_key(ws_values.cell(r, c).value) for c in range(1, ws_values.max_column + 1)]
        if any(t in {"parameters", "paramaters"} for t in row_texts):
            header_row = r
            break
    if header_row is None:
        return section.cli_row

    last = header_row
    for r in range(header_row + 1, max_row + 1):
        row_join = " | ".join(as_text(ws_values.cell(r, c).value) for c in range(1, ws_values.max_column + 1))
        if TARGET_RE.search(row_join):
            break
        for c in range(1, ws_values.max_column + 1):
            if clean_param(ws_values.cell(r, c).value) in known_params:
                last = r
                break
    return last

# -----------------------------
# Target table discovery and parsing
# -----------------------------
def find_target_markers(ws_values, start_row: int = 1, end_row: Optional[int] = None) -> List[Tuple[int, int]]:
    """Find Target Values markers without scanning huge data sheets end-to-end.

    When end_row is not supplied, this is usually a referred data sheet.  In CIQ
    workbooks the Target Values marker is normally near the top, while data may
    extend hundreds of thousands of rows.  To avoid scanning 400k rows just to
    locate the marker, the default search is capped by TL1_TARGET_MARKER_SCAN_MAX_ROWS.
    Set TL1_TARGET_MARKER_SCAN_MAX_ROWS=0 to force a full scan.
    """
    if end_row is None:
        scan_limit_env = os.environ.get("TL1_TARGET_MARKER_SCAN_MAX_ROWS", str(DEFAULT_TARGET_MARKER_SCAN_MAX_ROWS)).strip()
        try:
            scan_limit = int(scan_limit_env)
        except ValueError:
            scan_limit = DEFAULT_TARGET_MARKER_SCAN_MAX_ROWS
        if scan_limit <= 0:
            end_row = ws_values.max_row
        else:
            end_row = min(ws_values.max_row, start_row + scan_limit - 1)
            if ws_values.max_row > end_row:
                print(
                    f"Scanning first {end_row:,} rows for Target Values on large/referred sheet "
                    f"'{ws_values.title}' ({ws_values.max_row:,} total rows). "
                    "Set TL1_TARGET_MARKER_SCAN_MAX_ROWS=0 to scan all rows."
                )

    markers = []
    max_r = min(end_row, ws_values.max_row)
    for row in ws_values.iter_rows(min_row=start_row, max_row=max_r):
        for cell in row:
            if cell.value is None:
                continue
            text = as_text(cell.value)
            if TARGET_RE.search(text):
                markers.append((cell.row, cell.column))
    return markers

def find_referred_sheet_name(ws_values, start_row: int, end_row: Optional[int] = None) -> Optional[str]:
    if end_row is None:
        end_row = ws_values.max_row
    for r in range(start_row, min(end_row, ws_values.max_row) + 1):
        for c in range(1, ws_values.max_column + 1):
            text = as_text(ws_values.cell(r, c).value)
            m = REFER_RE.search(text)
            if m:
                return m.group(1).strip()
    return None


def resolve_referred_sheet(wb_values, referred: str) -> Optional[str]:
    if not referred:
        return None
    referred_norm = norm_key(referred)
    for name in wb_values.sheetnames:
        if norm_key(name) == referred_norm:
            return name
    # Loose match: ignore punctuation/spaces.
    compact = re.sub(r"[^a-z0-9]", "", referred_norm)
    for name in wb_values.sheetnames:
        if re.sub(r"[^a-z0-9]", "", norm_key(name)) == compact:
            return name
    return None


def _param_match(value: Any, known_params: set[str]) -> Optional[str]:
    text = clean_param(value)
    if text in known_params:
        return text

    # Accept direct CLI-name headers in target tables, including optional brackets.
    # Example: [RTALIDX] should match the RTALIDX -> RTALIDX alias in known_params.
    cli_text = clean_cli_name(text)
    if cli_text in known_params:
        return cli_text

    # Accept case-insensitive parameter/header match.
    known_lower = {k.lower(): k for k in known_params}
    if text.lower() in known_lower:
        return known_lower[text.lower()]
    if cli_text.lower() in known_lower:
        return known_lower[cli_text.lower()]

    return None


def _region_end_row(ws_values, marker_row: int, next_section_row: Optional[int] = None) -> int:
    limit = min(ws_values.max_row, (next_section_row - 1) if next_section_row else ws_values.max_row)
    return limit




def target_column_bounds(ws_values, marker_row: int, marker_col: int) -> Tuple[int, int]:
    """Return inclusive target value column bounds, stopping before Current Values."""
    end_col = ws_values.max_column
    for c in range(marker_col + 1, ws_values.max_column + 1):
        if CURRENT_RE.search(as_text(ws_values.cell(marker_row, c).value)):
            end_col = c - 1
            break
    return marker_col, end_col


def normalize_row_action(value: Any) -> Optional[str]:
    """Return a normalized row action marker such as CHG or CRTE.

    CIQ target tables sometimes have an extra leading column whose cell contains
    CHG on rows that must use a change command instead of the section's normal
    create command. The action column may have a blank/unmapped header.
    """
    text = as_text(value).strip().upper()
    if text in {"CHG", "CHANGE"}:
        return "CHG"
    if text in {"CRTE", "CREATE"}:
        return "CRTE"
    if text in {"REG", "REGISTER"}:
        return "REG"
    return None


def derive_command_for_row(base_command: str, action: Optional[str]) -> str:
    """Derive the TL1 command prefix for a row action marker.

    Examples:
        CRTE-FOO + CHG  -> CHG-FOO
        CRTE-FOO + CRTE -> CRTE-FOO
        REG-FOO  + CHG  -> CHG-FOO

    If no supported action marker is present, return the section command.
    """
    cmd = (base_command or "").strip().upper()
    if not action:
        return cmd

    action = action.upper()
    if "-" in cmd:
        _prefix, suffix = cmd.split("-", 1)
    else:
        suffix = cmd

    if action == "CHG":
        return f"CHG-{suffix}"
    if action == "CRTE":
        return f"CRTE-{suffix}"
    if action == "REG":
        return f"REG-{suffix}"
    return cmd


def detect_row_action_near_target_columns(
    ws_values,
    row: int,
    param_cols: Dict[int, str],
    marker_col: int,
) -> Optional[str]:
    """Detect CHG/CRTE/REG markers in nearby unmapped cells on a target row.

    The common layout has a blank first/action column immediately to the left
    of the first mapped parameter column, but this scans the small gap between
    the Target Values marker and the first mapped parameter as well.
    """
    if not param_cols:
        return None

    first_param_col = min(param_cols.keys())

    # Prefer the immediately preceding cell, then scan left back to marker_col.
    candidate_cols = []
    if first_param_col > 1:
        candidate_cols.append(first_param_col - 1)

    left_bound = max(1, min(marker_col, first_param_col - 1))
    for c in range(first_param_col - 2, left_bound - 1, -1):
        if c not in candidate_cols:
            candidate_cols.append(c)

    # Also inspect a few leading columns because some sheets place the action
    # column outside the merged Target Values area.
    for c in range(1, min(first_param_col, 5)):
        if c not in candidate_cols:
            candidate_cols.append(c)

    for c in candidate_cols:
        action = normalize_row_action(ws_values.cell(row, c).value)
        if action:
            return action

    return None



def row_starts_documentation_block(ws_values, row: int, scan_cols: Optional[Sequence[int]] = None) -> bool:
    """Detect detail/reference tables that follow a real Target Values table.

    CIQ sheets commonly append documentation blocks after configuration rows:
      Resource Details | Details hidden
      Call Timer Details | ...
      License Details | ...
      Resources (...) | Description

    These are not configuration records and must terminate target parsing.
    """
    if scan_cols is None:
        scan_cols = range(1, ws_values.max_column + 1)

    vals = [
        re.sub(r"\s+", " ", as_text(ws_values.cell(row, c).value)).strip().lower()
        for c in scan_cols
        if as_text(ws_values.cell(row, c).value)
    ]
    if not vals:
        return False

    joined = " | ".join(vals)

    strong_markers = (
        "resource details",
        "details hidden",
        "call timer details",
        "license details",
        "timer details",
    )
    if any(marker in joined for marker in strong_markers):
        return True

    # Generic reference/detail-table header after target rows.
    if any(v == "description" for v in vals):
        if any(
            ("resource" in v or "timer" in v or "item" == v or "value range" in v)
            for v in vals
        ):
            return True

    return False


def parse_horizontal_targets(
    ws_values,
    marker_row: int,
    marker_col: int,
    known_params: set[str],
    max_end_row: Optional[int] = None,
) -> List[Dict[str, Any]]:
    """Parse horizontal Target Values tables efficiently, including very large tables.

    The older loop scanned every cell in every row to look for stop markers. On
    400k-row target sheets this made runtime explode.  This version only reads
    the target columns plus a small leading window used to detect a new section.
    """
    end_row = _region_end_row(ws_values, marker_row, max_end_row)
    start_col, end_col = target_column_bounds(ws_values, marker_row, marker_col)

    header_row = None
    param_cols: Dict[int, str] = {}

    # Header is normally 1-3 rows below Target Values marker.
    for r in range(marker_row + 1, min(end_row, marker_row + 5) + 1):
        matches: Dict[int, str] = {}
        for c in range(start_col, end_col + 1):
            p = _param_match(ws_values.cell(r, c).value, known_params)
            if p:
                matches[c] = p
        if len(matches) >= 2:
            header_row = r
            param_cols = matches
            break

    if header_row is None:
        return []

    records: List[Dict[str, Any]] = []
    blank_run = 0
    watch_cols = sorted(set(list(param_cols.keys()) + list(range(1, min(ws_values.max_column, 8) + 1))))

    for r in range(header_row + 1, end_row + 1):
        # Only inspect a small subset of columns for section/current markers.
        marker_text = " | ".join(as_text(ws_values.cell(r, c).value) for c in watch_cols)
        if CLI_RE.search(marker_text):
            break
        if TARGET_RE.search(marker_text) and r != marker_row:
            break
        if CURRENT_RE.search(marker_text):
            break

        # Stop before post-target documentation/reference tables.
        if records and row_starts_documentation_block(ws_values, r, watch_cols):
            break

        rec: Dict[str, Any] = {}
        has_any_value = False
        for c, p in param_cols.items():
            value = ws_values.cell(r, c).value
            if value is not None and as_text(value) != "":
                has_any_value = True
            rec[p] = value

        # Capture an optional row-level action marker from an unmapped control
        # column. This allows a row containing CHG to render as CHG-XXXX even
        # when the section CLI Command is CRTE-XXXX.
        row_action = detect_row_action_near_target_columns(
            ws_values, r, param_cols, marker_col
        )
        if row_action:
            rec["__ROW_ACTION__"] = row_action

        if has_any_value:
            if record_contains_metadata_row(rec):
                # Once real target rows have started, a metadata/details row
                # marks the end of the target table. Stop parsing so following
                # documentation rows (Resource Details, descriptions, timer
                # details, etc.) cannot become TL1 commands.
                if records:
                    break
                continue
            blank_run = 0
            records.append(rec)
            if len(records) % 50000 == 0:
                print(f"  Parsed {len(records):,} target rows from sheet '{ws_values.title}'...")
        else:
            blank_run += 1
            # On very large tables there may be occasional blank separators, but
            # repeated blanks should terminate the table.
            if blank_run >= MAX_BLANK_ROWS_IN_TARGET_TABLE:
                break

    return records

def parse_vertical_targets(
    ws_values,
    marker_row: int,
    marker_col: int,
    known_params: set[str],
    max_end_row: Optional[int] = None,
) -> List[Dict[str, Any]]:
    """Parse target table where parameter names are vertical and values are in column(s) to the right."""
    end_row = _region_end_row(ws_values, marker_row, max_end_row)
    start_col, end_col = target_column_bounds(ws_values, marker_row, marker_col)

    # Find dominant parameter column after marker.
    best_col = None
    best_count = 0
    scan_end = min(end_row, marker_row + 100)
    for c in range(start_col, end_col + 1):
        count = 0
        for r in range(marker_row + 1, scan_end + 1):
            if _param_match(ws_values.cell(r, c).value, known_params):
                count += 1
        if count > best_count:
            best_count = count
            best_col = c

    if best_col is None or best_count < 2:
        return []

    param_rows: List[Tuple[int, str]] = []
    blank_run = 0
    for r in range(marker_row + 1, end_row + 1):
        p = _param_match(ws_values.cell(r, best_col).value, known_params)
        if p:
            param_rows.append((r, p))
            blank_run = 0
        else:
            blank_run += 1
            if param_rows and blank_run >= 2:
                break

    if not param_rows:
        return []

    # Determine value columns. Ignore columns that contain status/current markers or repeat param names.
    candidate_cols = []
    for c in range(best_col + 1, end_col + 1):
        values = [ws_values.cell(r, c).value for r, _ in param_rows]
        nonblank = [v for v in values if not is_blank(v)]
        if not nonblank:
            continue
        if all(is_status(v) for v in nonblank):
            continue
        if any(CURRENT_RE.search(as_text(ws_values.cell(marker_row, c).value)) for _ in [0]):
            continue
        if all(_param_match(v, known_params) for v in nonblank if as_text(v)):
            continue
        candidate_cols.append(c)

    if not candidate_cols:
        return []

    # Stop candidate cols at the first obvious Current Values area.
    trimmed_cols = []
    for c in candidate_cols:
        marker_text = as_text(ws_values.cell(marker_row, c).value)
        if CURRENT_RE.search(marker_text):
            break
        trimmed_cols.append(c)
    candidate_cols = trimmed_cols or candidate_cols[:1]

    records: List[Dict[str, Any]] = []
    for c in candidate_cols:
        rec = {p: ws_values.cell(r, c).value for r, p in param_rows}
        if any(not is_blank(v) and not is_status(v) for v in rec.values()):
            if not record_contains_metadata_row(rec):
                records.append(rec)
    return records



def parse_embedded_parameter_target_values(
    ws_values,
    section: Section,
    next_cli_row: int,
) -> List[Dict[str, Any]]:
    """Parse a PARAMATERS table whose Target Value column is embedded in the table.

    Some CIQ sections do not have a separate "Target Values" block. Instead the
    PARAMATERS header contains a Target Value column and each parameter row has
    its desired value in that same column. Example:

        PARAMATERS | Type | Mandatory | CLI Name | ... | Target Value (Staging)
        mtc-conf-mtc-t3412-home | ... | [MTC_T3412_HOME] | ... | MME
        mtc-conf-mtc-t3412-roam | ... | [MTC_T3412_ROAM] | ... | MME

    This function converts the whole vertical parameter table into one target
    record so commands such as CHG-MMEMTC-CONF are generated as one combined
    TL1 command line.
    """
    max_row = min(ws_values.max_row, next_cli_row - 1 if next_cli_row else ws_values.max_row)

    header_row = None
    param_col = None
    cli_col = None
    target_col = None

    for r in range(section.cli_row, max_row + 1):
        row_texts = [norm_key(ws_values.cell(r, c).value) for c in range(1, ws_values.max_column + 1)]
        if not any(t in {"parameters", "paramaters"} for t in row_texts):
            continue

        header_row = r
        for c in range(1, ws_values.max_column + 1):
            raw = as_text(ws_values.cell(r, c).value)
            key = norm_key(raw)
            if key in {"parameters", "paramaters"}:
                param_col = c
            elif key == "cli name":
                cli_col = c
            elif TARGET_RE.search(raw):
                target_col = c
        break

    if header_row is None or cli_col is None or target_col is None:
        return []

    record: Dict[str, Any] = {}
    useful_rows = 0
    blank_run = 0

    for r in range(header_row + 1, max_row + 1):
        # Stop on a new CLI section or a true separate Target Values block.
        row_join = " | ".join(
            as_text(ws_values.cell(r, c).value)
            for c in range(1, min(ws_values.max_column, max(target_col, cli_col) + 2) + 1)
        )
        if CLI_RE.search(row_join):
            break

        cli_name = clean_cli_name(ws_values.cell(r, cli_col).value)
        param = clean_param(ws_values.cell(r, param_col).value) if param_col else ""
        if is_placeholder_param(param):
            param = ""
        raw_target = ws_values.cell(r, target_col).value

        if not cli_name and not param and is_blank(raw_target):
            blank_run += 1
            if useful_rows and blank_run >= 2:
                break
            continue
        blank_run = 0

        if not cli_name:
            # If the explicit CLI Name cell is blank, try to infer it from the
            # parameter/YANG-style header.
            cli_name = infer_cli_from_param(param, section.cli_order) or ""

        if not cli_name:
            continue

        # Only consume parameters that belong to this command template.
        if section.cli_order and cli_name not in {x.upper() for x in section.cli_order}:
            inferred = infer_cli_from_param(param, section.cli_order)
            if inferred:
                cli_name = inferred
            else:
                continue

        # Ignore helper/status text rather than generating it into TL1.
        if is_metadata_or_helper_value(raw_target):
            continue

        # Use the parameter key when available so existing mapping still applies.
        # Also register direct CLI aliases defensively.
        key = param or cli_name
        section.param_to_cli.setdefault(key, cli_name)
        section.param_to_cli.setdefault(cli_name, cli_name)
        section.param_to_cli.setdefault(f"[{cli_name}]", cli_name)

        record[key] = raw_target
        useful_rows += 1

    if not record:
        return []

    return [record]



def extract_act_lic_record(
    wb_values,
    section: Section,
    sections: List[Section],
) -> List[Dict[str, Any]]:
    """Extract ACT-LIC information/signature target values.

    ACT-LIC uses labeled target rows instead of a normal horizontal/vertical
    target table, e.g.:
        Target Value (...) (information) | SN=...
        Target Value (...) (signature)   | MCw...
    """
    if section.command.upper() != "ACT-LIC":
        return []

    ws = wb_values[section.sheet_name]
    next_cli_row = min(
        [s.cli_row for s in sections
         if s.sheet_name == section.sheet_name and s.cli_row > section.cli_row]
        or [ws.max_row + 1]
    )

    information = None
    signature = None

    for r in range(section.cli_row, min(next_cli_row - 1, ws.max_row) + 1):
        for c in range(1, ws.max_column + 1):
            label = as_text(ws.cell(r, c).value)
            if not label:
                continue

            low = re.sub(r"\s+", " ", label).strip().lower()
            if not TARGET_RE.search(label):
                continue

            is_info = "information" in low
            is_sig = "signature" in low
            if not (is_info or is_sig):
                continue

            # Value is normally in the next non-empty cell on the same row.
            value = None
            for vc in range(c + 1, ws.max_column + 1):
                candidate = ws.cell(r, vc).value
                if candidate is None or as_text(candidate) == "":
                    continue
                candidate_text = as_text(candidate)
                # Do not accidentally consume another label/status cell.
                if is_metadata_or_helper_value(candidate_text):
                    continue
                value = candidate_text
                break

            if value is None:
                continue

            if is_info and information is None:
                information = value
            elif is_sig and signature is None:
                signature = value

        if information is not None and signature is not None:
            break

    if information is None and signature is None:
        return []

    record: Dict[str, Any] = {}
    if information is not None:
        record["__ACT_LIC_INF__"] = information
    if signature is not None:
        record["__ACT_LIC_SIG__"] = signature
    return [record]


def quote_tl1_exact(value: Any) -> str:
    """Wrap a value in double quotes while otherwise preserving it exactly."""
    text = as_text(value)
    # Preserve embedded quotes safely for TL1 output.
    text = text.replace('"', '\\"')
    return f'"{text}"'


def records_to_act_lic_tl1(section: Section, records: List[Dict[str, Any]]) -> List[str]:
    """Render ACT-LIC as one exact-case INF/SIG command."""
    if section.command.upper() != "ACT-LIC" or not records:
        return []

    record = records[0]
    inf = record.get("__ACT_LIC_INF__")
    sig = record.get("__ACT_LIC_SIG__")

    parts: List[str] = []
    if inf is not None and as_text(inf) != "":
        parts.append(f"INF={quote_tl1_exact(inf)}")
    if sig is not None and as_text(sig) != "":
        parts.append(f"SIG={quote_tl1_exact(sig)}")

    if not parts:
        return []

    return [f"ACT-LIC:{','.join(parts)};"]



def extract_meas_kday_records(
    wb_values,
    section: Section,
    sections: List[Section],
) -> List[Dict[str, Any]]:
    """Build MEAS-KDAY records from paired *-day / *-hour parameters.

    Some PM File Storage Period sections intentionally leave CLI Name blank and
    encode the command TYPE in the parameter name itself:

        raw-1min-day   -> TYPE=RAW-1MIN, DAYS=<target>
        raw-1min-hour  -> TYPE=RAW-1MIN, HOURS=<target>

    The same rule applies generically to any <type>-day / <type>-hour pair.
    """
    if section.command.upper() != "CHG-MEAS-KDAY":
        return []

    ws = wb_values[section.sheet_name]
    next_cli_row = min(
        [s.cli_row for s in sections
         if s.sheet_name == section.sheet_name and s.cli_row > section.cli_row]
        or [ws.max_row + 1]
    )

    # Locate PARAMATERS header and its embedded Target Value column.
    header_row = None
    param_col = None
    target_col = None

    for r in range(section.cli_row, min(next_cli_row - 1, ws.max_row) + 1):
        row_texts = [norm_key(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
        if not any(t in {"parameters", "paramaters"} for t in row_texts):
            continue
        header_row = r
        for c in range(1, ws.max_column + 1):
            raw = as_text(ws.cell(r, c).value)
            key = norm_key(raw)
            if key in {"parameters", "paramaters"}:
                param_col = c
            elif TARGET_RE.search(raw):
                target_col = c
        break

    if header_row is None or param_col is None or target_col is None:
        return []

    grouped: Dict[str, Dict[str, Any]] = {}
    order: List[str] = []

    for r in range(header_row + 1, min(next_cli_row - 1, ws.max_row) + 1):
        param = clean_param(ws.cell(r, param_col).value)
        if not param:
            continue

        m = re.fullmatch(r"(.+)-(day|hour)", param.strip(), re.I)
        if not m:
            # Once grouped rows have started, unrelated rows are documentation
            # or the next structure; ignore them rather than inventing output.
            continue

        type_name = m.group(1)
        kind = m.group(2).lower()
        raw_target = ws.cell(r, target_col).value

        if is_metadata_or_helper_value(raw_target):
            continue

        if type_name not in grouped:
            grouped[type_name] = {"TYPE": type_name}
            order.append(type_name)

        grouped[type_name]["DAYS" if kind == "day" else "HOURS"] = raw_target

    if not grouped:
        return []

    # Direct aliases allow generic record_to_cli_values/render_cli_line reuse.
    for cli in ("TYPE", "DAYS", "HOURS"):
        section.param_to_cli.setdefault(cli, cli)

    return [grouped[name] for name in order]


def extract_target_records_for_section(wb_values, section: Section, sections: List[Section]) -> List[Dict[str, Any]]:
    if section.command.upper() == "ACT-LIC":
        return extract_act_lic_record(wb_values, section, sections)

    if section.command.upper() == "CHG-MEAS-KDAY":
        meas_records = extract_meas_kday_records(wb_values, section, sections)
        if meas_records:
            return meas_records

    known_params = set(section.param_to_cli.keys())
    if not known_params:
        return []

    ws = wb_values[section.sheet_name]
    next_cli_row = min(
        [s.cli_row for s in sections if s.sheet_name == section.sheet_name and s.cli_row > section.cli_row] or [ws.max_row + 1]
    )

    # 1) Try target markers in this section only.
    # Ignore the Target Value/Current Value columns that appear in the Parameters table header.
    param_table_end = find_parameter_table_end(ws, section, next_cli_row)
    markers = find_target_markers(ws, param_table_end + 1, next_cli_row - 1)
    records: List[Dict[str, Any]] = []
    for marker_row, marker_col in markers:
        records.extend(parse_horizontal_targets(ws, marker_row, marker_col, known_params, next_cli_row))
        if not records:
            records.extend(parse_vertical_targets(ws, marker_row, marker_col, known_params, next_cli_row))
        if records:
            break

    # 2) If target block is only a pointer, follow "Refer to <sheet> sheet".
    if not records:
        referred = find_referred_sheet_name(ws, section.cli_row, next_cli_row - 1)
        referred_sheet_name = resolve_referred_sheet(wb_values, referred) if referred else None
        if referred_sheet_name:
            ref_ws = wb_values[referred_sheet_name]
            ref_markers = find_target_markers(ref_ws)
            for marker_row, marker_col in ref_markers:
                records.extend(parse_horizontal_targets(ref_ws, marker_row, marker_col, known_params))
                if not records:
                    records.extend(parse_vertical_targets(ref_ws, marker_row, marker_col, known_params))
                if records:
                    break

    # 3) Fallback: some sections keep Target Value directly inside the
    # PARAMATERS table and have no separate Target Values block.
    if not records:
        records = parse_embedded_parameter_target_values(ws, section, next_cli_row)

    return records


# -----------------------------
# TL1 rendering
# -----------------------------

def is_metadata_or_helper_value(value: Any) -> bool:
    """Return True for CIQ header/status/helper cells that must never become TL1 values.

    This prevents output such as:
        CHG-AMFTA-INF:MNC="REVIEW STATUS",TAL=PENDING;
        CHG-AMFTA-INF:MCC="YANG PATH: ...";
    which can happen when a shifted Target Values marker causes the parser to
    treat section metadata rows as target data.
    """
    text = as_text(value).strip()
    if not text:
        return False

    low = re.sub(r"\s+", " ", text).strip().lower()

    exact_bad = {
        "review status",
        "config status",
        "description",
        "description:",
        "dependencies",
        "environment specific?",
        "parameters",
        "paramaters",
        "cli name",
        "remarks",
        "target value",
        "target values",
        "current value",
        "current values",
        "type",
        "mandatory",
        "details hidden",
        "resource details",
        "call timer details",
        "license details",
        "value range",
        "recommended",
    }
    if low in exact_bad:
        return True

    bad_prefixes = (
        "yang path:",
        "cli command:",
        "internal sanity check:",
        "target value ",
        "target values ",
        "current value ",
        "current values ",
        "resource details",
        "call timer details",
        "license details",
        "details hidden",
    )
    if low.startswith(bad_prefixes):
        return True

    # Header-like rows often contain these fragments and should be ignored when
    # they appear as target values.
    bad_fragments = (
        "review status",
        "config status",
        "yang path:",
        "cli command:",
    )
    return any(fragment in low for fragment in bad_fragments)


def record_contains_metadata_row(record: Dict[str, Any]) -> bool:
    """Detect and skip whole records accidentally parsed from CIQ metadata rows."""
    return any(is_metadata_or_helper_value(v) for v in record.values())


def record_to_cli_values(section: Section, record: Dict[str, Any], *, keep_arrows: bool = False) -> Dict[str, str]:
    """Build CLI-name -> formatted value from one target record."""
    if record_contains_metadata_row(record):
        return {}

    cli_values: Dict[str, str] = {}
    for param, raw_value in record.items():
        if param == "__ROW_ACTION__":
            continue
        cli_name = section.param_to_cli.get(param)

        # For MMECALL-PROF, the CIQ template sometimes has swapped/misaligned
        # CLI Name cells in the PARAMATERS table; the Target Values headers are
        # the reliable source. Force known call-prof-* headers to their correct
        # TL1 names before formatting. This prevents:
        #   call-prof-dcn_mmeg -> NEPSIDX
        #   call-prof-nepsidx  -> DCN_MMEG
        # and restores DCN_MMEG=NONE, NEPSIDX=1.
        if section.command.upper() == "CHG-MMECALL-PROF":
            inferred_cli = infer_cli_from_param(param, section.cli_order)
            if inferred_cli:
                cli_name = inferred_cli

        if not cli_name:
            continue
        if is_arrow_value(raw_value) and not keep_arrows:
            continue
        cleaned = clean_cell_value(raw_value)
        if cleaned is None:
            continue
        formatted = format_tl1_value(cli_name, cleaned, section.command)
        if cli_name.upper() == "NE" and section.command.upper() in {"CRTE-NE-GRP", "CHG-NE-GRP"}:
            formatted = normalize_ne_type(formatted)
        cli_values[cli_name] = formatted

    fixed_values = extract_fixed_cli_values(section.raw_cli)
    for fixed_name, fixed_value in fixed_values.items():
        cli_values[fixed_name] = fixed_value

    return cli_values


def render_cli_line(
    command: str,
    cli_order: Sequence[str],
    values: Dict[str, str],
    *,
    required_order: Optional[Sequence[str]] = None,
    keep_empty_fields: Optional[set[str]] = None,
) -> Optional[str]:
    """Render one TL1 line using the command/template order."""
    order = list(required_order or cli_order or values.keys())
    keep_empty_fields = {x.upper() for x in (keep_empty_fields or set())}
    parts: List[str] = []
    seen = set()
    for cli_name in order:
        cli_name = cli_name.upper()
        if cli_name in seen:
            continue
        seen.add(cli_name)
        if cli_name not in values:
            continue
        value = values[cli_name]

        # NE-GRP convention: do not emit BKUPGID when the configured value is 0.
        # Example:
        #   CRTE-NE-GRP:NE=UDM,GID=1,GNAME=LAB-UDM-GRP1,BKUPGID=0,SELM=CAPA,...
        # becomes:
        #   CRTE-NE-GRP:NE=UDM,GID=1,GNAME=LAB-UDM-GRP1,SELM=CAPA,...
        if command.upper() == "CRTE-NE-GRP" and cli_name == "BKUPGID" and str(value).strip() == "0":
            continue

        if value == "" and cli_name not in keep_empty_fields:
            continue
        parts.append(f"{cli_name}={value}")
    if not parts:
        return None
    return f"{command}:{','.join(parts)};"


def records_to_ne_grp_tl1(section: Section, records: List[Dict[str, Any]]) -> List[str]:
    """Special renderer for NE-GRP tables with arrow continuation rows.

    First full row creates CRTE-NE-GRP. Rows with arrow markers inherit the
    parent group fields and create CHG-NE-GRP rows using the current NEID/CAPA/etc.
    This matches CIQ group samples where continuation rows add group members.
    """
    if section.command.upper() != "CRTE-NE-GRP":
        return []

    parent_fields = ["NE", "GID", "GNAME", "BKUPGID", "DFTGID", "DFLTGID", "DEFAULTGID", "DEFGID", "SELM", "IFDIR"]
    child_fields = ["NEID", "PRI", "CAPA", "ROAM", "RCAPA", "ECR", "INTV"]
    crte_order = [x for x in section.cli_order if x in parent_fields + child_fields]
    if not crte_order:
        crte_order = ["NE", "GID", "GNAME", "BKUPGID", "SELM", "IFDIR", "NEID", "PRI", "CAPA", "ROAM", "RCAPA", "ECR", "INTV"]
    chg_order = ["NE", "GID", "MODE", "NEID", "PRI", "CAPA", "ROAM", "RCAPA", "ECR", "INTV"]

    lines: List[str] = []
    current_parent: Dict[str, str] = {}

    for rec in records:
        values = record_to_cli_values(section, rec, keep_arrows=False)
        if not values:
            continue

        # A new group starts whenever NE or GID is explicitly provided in the row.
        starts_new_group = any(k in values for k in ("NE", "GID", "GNAME"))
        if starts_new_group or not current_parent:
            current_parent.update({k: v for k, v in values.items() if k in parent_fields})
            line_values = {**current_parent, **{k: v for k, v in values.items() if k in child_fields}}
            line = render_cli_line("CRTE-NE-GRP", crte_order, line_values)
            if line:
                lines.append(line)
        else:
            line_values = {k: v for k, v in current_parent.items() if k in {"NE", "GID"}}
            line_values["MODE"] = "ADD"
            line_values.update({k: v for k, v in values.items() if k in child_fields})
            line = render_cli_line("CHG-NE-GRP", chg_order, line_values, required_order=chg_order)
            if line:
                lines.append(line)

    return lines


def _has_arrow_marker(record: Dict[str, Any]) -> bool:
    return any(is_arrow_value(v) for v in record.values())


def _service_fields_present(section: Section) -> bool:
    order = {x.upper() for x in section.cli_order}
    return "SVC" in order and ("SVCIP" in order or "SVCPORT" in order or "SVCID" in order)


def records_to_peer_service_tl1(section: Section, records: List[Dict[str, Any]]) -> List[str]:
    """Render CRTE-<NE> peer service-entry tables with arrow continuation rows.

    Generic for UDM/AUSF/NRF/SMF/PCF and similar network-element peer
    sections where the first row creates the peer and arrow rows add service
    entries using CHG-<NE>-INF:...,MODE=ADD,... .
    """
    cmd = section.command.upper()
    if not cmd.startswith("CRTE-") or cmd == "CRTE-NE-GRP" or not _service_fields_present(section):
        return []
    if not any(_has_arrow_marker(rec) for rec in records):
        return []

    service_fields = {"SVC", "SCHEME", "SVCIP", "SVCPORT", "SVCID"}
    object_name = cmd[5:]
    chg_command = f"CHG-{object_name}-INF"

    # Key fields to carry into CHG lines.
    # For peer service continuation rows, only the object key is carried.
    # Do NOT carry NAME or INSTID on CHG-<NE>-INF rows.
    # Example required output:
    # CHG-UDM-INF:UDM=1,MODE=ADD,SVC=nudm-ee,SVCIP=...,SVCPORT=...;
    object_key_candidates = {object_name, "ID"}
    parent_key_order = [x for x in section.cli_order if x.upper() in object_key_candidates]
    if not parent_key_order and section.cli_order:
        parent_key_order = [section.cli_order[0]]
    chg_order = [x.upper() for x in parent_key_order] + ["MODE", "SVC", "SCHEME", "SVCIP", "SVCPORT", "SVCID"]

    lines: List[str] = []
    current_parent: Dict[str, str] = {}
    previous_svc: Optional[str] = None

    for rec in records:
        values = record_to_cli_values(section, rec, keep_arrows=False)
        if not values:
            continue

        has_arrow = _has_arrow_marker(rec)
        parent_values = {k: v for k, v in values.items() if k.upper() not in service_fields}
        svc_values = {k: v for k, v in values.items() if k.upper() in service_fields}

        if not has_arrow or not current_parent:
            current_parent.update(parent_values)
            line = render_cli_line(cmd, section.cli_order, {**current_parent, **svc_values})
            if line:
                lines.append(line)
            previous_svc = svc_values.get("SVC") or previous_svc
            continue

        # Continuation service entry. Carry only the object key(s).
        # NAME and INSTID are intentionally omitted from CHG-<NE>-INF rows.
        line_values: Dict[str, str] = {}
        for key in parent_key_order:
            k = key.upper()
            if k in current_parent:
                line_values[k] = current_parent[k]
        line_values["MODE"] = "ADD"
        line_values.update(svc_values)

        # Avoid repeating SCHEME on the second IP row for the same service.
        current_svc = svc_values.get("SVC")
        if current_svc and previous_svc == current_svc and "SCHEME" in line_values:
            line_values.pop("SCHEME", None)
        previous_svc = current_svc or previous_svc

        line = render_cli_line(
            chg_command,
            chg_order,
            line_values,
            required_order=chg_order,
        )
        if line:
            lines.append(line)

    return lines



def records_to_mmertal_tl1(section: Section, records: List[Dict[str, Any]]) -> List[str]:
    """Special renderer for MMERTAL restricted TA list tables.

    Required behavior:
      - First row with numeric RTALIDX generates:
            CRTE-MMERTAL:RTALIDX=...,CAUSE=...,MCC=...,MNC=...,TAC=...;
      - Following rows after that index generate:
            CHG-MMERTAL-INF:RTALIDX=...,CAUSE=...,TYPE=ADD,MCC=...,MNC=...,TAC=...;
      - Arrow/ditto RTALIDX values inherit the current numeric RTALIDX.
      - TYPE=ADD is inserted only on CHG-MMERTAL-INF rows.
    """
    if section.command.upper() != "CRTE-MMERTAL":
        return []

    lines: List[str] = []
    current_idx: Optional[str] = None
    current_cause: Optional[str] = None
    created_for_idx: set[str] = set()

    crte_order = ["RTALIDX", "CAUSE", "MCC", "MNC", "TAC"]
    chg_order = ["RTALIDX", "CAUSE", "TYPE", "MCC", "MNC", "TAC"]

    for rec in records:
        values = record_to_cli_values(section, rec, keep_arrows=False)
        raw_values = {}
        for param, raw_value in rec.items():
            cli_name = section.param_to_cli.get(param)
            if cli_name:
                raw_values[cli_name.upper()] = raw_value

        # Detect explicit numeric RTALIDX before arrow removal.
        raw_idx = raw_values.get("RTALIDX")
        cleaned_raw_idx = clean_cell_value(raw_idx) if raw_idx is not None and not is_arrow_value(raw_idx) else None
        has_new_idx = cleaned_raw_idx not in (None, "")

        if has_new_idx:
            if "RTALIDX" in values:
                current_idx = values["RTALIDX"]
            else:
                current_idx = format_tl1_value("RTALIDX", cleaned_raw_idx, section.command)

        if "CAUSE" in values and values["CAUSE"] != "":
            current_cause = values["CAUSE"]

        # Skip rows until a numeric RTALIDX has appeared.
        if not current_idx:
            continue

        line_values: Dict[str, str] = {}
        line_values["RTALIDX"] = current_idx

        # CAUSE should be carried forward if blank/arrow after the first row.
        if "CAUSE" in values and values["CAUSE"] != "":
            line_values["CAUSE"] = values["CAUSE"]
        elif current_cause:
            line_values["CAUSE"] = current_cause

        for field in ("MCC", "MNC", "TAC"):
            if field in values:
                line_values[field] = values[field]

        # Do not generate a line without actual TA content unless the first row
        # has all required values in the CRTE line.
        has_ta_content = any(field in line_values and line_values[field] != "" for field in ("MCC", "MNC", "TAC"))
        if not has_ta_content:
            continue

        if current_idx not in created_for_idx:
            line = render_cli_line("CRTE-MMERTAL", crte_order, line_values, required_order=crte_order)
            if line:
                lines.append(line)
                created_for_idx.add(current_idx)
        else:
            line_values["TYPE"] = "ADD"
            line = render_cli_line("CHG-MMERTAL-INF", chg_order, line_values, required_order=chg_order)
            if line:
                lines.append(line)

    return lines


def record_to_tl1(section: Section, record: Dict[str, Any]) -> Optional[str]:
    """Convert one target record into one TL1 command line.

    If the target row contains an action marker such as CHG, override the
    section command prefix for that row while preserving the command object.
    """
    cli_values = record_to_cli_values(section, record, keep_arrows=False)
    row_action = normalize_row_action(record.get("__ROW_ACTION__"))
    command = derive_command_for_row(section.command, row_action)
    command_upper = command.upper()

    # Skip Service Overload Control entries that target the synthetic
    # "{all services}" selector.
    if command_upper == "CHG-OVLD-CNTL":
        svc = cli_values.get("SVC", "")
        if svc.strip().strip('"').lower() == "{all services}":
            return None

    # Skip VES control entries when EVENT is the synthetic ALL selector.
    if command_upper == "CHG-VES-CNTL":
        event = cli_values.get("EVENT", "")
        if event.strip().strip('"').upper() == "{ALL}":
            return None

    # MEAS-KDAY TL1 expects TYPE without the RAW- prefix.
    # RAW-1MIN -> 1MIN, RAW-5MIN -> 5MIN, RAW-HOUR -> HOUR, etc.
    if command_upper == "CHG-MEAS-KDAY" and "TYPE" in cli_values:
        type_value = cli_values["TYPE"]
        if type_value.upper().startswith("RAW-"):
            cli_values["TYPE"] = type_value[4:]

    return render_cli_line(command, section.cli_order, cli_values)


def build_tl1_output(wb_values, selected_sheets: Sequence[str], selected_section_numbers: Optional[set[int]] = None) -> Tuple[str, List[str]]:
    sections = discover_sections(wb_values, selected_sheets)
    output_blocks: List[str] = []
    warnings: List[str] = []

    for idx, section in enumerate(sections, start=1):
        if selected_section_numbers and idx not in selected_section_numbers:
            continue

        ws_for_size = wb_values[section.sheet_name]

        next_cli_row = min(
            [s.cli_row for s in sections
             if s.sheet_name == section.sheet_name and s.cli_row > section.cli_row]
            or [ws_for_size.max_row + 1]
        )

        # Skip the entire section before parsing any target values when either
        # Review Status or Config Status says Out of Scope.
        if is_out_of_scope_section(ws_for_size, section, next_cli_row):
            warnings.append(
                f"Skipped Out of Scope section #{idx}: {section.title} "
                f"[{section.sheet_name} row {section.cli_row}]"
            )
            continue

        if getattr(ws_for_size, "max_row", 0) and ws_for_size.max_row >= LARGE_SHEET_ROW_THRESHOLD:
            print(f"Processing large sheet '{section.sheet_name}' with {ws_for_size.max_row:,} rows, section #{idx}: {section.title}")

        records = extract_target_records_for_section(wb_values, section, sections)
        if len(records) >= 50000:
            print(f"  Rendering {len(records):,} TL1 candidate records for section #{idx}: {section.title}")

        if section.command.upper() == "ACT-LIC":
            lines = records_to_act_lic_tl1(section, records)
        elif section.command.upper() == "CRTE-NE-GRP":
            lines = records_to_ne_grp_tl1(section, records)
        elif section.command.upper() == "CRTE-MMERTAL":
            lines = records_to_mmertal_tl1(section, records)
        else:
            lines = records_to_peer_service_tl1(section, records)
            if not lines:
                lines = []
                for row_num, rec in enumerate(records, start=1):
                    line = record_to_tl1(section, rec)
                    if line:
                        lines.append(line)
                    if row_num % 50000 == 0:
                        print(f"  Rendered {row_num:,} records for section #{idx}...")

        if lines:
            block = [f"### {section.title}"] + lines
            output_blocks.append("\n".join(block))
        else:
            warnings.append(f"No TL1 rows generated for section #{idx}: {section.title} [{section.sheet_name} row {section.cli_row}]")

    return "\n\n".join(output_blocks).rstrip() + "\n", warnings


# -----------------------------
# User interface
# -----------------------------
def choose_file_dialog() -> Optional[str]:
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        path = filedialog.askopenfilename(
            title="Select Excel workbook",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        root.destroy()
        return path or None
    except Exception:
        return None


def choose_folder_dialog() -> Optional[str]:
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        path = filedialog.askdirectory(title="Select output folder")
        root.destroy()
        return path or None
    except Exception:
        return None


def prompt_path(prompt: str, must_exist: bool = False, is_dir: bool = False) -> str:
    while True:
        p = input(prompt).strip().strip('"')
        if not p:
            continue
        path = Path(p)
        if must_exist:
            if is_dir and not path.is_dir():
                print("Folder not found. Try again.")
                continue
            if not is_dir and not path.is_file():
                print("File not found. Try again.")
                continue
        return str(path)


def ask_workbook_path() -> str:
    print("Select Excel workbook...")
    path = choose_file_dialog()
    if path:
        return path
    return prompt_path("Enter Excel workbook path: ", must_exist=True)




def split_text_by_max_lines(text: str, max_lines: int = MAX_TL1_LINES_PER_FILE) -> List[str]:
    """Split output text into chunks with at most max_lines lines per file.

    Blank separator lines and section headers count as lines, so every generated
    file is guaranteed to stay within the configured maximum line count.
    """
    lines = text.splitlines()
    if not lines:
        return [""]

    chunks: List[str] = []
    for start in range(0, len(lines), max_lines):
        chunk_lines = lines[start:start + max_lines]
        chunks.append("\n".join(chunk_lines).rstrip() + "\n")
    return chunks



def timestamped_run_output_path(output_path: str) -> Path:
    """Return an output path inside a YYYYMMDD_HHMMSS run folder.

    The caller supplies the desired parent directory and filename. This helper
    guarantees that generation occurs inside a timestamped child folder:

        <selected-output-folder>/YYYYMMDD_HHMMSS/<filename>

    If the supplied parent is already a timestamp run folder, it is reused.
    This keeps CLI and GUI behavior identical and prevents accidental double
    timestamp nesting.
    """
    path = Path(output_path)
    parent = path.parent

    if re.fullmatch(r"\d{8}_\d{6}", parent.name):
        run_folder = parent
    else:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        run_folder = parent / timestamp

        # A second execution within the exact same second should not mix files.
        if run_folder.exists():
            suffix = 2
            base = run_folder
            while run_folder.exists():
                run_folder = parent / f"{base.name}_{suffix:02d}"
                suffix += 1

    run_folder.mkdir(parents=True, exist_ok=True)
    return run_folder / path.name



def write_warnings_file(written_files: Sequence[str], warnings: Sequence[str]) -> Optional[str]:
    """Write warnings beside the actual TL1 output inside the run timestamp folder."""
    if not warnings or not written_files:
        return None

    first_output = Path(written_files[0])
    warn_path = first_output.with_suffix(".warnings.txt")
    warn_path.write_text("\n".join(warnings) + "\n", encoding="utf-8")
    return str(warn_path)


def write_tl1_files(output_path: str, text: str, max_lines: int = MAX_TL1_LINES_PER_FILE) -> List[str]:
    """Write TL1 output, splitting into numbered files when needed.

    If the output has <= max_lines lines, writes the requested output_path.
    If it has more, writes:
        <stem>_part001.tl1
        <stem>_part002.tl1
        ...
    and does not create the oversized unsplit file.
    """
    path = timestamped_run_output_path(output_path)

    chunks = split_text_by_max_lines(text, max_lines=max_lines)
    written: List[str] = []

    if len(chunks) == 1:
        with open(path, "w", encoding="utf-8", newline="\n") as f:
            f.write(chunks[0])
        written.append(str(path))
        return written

    for idx, chunk in enumerate(chunks, start=1):
        part_path = path.with_name(f"{path.stem}_part{idx:03d}{path.suffix}")
        with open(part_path, "w", encoding="utf-8", newline="\n") as f:
            f.write(chunk)
        written.append(str(part_path))

    return written

def ask_output_path(default_stem: str) -> str:
    """Choose the parent output folder and desired TL1 filename.

    The timestamped run directory is created centrally by write_tl1_files(), so
    every caller follows the same output-folder behavior.
    """
    print("Select output folder...")
    folder = choose_folder_dialog()
    if not folder:
        folder = prompt_path("Enter output folder path: ", must_exist=True, is_dir=True)

    name = input(f"Enter output TL1 filename stem [{default_stem}]: ").strip().strip('"')
    if name:
        stem = Path(name).stem if name.lower().endswith(".tl1") else name
    else:
        stem = default_stem

    filename = f"{stem}_{VERSION_TAG}.tl1"
    return str(Path(folder) / filename)

def ask_sheet_selection(wb_values) -> List[str]:
    print("\nWorkbook sheets:")
    for i, name in enumerate(wb_values.sheetnames, start=1):
        print(f"  {i}. {name}")

    choice = input("\nProcess whole workbook or specific sheet(s)? [A=all / S=specific] (default A): ").strip().lower()
    if choice not in {"s", "specific"}:
        large_sheets = [name for name in wb_values.sheetnames if wb_values[name].max_row >= LARGE_SHEET_ROW_THRESHOLD]
        if large_sheets:
            print("\nLarge sheets detected:")
            for name in large_sheets:
                print(f"  - {name}: {wb_values[name].max_row:,} rows")
            confirm = input("Process ALL sheets including these large sheets? [y/N] (default N): ").strip().lower()
            if confirm not in {"y", "yes"}:
                print("Large sheets will be skipped during CLI-section discovery. They can still be read when referenced by a selected section.")
                return [name for name in wb_values.sheetnames if name not in set(large_sheets)]
        return list(wb_values.sheetnames)

    raw = input("Enter sheet number(s) or name(s), comma-separated: ").strip()
    selected = []
    for token in [x.strip() for x in raw.split(",") if x.strip()]:
        range_match = re.fullmatch(r"(\d+)\s*-\s*(\d+)", token)
        if range_match:
            start_idx, end_idx = int(range_match.group(1)), int(range_match.group(2))
            if start_idx > end_idx:
                start_idx, end_idx = end_idx, start_idx
            for idx in range(start_idx, end_idx + 1):
                if 1 <= idx <= len(wb_values.sheetnames):
                    selected.append(wb_values.sheetnames[idx - 1])
                else:
                    print(f"WARNING: sheet number out of range: {idx}")
        elif token.isdigit():
            idx = int(token)
            if 1 <= idx <= len(wb_values.sheetnames):
                selected.append(wb_values.sheetnames[idx - 1])
        elif token in wb_values.sheetnames:
            selected.append(token)
        else:
            print(f"WARNING: sheet not found: {token}")

    return selected or list(wb_values.sheetnames)


def ask_section_selection(sections: List[Section]) -> Optional[set[int]]:
    if not sections:
        return None
    print("\nDiscovered CLI sections:")
    for i, sec in enumerate(sections, start=1):
        print(f"  {i}. [{sec.sheet_name}] {sec.title} -> {sec.command}")
    raw = input("\nProcess all sections or selected section number(s)? [A=all / numbers comma-separated] (default A): ").strip()
    if not raw or raw.lower() in {"a", "all"}:
        return None
    nums = set()
    for token in raw.split(","):
        token = token.strip()
        range_match = re.fullmatch(r"(\d+)\s*-\s*(\d+)", token)
        if range_match:
            start_idx, end_idx = int(range_match.group(1)), int(range_match.group(2))
            if start_idx > end_idx:
                start_idx, end_idx = end_idx, start_idx
            nums.update(range(start_idx, end_idx + 1))
        elif token.isdigit():
            nums.add(int(token))
    return nums or None


def main() -> None:
    print(f"Excel CIQ -> TL1 generator {VERSION_TAG}")
    workbook_path = ask_workbook_path()
    print(f"Loading workbook: {workbook_path}")

    # data_only=True uses cached formula results saved by Excel. This is important because many CIQ templates use formulas like =C12.
    # For very large workbooks, set TL1_READ_ONLY=1 before running the script to reduce memory usage.
    # Normal mode remains the default because this parser uses random cell access heavily.
    read_only_mode = os.environ.get("TL1_READ_ONLY", "").strip().lower() in {"1", "true", "yes", "y"}
    if read_only_mode:
        print("Large-workbook mode enabled: loading workbook with openpyxl read_only=True.")
    wb_values = load_workbook(workbook_path, data_only=True, read_only=read_only_mode)

    selected_sheets = ask_sheet_selection(wb_values)
    sections = discover_sections(wb_values, selected_sheets)
    if not sections:
        print("ERROR: No sections found. Expected cells containing 'CLI Command:'.")
        sys.exit(2)

    selected_section_numbers = ask_section_selection(sections)
    default_stem = Path(workbook_path).stem + "_output"
    output_path = ask_output_path(default_stem)

    text, warnings = build_tl1_output(wb_values, selected_sheets, selected_section_numbers)

    written_files = write_tl1_files(output_path, text, max_lines=MAX_TL1_LINES_PER_FILE)

    print("\nDone. TL1 file generated:")
    for path in written_files:
        print(f"  {path}")
    if len(written_files) > 1:
        print(f"Split output into {len(written_files)} files with maximum 10000 lines per file.")

    line_count = sum(1 for line in text.splitlines() if line.strip() and not line.startswith("###"))
    total_lines = len(text.splitlines())
    print(f"Generated TL1 command lines: {line_count}")
    print(f"Generated total output lines: {total_lines}")

    warn_path = write_warnings_file(written_files, warnings)
    if warn_path:
        print(f"Warnings written to: {warn_path}")


if __name__ == "__main__":
    main()

"""Data format conversion tools (JSON ↔ YAML ↔ CSV ↔ TOML ↔ XLSX)."""
import json
from pathlib import Path
from typing import Dict, Optional

from mcp.server.fastmcp import FastMCP

from utils.file_utils import resolve_input, resolve_output, make_output_path

SUPPORTED_FORMATS = ["json", "yaml", "csv", "toml", "xlsx"]

_EXT_MAP = {
    ".json": "json",
    ".yaml": "yaml", ".yml": "yaml",
    ".csv": "csv",
    ".toml": "toml",
    ".xlsx": "xlsx", ".xls": "xlsx",
}


def _infer_format(path: Path) -> Optional[str]:
    return _EXT_MAP.get(path.suffix.lower())


def _read(path: Path, fmt: str, sheet_name: str = "Sheet1", delimiter: str = ","):
    """Load file into a Python object (dict/list) or pandas DataFrame."""
    import pandas as pd
    import yaml
    import toml

    if fmt == "json":
        return json.loads(path.read_text(encoding="utf-8"))
    if fmt == "yaml":
        return yaml.safe_load(path.read_text(encoding="utf-8"))
    if fmt == "toml":
        return toml.loads(path.read_text(encoding="utf-8"))
    if fmt == "csv":
        return pd.read_csv(path, delimiter=delimiter)
    if fmt == "xlsx":
        return pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")
    raise ValueError(f"Unknown format: {fmt}")


def _write(data, path: Path, fmt: str, indent: int = 2, delimiter: str = ","):
    """Write data (dict/list or DataFrame) to the given format."""
    import pandas as pd
    import yaml
    import toml

    path.parent.mkdir(parents=True, exist_ok=True)

    # Normalise to DataFrame for tabular targets
    if fmt in ("csv", "xlsx"):
        if isinstance(data, pd.DataFrame):
            df = data
        elif isinstance(data, list):
            df = pd.DataFrame(data)
        elif isinstance(data, dict):
            # Try to make a single-row DataFrame or flatten
            df = pd.DataFrame([data])
        else:
            raise ValueError("Cannot convert this data structure to a tabular format.")

        if fmt == "csv":
            df.to_csv(path, index=False, sep=delimiter)
        else:
            df.to_excel(path, index=False, engine="openpyxl")
        return len(df)

    # For dict/list targets convert DataFrame back to native Python
    if isinstance(data, pd.DataFrame):
        data = data.to_dict(orient="records")

    if fmt == "json":
        path.write_text(json.dumps(data, indent=indent, ensure_ascii=False), encoding="utf-8")
    elif fmt == "yaml":
        path.write_text(yaml.dump(data, allow_unicode=True, default_flow_style=False), encoding="utf-8")
    elif fmt == "toml":
        # TOML requires a dict at the top level
        if isinstance(data, list):
            data = {"items": data}
        path.write_text(toml.dumps(data), encoding="utf-8")
    return None


def register_data_tools(app: FastMCP) -> None:

    @app.tool()
    def convert_data(
        input_path: str,
        output_format: str,
        output_path: Optional[str] = None,
        input_format: Optional[str] = None,
        sheet_name: str = "Sheet1",
        csv_delimiter: str = ",",
        json_indent: int = 2,
    ) -> Dict:
        """
        Convert between data formats: json, yaml, csv, toml, xlsx.

        - JSON/YAML/TOML preserve nested structure.
        - CSV/XLSX work best with flat lists-of-dicts.
        - Converting nested data to CSV/XLSX flattens one level.
        """
        try:
            src = resolve_input(input_path)
            in_fmt = input_format or _infer_format(src)
            if not in_fmt:
                return {"success": False, "error": f"Cannot detect input format from '{src.suffix}'. Specify input_format."}

            out_fmt = output_format.lower()
            if out_fmt not in SUPPORTED_FORMATS:
                return {"success": False, "error": f"Unsupported output_format '{out_fmt}'. Choose from {SUPPORTED_FORMATS}"}

            ext_map = {"json": ".json", "yaml": ".yaml", "csv": ".csv", "toml": ".toml", "xlsx": ".xlsx"}
            if not output_path:
                output_path = make_output_path(input_path, f"_to_{out_fmt}", ext_map[out_fmt])
            dst = resolve_output(output_path)

            data = _read(src, in_fmt, sheet_name=sheet_name, delimiter=csv_delimiter)
            rows = _write(data, dst, out_fmt, indent=json_indent, delimiter=csv_delimiter)

            result = {
                "success": True,
                "input_path": str(src),
                "output_path": str(dst),
                "input_format": in_fmt,
                "output_format": out_fmt,
                "file_size_bytes": dst.stat().st_size,
            }
            if rows is not None:
                result["rows_converted"] = rows
            return result

        except FileNotFoundError as e:
            return {"success": False, "error": str(e)}
        except Exception as e:
            return {"success": False, "error": f"Data conversion failed: {e}"}

    @app.tool()
    def get_supported_formats() -> Dict:
        """List all supported data format conversions."""
        conversions = [
            [a, b] for a in SUPPORTED_FORMATS for b in SUPPORTED_FORMATS if a != b
        ]
        return {
            "formats": SUPPORTED_FORMATS,
            "conversions": conversions,
            "notes": {
                "csv_xlsx": "Best for flat list-of-dicts; nested data is flattened one level.",
                "toml": "Top-level must be a dict; lists are wrapped under an 'items' key when writing.",
            },
        }

    # ── XML ↔ JSON / CSV / YAML / XLSX ─────────────────────────────────────────────
    @app.tool()
    def xml_to_json(input_path: str, output_path: Optional[str] = None, indent: int = 2) -> Dict:
        """Convert an XML file to JSON using lxml."""
        try:
            from lxml import etree
            src = resolve_input(input_path)
            if not output_path:
                output_path = make_output_path(input_path, "", ".json")
            dst = resolve_output(output_path)

            def _el_to_dict(el):
                d = {**el.attrib}
                children = list(el)
                if children:
                    child_dict = {}
                    for ch in children:
                        tag = etree.QName(ch.tag).localname
                        val = _el_to_dict(ch)
                        if tag in child_dict:
                            if not isinstance(child_dict[tag], list):
                                child_dict[tag] = [child_dict[tag]]
                            child_dict[tag].append(val)
                        else:
                            child_dict[tag] = val
                    d.update(child_dict)
                elif el.text and el.text.strip():
                    if d:
                        d["_text"] = el.text.strip()
                    else:
                        return el.text.strip()
                return d

            tree = etree.parse(str(src))
            root = tree.getroot()
            tag = etree.QName(root.tag).localname
            data = {tag: _el_to_dict(root)}
            dst.write_text(json.dumps(data, indent=indent, ensure_ascii=False), encoding="utf-8")
            return {"success": True, "output_path": str(dst)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @app.tool()
    def json_to_xml(input_path: str, output_path: Optional[str] = None, root_tag: str = "root") -> Dict:
        """Convert a JSON file to XML."""
        try:
            from lxml import etree
            src = resolve_input(input_path)
            if not output_path:
                output_path = make_output_path(input_path, "", ".xml")
            dst = resolve_output(output_path)

            data = json.loads(src.read_text(encoding="utf-8"))

            def _dict_to_el(parent, obj):
                if isinstance(obj, dict):
                    for k, v in obj.items():
                        child = etree.SubElement(parent, str(k))
                        _dict_to_el(child, v)
                elif isinstance(obj, list):
                    for item in obj:
                        child = etree.SubElement(parent, "item")
                        _dict_to_el(child, item)
                else:
                    parent.text = str(obj)

            # If top-level is a dict with one key, use that as root
            if isinstance(data, dict) and len(data) == 1:
                root_tag = list(data.keys())[0]
                data = data[root_tag]

            root = etree.Element(root_tag)
            _dict_to_el(root, data)
            tree = etree.ElementTree(root)
            tree.write(str(dst), pretty_print=True, xml_declaration=True, encoding="UTF-8")
            return {"success": True, "output_path": str(dst)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @app.tool()
    def xml_to_csv(input_path: str, output_path: Optional[str] = None, record_tag: Optional[str] = None) -> Dict:
        """
        Flatten XML records to CSV. Finds repeating child elements as rows.
        record_tag: the element tag name that represents each row (auto-detected if omitted).
        """
        try:
            import pandas as pd
            from lxml import etree
            from xml.etree.ElementTree import iterparse

            src = resolve_input(input_path)
            if not output_path:
                output_path = make_output_path(input_path, "", ".csv")
            dst = resolve_output(output_path)

            tree = etree.parse(str(src))
            root = tree.getroot()
            children = list(root)
            if not children:
                return {"success": False, "error": "XML root has no children to flatten into rows."}

            tag = record_tag or etree.QName(children[0].tag).localname
            rows = []
            for el in root.iter(tag):
                row = {**el.attrib}
                for child in el:
                    row[etree.QName(child.tag).localname] = (child.text or "").strip()
                rows.append(row)

            df = pd.DataFrame(rows)
            df.to_csv(str(dst), index=False)
            return {"success": True, "output_path": str(dst), "rows": len(rows)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ── INI / ENV ──────────────────────────────────────────────────────────
    @app.tool()
    def ini_to_json(input_path: str, output_path: Optional[str] = None) -> Dict:
        """Convert an INI config file to JSON."""
        try:
            import configparser
            src = resolve_input(input_path)
            if not output_path:
                output_path = make_output_path(input_path, "", ".json")
            dst = resolve_output(output_path)
            cfg = configparser.ConfigParser()
            cfg.read(str(src), encoding="utf-8")
            data = {s: dict(cfg[s]) for s in cfg.sections()}
            dst.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
            return {"success": True, "output_path": str(dst), "sections": len(data)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @app.tool()
    def ini_to_yaml(input_path: str, output_path: Optional[str] = None) -> Dict:
        """Convert an INI config file to YAML."""
        try:
            import configparser, yaml
            src = resolve_input(input_path)
            if not output_path:
                output_path = make_output_path(input_path, "", ".yaml")
            dst = resolve_output(output_path)
            cfg = configparser.ConfigParser()
            cfg.read(str(src), encoding="utf-8")
            data = {s: dict(cfg[s]) for s in cfg.sections()}
            dst.write_text(yaml.dump(data, allow_unicode=True, default_flow_style=False), encoding="utf-8")
            return {"success": True, "output_path": str(dst), "sections": len(data)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @app.tool()
    def env_to_json(input_path: str, output_path: Optional[str] = None) -> Dict:
        """Convert a .env file (KEY=VALUE pairs) to JSON."""
        try:
            src = resolve_input(input_path)
            if not output_path:
                output_path = make_output_path(input_path, "", ".json")
            dst = resolve_output(output_path)
            data = {}
            for line in src.read_text(encoding="utf-8").splitlines():
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                if "=" in line:
                    k, _, v = line.partition("=")
                    data[k.strip()] = v.strip().strip('"').strip("'")
            dst.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
            return {"success": True, "output_path": str(dst), "keys": len(data)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ── CSV / Markdown tables ─────────────────────────────────────────────────
    @app.tool()
    def csv_to_markdown(input_path: str, output_path: Optional[str] = None, delimiter: str = ",") -> Dict:
        """Convert a CSV file to a Markdown table."""
        try:
            import pandas as pd
            src = resolve_input(input_path)
            if not output_path:
                output_path = make_output_path(input_path, "", ".md")
            dst = resolve_output(output_path)
            df = pd.read_csv(str(src), sep=delimiter)
            try:
                md = df.to_markdown(index=False)
            except Exception:
                lines = ["| " + " | ".join(str(c) for c in df.columns) + " |"]
                lines.append("| " + " | ".join(["---"] * len(df.columns)) + " |")
                for _, row in df.iterrows():
                    lines.append("| " + " | ".join(str(v) for v in row) + " |")
                md = "\n".join(lines)
            dst.write_text(md, encoding="utf-8")
            return {"success": True, "output_path": str(dst), "rows": len(df)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @app.tool()
    def html_table_to_csv(input_path: str, output_path: Optional[str] = None, table_index: int = 0) -> Dict:
        """
        Extract an HTML table and save as CSV.
        table_index: which table to extract (0-based, default first table).
        """
        try:
            import pandas as pd
            src = resolve_input(input_path)
            if not output_path:
                output_path = make_output_path(input_path, f"_table{table_index}", ".csv")
            dst = resolve_output(output_path)
            tables = pd.read_html(str(src))
            if not tables:
                return {"success": False, "error": "No tables found in the HTML file."}
            if table_index >= len(tables):
                return {"success": False, "error": f"Only {len(tables)} table(s) found; table_index {table_index} is out of range."}
            df = tables[table_index]
            df.to_csv(str(dst), index=False)
            return {"success": True, "output_path": str(dst), "rows": len(df), "tables_found": len(tables)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ── JSONL ──────────────────────────────────────────────────────────────
    @app.tool()
    def jsonl_to_json(input_path: str, output_path: Optional[str] = None) -> Dict:
        """Convert a JSONL (newline-delimited JSON) file to a JSON array."""
        try:
            src = resolve_input(input_path)
            if not output_path:
                output_path = make_output_path(input_path, "", ".json")
            dst = resolve_output(output_path)
            records = [json.loads(line) for line in src.read_text(encoding="utf-8").splitlines() if line.strip()]
            dst.write_text(json.dumps(records, indent=2, ensure_ascii=False), encoding="utf-8")
            return {"success": True, "output_path": str(dst), "records": len(records)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @app.tool()
    def json_to_jsonl(input_path: str, output_path: Optional[str] = None) -> Dict:
        """Convert a JSON array to JSONL (one object per line)."""
        try:
            src = resolve_input(input_path)
            if not output_path:
                output_path = make_output_path(input_path, "", ".jsonl")
            dst = resolve_output(output_path)
            data = json.loads(src.read_text(encoding="utf-8"))
            if not isinstance(data, list):
                data = [data]
            lines = [json.dumps(r, ensure_ascii=False) for r in data]
            dst.write_text("\n".join(lines) + "\n", encoding="utf-8")
            return {"success": True, "output_path": str(dst), "records": len(lines)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ── SQLite ───────────────────────────────────────────────────────────────
    @app.tool()
    def sqlite_to_csv(input_path: str, output_dir: Optional[str] = None, table_name: Optional[str] = None) -> Dict:
        """Export SQLite database tables to CSV files (one per table)."""
        try:
            import sqlite3, pandas as pd
            src = resolve_input(input_path)
            out_dir = Path(output_dir) if output_dir else src.parent / src.stem
            out_dir.mkdir(parents=True, exist_ok=True)

            conn = sqlite3.connect(str(src))
            if table_name:
                tables = [table_name]
            else:
                tables = [r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")]

            output_files = []
            for t in tables:
                df = pd.read_sql_query(f"SELECT * FROM [{t}]", conn)
                out = out_dir / f"{t}.csv"
                df.to_csv(str(out), index=False)
                output_files.append(str(out))
            conn.close()
            return {"success": True, "output_files": output_files, "tables": tables}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @app.tool()
    def sqlite_to_json(input_path: str, output_path: Optional[str] = None, table_name: Optional[str] = None) -> Dict:
        """Export SQLite database tables to a JSON file. Each table becomes a key."""
        try:
            import sqlite3, pandas as pd
            src = resolve_input(input_path)
            if not output_path:
                output_path = make_output_path(input_path, "", ".json")
            dst = resolve_output(output_path)

            conn = sqlite3.connect(str(src))
            if table_name:
                tables = [table_name]
            else:
                tables = [r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")]

            result = {}
            for t in tables:
                df = pd.read_sql_query(f"SELECT * FROM [{t}]", conn)
                result[t] = df.to_dict(orient="records")
            conn.close()
            dst.write_text(json.dumps(result, indent=2, ensure_ascii=False, default=str), encoding="utf-8")
            return {"success": True, "output_path": str(dst), "tables": tables}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @app.tool()
    def sqlite_to_xlsx(input_path: str, output_path: Optional[str] = None, table_name: Optional[str] = None) -> Dict:
        """Export SQLite database tables to Excel (one sheet per table)."""
        try:
            import sqlite3, pandas as pd
            src = resolve_input(input_path)
            if not output_path:
                output_path = make_output_path(input_path, "", ".xlsx")
            dst = resolve_output(output_path)

            conn = sqlite3.connect(str(src))
            if table_name:
                tables = [table_name]
            else:
                tables = [r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")]

            with pd.ExcelWriter(str(dst), engine="openpyxl") as writer:
                for t in tables:
                    df = pd.read_sql_query(f"SELECT * FROM [{t}]", conn)
                    df.to_excel(writer, sheet_name=t[:31], index=False)  # Excel sheet name limit: 31 chars
            conn.close()
            return {"success": True, "output_path": str(dst), "tables": tables}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ── Parquet ───────────────────────────────────────────────────────────────
    @app.tool()
    def parquet_to_csv(input_path: str, output_path: Optional[str] = None) -> Dict:
        """Convert an Apache Parquet file to CSV."""
        try:
            import pandas as pd
            src = resolve_input(input_path)
            if not output_path:
                output_path = make_output_path(input_path, "", ".csv")
            dst = resolve_output(output_path)
            df = pd.read_parquet(str(src))
            df.to_csv(str(dst), index=False)
            return {"success": True, "output_path": str(dst), "rows": len(df), "columns": len(df.columns)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @app.tool()
    def parquet_to_json(input_path: str, output_path: Optional[str] = None) -> Dict:
        """Convert an Apache Parquet file to JSON."""
        try:
            import pandas as pd
            src = resolve_input(input_path)
            if not output_path:
                output_path = make_output_path(input_path, "", ".json")
            dst = resolve_output(output_path)
            df = pd.read_parquet(str(src))
            dst.write_text(df.to_json(orient="records", indent=2), encoding="utf-8")
            return {"success": True, "output_path": str(dst), "rows": len(df)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @app.tool()
    def csv_to_parquet(input_path: str, output_path: Optional[str] = None) -> Dict:
        """Convert a CSV file to Apache Parquet format."""
        try:
            import pandas as pd
            src = resolve_input(input_path)
            if not output_path:
                output_path = make_output_path(input_path, "", ".parquet")
            dst = resolve_output(output_path)
            df = pd.read_csv(str(src))
            df.to_parquet(str(dst), index=False)
            return {"success": True, "output_path": str(dst), "rows": len(df), "file_size_bytes": dst.stat().st_size}
        except Exception as e:
            return {"success": False, "error": str(e)}

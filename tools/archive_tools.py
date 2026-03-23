"""Archive and compression tools: ZIP, TAR, 7Z."""
import os
import tarfile
import zipfile
from pathlib import Path
from typing import Dict, List, Optional

from mcp.server.fastmcp import FastMCP
from utils.file_utils import resolve_input, resolve_output


def register_archive_tools(app: FastMCP) -> None:

    @app.tool()
    def zip_files(
        input_paths: List[str],
        output_path: str,
        compression_level: int = 6,
    ) -> Dict:
        """
        Compress files or folders into a ZIP archive.

        input_paths: list of file or directory paths to include.
        compression_level: 0 (no compression) to 9 (maximum), default 6.
        """
        try:
            if not input_paths:
                return {"success": False, "error": "Provide at least one input path."}
            dst = resolve_output(output_path)
            compress = zipfile.ZIP_DEFLATED
            total_files = 0

            with zipfile.ZipFile(str(dst), "w", compression=compress, compresslevel=compression_level) as zf:
                for p in input_paths:
                    path = Path(p).resolve()
                    if not path.exists():
                        return {"success": False, "error": f"Path not found: {path}"}
                    if path.is_dir():
                        for file in path.rglob("*"):
                            if file.is_file():
                                zf.write(file, file.relative_to(path.parent))
                                total_files += 1
                    else:
                        zf.write(path, path.name)
                        total_files += 1

            return {
                "success": True,
                "output_path": str(dst),
                "files_added": total_files,
                "file_size_bytes": dst.stat().st_size,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    @app.tool()
    def unzip_files(
        input_path: str,
        output_dir: Optional[str] = None,
        password: Optional[str] = None,
    ) -> Dict:
        """
        Extract a ZIP archive to a directory.

        output_dir: destination directory (default: same dir as ZIP, folder named after archive).
        password: ZIP password if the archive is encrypted.
        """
        try:
            src = resolve_input(input_path)
            out_dir = Path(output_dir) if output_dir else src.parent / src.stem
            out_dir.mkdir(parents=True, exist_ok=True)

            pwd = password.encode() if password else None
            with zipfile.ZipFile(str(src), "r") as zf:
                zf.extractall(str(out_dir), pwd=pwd)
                names = zf.namelist()

            return {
                "success": True,
                "output_dir": str(out_dir),
                "files_extracted": len(names),
            }
        except zipfile.BadZipFile:
            return {"success": False, "error": "File is not a valid ZIP archive."}
        except RuntimeError as e:
            if "password" in str(e).lower():
                return {"success": False, "error": "ZIP is password-protected. Provide the 'password' parameter."}
            return {"success": False, "error": str(e)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @app.tool()
    def tar_files(
        input_paths: List[str],
        output_path: str,
        compression: str = "gz",
    ) -> Dict:
        """
        Compress files or folders into a TAR archive.

        compression: 'gz' (gzip, default), 'bz2' (bzip2), 'xz', or '' (uncompressed).
        output_path should end with .tar.gz / .tar.bz2 / .tar.xz / .tar accordingly.
        """
        try:
            if compression not in ("gz", "bz2", "xz", ""):
                return {"success": False, "error": "compression must be 'gz', 'bz2', 'xz', or ''."}
            if not input_paths:
                return {"success": False, "error": "Provide at least one input path."}

            dst = resolve_output(output_path)
            mode = f"w:{compression}" if compression else "w"
            total_files = 0

            with tarfile.open(str(dst), mode) as tf:
                for p in input_paths:
                    path = Path(p).resolve()
                    if not path.exists():
                        return {"success": False, "error": f"Path not found: {path}"}
                    tf.add(str(path), arcname=path.name)
                    total_files += sum(1 for _ in path.rglob("*")) if path.is_dir() else 1

            return {
                "success": True,
                "output_path": str(dst),
                "files_added": total_files,
                "file_size_bytes": dst.stat().st_size,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    @app.tool()
    def untar_files(
        input_path: str,
        output_dir: Optional[str] = None,
    ) -> Dict:
        """
        Extract a TAR archive (.tar, .tar.gz, .tar.bz2, .tar.xz) to a directory.
        """
        try:
            src = resolve_input(input_path)
            out_dir = Path(output_dir) if output_dir else src.parent / src.stem.replace(".tar", "")
            out_dir.mkdir(parents=True, exist_ok=True)

            with tarfile.open(str(src), "r:*") as tf:
                tf.extractall(str(out_dir))
                names = tf.getnames()

            return {
                "success": True,
                "output_dir": str(out_dir),
                "files_extracted": len(names),
            }
        except tarfile.TarError as e:
            return {"success": False, "error": f"Invalid TAR archive: {e}"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @app.tool()
    def create_7z(
        input_paths: List[str],
        output_path: str,
    ) -> Dict:
        """Compress files or folders into a 7Z archive."""
        try:
            import py7zr
            if not input_paths:
                return {"success": False, "error": "Provide at least one input path."}
            dst = resolve_output(output_path)
            total_files = 0

            with py7zr.SevenZipFile(str(dst), "w") as sz:
                for p in input_paths:
                    path = Path(p).resolve()
                    if not path.exists():
                        return {"success": False, "error": f"Path not found: {path}"}
                    if path.is_dir():
                        sz.writeall(str(path), path.name)
                        total_files += sum(1 for _ in path.rglob("*") if _.is_file())
                    else:
                        sz.write(str(path), path.name)
                        total_files += 1

            return {"success": True, "output_path": str(dst), "files_added": total_files, "file_size_bytes": dst.stat().st_size}
        except ImportError:
            return {"success": False, "error": "py7zr not installed. Run: pip install py7zr"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @app.tool()
    def extract_7z(
        input_path: str,
        output_dir: Optional[str] = None,
        password: Optional[str] = None,
    ) -> Dict:
        """Extract a 7Z archive."""
        try:
            import py7zr
            src = resolve_input(input_path)
            out_dir = Path(output_dir) if output_dir else src.parent / src.stem
            out_dir.mkdir(parents=True, exist_ok=True)

            kwargs = {"password": password} if password else {}
            with py7zr.SevenZipFile(str(src), "r", **kwargs) as sz:
                names = sz.getnames()
                sz.extractall(str(out_dir))

            return {"success": True, "output_dir": str(out_dir), "files_extracted": len(names)}
        except ImportError:
            return {"success": False, "error": "py7zr not installed. Run: pip install py7zr"}
        except Exception as e:
            return {"success": False, "error": str(e)}

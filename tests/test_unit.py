"""Fast tests that need no LibreOffice: config, paths, formats, tool selection, protocol."""

from __future__ import annotations

from pathlib import Path

import pytest

from libreoffice_mcp import formats, protocol, tools
from libreoffice_mcp.config import Settings, parse_settings
from libreoffice_mcp.errors import OfficeError, PathNotAllowedError
from libreoffice_mcp.paths import PathPolicy
from libreoffice_mcp.tools import registry

# -- formats -------------------------------------------------------------------


def test_extension_and_kind():
    assert formats.ext_of(r"C:\Docs\Report.FINAL.DOCX") == "docx"
    assert formats.ext_of("noext") == ""
    assert formats.kind_for("a.xlsx") == formats.CALC
    assert formats.kind_for("a.pptx") == formats.IMPRESS
    assert formats.kind_for("a.unknown") is None


def test_export_filters_respect_the_application():
    assert formats.export_filter(formats.WRITER, "pdf")[0] == "writer_pdf_Export"
    assert formats.export_filter(formats.WRITER, "docx")[0] == "MS Word 2007 XML"
    assert formats.export_filter(formats.CALC, "docx") is None  # a spreadsheet is not a document
    assert formats.export_filter(formats.IMPRESS, "pptx") is not None
    assert "pdf" in formats.supported_exports(formats.CALC)


# -- paths ---------------------------------------------------------------------


def test_relative_paths_use_default_dir(tmp_path: Path):
    policy = PathPolicy(tmp_path)
    assert policy.resolve("a/b.docx") == (tmp_path / "a" / "b.docx").resolve()


def test_roots_block_escapes(tmp_path: Path):
    root = tmp_path / "root"
    root.mkdir()
    (tmp_path / "secret.docx").write_text("x")
    policy = PathPolicy(root, (root,))
    assert policy.resolve("ok.docx").is_relative_to(root.resolve())
    with pytest.raises(PathNotAllowedError):
        policy.resolve("../secret.docx")
    with pytest.raises(PathNotAllowedError):
        policy.resolve(str(tmp_path / "secret.docx"))


def test_empty_and_missing_paths_are_clear_errors(tmp_path: Path):
    policy = PathPolicy(tmp_path)
    with pytest.raises(OfficeError, match="empty"):
        policy.resolve("  ")
    with pytest.raises(OfficeError, match=r"File not found.*relative paths"):
        policy.existing_file("nope.docx")
    (tmp_path / "x.exe").write_text("x")
    with pytest.raises(OfficeError, match="Unsupported file type"):
        policy.supported_file("x.exe")


# -- config --------------------------------------------------------------------


def test_defaults():
    _, s = parse_settings([])
    assert s.apps == formats.APPS and not s.read_only and s.transport == "stdio"


def test_apps_flag_and_unknown_app(capsys: pytest.CaptureFixture[str]):
    _, s = parse_settings(["--apps", "writer, impress"])
    assert s.apps == ("writer", "impress")
    with pytest.raises(SystemExit):
        parse_settings(["--apps", "writer,sheets"])
    assert "unknown app" in capsys.readouterr().err


def test_env_vars_are_fallbacks_and_flags_win(monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setenv("LIBREOFFICE_MCP_APPS", "calc")
    monkeypatch.setenv("LIBREOFFICE_MCP_READ_ONLY", "true")
    _, s = parse_settings([])
    assert s.apps == ("calc",) and s.read_only
    _, s = parse_settings(["--apps", "writer"])
    assert s.apps == ("writer",)


def test_roots_are_resolved(tmp_path: Path):
    _, s = parse_settings(["--root", str(tmp_path), "--default-dir", str(tmp_path)])
    assert s.roots == (tmp_path.resolve(),)


# -- tool selection ---------------------------------------------------------------


@pytest.fixture(scope="module")
def all_tools():
    tools.load_modules(formats.APPS)
    return registry.TOOLS


def names(settings: Settings) -> set[str]:
    return {t.name for t in registry.select(settings)}


def test_core_tools_always_present(all_tools):
    assert "office_create_document" in names(Settings(apps=("writer",)))
    assert "office_create_document" in names(Settings(apps=()))


def test_apps_filter_hides_other_apps(all_tools):
    only_writer = names(Settings(apps=("writer",)))
    assert "writer_add_paragraph" in only_writer
    assert not any(n.startswith(("calc_", "impress_")) for n in only_writer)
    assert not any(n.startswith("writer_") for n in names(Settings(apps=("calc",))))


def test_read_only_exposes_no_mutating_tool(all_tools):
    selected = registry.select(Settings(read_only=True))
    assert selected and all(t.mode == "read" for t in selected)
    assert "writer_get_text" in {t.name for t in selected}
    assert "writer_add_paragraph" not in {t.name for t in selected}


def test_include_exclude_globs(all_tools):
    assert names(Settings(include=("writer_get_*",))) <= {
        n for n in names(Settings()) if n.startswith("writer_get_")
    }
    s = names(Settings(exclude=("*delete*", "*replace*")))
    assert (
        "writer_delete_blocks" not in s
        and "writer_replace_text" not in s
        and "writer_add_heading" in s
    )


def test_every_tool_has_a_description_and_unique_name(all_tools):
    seen = set()
    for t in all_tools:
        assert t.name not in seen
        seen.add(t.name)
        assert (t.fn.__doc__ or "").strip(), f"{t.name} needs a docstring (the model reads it)"
        assert t.app in ("core", *formats.APPS)


# -- protocol --------------------------------------------------------------------


def test_protocol_round_trip_including_unicode():
    msg = {"id": 1, "text": "héllo 日本語\nline2"}
    wire = protocol.encode(msg)
    assert wire.endswith(b"\n") and wire.count(b"\n") == 1  # newlines inside strings are escaped
    assert protocol.decode(wire) == msg


def test_protocol_rejects_non_objects():
    with pytest.raises(ValueError):
        protocol.decode("[1, 2]")

from plugins.loader import load_all, get


def test_loader():
    reg = load_all()
    assert "send_pdf" in reg and "file_edit" in reg


def test_file_edit(tmp_path):
    p = tmp_path / "demo.txt"
    p.write_text("alpha beta alpha", encoding="utf-8")
    out = get("file_edit").run(path=str(p), find="alpha", replace="ALPHA", backup=True)
    assert out["ok"] and out["replacements"] == 2
    assert p.read_text(encoding="utf-8").count("ALPHA") == 2

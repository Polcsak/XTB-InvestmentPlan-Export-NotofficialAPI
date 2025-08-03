import ast
import types
from pathlib import Path
import os
import shutil

# Extract process_export function without importing heavy dependencies
source_path = Path(__file__).resolve().parents[1] / "XTB-script.py"
source = source_path.read_text()
module = types.ModuleType("xtb_script")
for node in ast.parse(source).body:
    if isinstance(node, ast.FunctionDef) and node.name == "process_export":
        code = compile(ast.Module([node], []), filename=str(source_path), mode="exec")
        exec(code, module.__dict__)
        break
xtb_script = module
xtb_script.os = os
xtb_script.shutil = shutil


class FakeSheet:
    def __init__(self, title):
        self.title = title


class FakeWorkbook:
    def __init__(self):
        sheet = FakeSheet("OPEN POSITION test")
        self._sheets = {sheet.title: sheet}

    @property
    def sheetnames(self):
        return list(self._sheets.keys())

    def __getitem__(self, name):
        return self._sheets[name]

    def save(self, path):
        # Update the sheetnames mapping in case titles changed
        self._sheets = {s.title: s for s in self._sheets.values()}


def test_process_export(tmp_path):
    downloads = tmp_path / "downloads"
    target = tmp_path / "target"
    downloads.mkdir()
    target.mkdir()

    sample_file = downloads / "account_50266286_test.xlsx"
    sample_file.write_bytes(b"")

    captured = {}

    def fake_load_workbook(path):
        wb = FakeWorkbook()
        captured["wb"] = wb
        return wb

    xtb_script.load_workbook = fake_load_workbook

    xtb_script.process_export(str(downloads), str(target))

    moved_file = target / "xtb_export.xlsx"
    assert moved_file.exists()
    assert "ExportXTB" in captured["wb"].sheetnames

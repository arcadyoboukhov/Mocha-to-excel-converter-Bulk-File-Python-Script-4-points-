from __future__ import annotations

import shutil
from pathlib import Path

import win32com.client as win32


LAYER_CONFIG = [
	{
		"control": "Layer_1_-_Spline_1_-_Control_Point",
		"tx": "Layer_1.Track.Translation_X.keyframes.append( Key(",
		"ty": "Layer_1.Track.Translation_Y.keyframes.append( Key(",
		"dest_x": "BD",
		"dest_y": "BE",
		"x_from": "AT",
		"copy_y": True,
	},
	{
		"control": "Layer_2_-_Spline_2_-_Control_Point",
		"tx": "Layer_2.Track.Translation_X.keyframes.append( Key(",
		"ty": "Layer_2.Track.Translation_Y.keyframes.append( Key(",
		"dest_x": "BG",
		"dest_y": "BH",
		"x_from": "AT",
		"copy_y": True,
	},
	{
		"control": "Layer_3_-_Spline_3_-_Control_Point",
		"tx": "Layer_3.Track.Translation_X.keyframes.append( Key(",
		"ty": "Layer_3.Track.Translation_Y.keyframes.append( Key(",
		"dest_x": "BJ",
		"dest_y": "BK",
		"x_from": "AT",
		"copy_y": True,
	},
	{
		"control": "Layer_4_-_Spline_4_-_Control_Point",
		"tx": "Layer_4.Track.Translation_X.keyframes.append( Key(",
		"ty": "Layer_4.Track.Translation_Y.keyframes.append( Key(",
		"dest_x": "BM",
		"dest_y": "BN",
		"x_from": "AT",
		"copy_y": True,
	},
]


def find_template(base_dir: Path) -> Path:
	direct = base_dir / "template.xlsx"
	if direct.exists():
		return direct

	matches = sorted(base_dir.rglob("template.xlsx"))
	if not matches:
		raise FileNotFoundError(f"Could not find template.xlsx in {base_dir} or subfolders")
	return matches[0]


def find_mocha_files(base_dir: Path) -> list[Path]:
	return sorted(base_dir.glob("*.mocha"))


def extract_matching_lines(text: str, marker: str) -> list[str]:
	return [line.rstrip("\r\n") for line in text.splitlines() if marker in line]


def clear_column_between(ws, left_col: str, right_col: str, max_row: int) -> None:
	ws.Range(f"{left_col}1:{right_col}{max_row}").ClearContents()


def write_and_split_lines(ws, start_col: str, lines: list[str], max_row: int) -> None:
	ws.Range(f"{start_col}1:{start_col}{max_row}").ClearContents()
	if not lines:
		return

	values = [[line] for line in lines]
	ws.Range(f"{start_col}1:{start_col}{len(lines)}").Value = values
	ws.Range(f"{start_col}1:{start_col}{len(lines)}").TextToColumns(
		Destination=ws.Range(f"{start_col}1"),
		DataType=1,
		TextQualifier=1,
		ConsecutiveDelimiter=True,
		Tab=True,
		Semicolon=True,
		Comma=True,
		Space=True,
		Other=True,
		OtherChar="_",
	)


def find_last_filled_row(ws, col: str) -> int:
	return ws.Cells(ws.Rows.Count, col).End(-4162).Row


def copy_values(ws, source_col: str, dest_col: str, last_row: int) -> None:
	ws.Range(f"{source_col}1:{source_col}{last_row}").Copy()
	ws.Range(f"{dest_col}1").PasteSpecial(Paste=-4163)


def process_single_mocha(mocha_path: Path, template_path: Path, excel) -> Path:
	output_path = mocha_path.with_suffix(".xlsx")
	shutil.copy2(template_path, output_path)

	workbook = excel.Workbooks.Open(str(output_path.resolve()))
	ws = workbook.Worksheets(1)
	text = mocha_path.read_text(encoding="utf-8", errors="ignore")

	formula_last_row = find_last_filled_row(ws, "AT")
	max_clear_row = max(formula_last_row, 8000)

	for idx, layer in enumerate(LAYER_CONFIG):
		ty_marker = layer["ty"]
		if (
			mocha_path.name.endswith(" - 2.mp4.mocha")
			and layer["control"] == "Layer_2_-_Spline_2_-_Control_Point"
		):
			ty_marker = "Layer_1.Track.Translation_Y.keyframes.append( Key("

		control_lines = extract_matching_lines(text, layer["control"])
		tx_lines = extract_matching_lines(text, layer["tx"])
		ty_lines = extract_matching_lines(text, ty_marker)

		write_and_split_lines(ws, "A", control_lines, max_clear_row)
		write_and_split_lines(ws, "X", tx_lines, max_clear_row)
		write_and_split_lines(ws, "AH", ty_lines, max_clear_row)

		excel.CalculateFull()

		copy_values(ws, layer.get("x_from", "AT"), layer["dest_x"], formula_last_row)
		if layer.get("copy_y", True):
			copy_values(ws, "AU", layer["dest_y"], formula_last_row)

		if idx < len(LAYER_CONFIG) - 1:
			clear_column_between(ws, "B", "P", max_clear_row)
			clear_column_between(ws, "Y", "AE", max_clear_row)
			clear_column_between(ws, "AI", "AQ", max_clear_row)

	excel.CutCopyMode = False
	workbook.Save()
	workbook.Close(SaveChanges=True)
	return output_path


def main() -> None:
	base_dir = Path(__file__).resolve().parent
	template_path = find_template(base_dir)
	mocha_files = find_mocha_files(base_dir)

	if not mocha_files:
		print(f"No .mocha files found in {base_dir}")
		return

	print(f"Using template: {template_path}")
	print(f"Found {len(mocha_files)} .mocha files")

	excel = win32.DispatchEx("Excel.Application")
	excel.Visible = False
	excel.DisplayAlerts = False

	try:
		for mocha_file in mocha_files:
			output_path = process_single_mocha(mocha_file, template_path, excel)
			print(f"Created: {output_path.name}")
	finally:
		excel.Quit()


if __name__ == "__main__":
	main()

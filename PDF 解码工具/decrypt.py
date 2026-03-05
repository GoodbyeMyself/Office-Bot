import os
from pathlib import Path
import tkinter as tk
from tkinter import filedialog

import pikepdf


def build_temp_path(output_path: Path) -> Path:
    """在输出目录中生成临时文件路径。"""
    # 先写入临时文件，成功后再替换目标文件，避免直接写坏目标文件。
    base = output_path.with_name(f"{output_path.stem}.tmp_unlocked{output_path.suffix}")
    if not base.exists():
        return base

    index = 1
    while True:
        # 若临时文件重名，追加序号避免冲突。
        candidate = output_path.with_name(
            f"{output_path.stem}.tmp_unlocked_{index}{output_path.suffix}"
        )
        if not candidate.exists():
            return candidate
        index += 1


def main() -> int:
    print("Please choose a PDF file.")

    # 仅显示文件选择窗口，不显示主窗口。
    root = tk.Tk()
    root.withdraw()
    file_path_str = filedialog.askopenfilename(
        title="Select PDF File",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
    )
    root.destroy()

    if not file_path_str:
        print("No file selected.")
        os.system("pause")
        return 1

    input_path = Path(file_path_str)
    # 解密后的文件统一输出到脚本同级 out 目录。
    output_dir = Path(__file__).resolve().parent / "out"
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / input_path.name
    temp_path = build_temp_path(output_path)

    try:
        with pikepdf.open(input_path) as pdf:
            pdf.save(temp_path)
        # 保存成功后再原子替换目标文件。
        temp_path.replace(output_path)
    except Exception as exc:
        # 异常时尽力清理临时文件。
        if temp_path.exists():
            try:
                temp_path.unlink()
            except OSError:
                pass
        print(f"Decrypt failed: {exc}")
        os.system("pause")
        return 1

    print(f"Done. Output file: {output_path}")
    os.system("pause")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

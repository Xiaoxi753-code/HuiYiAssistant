import json
import os
import shutil
import subprocess
import sys
import uuid
from pathlib import Path

try:
    from docxtpl import DocxTemplate
except ModuleNotFoundError:
    print("{\"error\": \"缺少依赖 docxtpl，请在 backend 目录执行 pip install docxtpl\"}")
    sys.exit(1)


def main():
    if len(sys.argv) < 5:
        print("{\"error\": \"Missing arguments\"}")
        sys.exit(1)

    template_path = Path(sys.argv[1])
    variables_json = sys.argv[2]
    output_dir = Path(sys.argv[3])
    soffice_path = Path(sys.argv[4])

    if not template_path.exists():
        print("{\"error\": \"Template not found\"}")
        sys.exit(1)

    if not soffice_path.exists():
        print(
            "{\"error\": \"LibreOffice 路径无效，请检查 SOFFICE_PATH 是否指向 soffice.exe\"}"
        )
        sys.exit(1)

    output_dir.mkdir(parents=True, exist_ok=True)
    temp_dir = output_dir / "tmp"
    temp_dir.mkdir(parents=True, exist_ok=True)

    unique_id = uuid.uuid4().hex
    filled_docx = temp_dir / f"filled_{unique_id}.docx"
    pdf_output_dir = temp_dir / f"pdf_{unique_id}"
    pdf_output_dir.mkdir(parents=True, exist_ok=True)

    try:
        variables = json.loads(variables_json) if variables_json else {}
    except json.JSONDecodeError:
        print("{\"error\": \"Invalid variables JSON\"}")
        sys.exit(1)

    try:
        doc = DocxTemplate(str(template_path))
        doc.render(variables)
        doc.save(str(filled_docx))

        cmd = [
            str(soffice_path),
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(pdf_output_dir),
            str(filled_docx),
        ]
        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        pdf_name = filled_docx.with_suffix(".pdf").name
        pdf_path = pdf_output_dir / pdf_name
        if not pdf_path.exists():
            print("{\"error\": \"PDF output missing\"}")
            sys.exit(1)

        final_pdf = output_dir / f"render_{unique_id}.pdf"
        shutil.move(str(pdf_path), str(final_pdf))

        print(json.dumps({"pdfPath": str(final_pdf)}))
    finally:
        if filled_docx.exists():
            try:
                filled_docx.unlink()
            except OSError:
                pass

        if pdf_output_dir.exists():
            shutil.rmtree(pdf_output_dir, ignore_errors=True)


if __name__ == "__main__":
    main()

import os
import random
import zipfile
import tempfile
import subprocess
import json
from pathlib import Path
from datetime import datetime, date

import requests
from flask import Flask, render_template, request, send_file, Response
from docx import Document
from docx.shared import Pt
from docxcompose.composer import Composer

app = Flask(__name__)

BANK_PATH = Path("bank")
GIST_ID = os.environ.get("GIST_ID")
GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN")

if not GIST_ID or not GITHUB_TOKEN:
    raise RuntimeError("Set GIST_ID and GITHUB_TOKEN environment variables")

API_URL = f"https://api.github.com/gists/{GIST_ID}"
HEADERS = {
    "Authorization": f"token {GITHUB_TOKEN}",
    "Accept": "application/vnd.github.v3+json"
}

# ---------- Счётчик и статистика ----------
def get_counter_data():
    resp = requests.get(API_URL, headers=HEADERS)
    resp.raise_for_status()
    content = resp.json()["files"]["counter.json"]["content"]
    data = json.loads(content)
    if "history" not in data:
        data["history"] = {}
    if "value" not in data:
        data["value"] = 0
    return data

def update_counter_data(new_data):
    payload = {"files": {"counter.json": {"content": json.dumps(new_data)}}}
    requests.patch(API_URL, headers=HEADERS, json=payload).raise_for_status()

def get_global_counter(k):
    data = get_counter_data()
    current = data["value"]
    today = date.today().isoformat()
    data["value"] = current + k
    history = data["history"]
    history[today] = history.get(today, 0) + k
    update_counter_data(data)
    return current + 1

def get_stats():
    data = get_counter_data()
    today = date.today().isoformat()
    today_count = data["history"].get(today, 0)
    current_month = today[:7]
    month_count = sum(
        count for day, count in data["history"].items()
        if day.startswith(current_month)
    )
    total = data["value"]
    return {"today": today_count, "month": month_count, "total": total}

# ---------- Конвертация docx → pdf (LibreOffice) ----------
def docx_to_pdf(docx_path, output_pdf_path):
    cmd = [
        "libreoffice", "--headless", "--convert-to", "pdf",
        "--outdir", str(output_pdf_path.parent),
        str(docx_path)
    ]
    for attempt in range(2):
        try:
            subprocess.run(cmd, check=True, timeout=120)
            expected_pdf = output_pdf_path.parent / (docx_path.stem + ".pdf")
            if expected_pdf.exists():
                os.rename(expected_pdf, output_pdf_path)
                return
        except Exception:
            if attempt == 1:
                raise RuntimeError("LibreOffice conversion failed")

# ---------- Заголовки вариантов ----------
def determine_header(selected_tasks):
    if not selected_tasks:
        return "Вариант №"
    if set(selected_tasks) == set(range(1, 19)):
        return "Вариант экзаменационной работы №"
    if set(selected_tasks) == set(range(1, 14)):
        return "Вариант № (часть 1, задания 1–13)"
    if set(selected_tasks) == set(range(14, 19)):
        return "Вариант № (часть 2, задания 14–18)"
    if set(selected_tasks) == set(list(range(1, 10)) + list(range(14, 17))):
        return "Вариант № (Алгебра)"
    if set(selected_tasks) == set(list(range(10, 14)) + list(range(17, 19))):
        return "Вариант № (Геометрия)"
    nums = sorted(selected_tasks)
    return f"Вариант № (задания {', '.join(map(str, nums))})"

def add_title_paragraph(master, text):
    para = master.add_paragraph()
    run = para.add_run(text)
    run.bold = True
    run.font.size = Pt(14)
    return para

# ---------- Генерация вариантов ----------
def compose_variants(k, selected_tasks, preview=False, merge=False):
    task_dirs = sorted(
        [d for d in BANK_PATH.iterdir() if d.is_dir() and d.name.startswith("task_")],
        key=lambda x: int(x.name.split("_")[1])
    )
    filtered_dirs = [d for d in task_dirs if int(d.name.split("_")[1]) in selected_tasks]
    if not filtered_dirs:
        raise Exception("Не выбрано ни одного задания")

    all_tasks = []
    for task_dir in filtered_dirs:
        docx_files = sorted(
            [f for f in task_dir.glob("variant_*.docx")],
            key=lambda x: int(x.stem.split("_")[1])
        )
        if not docx_files:
            raise Exception(f"В папке {task_dir.name} нет docx-файлов")
        random.shuffle(docx_files)
        all_tasks.append(docx_files)

    if not preview:
        start_global_id = get_global_counter(k)
    else:
        start_global_id = 0

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)
        result_files = []

        for exam_idx in range(k):
            global_id = start_global_id + exam_idx
            master = Document()
            composer = Composer(master)

            title = f"{determine_header(selected_tasks)} {global_id}"
            if preview:
                title += " (ПРИМЕР)"
            add_title_paragraph(master, title)

            for variants_list in all_tasks:
                variant_file = variants_list[exam_idx % len(variants_list)]
                sub_doc = Document(str(variant_file))
                composer.append(sub_doc)

            docx_path = tmpdir_path / f"variant_{global_id}.docx"
            master.save(str(docx_path))

            if merge:
                result_files.append(str(docx_path))
            else:
                pdf_path = tmpdir_path / f"variant_{global_id}.pdf"
                docx_to_pdf(str(docx_path), str(pdf_path))
                result_files.append(str(pdf_path))

        if merge:
            merged_doc = Document()
            merged_composer = Composer(merged_doc)
            for doc_path in result_files:
                merged_composer.append(Document(doc_path))
            merged_docx = tmpdir_path / "merged.docx"
            merged_doc.save(str(merged_docx))
            merged_pdf = tmpdir_path / "merged.pdf"
            docx_to_pdf(str(merged_docx), str(merged_pdf))
            final_path = Path(tempfile.gettempdir()) / f"variants_{datetime.now().strftime('%Y%m%d%H%M%S')}.pdf"
            merged_pdf.rename(final_path)
            return final_path
        else:
            zip_path = tmpdir_path / "variants.zip"
            with zipfile.ZipFile(zip_path, "w") as zf:
                for file_path in result_files:
                    zf.write(file_path, arcname=Path(file_path).name)
            final_zip = Path(tempfile.gettempdir()) / f"variants_{datetime.now().strftime('%Y%m%d%H%M%S')}.zip"
            zip_path.rename(final_zip)
            return final_zip

# ---------- Маршруты ----------
@app.route("/")
def index():
    stats = {}
    try:
        stats = get_stats()
    except Exception as e:
        app.logger.error(f"Stats error: {e}")
        stats = {"today": "—", "month": "—", "total": "—"}
    return render_template("index.html", stats=stats)

@app.route("/generate", methods=["POST"])
def generate():
    try:
        k = int(request.form.get("quantity", 0))
        merge = request.form.get("merge") == "1"
        if k < 1 or k > 50:
            return "Количество вариантов должно быть от 1 до 50", 400
        selected_tasks = [int(t) for t in request.form.getlist("task")]
        if not selected_tasks:
            return "Не выбрано ни одного задания", 400
        if not all(1 <= t <= 18 for t in selected_tasks):
            return "Некорректный номер задания", 400

        result_file = compose_variants(k, selected_tasks, preview=False, merge=merge)
        if merge:
            return send_file(result_file, as_attachment=True,
                             download_name="exam_variants.pdf", mimetype="application/pdf")
        else:
            return send_file(result_file, as_attachment=True,
                             download_name="exam_variants.zip")
    except Exception as e:
        app.logger.error(f"Generate error: {e}")
        return f"Ошибка при генерации: {e}", 500

@app.route("/preview", methods=["POST"])
def preview():
    try:
        selected_tasks = [int(t) for t in request.form.getlist("task")]
        if not selected_tasks:
            return "Не выбрано ни одного задания", 400
        if not all(1 <= t <= 18 for t in selected_tasks):
            return "Некорректный номер задания", 400
        result_file = compose_variants(1, selected_tasks, preview=True, merge=False)
        with zipfile.ZipFile(result_file, "r") as zf:
            pdf_name = zf.namelist()[0]
            pdf_data = zf.read(pdf_name)
        return Response(pdf_data, mimetype="application/pdf",
                        headers={"Content-Disposition": "inline; filename=preview.pdf"})
    except Exception as e:
        app.logger.error(f"Preview error: {e}")
        return f"Ошибка при предпросмотре: {e}", 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000, debug=False)
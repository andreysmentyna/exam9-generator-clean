import os
import random
import zipfile
import tempfile
import io
import json
from pathlib import Path
from datetime import datetime, date

import requests
from flask import Flask, render_template, request, send_file, Response
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

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

# ---------- Счётчик и статистика (GitHub Gist) ----------
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
    """Увеличивает счётчик на k и возвращает первый номер для текущей пачки."""
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

# ---------- Колонтитулы и заголовки ----------
def add_header_to_pdf(input_pdf_path, output_pdf_path, header_text):
    """Добавляет верхний колонтитул с header_text на каждую страницу."""
    reader = PdfReader(input_pdf_path)
    writer = PdfWriter()
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=A4)
    can.setFont("Helvetica", 10)
    can.drawCentredString(A4[0] / 2, A4[1] - 15 * mm, header_text)
    can.save()
    packet.seek(0)
    header_pdf = PdfReader(packet)
    header_page = header_pdf.pages[0]

    for page in reader.pages:
        page.merge_page(header_page)
        writer.add_page(page)

    with open(output_pdf_path, "wb") as f:
        writer.write(f)

def determine_header(selected_tasks):
    """Возвращает текст колонтитула в зависимости от выбранных заданий."""
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

# ---------- Объединение PDF ----------
def merge_pdfs(pdf_paths, output_path):
    """Объединяет PDF-файлы, вставляя пустую страницу между вариантами."""
    writer = PdfWriter()
    for i, path in enumerate(pdf_paths):
        reader = PdfReader(path)
        for page in reader.pages:
            writer.add_page(page)
        if i < len(pdf_paths) - 1:          # не после последнего
            writer.add_blank_page(width=A4[0], height=A4[1])
    with open(output_path, "wb") as f:
        writer.write(f)

# ---------- Генерация вариантов ----------
def compose_variants(k, selected_tasks, preview=False, merge=False):
    """
    Собирает k экзаменационных вариантов.
    preview – не увеличивает счётчик,
    merge – возвращает единый PDF вместо ZIP.
    """
    # Список папок заданий, отфильтрованный по выбранным номерам
    task_dirs = sorted(
        [d for d in BANK_PATH.iterdir() if d.is_dir() and d.name.startswith("task_")],
        key=lambda x: int(x.name.split("_")[1])
    )
    filtered_dirs = [d for d in task_dirs if int(d.name.split("_")[1]) in selected_tasks]
    if not filtered_dirs:
        raise Exception("Не выбрано ни одного задания")

    # Загружаем списки PDF-файлов по каждому заданию
    all_tasks = []
    for task_dir in filtered_dirs:
        pdf_files = sorted(
            [f for f in task_dir.glob("variant_*.pdf")],
            key=lambda x: int(x.stem.split("_")[1])
        )
        if not pdf_files:
            raise Exception(f"В папке {task_dir.name} нет PDF-файлов")
        random.shuffle(pdf_files)
        all_tasks.append(pdf_files)

    # Начальный глобальный номер (если не превью, увеличиваем счётчик)
    if not preview:
        start_global_id = get_global_counter(k)
    else:
        start_global_id = 0

    header_base = determine_header(selected_tasks)

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)
        final_pdfs = []

        for exam_idx in range(k):
            global_id = start_global_id + exam_idx
            writer = PdfWriter()

            # Добавляем страницы по одному заданию из каждого списка
            for variants_list in all_tasks:
                variant_file = variants_list[exam_idx % len(variants_list)]
                reader = PdfReader(str(variant_file))
                for page in reader.pages:
                    writer.add_page(page)

            # Временный PDF без колонтитула
            tmp_pdf = tmpdir_path / f"tmp_{global_id}.pdf"
            with open(tmp_pdf, "wb") as f:
                writer.write(f)

            # Накладываем колонтитул
            header_text = f"{header_base} {global_id}" if not preview else f"{header_base} (ПРИМЕР)"
            final_pdf = tmpdir_path / f"variant_{global_id}.pdf"
            add_header_to_pdf(str(tmp_pdf), str(final_pdf), header_text)
            final_pdfs.append(final_pdf)

        # ----- Финальная сборка: ZIP или объединённый PDF -----
        if merge:
            merged_pdf = tmpdir_path / "merged_variants.pdf"
            merge_pdfs(final_pdfs, str(merged_pdf))
            # Переносим в постоянное место
            final_merged = Path(tempfile.gettempdir()) / f"variants_{datetime.now().strftime('%Y%m%d%H%M%S')}.pdf"
            merged_pdf.rename(final_merged)
            return final_merged
        else:
            zip_path = tmpdir_path / "variants.zip"
            with zipfile.ZipFile(zip_path, "w") as zf:
                for pdf in final_pdfs:
                    zf.write(pdf, arcname=pdf.name)
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
            return send_file(
                result_file,
                as_attachment=True,
                download_name="exam_variants.pdf",
                mimetype="application/pdf"
            )
        else:
            return send_file(
                result_file,
                as_attachment=True,
                download_name="exam_variants.zip"
            )
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
        # preview всегда один вариант, без объединения
        zip_file = compose_variants(1, selected_tasks, preview=True, merge=False)
        with zipfile.ZipFile(zip_file, "r") as zf:
            pdf_name = zf.namelist()[0]
            pdf_data = zf.read(pdf_name)
        return Response(pdf_data, mimetype="application/pdf",
                        headers={"Content-Disposition": "inline; filename=preview.pdf"})
    except Exception as e:
        app.logger.error(f"Preview error: {e}")
        return f"Ошибка при предпросмотре: {e}", 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000, debug=False)
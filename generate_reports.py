from openpyxl import load_workbook
from datetime import datetime, timedelta
from pathlib import Path
from docx import Document

INPUT_FILE = "input/TranVanAn.xlsx"
OUTPUT_DIR = Path("Generate")
OUTPUT_DIR.mkdir(exist_ok=True)

TEMPLATE_DIR = Path("templates")

# Set the reference date and its Day number for the Fresher counter.
# Example: 28/04/2026 is Day 69, then 29/04/2026 becomes Day 70, etc.
DAY_COUNT_REFERENCE_DATE = datetime(2026, 4, 28)
DAY_COUNT_REFERENCE_NUMBER = 69


def load_template(template_name):
    return (TEMPLATE_DIR / template_name).read_text(encoding="utf-8")


def normalize_percent(percent):
    if percent is None:
        return None

    if isinstance(percent, str):
        try:
            return float(percent.strip().rstrip("%"))
        except ValueError:
            return None

    if isinstance(percent, (int, float)):
        return float(percent)

    return None


def format_percent_value(percent):
    normalized = normalize_percent(percent)
    if normalized is None:
        return "0-100%"
    if normalized >= 100:
        return "100%"
    if normalized.is_integer():
        normalized = int(normalized)
    return f"{normalized}-100%"


def is_task_complete(percent):
    normalized = normalize_percent(percent)
    return normalized is not None and normalized >= 100


# ========================
# Utils
# ========================
def excel_date_to_datetime(value):
    if value is None:
        return None

    if isinstance(value, datetime):
        return value

    if isinstance(value, (int, float)):
        base_date = datetime(1899, 12, 30)
        return base_date + timedelta(days=int(value))

    if isinstance(value, str):
        value = value.strip()

        for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(value, fmt)
            except ValueError:
                pass

    return None


def date_to_display(date_obj):
    return date_obj.strftime("%d/%m/%Y")


def date_to_folder(date_obj):
    return date_obj.strftime("%Y-%m-%d")


def normalize_percent(percent):
    if percent is None:
        return None

    if isinstance(percent, str):
        try:
            return float(percent.strip().rstrip("%"))
        except ValueError:
            return None

    if isinstance(percent, (int, float)):
        return float(percent)

    return None


def format_percent_value(percent):
    normalized = normalize_percent(percent)
    if normalized is None:
        return "0-100%"
    if normalized >= 100:
        return "100%"
    if normalized.is_integer():
        normalized = int(normalized)
    return f"{normalized}-100%"


def is_task_complete(percent):
    normalized = normalize_percent(percent)
    return normalized is not None and normalized >= 100


def date_to_day_number(date_obj):
    if date_obj is None:
        return None

    return DAY_COUNT_REFERENCE_NUMBER + (date_obj.date() - DAY_COUNT_REFERENCE_DATE.date()).days


def get_week_start(date_obj):
    # Monday is weekday 0
    return date_obj - timedelta(days=date_obj.weekday())


def get_daily_dir(date_obj):
    week_start = get_week_start(date_obj)
    week_folder = f"week-{date_to_folder(week_start)}"
    day_folder = date_to_folder(date_obj)

    folder = OUTPUT_DIR / week_folder / day_folder
    folder.mkdir(parents=True, exist_ok=True)

    return folder


def clean_text(value):
    if value is None:
        return ""

    return str(value).strip()


# ========================
# Read Excel
# ========================
def read_wbs(wb):
    ws = wb["WBS"]
    tasks = {}

    for row in ws.iter_rows(min_row=6, values_only=True):
        task_id = clean_text(row[1])
        task_name = clean_text(row[2])

        if not task_id:
            continue

        tasks[task_id] = {
            "task_id": task_id,
            "task_name": task_name,
            "estimate_h": row[3],
            "actual_h": row[6],
            "percent": row[9],
            "note": clean_text(row[10]),
        }

    return tasks


def read_daily(wb):
    ws = wb["Report Daily"]

    date_cols = {}

    for col in range(4, ws.max_column + 1):
        date_val = ws.cell(row=1, column=col).value
        date_obj = excel_date_to_datetime(date_val)

        if date_obj:
            date_cols[col] = date_obj

    data = []

    for row in range(5, ws.max_row + 1):
        task_id = clean_text(ws.cell(row=row, column=2).value)
        task_name = clean_text(ws.cell(row=row, column=3).value)

        if not task_id:
            continue

        for col, date_obj in date_cols.items():
            content = clean_text(ws.cell(row=row, column=col).value)

            # Important:
            # Nếu ngày đó không có nội dung trong Excel thì bỏ qua,
            # không tạo folder/file cho ngày đó.
            if not content:
                continue

            data.append({
                "date_obj": date_obj,
                "date": date_to_display(date_obj),
                "folder_date": date_to_folder(date_obj),
                "task_id": task_id,
                "task_name": task_name,
                "content": content,
            })

    return data


# ========================
# Generate TXT Reports
# ========================
def generate_reports_by_date(daily_data, tasks):
    grouped_by_date = {}

    for item in daily_data:
        key = item["folder_date"]
        grouped_by_date.setdefault(key, []).append(item)

    sorted_dates = sorted(grouped_by_date.keys())

    for i, date_key in enumerate(sorted_dates):
        items = grouped_by_date[date_key]

        # Group by task_id to prevent duplicates
        task_dict = {}
        for item in items:
            tid = item["task_id"]
            task_dict[tid] = item  # Keep the last if duplicates
        items = list(task_dict.values())

        date_obj = items[0]["date_obj"]
        date_display = items[0]["date"]

        # Build next-day work list from current unfinished tasks
        next_tasks_list = []
        for item in items:
            task = tasks.get(item["task_id"], {})
            percent = task.get("percent")
            if not is_task_complete(percent):
                next_tasks_list.append(f"- {item['task_name']}. `{format_percent_value(percent)}`")
        next_tasks = "\n".join(next_tasks_list) if next_tasks_list else "- N/A"

        daily_dir = get_daily_dir(date_obj)

        day_number = date_to_day_number(date_obj)
        fresher_day = f" Day {day_number}" if day_number is not None else ""

        eod_template = load_template("bao_cao_cuoi_ngay.txt")
        du2_template = load_template("bao_cao_du2.txt")

        # ===== Báo cáo cuối ngày =====
        done_tasks = []
        for item in items:
            task = tasks.get(item["task_id"], {})
            percent = task.get("percent")
            percent_str = f"{percent}%" if percent is not None else "100%"
            done_tasks.append(f"- {item['task_name']}. `{percent_str}`")
        done_tasks_str = "\n".join(done_tasks)

        eod_report = eod_template.format(
            date_display=date_display,
            fresher_day=fresher_day,
            done_tasks_str=done_tasks_str,
            next_tasks=next_tasks,
        )

        (daily_dir / "Báo cáo cuối ngày.txt").write_text(
            eod_report,
            encoding="utf-8",
        )

        # ===== Báo cáo DU2 =====
        du2_entries = []

        for item in items:
            task = tasks.get(item["task_id"], {})

            estimate_h = task.get("estimate_h", "")
            actual_h = task.get("actual_h", "")
            note = task.get("note") or "N/A"
            percent = task.get("percent")
            percent_str = f"{percent}%" if percent is not None else "100%"

            du2_entries.append(f" -   {item['task_name']}. `{percent_str}`")
            du2_entries.append(f"    - Time estimate: {estimate_h}h")
            du2_entries.append(f"    - Time complete: {actual_h}h")
            du2_entries.append(f"    - Difficulties at work: {note}")
            du2_entries.append("    - Lessons learned: N/A")
            du2_entries.append("")

        du2_report = du2_template.format(
            date_display=date_display,
            du2_entries="\n".join(du2_entries).rstrip(),
            next_tasks=next_tasks,
        )

        (daily_dir / "Báo cáo DU2.txt").write_text(
            du2_report,
            encoding="utf-8",
        )


# ========================
# Generate Word Report
# ========================
def generate_task_docx(tasks, daily_data):
    doc = Document()
    doc.add_heading("Quá trình thực hiện task", 0)

    grouped_by_task = {}

    for item in daily_data:
        grouped_by_task.setdefault(item["task_id"], []).append(item)

    for task_id, items in grouped_by_task.items():
        task = tasks.get(task_id, {})

        items = sorted(items, key=lambda x: x["date_obj"])

        task_name = task.get("task_name") or items[0]["task_name"]
        estimate_h = task.get("estimate_h", "")
        actual_h = task.get("actual_h", "")
        note = task.get("note") or "N/A"

        start_date = items[0]["date"]
        end_date = items[-1]["date"]

        doc.add_paragraph(f"Tên task: {task_name}")
        doc.add_paragraph(f"{start_date} ~ {end_date}")
        doc.add_paragraph(f"1. Nội dung task: {task_name}.")
        doc.add_paragraph(f"2. Estimated: {estimate_h}h")
        doc.add_paragraph(f"3. Time spent: {actual_h}h")
        doc.add_paragraph(f"4. Khó khăn/Lí do: {note}")
        doc.add_paragraph("5. Kinh nghiệm/Bài học rút ra: N/A")  # TODO: Map from Excel if lesson column is added later
        doc.add_paragraph("")

    doc.save(OUTPUT_DIR / "quá trình thực hiện task.docx")


# ========================
# Main
# ========================
def main():
    wb = load_workbook(INPUT_FILE, data_only=True)

    tasks = read_wbs(wb)
    daily_data = read_daily(wb)

    if not daily_data:
        print("Không có nội dung ngày nào trong sheet Report Daily.")
        return

    latest_date = max(item["date_obj"] for item in daily_data)
    latest_daily_data = [item for item in daily_data if item["date_obj"] == latest_date]

    generate_reports_by_date(latest_daily_data, tasks)
    generate_task_docx(tasks, daily_data)

    print("Done. Check folder Generate/")


if __name__ == "__main__":
    main()
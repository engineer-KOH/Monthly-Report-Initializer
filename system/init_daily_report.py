"""Automated monthly report initializer and calendar-based formatting tool."""

import calendar
from datetime import datetime, timedelta
import glob
import os
import re
import sys
from tkinter import messagebox

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

# Import jpholiday safely with fallback handling
try:
    import jpholiday
except ImportError:
    jpholiday = None


def is_holiday(dt: datetime) -> bool:
    """Check whether the specified date is a weekend or Japanese public holiday."""
    if dt is None:
        return False
    # Saturday (5), Sunday (6)
    if dt.weekday() >= 5:
        return True
    # Japanese public holidays via jpholiday
    if jpholiday and jpholiday.is_holiday(dt):
        return True
    return False


def main():
    """Main execution routine for initializing monthly report workbooks."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    root_dir = os.path.dirname(script_dir)

    if len(sys.argv) > 1:
        arg_path = sys.argv[1]
        if os.path.isabs(arg_path):
            input_file = arg_path
        else:
            input_file = os.path.join(root_dir, arg_path)
    else:
        masters = glob.glob(os.path.join(root_dir, "成果報告日報_*_マスター.xlsm"))
        if masters:
            input_file = masters[0]
        else:
            input_file = os.path.join(root_dir, "成果報告日報_末永_マスター.xlsm")

    if not os.path.exists(input_file):
        messagebox.showerror(
            "エラー",
            f"マスターファイルが見つかりません:\n{input_file}",
        )
        return

    base_name = os.path.splitext(os.path.basename(input_file))[0]
    name_match = re.search(r"成果報告日報_(.+?)_マスター", base_name)
    user_name = name_match.group(1) if name_match else "担当者"

    print(
        f"マスターファイル '{input_file}'（担当者: {user_name}）を読み込んでいます..."
    )

    wb = openpyxl.load_workbook(input_file, keep_vba=False)

    config_sheet_name = "年月入力"
    if config_sheet_name not in wb.sheetnames:
        messagebox.showerror(
            "エラー", f"'{config_sheet_name}' シートが存在しません。"
        )
        return

    ws_config = wb[config_sheet_name]
    config_value = ws_config["B2"].value
    if not config_value:
        messagebox.showerror(
            "エラー",
            f"'{config_sheet_name}' シートのB2セルに年月が入力されていません。",
        )
        return

    config_str = str(config_value).strip().replace("/", "-")
    try:
        target_date = datetime.strptime(config_str[:7], "%Y-%m")
        year = target_date.year
        month = target_date.month
    except ValueError:
        messagebox.showerror(
            "エラー",
            "年月のフォーマットは '2026-10' または '2026/10' の形式で入力してください。",
        )
        return

    _, last_day = calendar.monthrange(year, month)
    current_day = datetime(year, month, 1)
    end_day = datetime(year, month, last_day)

    weeks_data = []
    current_week = [None] * 5

    while current_day <= end_day:
        weekday = current_day.weekday()
        if weekday < 5:
            current_week[weekday] = current_day
        if weekday == 4 or current_day == end_day:
            if any(current_week):
                weeks_data.append(current_week)
            current_week = [None] * 5
        current_day += timedelta(days=1)

    sheet_index = 0
    columns = ["B", "C", "D", "E", "F"]

    # Hatch pattern (gray fill) for weekends and holidays
    hatch_fill = PatternFill(
        fill_type="lightGray",
        start_color="E0E0E0",
        end_color="E0E0E0"
    )

    for sheet_name in list(wb.sheetnames):
        ws = wb[sheet_name]

        if hasattr(ws, "_drawings"):
            ws._drawings.clear()
        if hasattr(ws, "legacy_drawing"):
            ws.legacy_drawing = None

        if sheet_name in [config_sheet_name, "sheetメニュ", "メニュ"]:
            continue

        if sheet_index >= len(weeks_data):
            break

        week = weeks_data[sheet_index]
        valid_days = [d for d in week if d is not None]
        new_sheet_name = (
            f"{valid_days[0].strftime('%m.%d')}-"
            f"{valid_days[-1].strftime('%m.%d')}"
        )
        ws.title = new_sheet_name

        for row in range(2, ws.max_row + 1):
            for col in range(2, ws.max_column + 1):
                cell = ws.cell(row=row, column=col)
                if not isinstance(cell, MergedCell):
                    cell.value = None

        for idx, day in enumerate(week):
            col_letter = columns[idx]
            cell_coord = f"{col_letter}1"

            if day is not None:
                ws[cell_coord].value = day.strftime("%Y-%m-%d")

                # Judge holidays and apply hatch pattern fills
                if is_holiday(day):
                    max_target_row = max(ws.max_row, 18)
                    for r in range(1, max_target_row + 1):
                        cell = ws[f"{col_letter}{r}"]
                        if not isinstance(cell, MergedCell):
                            cell.fill = hatch_fill
            else:
                ws[cell_coord].value = "ー"

        dv = DataValidation(
            type="list",
            formula1='"業務,自己学習,社内業務"',
            allow_blank=True,
            showInputMessage=True,
            promptTitle="以下の入力が必要です",
            prompt="自己学習⇒参照テキスト/サイト、参照章\n社内業務⇒使用ツール",
        )
        ws.add_data_validation(dv)
        dv.add("B4:F4")

        sheet_index += 1

    if config_sheet_name in wb.sheetnames:
        wb.remove(wb[config_sheet_name])

    for menu_sheet_name in ["sheetメニュ", "メニュ"]:
        if menu_sheet_name in wb.sheetnames:
            wb.remove(wb[menu_sheet_name])

    output_dir = os.path.dirname(os.path.abspath(input_file))
    new_filename = f"成果報告日報_{user_name}_{year}-{month:02d}.xlsx"
    output_file_path = os.path.join(output_dir, new_filename)

    try:
        wb.save(output_file_path)
        success_msg = (
            f"✨ {user_name} さんの {year}年{month}月用日報の初期化が完了しました！\n\n"
            f"・出力ファイル名: {new_filename}\n"
            f"・土日および祝日（文化の日・勤労感謝の日等）のハッチ処理を完了しました。\n"
            f"・VBAマクロ/不要ボタンを除去したクリーンな .xlsx を生成しました。"
        )
        messagebox.showinfo("社内DX初期化完了", success_msg)

    except PermissionError:
        messagebox.showerror(
            "保存エラー",
            f"出力対象のファイル '{new_filename}' が開かれているため保存できません。\n"
            f"ファイルを閉じてから再実行してください。",
        )
    except Exception as e:
        messagebox.showerror(
            "保存エラー", f"ファイルの出力に失敗しました:\n{e}"
        )


if __name__ == "__main__":
    main()
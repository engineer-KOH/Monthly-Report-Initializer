import calendar
from datetime import datetime, timedelta
import glob
import os
import re
import sys
from tkinter import messagebox
import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.worksheet.datavalidation import DataValidation


def main():
    # スクリプト自身のフォルダ（system）と親フォルダ（ルート階層）を取得
    script_dir = os.path.dirname(os.path.abspath(__file__))
    root_dir = os.path.dirname(script_dir)

    # 1. 対象のマスターファイルを動的に特定（ルートフォルダ直下を参照）
    if len(sys.argv) > 1:
        arg_path = sys.argv[1]
        # 引数が絶対パスか相対パスかに応じて読み込みパスを確定
        if os.path.isabs(arg_path):
            input_file = arg_path
        else:
            input_file = os.path.join(root_dir, arg_path)
    else:
        # 引数指定がない場合はルートフォルダからマスターファイルを自動検索
        masters = glob.glob(os.path.join(root_dir, "成果報告日報_*_マスター.xlsm"))
        if masters:
            input_file = masters[0]
        else:
            input_file = os.path.join(root_dir, "成果報告日報_氏名_マスター.xlsm")  # フォールバック

    if not os.path.exists(input_file):
        messagebox.showerror(
            "エラー",
            f"マスターファイルが見つかりません:\n{input_file}",
        )
        return

    # 🌟 ファイル名から氏名を動的に抽出（例: "成果報告日報_田中_マスター.xlsm" -> "田中"）
    base_name = os.path.splitext(os.path.basename(input_file))[0]
    name_match = re.search(r"成果報告日報_(.+?)_マスター", base_name)
    user_name = name_match.group(1) if name_match else "担当者"

    print(
        f"マスターファイル '{input_file}'（担当者: {user_name}）を読み込んでいます..."
    )

    # keep_vba=False でVBAマクロを除去
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

    # 日付計算
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

    for sheet_name in list(wb.sheetnames):
        ws = wb[sheet_name]

        # 不要な描画オブジェクトのクリア
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
        new_sheet_name = f"{valid_days[0].strftime('%m.%d')}-{valid_days[-1].strftime('%m.%d')}"
        ws.title = new_sheet_name

        # データクリア
        for row in range(2, ws.max_row + 1):
            for col in range(2, ws.max_column + 1):
                cell = ws.cell(row=row, column=col)
                if not isinstance(cell, MergedCell):
                    cell.value = None

        # 日付書き込み
        for idx, day in enumerate(week):
            cell_coord = f"{columns[idx]}1"
            if day is not None:
                ws[cell_coord].value = day.strftime("%Y-%m-%d")
            else:
                ws[cell_coord].value = "ー"

        # プルダウン再設定
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

    # 生成後ファイルから管理用不要シートを削除
    if config_sheet_name in wb.sheetnames:
        wb.remove(wb[config_sheet_name])

    for menu_sheet_name in ["sheetメニュ", "メニュ"]:
        if menu_sheet_name in wb.sheetnames:
            wb.remove(wb[menu_sheet_name])

    # 🌟 保存先をマスターファイルが存在するフォルダ（ルートフォルダ直下）に明示指定
    output_dir = os.path.dirname(os.path.abspath(input_file))
    new_filename = f"成果報告日報_{user_name}_{year}-{month:02d}.xlsx"
    output_file_path = os.path.join(output_dir, new_filename)

    try:
        wb.save(output_file_path)

        success_msg = (
            f"✨ {user_name} さんの {year}年{month}月用日報の初期化が完了しました！\n\n"
            f"・出力ファイル名: {new_filename}\n"
            f"・初期化専用シートを自動削除しました。\n"
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
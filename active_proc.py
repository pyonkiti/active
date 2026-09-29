# ActiveのExcelファイルを集計してExcelファイルに出力する機能です

import time
import tkinter          as tk
import win32com.client  as win32
import common.common    as COM                    # 共通モジュール
import common.config    as config                 # 共通モジュール
from tkinter            import messagebox
from tkinter            import filedialog
from datetime           import date, datetime
from decimal            import Decimal

# --------------------------------------------------------------------
# メイン系
# --------------------------------------------------------------------
class MAIN:
    # ----------------------------------------------------------------
    # 初期処理
    # ----------------------------------------------------------------
    def proc_init():

        # YAMLファイルの読み込み
        msg = COMMON.get_yaml()
        if msg != None: return msg

        # YAMLファイルの登録データチェック
        msg = COMMON.check_yaml()
        if msg != None: return msg

        return None

    # ----------------------------------------------------------------
    # ヘッダー処理
    # ----------------------------------------------------------------
    def proc_header():

        # ダイアログボックスを表示
        file_path = filedialog.askopenfilename(
            title="Excelファイルを選択してください",
            filetypes=[("Excelファイル", "*.xlsx *.xls"), ("すべてのファイル", "*.*")]
        )
        if not file_path: return "cancel"

        # 拡張子からExcelファイルか確認
        ret = COMMON.check_filetype(file_path)
        if not ret: return "Excel以外のファイルが選択されています。"

        config.path_excel_in = str(file_path)           # Excelファイル（取り込み元）
        
        return None
        
    # ----------------------------------------------------------------
    # メイン処理
    # ----------------------------------------------------------------
    def proc_main():

        # Excelを起動
        ret = PROC.init_excel()
        if not ret: return "Excelの起動に失敗しました。"

        # Excelを読み込む
        excel_arry = PROC.read_excel()

        # Excelの集計
        excel_arry_ret = PROC.sum_excel(excel_arry)
        
        # Excelを書き込む
        PROC.write_excel(excel_arry_ret)

        # 全ての処理が終わったらExcelアプリケーションを終了する
        excel.Quit()

        return None

    # ----------------------------------------------------------------
    # フッター処理
    # ----------------------------------------------------------------
    def proc_footer():
        msg = f"処理数 : {config.sum_count} 件"
        messagebox.showinfo("確認", f"処理が完了しました。 {msg}")

# --------------------------------------------------------------------
# 処理系
# --------------------------------------------------------------------
class PROC:
    # ----------------------------------------------------------------
    # Excelを起動
    # ----------------------------------------------------------------
    def init_excel():

        global excel

        excel = win32.gencache.EnsureDispatch("Excel.Application")
        excel.Visible = False    
        excel.DisplayAlerts = False        

        return True
        
    # ----------------------------------------------------------------
    # Excelを読み込む
    # ----------------------------------------------------------------
    def read_excel():

        try:
            # メッセージ表示の初期処理
            TKINTER.tkinter_init()

            status_label.config(text = "Excelファイルを読み込んでいます・・・")
            processing_window.update()
            time.sleep(0.2)
            
            wb = excel.Workbooks.Open(config.path_excel_in)
            ws = wb.Worksheets(None) if None else wb.Worksheets(1)

            row_excel   = 2                 # タイトルは読み飛ばす
            count_blank = 0
            excel_arry = []

            # ブレイク条件（連続空白）
            while count_blank < 5:

                # 無限ループ回避
                if row_excel > 3000: break
                
                # ブレイク条件（商談コード）
                if ws.Range(f"A{row_excel}").Value is None:
                    
                    count_blank += 1
                    continue

                sort, bumon_nm = COMMON.get_bumon_name(ws.Range(f"M{row_excel}").Value)     # 社員名から部門名を取得
                uriage_kin     = COMMON.change_decimal(ws.Range(f"N{row_excel}").Value)     # 売上金額
                net_kin        = COMMON.change_decimal(ws.Range(f"S{row_excel}").Value)     # NET
                kensyu_ymd     = COMMON.check_date(ws.Range(f"Z{row_excel}").Value)         # 検収予定日
                ki             = COMMON.get_ki(kensyu_ymd)                                  # 検収予定日から期を取得

                excel_arry.append([ki, sort, bumon_nm, uriage_kin, net_kin])
                row_excel += 1

            config.sum_count = row_excel - 1
            wb.Close(SaveChanges=False)
            wb = None

            return excel_arry

        finally:
            if wb is not None:
                wb.Close(SaveChanges = False)

            processing_window.withdraw()
     
    # --------------------------------------------------------------------
    # Excelで読み込んだ結果を集計
    # --------------------------------------------------------------------
    def sum_excel(excel_arry):

        excel_arry_sum = {}

        for row in excel_arry:
            
            key = (row[0], row[1])                                  # 第一要素と第２要素をキー

            if key not in excel_arry_sum:
                excel_arry_sum[key] = [row[0], row[1], row[2], row[3], row[4]]
            else:
                excel_arry_sum[key][3] += row[3]                    # 売上金額
                excel_arry_sum[key][4] += row[4]                    # NET

        excel_arry_ret = list(excel_arry_sum.values())
        excel_arry_ret.sort(key = lambda row: (row[0], row[1]))     # 第一要素、第三要素の順番で昇順ソート

        return excel_arry_ret
    
    # ----------------------------------------------------------------
    # Excelを書き込む
    # ----------------------------------------------------------------
    def write_excel(excel_arry_ret):

        wb = None
        try:
            wb = excel.Workbooks.Add()
            ws = wb.Worksheets(1)

            idx = 0
            mongon = ["売上", "粗利", "前月売上差分"]
            ws.Range("A1").Value = "母数集計"
            
            for data in excel_arry_ret:

                kingak = [COMMON.get_senen_cut(data[3]), COMMON.get_senen_cut(data[4]), None]
                for inext in range(3):

                    ws.Range(f"{COMMON.col_letter(idx + inext)}4").Value = data[2]              # 部門名
                    ws.Range(f"{COMMON.col_letter(idx + inext)}5").Value = data[0]              # 期
                    ws.Range(f"{COMMON.col_letter(idx + inext)}6").Value = mongon[inext]        # 固定文言
                    ws.Range(f"{COMMON.col_letter(idx + inext)}7").Value = kingak[inext]        # 売上金額 / NET
                    ws.Range(f"{COMMON.col_letter(idx + inext)}5").HorizontalAlignment = -4131  # 期を右寄せ
                    
                idx += 3

            wb.SaveAs(config.path_excel)
            wb.Close(SaveChanges = False)
            wb = None

            return True
        
        finally:
            if wb is not None:
                wb.Close(SaveChanges = False)

# --------------------------------------------------------------------
# メッセージ表示系
# --------------------------------------------------------------------
class TKINTER():
    # ----------------------------------------------------------------
    # メッセージＢＯＸの形成
    # ----------------------------------------------------------------
    def tkinter_init():
            
        processing_window.overrideredirect(True)
        processing_window.geometry("250x70")
        processing_window.resizable(False, False)
        processing_window.update_idletasks()

        status_label.pack(expand = True)

        width = processing_window.winfo_width()
        height = processing_window.winfo_height()
        x = (processing_window.winfo_screenwidth() - width) // 2
        y = (processing_window.winfo_screenheight() - height) // 2
        processing_window.geometry(f"{width}x{height}+{x}+{y}")

        processing_window.deiconify()
        processing_window.update()

# ----------------------------------------------------------------
# 初期処理
# ----------------------------------------------------------------
def init():
    try:
        global COMMON, root, processing_window, status_label

        COMMON = COM.COMMON()

        root = tk.Tk()
        root.withdraw()
        processing_window = tk.Toplevel(root)
        processing_window.withdraw()
        status_label = tk.Label(processing_window, text = "準備中です・・・", font = ("Meiryo", 9))

        # メッセージ表示の初期処理
        TKINTER.tkinter_init()

        status_label.config(text = "ライブラリをロードしています・・・")
        processing_window.update()
        time.sleep(0.5)
        processing_window.withdraw()

        return True

    except Exception as e:
        messagebox.showerror("エラー", f"{COMMON.log_exception(e)}")
        processing_window.destroy()
        root.destroy()
    finally:
        None

# --------------------------------------------------------------------
# メイン処理
# --------------------------------------------------------------------
def main():
    try:
        # 初期処理
        msg = MAIN.proc_init()
        if msg != None:
            messagebox.showerror("エラー", msg); return

        # ヘッダー処理
        msg = MAIN.proc_header()
        if msg != None:
            if msg == "cancel": 
                return
            else:
                messagebox.showerror("エラー", msg); return

        if messagebox.askyesno("確認", "処理を開始しますか？\nExcelが起動されていれば強制的に終了されます。"):
            
            # メイン処理
            msg = MAIN.proc_main()
            if msg != None:
                messagebox.showerror("エラー", msg); return
            
            # フッター処理
            MAIN.proc_footer()
        
    except Exception as e:
        messagebox.showerror("エラー", f"{COMMON.log_exception(e)}")
    finally:
        processing_window.destroy()
        root.destroy()

# --------------------------------------------------------------------
# グローバル処理
# --------------------------------------------------------------------

# 初期処理
if not init(): exit(1)

# メイン処理
main()

import os
import re
import yaml
import common.config    as config
from pathlib            import Path
from decimal            import Decimal
from datetime           import date, datetime

# --------------------------------------------------------------------
# 共通処理
# --------------------------------------------------------------------
class COMMON:

    # --------------------------------------------------------------------
    # YAMLファイルの読み込み
    # --------------------------------------------------------------------
    def get_yaml(self):

        if not Path(config.path_yaml).exists():
            return "yamlファイルが存在しません。"
        
        with open(Path(config.path_yaml), mode = "r", encoding = "utf-8") as f:
            yaml_fil = yaml.safe_load(f)

        config.bumon = yaml_fil["bumon"]                    # 部門
        config.path_excel = yaml_fil["path_excel"]          # Excelの出力先ファイル

        return None

    # --------------------------------------------------------------------
    # 社員名の中に含まれる、全角半角スペースを取り除く
    # --------------------------------------------------------------------
    def remove_spaces(text):

        text = text if (text is not None) else ""                  # None → "" に変換
        return re.sub(r'[\u3000 ]+', '', text)
        
    # --------------------------------------------------------------------
    # YAMLファイルの登録データチェック
    # --------------------------------------------------------------------
    def check_yaml(self):

        # （上下水）社員名からスペースを取り除く
        ary = config.bumon["jyougesui"]["syain"]
        config.bumon["jyougesui"]["syain"] = [COMMON.remove_spaces(item) for item in ary]

        # （IoT）社員名からスペースを取り除く
        ary = config.bumon["iot"]["syain"]
        config.bumon["iot"]["syain"] = [COMMON.remove_spaces(item) for item in ary]

        # （東京）社員名からスペースを取り除く
        ary = config.bumon["tokyo"]["syain"]
        config.bumon["tokyo"]["syain"] = [COMMON.remove_spaces(item) for item in ary]

        # （上下水道事務）社員名からスペースを取り除く
        ary = config.bumon["jimu"]["syain"]
        config.bumon["jimu"]["syain"] = [COMMON.remove_spaces(item) for item in ary]

        # パスの存在チェック
        path_excel_dir = os.path.dirname(config.path_excel)
        if not os.path.exists(path_excel_dir):
            return "path_excelで指定したパスが存在しません。"
        
        # 社員重複チェック
        syain_sum = config.bumon["jyougesui"]["syain"] + config.bumon["iot"]["syain"] + config.bumon["tokyo"]["syain"] + config.bumon["jimu"]["syain"]
        if len(syain_sum) != len(set(syain_sum)):
            return "YAMLファイルに重複した社員が登録されています。"

        return None

    # --------------------------------------------------------------------
    # 拡張子からExcelファイルか確認
    # --------------------------------------------------------------------
    def check_filetype(self, file):

        ext = os.path.splitext(file)[1].lower()
        return ext in ('.xlsx', '.xls')

    # --------------------------------------------------------------------
    # 社員名から部門を取得する
    # --------------------------------------------------------------------
    def get_bumon_name(self, syain_old):

        syain_old = syain_old if (syain_old is not None) else ""      # None → "" に変換

        # スペースを除去
        syain = COMMON.remove_spaces(syain_old)
        # syain = syain_old

        # 上下水に含まれるか
        if syain in config.bumon["jyougesui"]["syain"]:
            return config.bumon["jyougesui"]["sort"], config.bumon["jyougesui"]["name"]

        # IoTに含まれるか
        if syain in config.bumon["iot"]["syain"]:
                return config.bumon["iot"]["sort"], config.bumon["iot"]["name"]
        
        # 東京に含まれるか
        if syain in config.bumon["tokyo"]["syain"]:
                return config.bumon["tokyo"]["sort"], config.bumon["tokyo"]["name"]
        
        # 上下水道事務に含まれるか
        if syain in config.bumon["jimu"]["syain"]:
                return config.bumon["jimu"]["sort"], config.bumon["jimu"]["name"]
        
        return 9, "部門なし"

    # --------------------------------------------------------------------
    # 文字列をDecimal型に変換
    # --------------------------------------------------------------------
    def change_decimal(self, data):

        data_none = data if (data is not None) else ""                  # None → "" に変換
        data_strp = str(data_none).strip()                              # 左右のスペースを削除
        data_zero = data_strp if (data_strp != "") else "0"             # "" → "0" に変換
        data_kanm = data_zero.replace(",", "")                          # "1,000" → "1000" に変換
        
        try:
            return Decimal(data_kanm)
        except Exception:
            return Decimal(0)
        finally:
            None
        
    # --------------------------------------------------------------------
    # 日付の論理チェック
    # --------------------------------------------------------------------
    def check_date(self, input_date = "1971-05-31"):

        if input_date == None: return None

        # 日付の論理チェック
        if isinstance(input_date, datetime):
            ret = input_date.date().strftime("%Y-%m-%d")
            return ret

        # 日付の論理チェック
        if isinstance(input_date, date):
            ret = input_date.strftime("%Y-%m-%d")
            return ret

        # 文字列以外はエラー
        if not isinstance(input_date, str):
            return None

        match = re.match(r"^(\d{1,4})([/\-])(\d{1,2})\2(\d{1,2})$", input_date)
        if not match: return None
            
        year_str, _, month_str, day_str = match.groups()

        if len(year_str) != 4: return None
        
        try:
            yy = int(year_str)
            mm = int(month_str)
            dd = int(day_str)
            valid_date = date(yy, mm, dd)
        except ValueError:
            return None

        return valid_date.strftime("%Y-%m-%d")

    # --------------------------------------------------------------------
    # 日付から期を取得する
    # --------------------------------------------------------------------
    def get_ki(self, input_date = "1971-05-31"):

        try:
            target_date = datetime.strptime(input_date, "%Y-%m-%d").date()
        except ValueError:
            return None

        if target_date < date(1971, 6, 1): return None
            
        # 期の計算
        if target_date.month >= 6:
            start_year = target_date.year
        else:
            start_year = target_date.year - 1

        ki = 1 + (start_year - 1971)

        return ki

    # ----------------------------------------------------------------
    # 現期を取得 （未使用）
    # ----------------------------------------------------------------
    def get_start_ki(self):

        today = datetime.now().strftime("%Y-%m-%d")
        ki = COMMON.get_ki(today)
        
        return ki

    # --------------------------------------------------------------------
    # 期が指定した範囲内にあるかチェック （未使用）
    # --------------------------------------------------------------------
    def check_ki(self, ki):

        if ki in [config.start_ki, config.start_ki + 1, config.start_ki + 2]:
            return True
        else:
            return False
        
    # ----------------------------------------------------------------
    # Excelの列名を順に取得
    # 0-indexed の列番号を列名(A, B, ..., Z, AA, AB, ...)に変換する
    # ----------------------------------------------------------------
    def col_letter(self, n):

        n += 1                                      # 1-indexed に変換
        result = ""
        while n > 0:
            n, rem = divmod(n - 1, 26)
            result = chr(65 + rem) + result
        return result

    # --------------------------------------------------------------------
    # 千円未満切り捨て
    # --------------------------------------------------------------------
    def get_senen_cut(self, kingaku):

        result = int(kingaku / 1000)
        return result

    # --------------------------------------------------------------------
    # 詳細ログをGUI表示
    # --------------------------------------------------------------------
    def log_exception(self, e: Exception, show_locals: bool = True) -> str:

        lines = []
        lines.append(f"例外クラス : {type(e).__module__}.{type(e).__qualname__}")
        lines.append(f"メッセージ : {e}")
        lines.append("-" * 70)
        lines.append("呼び出し経路:")

        tb = e.__traceback__
        depth = 0
        last_frame_info = None
        while tb:
            frame = tb.tb_frame
            code = frame.f_code
            qualname = getattr(code, "co_qualname", code.co_name)
            filename = code.co_filename
            lineno = tb.tb_lineno

            try:
                with open(filename, encoding="utf-8") as f:
                    source_line = f.readlines()[lineno - 1].strip()
            except Exception:
                source_line = "(ソース取得不可)"

            lines.append(f"  [{depth}] {qualname}  ({filename}:{lineno})")
            lines.append(f"       >> {source_line}")

            if show_locals:
                for k, v in frame.f_locals.items():
                    if k in ("self", "cls"):
                        continue
                    try:
                        v_repr = repr(v)
                    except Exception:
                        v_repr = "<repr失敗>"
                    if len(v_repr) > 200:
                        v_repr = v_repr[:200] + "...(省略)"
                    lines.append(f"         {k} = {v_repr}")

            last_frame_info = (qualname, filename, lineno)
            depth += 1
            tb = tb.tb_next

        lines.append("-" * 70)
        if last_frame_info:
            qualname, filename, lineno = last_frame_info
            lines.append(f" 発生箇所: {qualname}  ({filename}:{lineno})")

        if e.__cause__:
            lines.append("-" * 70)
            lines.append(" 直接の原因となった例外（raise ... from ...）:")
            lines.append(f"   {type(e.__cause__).__name__}: {e.__cause__}")
        elif e.__context__ and e.__context__ is not e:
            lines.append("-" * 70)
            lines.append(" 処理中に暗黙的に発生していた元の例外:")
            lines.append(f"   {type(e.__context__).__name__}: {e.__context__}")

        return "\n".join(lines)

#!/usr/bin/env python3
"""鉄骨設計 Action List を Excel ファイルとして出力するスクリプト"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

DATA = [
    ("1. 企画・基本計画", [
        "建物用途・規模・階数の確認",
        "適用法規の確認（建築基準法・消防法等）",
        "構造種別の決定（S造・SRC造等）",
        "地盤調査報告書の確認（ボーリングデータ・N値）",
        "設計方針の決定（ルート判定：ルート1/2/3）",
        "耐震等級の確認・設定",
    ]),
    ("2. 荷重条件の設定", [
        "固定荷重（DL）の算定",
        "積載荷重（LL）の設定（床用・架構用・地震用）",
        "積雪荷重（SL）の設定（一般/多雪区域の確認）",
        "風荷重（WL）の算定（速度圧・風力係数）",
        "地震荷重の算定（Ci・Ai分布・地域係数Z）",
        "荷重組合せの設定（長期・短期・各ケース）",
    ]),
    ("3. 構造計画・架構計画", [
        "柱配置・スパン割りの決定",
        "架構形式の決定（純ラーメン・ブレース付等）",
        "ブレース配置計画（偏心率の考慮）",
        "床構造の決定（デッキスラブ・ALC等）",
        "柱脚形式の決定（露出型・根巻型・埋込型）",
        "基礎形式の決定（直接基礎・杭基礎）",
        "エキスパンションジョイントの要否検討",
    ]),
    ("4. 部材断面の仮定・設計", [
        "鋼材種別の選定（SN400B・SN490B・BCR295等）",
        "柱断面の仮定（角形鋼管・H形鋼・円形鋼管）",
        "梁断面の仮定（H形鋼・BH）",
        "ブレース断面の仮定",
        "小梁・間柱の断面設計",
        "幅厚比の確認（FA・FB・FC・FD）",
    ]),
    ("5. 構造解析", [
        "解析モデルの作成（一貫計算ソフト入力）",
        "長期応力解析の実施",
        "短期応力解析の実施（地震時・風荷重時）",
        "層間変形角の確認（1/200以下）",
        "剛性率の確認（0.6以上）",
        "偏心率の確認（0.15以下）",
        "塔状比の確認",
    ]),
    ("6. 断面検定", [
        "柱の断面検定（軸力+曲げ）",
        "梁の断面検定（曲げ+せん断）",
        "梁の横座屈の検討",
        "ブレースの断面検定（引張・座屈）",
        "柱梁耐力比の確認（冷間成形角形鋼管の場合）",
        "たわみの確認（梁：L/300以下等）",
        "検定比のNG部材がないことを確認",
    ]),
    ("7. 接合部設計", [
        "柱梁接合部（仕口）の設計",
        "パネルゾーンの検討",
        "ダイアフラムの設計（通し・内・外ダイアフラム）",
        "高力ボルト接合部の設計（摩擦接合・引張接合）",
        "溶接接合部の設計（完全溶込み・隅肉溶接）",
        "ブレース接合部（ガセットプレート）の設計",
        "継手位置・形式の決定",
        "スカラップ形状の確認（ノンスカラップ工法の検討）",
    ]),
    ("8. 柱脚設計", [
        "アンカーボルトの設計（本数・径・埋込み長さ）",
        "ベースプレートの設計（板厚・寸法）",
        "基礎コンクリートの支圧応力度の検討",
        "柱脚の回転剛性の設定と反映",
        "根巻きコンクリートの設計（根巻型の場合）",
    ]),
    ("9. 基礎設計", [
        "基礎梁の断面設計",
        "フーチングの設計（直接基礎の場合）",
        "杭の選定・設計（杭基礎の場合）",
        "地中梁の設計",
        "転倒・滑動の検討",
    ]),
    ("10. 保有水平耐力計算（ルート3の場合）", [
        "保有水平耐力（Qu）の算定",
        "必要保有水平耐力（Qun）の算定",
        "構造特性係数（Ds）の算定",
        "崩壊メカニズムの確認（全体崩壊型）",
        "Qu ≧ Qun の確認",
        "部材のランク確認（FA～FD）",
    ]),
    ("11. 図面・計算書の作成", [
        "構造図の作成（伏図・軸組図）",
        "部材リストの作成",
        "接合部詳細図の作成",
        "柱脚詳細図の作成",
        "構造計算書のとりまとめ",
        "構造計算概要書の作成",
    ]),
    ("12. 確認申請・審査対応", [
        "確認申請書類の作成・提出",
        "適合性判定の要否確認（ルート2-3以上）",
        "審査機関からの質疑対応",
        "確認済証の取得",
    ]),
    ("13. 施工段階の確認", [
        "工作図（ファブ図）の確認・承認",
        "鋼材のミルシート確認",
        "溶接施工要領書の確認",
        "超音波探傷検査（UT検査）の実施確認",
        "高力ボルト締付け管理の確認",
        "建方精度の確認（柱の倒れ等）",
        "中間検査・完了検査の対応",
    ]),
]

def main():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "鉄骨設計 Action List"

    # Column widths
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 55
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 20

    # Styles
    title_font = Font(name="Meiryo", size=16, bold=True, color="1A237E")
    header_font = Font(name="Meiryo", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1565C0", end_color="1565C0", fill_type="solid")
    phase_font = Font(name="Meiryo", size=10, bold=True, color="1A237E")
    phase_fill = PatternFill(start_color="E3F2FD", end_color="E3F2FD", fill_type="solid")
    item_font = Font(name="Meiryo", size=10)
    thin_border = Border(
        left=Side(style="thin", color="BDBDBD"),
        right=Side(style="thin", color="BDBDBD"),
        top=Side(style="thin", color="BDBDBD"),
        bottom=Side(style="thin", color="BDBDBD"),
    )
    center = Alignment(horizontal="center", vertical="center")
    left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # Title row
    ws.merge_cells("A1:F1")
    cell = ws["A1"]
    cell.value = "鉄骨設計 Action List"
    cell.font = title_font
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36

    # Header row
    headers = ["No.", "フェーズ", "Action 項目", "チェック", "担当者", "備考"]
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = thin_border
    ws.row_dimensions[3].height = 24

    row = 4
    item_no = 1

    for phase_title, items in DATA:
        # Phase header row
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        cell = ws.cell(row=row, column=1, value=phase_title)
        cell.font = phase_font
        cell.fill = phase_fill
        cell.alignment = Alignment(horizontal="left", vertical="center")
        for c in range(1, 7):
            ws.cell(row=row, column=c).border = thin_border
            ws.cell(row=row, column=c).fill = phase_fill
        ws.row_dimensions[row].height = 26
        row += 1

        for item in items:
            ws.cell(row=row, column=1, value=item_no).font = item_font
            ws.cell(row=row, column=1).alignment = center

            ws.cell(row=row, column=2, value=phase_title.split(". ", 1)[1] if ". " in phase_title else phase_title).font = item_font
            ws.cell(row=row, column=2).alignment = left_wrap

            ws.cell(row=row, column=3, value=item).font = item_font
            ws.cell(row=row, column=3).alignment = left_wrap

            ws.cell(row=row, column=4, value="☐").font = Font(name="Meiryo", size=12)
            ws.cell(row=row, column=4).alignment = center

            # D, E columns left blank for user input
            for c in range(1, 7):
                ws.cell(row=row, column=c).border = thin_border

            ws.row_dimensions[row].height = 22
            item_no += 1
            row += 1

    # Freeze panes (header row)
    ws.freeze_panes = "A4"

    # Print settings
    ws.print_area = f"A1:F{row - 1}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1

    output_path = "/home/user/Open_Test/steel-frame-action-list.xlsx"
    wb.save(output_path)
    print(f"Saved: {output_path}")

if __name__ == "__main__":
    main()

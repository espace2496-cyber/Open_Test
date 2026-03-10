#!/usr/bin/env python3
"""鉄骨設計 LE Action List Excel生成スクリプト"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "鉄骨設計LE_ActionList"

# === スタイル定義 ===
thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)
header_font = Font(name='Meiryo', bold=True, size=11, color='FFFFFF')
title_font = Font(name='Meiryo', bold=True, size=14, color='1A237E')
phase_font = Font(name='Meiryo', bold=True, size=11, color='FFFFFF')
normal_font = Font(name='Meiryo', size=10)
wrap_align = Alignment(wrap_text=True, vertical='center')
center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

phase_colors = {
    1: '1565C0', 2: '00838F', 3: '2E7D32',
    4: 'EF6C00', 5: '6A1B9A', 6: 'C62828'
}

importance_fills = {
    '最重要': PatternFill(start_color='FF8A80', end_color='FF8A80', fill_type='solid'),
    '重要':   PatternFill(start_color='FFD180', end_color='FFD180', fill_type='solid'),
    '標準':   PatternFill(start_color='B9F6CA', end_color='B9F6CA', fill_type='solid'),
}

# === タイトル行 ===
ws.merge_cells('A1:J1')
ws['A1'] = 'Plant設計 鉄骨設計 Lead Engineer (LE) Action List'
ws['A1'].font = title_font
ws['A1'].alignment = Alignment(vertical='center')
ws.row_dimensions[1].height = 35

ws.merge_cells('A2:J2')
ws['A2'] = 'プロジェクト名:                    作成日:          Rev:'
ws['A2'].font = Font(name='Meiryo', size=10, color='555555')
ws.row_dimensions[2].height = 22

# === ヘッダー行 ===
headers = [
    ('No.', 6),
    ('フェーズ', 22),
    ('アクション項目', 50),
    ('詳細内容 / 備考', 45),
    ('重要度', 10),
    ('実施時期', 18),
    ('担当者', 14),
    ('R\n(実行責任)', 12),
    ('A\n(説明責任)', 12),
    ('C\n(協議)', 14),
    ('I\n(報告先)', 14),
    ('ステータス', 10),
]

header_fill = PatternFill(start_color='1A237E', end_color='1A237E', fill_type='solid')
row = 4
for col_idx, (h, w) in enumerate(headers, 1):
    cell = ws.cell(row=row, column=col_idx, value=h)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center_align
    cell.border = thin_border
    ws.column_dimensions[get_column_letter(col_idx)].width = w
ws.row_dimensions[row].height = 36

# === データ定義 ===
# (phase, action, detail, importance, timing, person, R, A, C, I)
data = [
    # ====== Phase 1: プロジェクト概要の把握 ======
    (1, 'プロジェクト概要の把握', 'プロジェクト基本情報の確認',
     'プラント種別、規模、処理能力、所在地、全体工期、予算規模の確認',
     '最重要', '着任即日', 'LE', 'LE', 'PM', 'PM', 'チーム全員'),
    (1, '', '契約図書・仕様書（Job Spec）の入手・通読',
     'Contract Document, Project Specification, Special Requirementsの精読',
     '最重要', '着任〜1週間', 'LE', 'LE', 'PM', 'PM/契約担当', 'チーム全員'),
    (1, '', '適用基準・コードの確認',
     'AISC 360/341, AWS D1.1, ASCE 7, IBC, 建築基準法, 高圧ガス保安法 等',
     '最重要', '着任〜1週間', 'LE', 'LE', 'PM', 'QA', 'チーム全員'),
    (1, '', 'Design Criteria / Design Basis の確認',
     'クライアント発行のDesign Criteria、荷重組合せ、安全率等',
     '最重要', '着任〜1週間', 'LE', 'LE', 'PM', 'クライアント', 'チーム全員'),
    (1, '', '前任LE / PM からの引継ぎ',
     '設計進捗、課題事項、クライアント要望、過去の決定事項の確認',
     '最重要', '着任即日', 'LE', 'LE', 'PM', '前任LE', ''),
    (1, '', '過去類似PJ実績・Lessons Learned 収集',
     '同種プラント・同クライアントの過去PJ資料、反省点の確認',
     '重要', '着任〜2週間', 'LE', 'LE', 'PM', 'ベテラン技術者', ''),
    (1, '', 'クライアント組織・承認者の確認',
     'クライアント側のPM, 設計審査担当, 承認権限者の把握',
     '重要', '着任〜1週間', 'LE', 'LE', 'PM', 'PM', ''),
    (1, '', 'プロジェクト全体スケジュールの把握',
     'EPC全体工程、調達LLI納期、建設工程との整合確認',
     '重要', '着任〜1週間', 'LE', 'LE', 'PM', 'スケジュール担当', 'チーム全員'),
    (1, '', 'Vendor Document リストの確認',
     '機器・配管等のVendor図書受領予定の確認',
     '標準', '着任〜2週間', 'LE', 'LE', 'PM', '調達担当', ''),

    # ====== Phase 2: 設計条件の整理 ======
    (2, '設計条件の整理', '荷重条件の整理（Dead Load）',
     '鉄骨自重、グレーチング、手摺、階段、配管サポート等の単位荷重設定',
     '最重要', '1〜2週目', 'LE/構造担当', 'LE', 'PM', '配管/機器部門', 'チーム全員'),
    (2, '', '荷重条件の整理（Live Load）',
     '床積載荷重、プラットフォーム荷重、メンテナンス荷重の設定',
     '最重要', '1〜2週目', 'LE/構造担当', 'LE', 'PM', 'プロセス部門', 'チーム全員'),
    (2, '', '荷重条件の整理（Wind Load）',
     '基本風速、地表面粗度区分、風力係数、ガスト係数の設定',
     '最重要', '1〜2週目', 'LE/構造担当', 'LE', 'PM', '気象データ提供元', 'チーム全員'),
    (2, '', '荷重条件の整理（Seismic Load）',
     '地震地域係数、地盤種別、重要度係数、応答スペクトル、SDS/SD1',
     '最重要', '1〜2週目', 'LE/構造担当', 'LE', 'PM', '地盤調査会社', 'チーム全員'),
    (2, '', '荷重条件の整理（温度荷重）',
     '設計温度範囲、熱膨張による部材応力・変位の検討要否',
     '重要', '1〜2週目', 'LE/構造担当', 'LE', 'PM', 'プロセス部門', ''),
    (2, '', '荷重条件の整理（配管反力・機器荷重）',
     'Equipment Load Data Sheet, Piping Stress Analysisからの荷重取得方法',
     '最重要', '2〜3週目', 'LE/構造担当', 'LE', 'PM', '配管応力/機器部門', 'チーム全員'),
    (2, '', '荷重条件の整理（特殊荷重）',
     'Blast Load, Snow Load, Ice Load, Crane Load 等の適用有無確認',
     '重要', '1〜2週目', 'LE', 'LE', 'PM', 'クライアント/HSE', ''),
    (2, '', '荷重組合せ（Load Combination）の作成',
     'ASD/LRFD別、通常時/地震時/風時/据付時の荷重組合せ表作成',
     '最重要', '2〜3週目', 'LE', 'LE', 'PM', 'QA', 'チーム全員'),
    (2, '', '地盤条件・基礎設計条件の確認',
     '土質調査報告書、地耐力、杭仕様、液状化判定結果の入手',
     '重要', '1〜2週目', 'LE', 'LE', 'PM', '土木部門/地盤調査会社', ''),
    (2, '', '耐火要件の確認',
     'Fireproofing要否、耐火被覆厚、耐火時間（1hr/2hr/3hr）の確認',
     '重要', '2〜3週目', 'LE', 'LE', 'PM', 'HSE/防災部門', ''),
    (2, '', '防食要件の確認',
     '塗装仕様（系統別）、溶融亜鉛めっき要否、CUI対策の確認',
     '重要', '2〜3週目', 'LE', 'LE', 'PM', '防食担当/クライアント', ''),
    (2, '', '材料規格の選定',
     'SS400/SM490/SN490, ASTM A36/A992/A500, ボルトF10T/A325/A490 等',
     '最重要', '1〜2週目', 'LE', 'LE', 'PM', '調達/QA', 'チーム全員'),
    (2, '', '設計ソフトウェア・ツールの確認',
     'STAAD.Pro, SAP2000, ETABS, Tekla, AutoCAD, MathCAD等のライセンス確認',
     '重要', '1週目', 'LE', 'LE', 'PM', 'IT/CAD管理者', 'チーム全員'),
    (2, '', '接合部設計方針の策定',
     'ボルト接合/溶接接合の使い分け、標準接合ディテールの選定',
     '重要', '2〜3週目', 'LE', 'LE', 'PM', 'Fabricator', ''),
    (2, '', 'Deflection / Drift 許容値の設定',
     '梁たわみ L/240〜L/360、層間変形角 1/200〜1/300 等の基準値設定',
     '重要', '1〜2週目', 'LE', 'LE', 'PM', 'クライアント', 'チーム全員'),
    (2, '', '振動検討要否の確認',
     '回転機器基礎、Flare構造、塔槽類支持架構の動的解析要否判定',
     '重要', '2〜3週目', 'LE', 'LE', 'PM', '機器/プロセス部門', ''),

    # ====== Phase 3: 体制構築とスケジュール策定 ======
    (3, '体制構築とスケジュール策定', 'チーム編成・人員計画の作成',
     '必要人員数、スキルレベル、配置時期、外注比率の計画',
     '最重要', '着任〜1週間', 'LE', 'LE', 'PM/部門長', 'HR', 'チーム全員'),
    (3, '', 'Man-Hour見積り・工数配分の策定',
     '構造物別・フェーズ別（基本/詳細/照査）のMH配分、進捗管理基準',
     '最重要', '1〜2週目', 'LE', 'LE', 'PM', 'プロジェクトコントロール', ''),
    (3, '', '設計マスタースケジュールの作成',
     'IFA/IFR/IFC等マイルストーン、クライアント承認期間を含む工程表',
     '最重要', '1〜2週目', 'LE', 'LE', 'PM', 'スケジュール担当', 'チーム全員'),
    (3, '', '図面リスト（Drawing List）の作成',
     'GA図, Detail図, Foundation Plan等の全図面リストとスケジュール',
     '重要', '2〜3週目', 'LE/シニア担当', 'LE', 'PM', 'Document Control', 'チーム全員'),
    (3, '', '設計計算書リスト（Calculation List）の作成',
     '構造物別の計算書一覧、担当者割当、提出スケジュール',
     '重要', '2〜3週目', 'LE/シニア担当', 'LE', 'PM', 'Document Control', 'チーム全員'),
    (3, '', '他部門とのInterface Matrix作成',
     '配管・機器・電気・計装・土建・プロセスとの入出力情報整理',
     '最重要', '1〜2週目', 'LE', 'LE', 'PM', '各部門LE', 'チーム全員'),
    (3, '', '外注設計スコープの明確化',
     '外注範囲、成果物定義、品質要求、スケジュール、検収基準',
     '重要', '2〜4週目', 'LE', 'LE', 'PM', '外注先/調達', ''),
    (3, '', '設計進捗報告の仕組み構築',
     '週次/月次レポートフォーマット、KPI（S-curve, MH消化率）設定',
     '重要', '2〜3週目', 'LE', 'LE', 'PM', 'プロジェクトコントロール', ''),
    (3, '', '会議体の設定',
     '週次チーム会議、IDC会議、クライアント定例の頻度・参加者設定',
     '標準', '1〜2週目', 'LE', 'LE', 'PM', '各部門LE', ''),
    (3, '', 'コミュニケーションルールの策定',
     'メール/RFI/TQ発行ルール、ファイル共有方法、命名規則',
     '標準', '1〜2週目', 'LE', 'LE', 'PM', 'Document Control', 'チーム全員'),

    # ====== Phase 4: 設計基盤の準備 ======
    (4, '設計基盤の準備', '標準図・標準ディテールの選定',
     '接合部詳細、ベースプレート、アンカーボルト、ブレース接合等',
     '重要', '2〜3週目', 'LE/シニア担当', 'LE', 'PM', 'Fabricator', 'チーム全員'),
    (4, '', 'CADテンプレート・図枠の準備',
     '図枠、レイヤー設定、線種、文字スタイル、縮尺設定の標準化',
     '重要', '1〜2週目', 'CAD担当', 'LE', 'PM', 'CAD管理者', 'チーム全員'),
    (4, '', '構造解析モデルの基本方針策定',
     '2D/3Dモデル使い分け、境界条件、メッシュ方針、解析ケース定義',
     '最重要', '2〜3週目', 'LE', 'LE', 'PM', 'QA', 'チーム全員'),
    (4, '', '構造物タイプ分類と設計方針',
     'Pipe Rack, Equipment Support, Access Platform, Stack Support等の分類・方針',
     '重要', '2〜3週目', 'LE', 'LE', 'PM', '各部門LE', 'チーム全員'),
    (4, '', 'Key Plan / Plot Planからの構造物リスト作成',
     '全鉄骨構造物の洗い出し、優先順位付け、概算重量見積り',
     '最重要', '1〜3週目', 'LE/シニア担当', 'LE', 'PM', 'レイアウト担当', 'チーム全員'),
    (4, '', '設計計算書テンプレートの整備',
     '表紙、目次、設計条件、計算本体、結論の標準フォーマット作成',
     '重要', '2〜3週目', 'LE/シニア担当', 'LE', 'PM', 'QA', 'チーム全員'),
    (4, '', '番号体系（マーク番号・図番ルール）の設定',
     '構造物マーク、部材マーク、図面番号、計算書番号の付番ルール',
     '重要', '1〜2週目', 'LE', 'LE', 'PM', 'Document Control', 'チーム全員'),
    (4, '', 'BIMモデル方針の確認',
     '3Dモデル作成範囲、LOD、他部門モデルとの統合方法の確認',
     '重要', '2〜3週目', 'LE/3D担当', 'LE', 'PM', '3D統合担当', ''),
    (4, '', '既製品・標準品のカタログ整備',
     'グレーチング、手摺、階段、チェッカープレート等のメーカー選定',
     '標準', '3〜4週目', '担当者', 'LE', 'PM', '調達', ''),
    (4, '', '設計用スプレッドシート・ツールの整備',
     'ベースプレート計算、ボルト本数算定、部材選定ツール等の準備',
     '標準', '2〜4週目', 'LE/シニア担当', 'LE', 'PM', '', 'チーム全員'),

    # ====== Phase 5: 品質管理・レビュー体制 ======
    (5, '品質管理・レビュー体制', '設計チェックリストの整備',
     '図面チェックリスト、計算書チェックリスト、モデルチェックリスト',
     '最重要', '2〜3週目', 'LE', 'LE', 'PM/QA', 'QA', 'チーム全員'),
    (5, '', '設計照査（Design Review）プロセスの確立',
     'Self Check → Checker → LE承認の3段階照査フロー策定',
     '最重要', '2〜3週目', 'LE', 'LE', 'PM/QA', 'QA', 'チーム全員'),
    (5, '', 'Interdisciplinary Check（IDC）手順の確認',
     '他部門との相互チェック方法、タイミング、記録様式の設定',
     '重要', '2〜3週目', 'LE', 'LE', 'PM', '各部門LE', 'チーム全員'),
    (5, '', 'クライアント承認プロセスの確認',
     'IFA/IFR/IFC提出フロー、承認期間、コメント回答期限の確認',
     '最重要', '1〜2週目', 'LE', 'LE', 'PM', 'クライアント/DC', 'チーム全員'),
    (5, '', '設計変更管理（MOC）手順の確認',
     'Design Change Notice発行基準、影響評価、承認フロー',
     '重要', '2〜3週目', 'LE', 'LE', 'PM', 'QA', ''),
    (5, '', 'NCR（不適合報告）手順の確認',
     '設計不適合の検出、記録、是正処置、再発防止のプロセス',
     '重要', '2〜3週目', 'LE', 'LE', 'PM/QA', 'QA', ''),
    (5, '', 'QMS要求事項の確認',
     'ISO 9001準拠要求、PJ固有品質計画書（PQP）の確認',
     '標準', '2〜3週目', 'LE', 'LE', 'PM/QA', 'QA', ''),
    (5, '', 'Technical Query (TQ) / RFI 管理',
     'TQ/RFI発行・回答の追跡管理方法、台帳フォーマット整備',
     '重要', '1〜2週目', 'LE', 'LE', 'PM', 'Document Control', 'チーム全員'),
    (5, '', 'Lesson Learned 収集・記録の仕組み',
     'PJ進行中の気づき・改善点の記録方法、定期レビューの設定',
     '標準', '3〜4週目', 'LE', 'LE', 'PM', '', 'チーム全員'),

    # ====== Phase 6: キックオフ・初期アクション ======
    (6, 'キックオフ・初期アクション', 'チーム内キックオフミーティングの実施',
     '設計方針、スケジュール、役割分担、品質基準の共有',
     '最重要', '1〜2週目', 'LE', 'LE', 'PM', '', 'チーム全員'),
    (6, '', '他部門との設計条件確認会議',
     '配管・機器・電気・土建の各LEとの荷重条件・Interface確認',
     '最重要', '2〜3週目', 'LE', 'LE', 'PM', '各部門LE', ''),
    (6, '', 'クライアント キックオフ / Design Review Meeting',
     '設計方針・基準・スケジュールのクライアント確認・承認',
     '最重要', '2〜4週目', 'LE', 'LE', 'PM', 'クライアント', ''),
    (6, '', 'Long Lead Item 関連構造物の特定・着手',
     '大型機器基礎、Pipe Rack等のLLI関連鉄骨の優先設計着手',
     '最重要', '2〜3週目', 'LE/担当者', 'LE', 'PM', '調達/機器部門', ''),
    (6, '', 'MTO（Material Take Off）初版の計画',
     '概算鋼材量算出、材料発注リードタイム逆算、MTO提出計画',
     '重要', '3〜4週目', 'LE/担当者', 'LE', 'PM', '調達', ''),
    (6, '', 'クライアントへの設計方針説明資料作成',
     'Design Philosophy / Basis of Design の作成、提出、承認取得',
     '最重要', '2〜4週目', 'LE', 'LE', 'PM', 'クライアント', ''),
    (6, '', 'リスク・未確定事項（TBD/TBC）一覧の作成',
     '未受領情報、未確定条件、仮定事項の一覧と解決予定日の管理',
     '最重要', '1〜2週目', 'LE', 'LE', 'PM', '各部門LE', 'チーム全員'),
    (6, '', 'Hold / Assumption リストの作成',
     '設計保留事項と仮定条件の正式記録、クライアント確認',
     '重要', '2〜3週目', 'LE', 'LE', 'PM', 'クライアント', ''),
    (6, '', '初回進捗報告の実施',
     '着任後の状況把握結果、課題、今後の計画をPMに報告',
     '重要', '2〜3週目', 'LE', 'LE', 'PM', '', 'PM/チーム全員'),
    (6, '', 'Safety / HSE 要求事項の確認',
     '設計段階でのHAZOP対応、安全距離、避難経路確保等の確認',
     '重要', '2〜4週目', 'LE', 'LE', 'PM', 'HSE部門', ''),
]

# === データ書込み ===
current_row = 5
item_no = 1
for phase, phase_name, action, detail, importance, timing, person, r, a, c, i in data:
    row_data = [item_no, phase_name if phase_name else '', action, detail,
                importance, timing, person, r, a, c, i, '未着手']

    phase_fill = PatternFill(start_color=phase_colors[phase], end_color=phase_colors[phase], fill_type='solid')

    for col_idx, val in enumerate(row_data, 1):
        cell = ws.cell(row=current_row, column=col_idx, value=val)
        cell.font = normal_font
        cell.alignment = wrap_align if col_idx in (3, 4, 10, 11) else center_align
        cell.border = thin_border

        # フェーズ列の色
        if col_idx == 2 and val:
            cell.fill = phase_fill
            cell.font = Font(name='Meiryo', bold=True, size=10, color='FFFFFF')

        # 重要度の色分け
        if col_idx == 5 and val in importance_fills:
            cell.fill = importance_fills[val]

    ws.row_dimensions[current_row].height = 36
    current_row += 1
    item_no += 1

# === 凡例シート ===
ws2 = wb.create_sheet('凡例・RACI説明')
legend_data = [
    ['RACI マトリクス 凡例', '', '', ''],
    ['記号', '役割', '英語', '説明'],
    ['R', '実行責任者', 'Responsible', '実際にタスクを実行する担当者'],
    ['A', '説明責任者', 'Accountable', '最終的な承認・決定権限を持つ者（LE自身が多い）'],
    ['C', '協議先', 'Consulted', '事前に意見を求める関係者（双方向コミュニケーション）'],
    ['I', '報告先', 'Informed', '結果を事後報告する関係者（一方向コミュニケーション）'],
    ['', '', '', ''],
    ['重要度 凡例', '', '', ''],
    ['最重要', 'LE着任直後に最優先で実施すべき項目', '', ''],
    ['重要', '初期段階で必ず実施すべき項目', '', ''],
    ['標準', '通常の優先度で実施する項目', '', ''],
]
for r_idx, row_vals in enumerate(legend_data, 1):
    for c_idx, val in enumerate(row_vals, 1):
        cell = ws2.cell(row=r_idx, column=c_idx, value=val)
        cell.font = Font(name='Meiryo', size=10, bold=(r_idx in (1, 2, 8)))
        if r_idx in (1, 8):
            cell.font = Font(name='Meiryo', size=12, bold=True, color='1A237E')
ws2.column_dimensions['A'].width = 12
ws2.column_dimensions['B'].width = 40
ws2.column_dimensions['C'].width = 16
ws2.column_dimensions['D'].width = 50

# === 保存 ===
output_path = '/home/user/Open_Test/Steel_Design_LE_ActionList.xlsx'
wb.save(output_path)
print(f'Excel file saved: {output_path}')

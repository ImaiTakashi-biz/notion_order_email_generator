import os
import re
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, portrait
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont, TTFError
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer

import config

# --- 定数 ---
FONT_NAME = "MSPGothic"
FONT_PATH = "C:\\Windows\\Fonts\\msgothic.ttc"
FALLBACK_FONT_NAME = "Helvetica"

_REGISTERED_FONT_NAME = FALLBACK_FONT_NAME

# --- 日本語フォントの登録 ---
def register_japanese_font() -> str:
    """日本語フォントを登録し、使用するフォント名を返す"""
    global _REGISTERED_FONT_NAME
    if FONT_NAME in pdfmetrics.getRegisteredFontNames():
        _REGISTERED_FONT_NAME = FONT_NAME
        return _REGISTERED_FONT_NAME
    try:
        if not os.path.exists(FONT_PATH):
            raise FileNotFoundError(FONT_PATH)
        pdfmetrics.registerFont(TTFont(FONT_NAME, FONT_PATH, subfontIndex=1))
        _REGISTERED_FONT_NAME = FONT_NAME
    except (FileNotFoundError, TTFError) as err:
        print(f"⚠️ フォント登録に失敗しました ({err})。既定フォントを使用します。")
        _REGISTERED_FONT_NAME = FALLBACK_FONT_NAME
    return _REGISTERED_FONT_NAME

# --- スタイルの定義 ---
def get_custom_styles() -> Dict[str, ParagraphStyle]:
    """PDF用のカスタムスタイルを返す"""
    font_name = register_japanese_font()
    styles = getSampleStyleSheet()
    
    leading = 15
    # タイトル: サイズ22
    styles.add(ParagraphStyle(name='Title_J', parent=styles['h1'], fontName=font_name, fontSize=22, alignment=1, spaceAfter=10))
    # その他: サイズ11
    styles.add(ParagraphStyle(name='Normal_J', parent=styles['Normal'], fontName=font_name, fontSize=11, leading=leading))
    styles.add(ParagraphStyle(name='Right_J', parent=styles['Normal'], fontName=font_name, fontSize=11, alignment=2, leading=leading))
    styles.add(ParagraphStyle(name='Center_J', parent=styles['Normal'], fontName=font_name, fontSize=11, alignment=1, leading=leading))
    # 仕入先宛先先担当者: サイズ14
    styles.add(ParagraphStyle(name='Supplier_J', parent=styles['Normal'], fontName=font_name, fontSize=14, leading=17))
    # 担当者名（インデント付き）用スタイル: 全角一文字分のインデント
    styles.add(ParagraphStyle(name='Supplier_Indent_J', parent=styles['Supplier_J'], leftIndent=44))

    return styles

def create_order_pdf(
    supplier_name: str,
    items: List[Dict[str, Any]],
    sales_contact: str,
    sender_info: Dict[str, str],
    selected_department: Optional[str] = None,
    save_dir: Optional[str] = None
) -> Optional[str]:
    """
    reportlabを使用してExcelレイアウト風の注文書PDFを直接生成する。
    """
    try:
        # --- ファイルパスの準備 ---
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        safe_supplier_name = re.sub(r'[\\/:*?"<>|]', '_', supplier_name)
        pdf_filename = f"{timestamp}_{safe_supplier_name}_注文書.pdf"
        
        base_save_dir = save_dir if save_dir is not None else config.PDF_SAVE_DIR
        if not base_save_dir: return None

        target_save_dir = base_save_dir
        if selected_department:
            department_dir = os.path.join(base_save_dir, selected_department)
            if not os.path.exists(department_dir): os.makedirs(department_dir, exist_ok=True)
            target_save_dir = department_dir

        if not os.path.exists(target_save_dir): os.makedirs(target_save_dir)
        pdf_path = os.path.join(target_save_dir, pdf_filename)

        # --- ドキュメントとスタイルの準備 ---
        doc = SimpleDocTemplate(pdf_path, pagesize=portrait(A4), topMargin=15*mm, bottomMargin=15*mm, leftMargin=15*mm, rightMargin=15*mm)
        styles = get_custom_styles()
        story = []

        # --- 差出人情報 ---
        company_info = config.AppConstants.COMPANY_INFO
        guidance_number = sender_info.get('guidance_number', '')
        tel_line = f"TEL: {company_info['tel_base']}" + (f"（ガイダンス{guidance_number}番）" if guidance_number else "")
        sender_p_list = [
            Paragraph(company_info['name'], styles['Normal_J']),
            Paragraph(f"{company_info['postal_code']} {company_info['address']}", styles['Normal_J']),
            Paragraph(tel_line, styles['Normal_J']),
            Paragraph(f"URL: {company_info['url']}", styles['Normal_J']),
            Paragraph(f"担当: {selected_department or ''} {sender_info.get('name', '')}", styles['Normal_J']),
            Paragraph(f"Email: {sender_info.get('email', '')}", styles['Normal_J'])
        ]

        # --- レイアウトの構築 ---
        # 1. 発行日 (一番上、右寄せ)
        issue_date_table = Table([[Paragraph(f"発行日: {datetime.now().strftime('%Y/%m/%d')}", styles['Right_J'])]], colWidths=[180*mm])
        story.append(issue_date_table)

        # 2. タイトル
        story.append(Paragraph("注 文 書", styles['Title_J']))
        story.append(Spacer(1, 8*mm))

        # 3. 宛先と差出人
        # 宛先をParagraphのリストとして作成
        supplier_p_list = [
            Paragraph(f"{supplier_name} 御中", styles['Supplier_J']),
            Spacer(1, 14), # 1行分の改行スペース
            Paragraph(f"{sales_contact} 様", styles['Supplier_Indent_J']) # インデント付きスタイルを適用
        ]

        header_data = [
            [supplier_p_list, ''], # 1行目: 宛先, (空)
            ['', sender_p_list]      # 2行目: (空), 差出人
        ]
        header_table = Table(header_data, colWidths=[110*mm, 70*mm])
        header_table.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('SPAN', (0,0), (0,1)), # 宛先のセルを縦に結合
        ]))
        story.append(header_table)
        story.append(Spacer(1, 10*mm))

        # 4. 挨拶文
        story.append(Paragraph("以下の通りご注文申し上げます。", styles['Normal_J']))
        story.append(Paragraph("2日以内に納期回答をご記入の上、ご返信頂けますよう宜しくお願い致します。", styles['Normal_J']))
        story.append(Spacer(1, 5*mm))

        # 5. 注文明細テーブル
        table_header = [Paragraph(h, styles['Center_J']) for h in ["品番（品名）", "メーカー", "数量", "回答納期", "備考"]]
        table_data = [table_header]
        
        url_pattern = r'(https?://[\w\-./?%&=]+)'

        for item in items:
            remarks_text = str(item.get('remarks', ''))
            # URLを検出し、ハイパーリンクに変換
            linked_remarks = re.sub(url_pattern, r'<a href="\1" color="blue">\1</a>', remarks_text)

            row = [
                Paragraph(str(item.get('db_part_number', '')), styles['Normal_J']),
                Paragraph(str(item.get('maker_name', '')), styles['Normal_J']),
                Paragraph(str(item.get('quantity', 0)), styles['Center_J']),
                '',  # 回答納期 (空欄)
                Paragraph(linked_remarks, styles['Normal_J'])   # 備考
            ]
            table_data.append(row)

        item_table = Table(table_data, colWidths=[65*mm, 40*mm, 15*mm, 25*mm, 35*mm], repeatRows=1)
        item_table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONT', (0,0), (-1,0), _REGISTERED_FONT_NAME, 11), # ヘッダーフォント
        ]))
        story.append(item_table)
        
        doc.build(story)
        
        return pdf_path

    except Exception as e:
        print(f"❌ PDF作成中に予期せぬエラーが発生しました (reportlab): {e}")
        return None

def generate_order_pdf_flow(
    supplier_name: str,
    items: List[Dict[str, Any]],
    sender_info: Dict[str, str],
    selected_department: Optional[str] = None,
    save_dir: Optional[str] = None
) -> Tuple[Optional[str], Optional[Dict[str, Any]], Optional[str]]:
    """
    UIからの情報を受け取り、PDF作成のフロー全体を管理する
    """
    try:
        if not items: return None, None, "対象アイテムが見つかりません。"

        pdf_path = create_order_pdf(
            supplier_name, items, items[0]["sales_contact"], sender_info,
            selected_department=selected_department, save_dir=save_dir
        )
        
        if not pdf_path: return None, None, "PDF作成中にエラーが発生しました。コンソールログを確認してください。"

        return pdf_path, items[0], None

    except Exception as e:
        error_message = f"予期せぬエラーが発生しました: {e}"
        print(f"❌ PDF生成フロー中に{error_message}")
        return None, None, error_message

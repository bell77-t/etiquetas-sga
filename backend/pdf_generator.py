"""
pdf_generator.py - Generador de Etiquetas Fitosanitarias SGA en PDF Vectorial
Utiliza ReportLab para generar documentos imprimibles en hojas Carta (multietiqueta) o Rollos Térmicos (100x150 mm).
"""

import io
from pathlib import Path
from typing import List, Dict, Any, Optional

import qrcode
from PIL import Image as PILImage

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage, PageBreak, KeepTogether
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

def get_qr_image(data_url: str, size_pt: float = 45) -> RLImage:
    """Genera un código QR en memoria para ReportLab."""
    qr = qrcode.QRCode(box_size=3, border=1)
    qr.add_data(data_url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return RLImage(buf, width=size_pt, height=size_pt)

def get_pictogram_image(picto_info: Dict[str, Any], picto_dir: Path, size_pt: float = 38) -> Optional[RLImage]:
    """Carga la imagen del pictograma GHS si existe."""
    if not picto_info or not picto_info.get("has_image"):
        return None
    
    fname = picto_info.get("filename")
    if not fname:
        return None
    
    fpath = picto_dir / fname
    if not fpath.exists():
        return None
    
    try:
        return RLImage(str(fpath), width=size_pt, height=size_pt)
    except Exception:
        return None

def build_label_table(label: Dict[str, Any], picto_dir: Path, qr_base_url: str, total_width_mm: float = 190.0) -> Table:
    """Construye una tabla ReportLab con la réplica exacta de la etiqueta SGA A2:F12."""
    styles = getSampleStyleSheet()
    
    # Estilos tipográficos
    style_header = ParagraphStyle('LHeader', fontName='Helvetica-Bold', fontSize=8, textColor=colors.HexColor('#475569'), alignment=0)
    style_header_c = ParagraphStyle('LHeaderC', fontName='Helvetica-Bold', fontSize=8, textColor=colors.HexColor('#475569'), alignment=1)
    style_val_bold = ParagraphStyle('LValB', fontName='Helvetica-Bold', fontSize=9, textColor=colors.HexColor('#0f172a'), alignment=0)
    style_val_bold_c = ParagraphStyle('LValBC', fontName='Helvetica-Bold', fontSize=9, textColor=colors.HexColor('#0f172a'), alignment=1)
    style_banner_h = ParagraphStyle('LBannerH', fontName='Helvetica-Bold', fontSize=8, textColor=colors.HexColor('#334155'), alignment=1)
    style_frase = ParagraphStyle('LFrase', fontName='Helvetica', fontSize=6.8, leading=8.5, textColor=colors.HexColor('#1e293b'), alignment=0)
    
    base = label.get("base_info", {})
    prod_name = label.get("producto", "N/A")
    vol_tanque = label.get("litros_tanque", 1000)
    dosis = label.get("dosis", 0)
    cant_dosificar = label.get("cantidad_dosificar", 0)
    um = base.get("um", "L/KG")
    sector = label.get("sector_bloque", "N/A")
    cultivo = label.get("cultivo", "VARIOS")
    reentrada = label.get("reentrada", 0)
    fecha = label.get("fecha", {}).get("display", "N/A")
    semana = label.get("semana", "9")
    t_num = label.get("tanque_numero", 1)
    t_tot = label.get("total_tanques", 1)
    advertencia = base.get("palabra_advertencia", "PELIGRO").strip().upper()
    if not advertencia:
        advertencia = "PELIGRO"
        
    color_adv = colors.HexColor('#dc2626') if advertencia == "PELIGRO" else colors.HexColor('#d97706')
    style_adv = ParagraphStyle('LAdv', fontName='Helvetica-Bold', fontSize=10, textColor=color_adv, alignment=1)
    
    # Pictogramas
    pictos = base.get("pictogramas", [])
    picto_imgs = []
    for p in pictos[:4]:
        img_obj = get_pictogram_image(p, picto_dir, size_pt=34)
        if img_obj:
            picto_imgs.append(img_obj)
            
    # Contenedor de pictogramas (fila horizontal)
    if picto_imgs:
        picto_table_data = [picto_imgs]
        picto_cell = Table(picto_table_data, colWidths=[38 for _ in picto_imgs])
        picto_cell.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('LEFTPADDING', (0,0), (-1,-1), 1),
            ('RIGHTPADDING', (0,0), (-1,-1), 1),
            ('TOPPADDING', (0,0), (-1,-1), 1),
            ('BOTTOMPADDING', (0,0), (-1,-1), 1),
        ]))
    else:
        picto_cell = Paragraph("<font size=7 color='#64748b'>SIN PICTOGRAMAS ASIGNADOS</font>", style_header_c)

    categoria = label.get("categoria", "")

    # Frases H y P
    frase_h = base.get("frase_h") or "No clasificado como peligroso / Sin frases H registradas."
    frase_p = base.get("frase_p") or "P102 Manténgase fuera del alcance de los niños.\nP270 No comer, beber ni fumar durante su utilización."

    col_w = (total_width_mm * mm) / 6.0
    
    table_data = [
        # Fila 0: Titulo SGA
        [Paragraph("<b>SISTEMA GLOBALMENTE ARMONIZADO &bull; MEZCLA FITOSANITARIA</b>", ParagraphStyle('T', fontName='Helvetica-Bold', fontSize=8, textColor=colors.white, alignment=1)), "", "", "", "", ""],
        # Fila 1: Encabezados Superiores
        [Paragraph("PRODUCTO", style_header), "", "", "", Paragraph("FECHA APLICACIÓN", style_header_c), Paragraph("SEMANA", style_header_c)],
        # Fila 2: Valores Superiores
        [Paragraph(f"<b>{prod_name}</b>", style_val_bold), "", "", "", Paragraph(f"<b>{fecha}</b>", style_val_bold_c), Paragraph(f"<b>{semana}</b>", style_val_bold_c)],
        # Fila 3: Encabezados Dosis
        [Paragraph("DOSIS x LITRO", style_header_c), Paragraph("LITROS AGUA (TANQUE)", style_header_c), Paragraph("CANTIDAD A DOSIFICAR", style_header_c), "", Paragraph("TANQUE NRO.", style_header_c), Paragraph("REENTRADA", style_header_c)],
        # Fila 4: Valores Dosis
        [Paragraph(f"<b>{dosis}</b>", style_val_bold_c), Paragraph(f"<b>{vol_tanque:g} L</b>", style_val_bold_c), Paragraph(f"<b>{cant_dosificar:.2f} {um}</b>", style_val_bold_c), "", Paragraph(f"<b>{t_num} de {t_tot}</b>", style_val_bold_c), Paragraph(f"<b>{reentrada} HORAS</b>", style_val_bold_c)],
        # Fila 5: Encabezados Sector / Cultivo / Categoría
        [Paragraph("SECTOR / BLOQUE", style_header), "", Paragraph("CULTIVO", style_header), "", Paragraph("CATEGORÍA / TIPO", style_header_c), ""],
        # Fila 6: Valores Sector / Cultivo / Categoría
        [Paragraph(f"<b>{sector}</b>", style_val_bold), "", Paragraph(f"<b>{cultivo}</b>", style_val_bold), "", Paragraph(f"<b>{categoria or 'FITOSANITARIO'}</b>", style_val_bold_c), ""],
        # Fila 7: Pictogramas + Advertencia
        [picto_cell, "", "", "", Paragraph(f"<b>{advertencia}</b>", style_adv), ""],
        # Fila 8: FRASE H Banner
        [Paragraph("FRASE H", style_banner_h), "", "", "", "", ""],
        # Fila 9: FRASE H Contenido
        [Paragraph(frase_h.replace("\n", "<br/>"), style_frase), "", "", "", "", ""],
        # Fila 10: FRASE P Banner
        [Paragraph("FRASE P", style_banner_h), "", "", "", "", ""],
        # Fila 11: FRASE P Contenido
        [Paragraph(frase_p.replace("\n", "<br/>"), style_frase), "", "", "", "", ""],
    ]

    t = Table(table_data, colWidths=[col_w]*6)
    t.setStyle(TableStyle([
        # Borde exterior de la etiqueta
        ('BOX', (0,0), (-1,-1), 1.2, colors.HexColor('#0f172a')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#cbd5e1')),
        
        # Banner Titulo
        ('SPAN', (0,0), (5,0)),
        ('BACKGROUND', (0,0), (5,0), colors.HexColor('#0f172a')),
        ('TOPPADDING', (0,0), (5,0), 3),
        ('BOTTOMPADDING', (0,0), (5,0), 3),

        # Spans Fila 1 & 2
        ('SPAN', (0,1), (3,1)),
        ('SPAN', (0,2), (3,2)),
        ('BACKGROUND', (0,1), (5,1), colors.HexColor('#f8fafc')),

        # Spans Fila 3 & 4
        ('SPAN', (2,3), (3,3)),
        ('SPAN', (2,4), (3,4)),
        ('BACKGROUND', (0,3), (5,3), colors.HexColor('#f8fafc')),

        # Spans Fila 5 & 6
        ('SPAN', (0,5), (1,5)),
        ('SPAN', (0,6), (1,6)),
        ('SPAN', (2,5), (3,5)),
        ('SPAN', (2,6), (3,6)),
        ('SPAN', (4,5), (5,5)),
        ('SPAN', (4,6), (5,6)),
        ('BACKGROUND', (0,5), (5,5), colors.HexColor('#f8fafc')),
        ('ALIGN', (4,6), (5,6), 'CENTER'),
        ('VALIGN', (4,6), (5,6), 'MIDDLE'),

        # Spans Fila 7 (Pictogramas + Advertencia)
        ('SPAN', (0,7), (3,7)),
        ('SPAN', (4,7), (5,7)),
        ('VALIGN', (0,7), (5,7), 'MIDDLE'),
        ('ALIGN', (4,7), (5,7), 'CENTER'),

        # Banner FRASE H
        ('SPAN', (0,8), (5,8)),
        ('BACKGROUND', (0,8), (5,8), colors.HexColor('#f1f5f9')),
        ('TOPPADDING', (0,8), (5,8), 2),
        ('BOTTOMPADDING', (0,8), (5,8), 2),

        # FRASE H Texto
        ('SPAN', (0,9), (5,9)),
        ('TOPPADDING', (0,9), (5,9), 3),
        ('BOTTOMPADDING', (0,9), (5,9), 3),

        # Banner FRASE P
        ('SPAN', (0,10), (5,10)),
        ('BACKGROUND', (0,10), (5,10), colors.HexColor('#f1f5f9')),
        ('TOPPADDING', (0,10), (5,10), 2),
        ('BOTTOMPADDING', (0,10), (5,10), 2),

        # FRASE P Texto
        ('SPAN', (0,11), (5,11)),
        ('TOPPADDING', (0,11), (5,11), 3),
        ('BOTTOMPADDING', (0,11), (5,11), 3),
    ]))
    return t

def generate_pdf(
    labels: List[Dict[str, Any]],
    picto_dir: Path,
    qr_base_url: str = "http://127.0.0.1:8000",
    layout: str = "letter"
) -> io.BytesIO:
    """Genera el archivo PDF con todas las etiquetas dadas."""
    buf = io.BytesIO()

    if layout == "thermal":
        # Rollo térmico estándar 100 x 150 mm (4 x 6 pulgadas)
        page_w = 100 * mm
        page_h = 150 * mm
        doc = SimpleDocTemplate(
            buf,
            pagesize=(page_w, page_h),
            leftMargin=3*mm,
            rightMargin=3*mm,
            topMargin=4*mm,
            bottomMargin=4*mm
        )
        elements = []
        for idx, lbl in enumerate(labels):
            tbl = build_label_table(lbl, picto_dir, qr_base_url, total_width_mm=94.0)
            elements.append(tbl)
            if idx < len(labels) - 1:
                elements.append(PageBreak())
        doc.build(elements)

    else:
        # Formato Carta / A4 (2 etiquetas por hoja con línea de corte)
        doc = SimpleDocTemplate(
            buf,
            pagesize=letter,
            leftMargin=8*mm,
            rightMargin=8*mm,
            topMargin=10*mm,
            bottomMargin=10*mm
        )
        elements = []
        for idx, lbl in enumerate(labels):
            tbl = build_label_table(lbl, picto_dir, qr_base_url, total_width_mm=195.0)
            elements.append(tbl)
            
            # Separador o salto de página (2 por hoja)
            if (idx + 1) % 2 == 0 and idx < len(labels) - 1:
                elements.append(PageBreak())
            elif idx < len(labels) - 1:
                elements.append(Spacer(1, 10*mm))
                
        doc.build(elements)

    buf.seek(0)
    return buf

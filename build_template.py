import os
import zipfile
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

sections_info = [
    (1, "Documentación"),
    (2, "Transformador principal"),
    (3, "Sistema de puesta a tierra"),
    (4, "Sistema contra descargas atmosféricas"),
    (5, "Espacios físicos cuartos eléctricos"),
    (6, "Transformadores de baja tensión"),
    (7, "Tableros y protecciones"),
    (8, "Conductores"),
    (9, "Canalizaciones, bandejas, cajas de registro y conexiones"),
    (10, "Iluminación"),
    (11, "Tomacorrientes y apagadores"),
    (12, "Equipos de aire acondicionado, ventilación y refrigeración"),
    (13, "Motores"),
    (14, "UPS"),
    (15, "Sistemas de emergencia"),
    (16, "16.1. Lugares Clase I"),
    (17, "16.2. Lugares Clase II"),
    (18, "16.3. Lugares Clase III"),
    (19, "Sitios de Reunión"),
    (20, "Teatros y lugares similares"),
    (21, "Carnavales, circos, ferias y lugares similares"),
    (22, "Piscinas"),
    (23, "Instalaciones de asistencia médica")
]

def add_checkbox_control(doc, sec_id, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)

    sdt = parse_xml(f'''
        <w:sdt {nsdecls("w")}>
            <w:sdtPr>
                <w:rPr>
                    <w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic" w:hint="default"/>
                </w:rPr>
                <w:alias w:val="{title}"/>
                <w:tag w:val="CHK_{sec_id}"/>
                <w:id w:val="{10000 + sec_id}"/>
                <w14:checkbox {nsdecls("w14")}>
                    <w14:checked w14:val="0"/>
                    <w14:checkedState w14:val="2612" w14:font="MS Gothic"/>
                    <w14:uncheckedState w14:val="25A2" w14:font="MS Gothic"/>
                </w14:checkbox>
            </w:sdtPr>
            <w:sdtContent>
                <w:r>
                    <w:rPr>
                        <w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic" w:hint="default"/>
                    </w:rPr>
                    <w:t>☐</w:t>
                </w:r>
            </w:sdtContent>
        </w:sdt>
    ''')

    p._p.append(sdt)

    run = p.add_run(f"  {sec_id}. {title}")
    run.font.name = "Calibri"
    run.font.size = Pt(11)

def add_bookmark(p, bm_id, bm_name):
    bm_start = parse_xml(f'<w:bookmarkStart {nsdecls("w")} w:id="{bm_id}" w:name="{bm_name}"/>')
    bm_end = parse_xml(f'<w:bookmarkEnd {nsdecls("w")} w:id="{bm_id}"/>')
    p._p.append(bm_start)
    p._p.append(bm_end)

def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def create_template():
    doc = Document()

    for sec in doc.sections:
        sec.top_margin = Inches(1)
        sec.bottom_margin = Inches(1)
        sec.left_margin = Inches(1)
        sec.right_margin = Inches(1)

    title_p = doc.add_paragraph()
    title_run = title_p.add_run("CLASIFICACIÓN DE NO CONFORMIDADES")
    title_run.font.name = "Calibri"
    title_run.font.size = Pt(16)
    title_run.font.bold = True
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    intro_p = doc.add_paragraph(
        "Con el propósito de detallar cada categoría, se exponen a continuación las definiciones "
        "conceptuales de las deficiencias o no conformidades identificadas en el proceso de verificación "
        "eléctrica, las cuales han sido clasificadas de acuerdo con su orden de prioridad."
    )
    intro_p.paragraph_format.space_after = Pt(10)

    prio_table = doc.add_table(rows=4, cols=2)
    prio_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    prio_table.style = 'Table Grid'

    hdr_cells = prio_table.rows[0].cells
    hdr_cells[0].text = "Prioridad"
    hdr_cells[1].text = "Definición"
    set_cell_background(hdr_cells[0], "003366")
    set_cell_background(hdr_cells[1], "003366")
    for cell in hdr_cells:
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)

    data = [
        ("Alta", "Condición que presenta una amenaza directa e inmediata a la vida (choque eléctrico, quemaduras por arco) o a la propiedad (riesgo inminente de incendio). Requiere acción correctiva prioritaria o desenergización"),
        ("Moderada", "No conformidad técnica con la NFPA 70 que afecta el estándar de la instalación, la mantenibilidad, el ciclo de vida o la eficiencia, pero que no presenta un peligro inminente para la seguridad de las personas o bienes."),
        ("Leve", "Incumplimientos documentales, de acceso o de trazabilidad que comprometen o impiden el proceso de inspección, o la verificación legal/técnica de la instalación.")
    ]

    for idx, (prio, desc) in enumerate(data, start=1):
        row_cells = prio_table.rows[idx].cells
        row_cells[0].text = prio
        row_cells[1].text = desc

        if prio == "Alta":
            set_cell_background(row_cells[0], "FFC7CE")
        elif prio == "Moderada":
            set_cell_background(row_cells[0], "FFEB9C")
        elif prio == "Leve":
            set_cell_background(row_cells[0], "C6EFCE")

        row_cells[0].paragraphs[0].runs[0].font.bold = True

    prio_table.columns[0].width = Inches(1.5)
    prio_table.columns[1].width = Inches(5.0)

    note_p = doc.add_paragraph()
    note_p.paragraph_format.space_before = Pt(10)
    note_p.paragraph_format.space_after = Pt(16)
    note_run_bold = note_p.add_run("Nota: ")
    note_run_bold.font.bold = True
    note_p.add_run(
        "Las deficiencias Altas deben tener la prioridad más alta en el plan de acciones correctivas "
        "con una recomendación explícita de plazo de intervención (p. ej., inmediato o corto plazo), "
        "mientras que las Moderadas y Leves se pueden programar a mediano plazo, dentro de los planes de "
        "mantenimiento preventivo o gestión documental del inmueble."
    )

    asp_p = doc.add_paragraph()
    asp_run = asp_p.add_run("ASPECTOS A EVALUAR")
    asp_run.font.name = "Calibri"
    asp_run.font.size = Pt(14)
    asp_run.font.bold = True

    asp_desc = doc.add_paragraph(
        "A continuación, se detalla la lista de verificación que contiene los componentes de la "
        "instalación eléctrica sujetos a examen. Marque las casillas de verificación para incluir "
        "los bloques correspondientes en el documento:"
    )
    asp_desc.paragraph_format.space_after = Pt(10)

    for sec_id, title in sections_info:
        add_checkbox_control(doc, sec_id, title)

    doc.add_paragraph().paragraph_format.space_after = Pt(20)

    anchor_p = doc.add_paragraph()
    anchor_p.paragraph_format.space_before = Pt(10)
    anchor_p.paragraph_format.space_after = Pt(10)
    add_bookmark(anchor_p, 1, "MARCADOR_INSERCION")

    doc.save("template_base.docx")
    print("Base template created successfully as template_base.docx")

if __name__ == "__main__":
    create_template()

from pathlib import Path
from PIL import Image
from docx import Document
from docx.shared import Inches
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx2pdf import convert

EMU_PER_INCH = 914400

def set_table_no_borders(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        elem = OxmlElement(f"w:{edge}")
        elem.set(qn("w:val"), "nil")
        borders.append(elem)
    tblPr.append(borders)

def fit_size_keep_aspect(img_path, max_w_in, max_h_in):
    with Image.open(img_path) as im:
        w_px, h_px = im.size

    ar = w_px / h_px

    w1 = max_w_in
    h1 = w1 / ar
    if h1 <= max_h_in:
        return w1, h1

    h2 = max_h_in
    w2 = h2 * ar
    return w2, h2

def crear_documento(img, output_dir, tipo_salida, numero, margin_in, gap_in):
    doc = Document()

    # Configuración de página
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width = Inches(11)
    sec.page_height = Inches(8.5)

    sec.left_margin = Inches(margin_in)
    sec.right_margin = Inches(margin_in)
    sec.top_margin = Inches(margin_in)
    sec.bottom_margin = Inches(margin_in)

    page_w = sec.page_width / EMU_PER_INCH
    page_h = sec.page_height / EMU_PER_INCH

    avail_w = page_w - 2 * margin_in
    avail_h = page_h - 2 * margin_in

    cell_w = (avail_w - gap_in) / 2
    cell_h = avail_h

    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_no_borders(table)

    for c in range(2):
        cell = table.cell(0, c)
        cell.width = Inches(cell_w)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        w_fit, h_fit = fit_size_keep_aspect(str(img), cell_w, cell_h)
        run = p.add_run()
        run.add_picture(str(img), width=Inches(w_fit), height=Inches(h_fit))

    output_docx = output_dir / f"guias {numero}.docx"
    doc.save(output_docx)
    print(f"DOCX creado: {output_docx}")

    if tipo_salida == 2:
        output_pdf = output_dir / f"guias {numero}.pdf"
        convert(str(output_docx), str(output_pdf))
        print(f"PDF creado: {output_pdf}")

        # eliminar docx si solo quiere pdf
        output_docx.unlink(missing_ok=True)
        print(f"DOCX eliminado: {output_docx}")

def main():
    print("Seleccione el tipo de salida:")
    print("1 -> Word (.docx)")
    print("2 -> PDF")

    while True:
        try:
            tipo = int(input("Ingrese opción (1 o 2): "))
            if tipo in [1, 2]:
                break
            else:
                print("Opción inválida.")
        except:
            print("Ingrese un número válido.")

    images_dir = Path("./")
    output_dir = Path("./salida")
    output_dir.mkdir(parents=True, exist_ok=True)

    valid_ext = {".jpg", ".jpeg", ".png", ".bmp", ".webp", ".tif", ".tiff"}
    imgs = sorted(images_dir.glob("*.*"))

    numero = 1

    for img in imgs:
        if img.suffix.lower() not in valid_ext:
            continue

        crear_documento(
            img=img,
            output_dir=output_dir,
            tipo_salida=tipo,
            numero=numero,
            margin_in=0.3,
            gap_in=0.2
        )
        numero += 1

if __name__ == "__main__":
    main()  
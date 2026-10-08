import zipfile
import xml.etree.ElementTree as ET
from oletools import olevba

def test_docm_structure():
    docm_path = "Lista_de_Chequeo_Electrica.docm"
    assert zipfile.is_zipfile(docm_path), "File is not a valid zip archive"

    with zipfile.ZipFile(docm_path, 'r') as z:
        names = z.namelist()
        assert '[Content_Types].xml' in names
        assert 'word/document.xml' in names
        assert 'word/_rels/document.xml.rels' in names
        assert 'word/vbaProject.bin' in names

        # Check Content Types
        ct_data = z.read('[Content_Types].xml').decode('utf-8')
        assert 'application/vnd.ms-word.document.macroEnabled.main+xml' in ct_data

        # Check Rels
        rels_data = z.read('word/_rels/document.xml.rels').decode('utf-8')
        assert 'vbaProject.bin' in rels_data
        assert 'vbaProject' in rels_data

        # Check Document XML
        doc_xml = z.read('word/document.xml').decode('utf-8')
        assert 'CLASIFICACIÓN DE NO CONFORMIDADES' in doc_xml
        assert 'ASPECTOS A EVALUAR' in doc_xml
        assert 'MARCADOR_INSERCION' in doc_xml

        # Verify 23 Checkbox Controls
        for i in range(1, 24):
            assert f'CHK_{i}' in doc_xml, f'Missing tag CHK_{i} in document.xml'

    print("✓ OpenXML Document structure and Content Controls verified!")

def test_vba_project():
    docm_path = "Lista_de_Chequeo_Electrica.docm"
    vba_parser = olevba.VBA_Parser(docm_path)
    assert vba_parser.detect_vba_macros(), "No VBA macros detected in docm"

    modules = {}
    for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_macros():
        modules[vba_filename] = vba_code

    print("Extracted VBA Modules:", list(modules.keys()))

    has_this_doc = any("ThisDocument" in k for k in modules.keys())
    has_mod_chk = any("modChecklist" in k for k in modules.keys())

    assert has_this_doc, "ThisDocument module missing"
    assert has_mod_chk, "modChecklist module missing"

    this_doc_code = next(v for k, v in modules.items() if "ThisDocument" in k)
    mod_chk_code = next(v for k, v in modules.items() if "modChecklist" in k)

    # Check key methods in ThisDocument
    assert "Document_ContentControlOnExit" in this_doc_code
    assert "ProcesarCambioCasilla" in this_doc_code

    # Check key methods and features in modChecklist
    assert "ProcesarCambioCasilla" in mod_chk_code
    assert "InsertarSeccion" in mod_chk_code
    assert "EliminarSeccion" in mod_chk_code
    assert "ADVERTENCIA: ¿Está seguro de que desea eliminar la sección" in mod_chk_code
    assert "SEC_" in mod_chk_code

    # Verify all 23 section titles are defined in VBA code
    expected_titles = [
        "Documentación",
        "Transformador principal",
        "Sistema de puesta a tierra",
        "Sistema contra descargas atmosféricas",
        "Espacios físicos cuartos eléctricos",
        "Transformadores de baja tensión",
        "Tableros y protecciones",
        "Conductores",
        "Canalizaciones, bandejas, cajas de registro y conexiones",
        "Iluminación",
        "Tomacorrientes y apagadores",
        "Equipos de aire acondicionado, ventilación y refrigeración",
        "Motores",
        "UPS",
        "Sistemas de emergencia",
        "16.1. Lugares Clase I",
        "16.2. Lugares Clase II",
        "16.3. Lugares Clase III",
        "Sitios de Reunión",
        "Teatros y lugares similares",
        "Carnavales, circos, ferias y lugares similares",
        "Piscinas",
        "Instalaciones de asistencia médica"
    ]

    for title in expected_titles:
        assert title in mod_chk_code, f"Title '{title}' missing in modChecklist VBA code"

    print("✓ VBA Project, macros, warning prompts and 23 sections verified!")

if __name__ == "__main__":
    test_docm_structure()
    test_vba_project()
    print("All automated tests passed successfully!")

import os
import zipfile

def make_docm(docx_path: str, vba_bin_path: str, docm_path: str):
    with open(vba_bin_path, "rb") as f:
        vba_data = f.read()

    with zipfile.ZipFile(docx_path, 'r') as in_zip, zipfile.ZipFile(docm_path, 'w', zipfile.ZIP_DEFLATED) as out_zip:
        for item in in_zip.infolist():
            data = in_zip.read(item.filename)

            if item.filename == '[Content_Types].xml':
                content_str = data.decode('utf-8')
                content_str = content_str.replace(
                    'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
                    'application/vnd.ms-word.document.macroEnabled.main+xml'
                )
                out_zip.writestr(item.filename, content_str.encode('utf-8'))

            elif item.filename == 'word/_rels/document.xml.rels':
                content_str = data.decode('utf-8')
                vba_rel = '<Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>'
                content_str = content_str.replace('</Relationships>', f'{vba_rel}</Relationships>')
                out_zip.writestr(item.filename, content_str.encode('utf-8'))

            else:
                out_zip.writestr(item, data)

        out_zip.writestr('word/vbaProject.bin', vba_data)

if __name__ == "__main__":
    make_docm("template_base.docx", "vbaProject.bin", "Lista_de_Chequeo_Electrica.docm")
    print("Generated Lista_de_Chequeo_Electrica.docm successfully!")

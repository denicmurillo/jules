import struct
import olefile
import re
from ole_writer import OLEWriter

def compress_vba_source(data_input) -> bytes:
    """
    Compresses source bytes/text using MS-OVBA LZNT1 Compression Algorithm (Literal-tokens compressed chunks).
    """
    if isinstance(data_input, str):
        cleaned_lines = []
        for line in data_input.splitlines():
            if line.startswith("VERSION ") or line.startswith("BEGIN") or line.startswith("END") or line.startswith("Attribute "):
                continue
            cleaned_lines.append(line)
        cleaned_text = "\r\n".join(cleaned_lines).lstrip()
        data = cleaned_text.encode('cp1252', errors='replace')
    else:
        data = bytes(data_input)

    compressed = bytearray()
    compressed.append(0x01) # Header signature byte

    pos = 0
    length = len(data)

    while pos < length:
        chunk_raw = data[pos:pos + 3600]
        pos += len(chunk_raw)

        payload = bytearray()
        i = 0
        raw_len = len(chunk_raw)

        while i < raw_len:
            group = chunk_raw[i:i + 8]
            i += len(group)
            payload.append(0x00) # Flag byte
            payload.extend(group)

        payload_len = len(payload)
        header = 0xB000 | (payload_len - 1)

        compressed.extend(struct.pack('<H', header))
        compressed.extend(payload)

    return bytes(compressed)

def build_vba_project_bin(this_doc_src: str, mod_checklist_src: str) -> bytes:
    vba_project_stream = bytes([0xCC, 0x61, 0x88, 0x00, 0x00, 0x00, 0x00])

    project_stream = (
        'ID="{00000000-0000-0000-0000-000000000000}"\r\n'
        'Document=ThisDocument/&H00000000\r\n'
        'Module=modChecklist\r\n'
        'Name="Project"\r\n'
        'HelpContextID="0"\r\n'
        'VersionCompatible32="393222000"\r\n'
        'CMG="000204"\r\n'
        'DPB="000204"\r\n'
        'GC="000204"\r\n'
        '\r\n'
        '[Host Extender Info]\r\n'
        '&H00000001={3832D640-042B-11D2-B826-006097C4E476};GWorkspace;Control\r\n'
        '\r\n'
        '[Workspace]\r\n'
        'ThisDocument=0, 0, 0, 0, C\r\n'
        'modChecklist=0, 0, 0, 0, C\r\n'
    ).encode('cp1252')

    projectwm_stream = bytearray()
    for name in ["ThisDocument", "modChecklist"]:
        projectwm_stream.extend(name.encode('cp1252') + b'\x00')
        projectwm_stream.extend(name.encode('utf-16le') + b'\x00\x00')
    projectwm_stream.extend(b'\x00\x00')

    this_doc_compressed = compress_vba_source(this_doc_src)
    mod_checklist_compressed = compress_vba_source(mod_checklist_src)

    dir_buf = bytearray()

    def add_rec(tag, data=b''):
        dir_buf.extend(struct.pack('<H', tag))
        dir_buf.extend(struct.pack('<I', len(data)))
        dir_buf.extend(data)

    add_rec(0x0001, struct.pack('<I', 1)) # SYSKIND = Win32
    add_rec(0x0002, struct.pack('<I', 0x0409)) # LCID
    add_rec(0x0014, struct.pack('<I', 0x0409)) # LCIDINVOKE
    add_rec(0x0003, struct.pack('<H', 1252)) # CODEPAGE
    add_rec(0x0004, b'Project') # NAME

    # DOCSTRING & DOCSTRINGUNICODE
    add_rec(0x0005, b'')
    dir_buf.extend(struct.pack('<H', 0x0040))
    dir_buf.extend(struct.pack('<I', 0))

    # HELPFILEPATH & HELPFILEPATHUNICODE
    add_rec(0x0006, b'')
    dir_buf.extend(struct.pack('<H', 0x003D))
    dir_buf.extend(struct.pack('<I', 0))
    dir_buf.extend(b'')

    add_rec(0x0007, struct.pack('<I', 0)) # HELPCONTEXT
    add_rec(0x0008, struct.pack('<I', 0)) # LIBFLAGS
    add_rec(0x0009, struct.pack('<I', 9) + struct.pack('<H', 1)) # VERSION

    # CONSTANTS & CONSTANTSUNICODE
    add_rec(0x000C, b'')
    dir_buf.extend(struct.pack('<H', 0x003C))
    dir_buf.extend(struct.pack('<I', 0))

    # References
    def add_reference_registered(ref_name, libid_str):
        if ref_name:
            name_bytes = ref_name.encode('cp1252')
            name_utf16 = ref_name.encode('utf-16le')
            add_rec(0x0016, name_bytes)
            dir_buf.extend(struct.pack('<H', 0x003E))
            dir_buf.extend(struct.pack('<I', len(name_utf16)))
            dir_buf.extend(name_utf16)

        libid_bytes = libid_str.encode('cp1252')
        rec_data = bytearray()
        rec_data.extend(struct.pack('<I', len(libid_bytes)))
        rec_data.extend(libid_bytes)
        rec_data.extend(struct.pack('<I', 0)) # reserved1
        rec_data.extend(struct.pack('<H', 0)) # reserved2
        add_rec(0x000D, rec_data)

    stdole_libid = r"*\G{00020430-0000-0000-C000-000000000468}#2.0#0#C:\Windows\SysWOW64\stdole2.tlb#OLE Automation"
    add_reference_registered("stdole", stdole_libid)

    word_libid = r"*\G{00020905-0000-0000-C000-000000000468}#8.7#0#C:\Program Files\Microsoft Office\root\Office16\MSWORD.OLB#Microsoft Word 16.0 Object Library"
    add_reference_registered("Word", word_libid)

    add_rec(0x000F, struct.pack('<H', 2)) # Module Count = 2
    add_rec(0x0013, struct.pack('<H', 0xFFFF)) # PROJECTCOOKIE record

    def add_module_record(mod_name, mod_type_tag, cookie_val):
        mod_bytes = mod_name.encode('cp1252')
        mod_utf16 = mod_name.encode('utf-16le')

        # 1. MODULENAME (0x0019)
        add_rec(0x0019, mod_bytes)

        # 2. MODULENAMEUNICODE (0x0047)
        dir_buf.extend(struct.pack('<H', 0x0047))
        dir_buf.extend(struct.pack('<I', len(mod_utf16)))
        dir_buf.extend(mod_utf16)

        # 3. MODULESTREAMNAME (0x001A) & MODULESTREAMNAMEUNICODE (0x0032)
        stream_rec = bytearray()
        stream_rec.extend(mod_bytes)
        stream_rec.extend(struct.pack('<H', 0x0032))
        stream_rec.extend(struct.pack('<I', len(mod_utf16)))
        stream_rec.extend(mod_utf16)
        add_rec(0x001A, stream_rec)

        # 4. MODULEDOCSTRING (0x001C) & MODULEDOCSTRINGUNICODE (0x0048)
        doc_rec = bytearray()
        doc_rec.extend(b'')
        doc_rec.extend(struct.pack('<H', 0x0048))
        doc_rec.extend(struct.pack('<I', 0))
        add_rec(0x001C, doc_rec)

        # 5. MODULEOFFSET (0x0031)
        add_rec(0x0031, struct.pack('<I', 0))

        # 6. MODULEHELPCONTEXT (0x001E)
        add_rec(0x001E, struct.pack('<I', 0))

        # 7. MODULECOOKIE (0x002C)
        add_rec(0x002C, struct.pack('<H', cookie_val))

        # 8. MODULETYPE
        add_rec(mod_type_tag)

        # 9. MODULETERMINATOR
        add_rec(0x002B)

    add_module_record('ThisDocument', 0x0022, 0)
    add_module_record('modChecklist', 0x0021, 1)

    add_rec(0x0010) # PROJECTTERMINATOR

    dir_compressed = compress_vba_source(dir_buf)

    writer = OLEWriter()
    writer.add_storage("VBA")
    writer.add_stream("PROJECT", project_stream)
    writer.add_stream("PROJECTwm", bytes(projectwm_stream))
    writer.add_stream("VBA/_VBA_PROJECT", vba_project_stream)
    writer.add_stream("VBA/dir", dir_compressed)
    writer.add_stream("VBA/ThisDocument", this_doc_compressed)
    writer.add_stream("VBA/modChecklist", mod_checklist_compressed)

    return writer.build()

if __name__ == "__main__":
    with open("src/ThisDocument.cls", "r", encoding="utf-8") as f:
        this_doc = f.read()
    with open("src/modChecklist.bas", "r", encoding="utf-8") as f:
        mod_chk = f.read()

    vba_bin = build_vba_project_bin(this_doc, mod_chk)
    with open("vbaProject.bin", "wb") as f:
        f.write(vba_bin)
    print(f"Generated vbaProject.bin ({len(vba_bin)} bytes)")

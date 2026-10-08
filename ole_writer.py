import struct
import io

class OLEWriter:
    SECTOR_SIZE = 512
    MINI_SECTOR_SIZE = 64
    MINI_STREAM_CUTOFF = 4096

    def __init__(self):
        self.entries = [] # list of (path, data, type)

    def add_stream(self, path: str, data: bytes):
        self.entries.append((path, data, 2))

    def add_storage(self, path: str):
        self.entries.append((path, b'', 1))

    def build(self) -> bytes:
        dir_entries = []

        # 0: Root Entry
        dir_entries.append({
            'name': 'Root Entry',
            'type': 5,
            'data': b'',
            'child': -1,
            'left': -1,
            'right': -1
        })

        nodes = {}
        nodes[''] = 0 # Root

        for path, data, etype in self.entries:
            parts = path.split('/')
            curr_path = ''
            for idx, part in enumerate(parts):
                parent_path = curr_path
                curr_path = parent_path + ('/' if parent_path else '') + part
                if curr_path not in nodes:
                    is_last = (idx == len(parts) - 1)
                    node_type = etype if is_last else 1
                    node_data = data if (is_last and etype == 2) else b''
                    node_idx = len(dir_entries)
                    dir_entries.append({
                        'name': part,
                        'type': node_type,
                        'data': node_data,
                        'child': -1,
                        'left': -1,
                        'right': -1
                    })
                    nodes[curr_path] = node_idx

                    parent_idx = nodes[parent_path]
                    if dir_entries[parent_idx]['child'] == -1:
                        dir_entries[parent_idx]['child'] = node_idx
                    else:
                        sibling = dir_entries[parent_idx]['child']
                        while dir_entries[sibling]['right'] != -1:
                            sibling = dir_entries[sibling]['right']
                        dir_entries[sibling]['right'] = node_idx

        sectors = [] # list of 512-byte sector data
        fat = [] # list of sector pointers

        def allocate_fat_sectors(data_bytes):
            if not data_bytes:
                return 0xFFFFFFFE, 0
            start_sec = len(sectors)
            num_secs = (len(data_bytes) + 511) // 512
            for i in range(num_secs):
                chunk = data_bytes[i*512:(i+1)*512]
                chunk = chunk.ljust(512, b'\x00')
                sectors.append(chunk)
                if i < num_secs - 1:
                    fat.append(len(sectors))
                else:
                    fat.append(0xFFFFFFFE) # ENDOFCHAIN
            return start_sec, len(data_bytes)

        # Separate regular streams (>= 4096) and mini streams (< 4096)
        ministream_bytes = bytearray()
        minifat = []

        def allocate_mini_sectors(data_bytes):
            if not data_bytes:
                return 0xFFFFFFFE, 0
            start_msec = len(ministream_bytes) // 64
            num_msecs = (len(data_bytes) + 63) // 64
            for i in range(num_msecs):
                chunk = data_bytes[i*64:(i+1)*64]
                chunk = chunk.ljust(64, b'\x00')
                ministream_bytes.extend(chunk)
                curr_msec = start_msec + i
                if i < num_msecs - 1:
                    minifat.append(curr_msec + 1)
                else:
                    minifat.append(0xFFFFFFFE)
            return start_msec, len(data_bytes)

        for entry in dir_entries:
            if entry['type'] == 2:
                d = entry['data']
                if len(d) >= self.MINI_STREAM_CUTOFF:
                    start_sec, size = allocate_fat_sectors(d)
                    entry['start_sec'] = start_sec
                    entry['size'] = size
                else:
                    start_msec, size = allocate_mini_sectors(d)
                    entry['start_sec'] = start_msec
                    entry['size'] = size
            else:
                entry['start_sec'] = 0xFFFFFFFE
                entry['size'] = 0

        # Allocate MiniStream in FAT
        if ministream_bytes:
            root_start_sec, root_size = allocate_fat_sectors(bytes(ministream_bytes))
            dir_entries[0]['start_sec'] = root_start_sec
            dir_entries[0]['size'] = root_size
        else:
            dir_entries[0]['start_sec'] = 0xFFFFFFFE
            dir_entries[0]['size'] = 0

        # Allocate MiniFAT in FAT
        if minifat:
            minifat_buf = bytearray()
            for val in minifat:
                minifat_buf.extend(struct.pack('<I', val))
            minifat_start_sec, minifat_size = allocate_fat_sectors(bytes(minifat_buf))
            num_minifat_secs = (len(minifat_buf) + 511) // 512
        else:
            minifat_start_sec = 0xFFFFFFFE
            num_minifat_secs = 0

        # Serialize Directory Sector(s)
        dir_bytes = bytearray()
        for entry in dir_entries:
            e_buf = bytearray(128)
            name_utf16 = entry['name'].encode('utf-16le') + b'\x00\x00'
            e_buf[0:len(name_utf16)] = name_utf16
            struct.pack_into('<H', e_buf, 64, len(name_utf16))
            e_buf[66] = entry['type']
            e_buf[67] = 0x01 # Black node color
            struct.pack_into('<i', e_buf, 68, entry['left'])
            struct.pack_into('<i', e_buf, 72, entry['right'])
            struct.pack_into('<i', e_buf, 76, entry['child'])
            struct.pack_into('<I', e_buf, 116, entry['start_sec'])
            struct.pack_into('<I', e_buf, 120, entry['size'])
            struct.pack_into('<I', e_buf, 124, 0)
            dir_bytes.extend(e_buf)

        dir_start_sec, dir_size = allocate_fat_sectors(bytes(dir_bytes))

        # FAT sectors
        num_fat_secs = 1
        while True:
            total_secs = len(sectors) + num_fat_secs
            needed_fat_entries = total_secs
            if num_fat_secs * 128 >= needed_fat_entries:
                break
            num_fat_secs += 1

        fat_start_sec = len(sectors)
        for i in range(num_fat_secs):
            sectors.append(b'\x00' * 512)
            fat.append(0xFFFFFFFD) # FATSECT marker

        fat_bytes = bytearray()
        for entry_val in fat:
            fat_bytes.extend(struct.pack('<I', entry_val))
        while len(fat_bytes) < num_fat_secs * 512:
            fat_bytes.extend(struct.pack('<I', 0xFFFFFFFF))

        for i in range(num_fat_secs):
            sectors[fat_start_sec + i] = fat_bytes[i*512:(i+1)*512]

        # Header Sector (512 bytes)
        hdr = bytearray(512)
        hdr[0:8] = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
        struct.pack_into('<H', hdr, 24, 0x003E)
        struct.pack_into('<H', hdr, 26, 0x0003)
        struct.pack_into('<H', hdr, 28, 0xFFFE)
        struct.pack_into('<H', hdr, 30, 9) # Sector shift (512)
        struct.pack_into('<H', hdr, 32, 6) # Mini sector shift (64)
        struct.pack_into('<I', hdr, 44, num_fat_secs)
        struct.pack_into('<I', hdr, 48, dir_start_sec)
        struct.pack_into('<I', hdr, 56, 4096) # Mini Stream Cutoff
        struct.pack_into('<I', hdr, 60, minifat_start_sec)
        struct.pack_into('<I', hdr, 64, num_minifat_secs)
        struct.pack_into('<I', hdr, 68, 0xFFFFFFFE)
        struct.pack_into('<I', hdr, 72, 0)

        for i in range(109):
            val = fat_start_sec + i if i < num_fat_secs else 0xFFFFFFFF
            struct.pack_into('<I', hdr, 76 + i*4, val)

        out = bytearray()
        out.extend(hdr)
        for sec in sectors:
            out.extend(sec)

        return bytes(out)

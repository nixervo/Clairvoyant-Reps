import os, struct

with open(os.path.join(os.path.dirname(__file__), 'members.bin'), 'rb') as f:
    data = f.read()

def decode_u29(b, off):
    val = b[off]
    if not (val & 0x80):
        return val, 1
    val = (val & 0x7f) << 7
    val |= (b[off + 1] & 0x7f)
    if not (b[off + 1] & 0x80):
        return val, 2
    val = (val & 0x3fff) << 7
    val |= (b[off + 2] & 0x7f)
    if not (b[off + 2] & 0x80):
        return val, 3
    val = (val & 0x1fffff) << 8
    val |= b[off + 3]
    return val, 4

# Read 4-byte member_id anywhere after "member_id\x04" marker
results = []
i = 0
while i < len(data) - 20:
    marker = data.find(b'member_id\x04', i)
    if marker < 0:
        break
    id_val, consumed = decode_u29(data, marker + 10)
    # Read member_name from the same object
    # After member_id, scan for member_name marker
    name_start = data.find(b'member_name\x06', marker, marker + 80)
    if name_start >= 0:
        sl, sc = decode_u29(data, name_start + 12)
        if sl & 1:
            length = sl >> 1
            name = data[name_start + 12 + sc:name_start + 12 + sc + length].decode('utf-8', errors='replace')
            results.append((id_val, name))
    i = marker + 1

print("%9s  %s" % ("ID", "Name"))
print("-" * 42)
for id_val, name in sorted(results):
    print("%9d  %s" % (id_val, name))

import os

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

# Find the first "member_name" string to locate the trait definition
first_name = data.find(b'member_name')
print("First 'member_name' at:", first_name)

# After the trait is defined, objects are encoded as:
# 0x0A (object) + reference_byte + values...
# The trait has sequential fields: member_id, member_name, member_level, etc.
# objects then have: int, string, int, int, int, int, int

# Find all encoded strings (0x06) near member positions to get names
# and find the preceding 0x04 (integer) for member_id
# Strategy: scan for 0x0A 0x0B (object start + closed) and extract

# Actually simpler: find all name strings and read their preceding member_id
all_names = []
for i, b in enumerate(data):
    if b == 0x06:  # string marker
        sl, sc = decode_u29(data, i + 1)
        if sl & 1:  # new string
            length = sl >> 1
            name = data[i + 1 + sc:i + 1 + sc + length].decode('utf-8', errors='replace')
            if len(name) > 2 and len(name) < 45 and name.isprintable():
                all_names.append((i, name))

# Now find member_id for each name - scan backwards for 0x04
seen = set()
print("\n%9s  %s" % ("ID", "Name"))
print("-" * 42)
for pos, name in sorted(all_names):
    # Find 0x04 (integer marker) before this string
    for back in range(pos - 1, max(pos - 60, 0), -1):
        if data[back] == 0x04:
            id_val, _ = decode_u29(data, back + 1)
            if 100 < id_val < 10000000 and id_val not in seen:
                seen.add(id_val)
                print("%9d  %s" % (id_val, name))
                break

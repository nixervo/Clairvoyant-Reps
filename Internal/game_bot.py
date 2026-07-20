"""
Ninja Rift Game Bot — AMF Remoting Client
==========================================
Spoofs the Adobe AIR game client by sending AMF (Action Message Format)
binary requests to the game server at ninjarift.org/amf_nl/

How the game communicates:
  1. Adobe AIR client sends HTTP POST to https://ninjarift.org/amf_nl/
  2. Body is AMF binary — a compact serialization format
  3. No authentication (no cookies, no tokens, no headers)
  4. Server returns AMF binary response with chunked transfer encoding
  5. The game uses a pattern: ServiceName.executeService with
     the actual method name (like "getClansForRequest") embedded
     inside AMF3-encoded parameter objects

What this script does:
  1. Builds AMF3 requests just like the game client does
  2. Sends them to the server
  3. Decodes the AMF response back to Python dicts/lists
  4. Pretty-prints the data

Usage:
  python game_bot.py                       # Show menu
  python game_bot.py clans                 # All clan rankings (2186 clans)
  python game_bot.py crews                 # All crew rankings
  python game_bot.py clan                  # Player's clan details
  python game_bot.py clan_members          # Clan member list + stats
  python game_bot.py clan_history          # Clan war/season history
  python game_bot.py clans_battle          # Clans available for war
  python game_bot.py clan_members_battle   # Clan battle member stats
  python game_bot.py stamina               # Character stamina
  python game_bot.py crew                  # Player's crew details
  python game_bot.py crew_members_battle   # Crew boss/battle stats
  python game_bot.py castles               # All castle ownership
  python game_bot.py castle <0-6>          # Specific castle details
  python game_bot.py daily                 # Daily tasks + rewards
  python game_bot.py skills                # Character skills
  python game_bot.py clan_status           # Clan season info
  python game_bot.py crew_status           # Crew season/phase info

Dependencies:
  pip install pyamf
"""

import struct
import sys
import json
import urllib.request
from collections import OrderedDict

# Fix PyAMF Python 2→3 compatibility (missing StringIO module)
# Only needed when cpyamf C extension is unavailable (e.g. GitHub Actions Linux)
import io
try:
    import cpyamf  # noqa: F401
except ImportError:
    sys.modules["StringIO"] = io
    sys.modules["cStringIO"] = io


# ─── Configuration ───────────────────────────────────────────────────────────
AMF_URL = "https://ninjarift.org/amf_nl/"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows; U; en) AppleWebKit/533.19.4 "
                  "(KHTML, like Gecko) AdobeAIR/51.2",
    "Content-Type": "application/x-amf",
    "Referer": "app:/NinjaRift.swf",
    "x-flash-version": "51,2,1,2",
    "Accept": "*/*",
    "Accept-Encoding": "gzip, deflate",
    "Connection": "Keep-Alive",
}


# ─── AMF3 Codec (manual encoder/decoder) ─────────────────────────────────────
#
# AMF3 is Adobe's compact binary format. Key types:
#   0x04 — Integer         0x05 — Double          0x06 — String
#   0x09 — Dynamic Object  0x0A — Typed Object    0x0A — Array
#   0x0C — ByteArray       0x01 — null            0x02 — false
#   0x03 — true
#
# Integers use variable-length U29 encoding (1-4 bytes).
# Strings use U29 header + UTF-8 data, with a reference table for reuse.

def decode_u29(data, offset):
    """Read a U29 (Unsigned 29-bit integer) from AMF3 stream.
    Each byte: 7 data bits + 1 continuation bit (MSB=1 means more bytes)."""
    val = data[offset]
    if val < 128:
        return val, offset + 1
    result = val & 0x7F
    offset += 1
    val = data[offset]
    result = (result << 7) | (val & 0x7F)
    if val < 128:
        return result, offset + 1
    offset += 1
    val = data[offset]
    result = (result << 7) | (val & 0x7F)
    if val < 128:
        return result, offset + 1
    offset += 1
    val = data[offset]
    result = (result << 8) | val
    return result, offset + 1


def encode_u29(value):
    """Write a U29 integer. Returns bytes."""
    if value < 0x80:
        return bytes([value])
    if value < 0x4000:
        return bytes([(value >> 7) | 0x80, value & 0x7F])
    if value < 0x200000:
        return bytes([(value >> 14) | 0x80,
                      ((value >> 7) & 0x7F) | 0x80,
                      value & 0x7F])
    return bytes([(value >> 22) | 0x80,
                  ((value >> 15) & 0x7F) | 0x80,
                  ((value >> 8) & 0x7F) | 0x80,
                  value & 0xFF])


# ─── Low-Level AMF3 Decoder ──────────────────────────────────────────────────
# For reference: PyAMF handles this automatically. This manual decoder
# exists so you can see exactly how the binary wire format works.

def _amf3_decode_string(data, pos, refs):
    """Decode an AMF3 string at position pos. Returns (value, new_pos)."""
    ref, pos = decode_u29(data, pos)
    if ref & 1:  # Reference to previous string
        idx = ref >> 1
        return refs[idx], pos
    length = ref >> 1
    if length == 0:
        refs.append("")
        return "", pos
    val = data[pos:pos + length].decode("utf-8", errors="replace")
    pos += length
    refs.append(val)
    return val, pos


def _amf3_decode_value(data, pos, refs):
    """Decode a single AMF3 value. Returns (value, new_pos)."""
    if pos >= len(data):
        return None, pos
    marker = data[pos]
    pos += 1

    if marker == 0x01:  # null
        return None, pos
    elif marker == 0x02:  # false
        return False, pos
    elif marker == 0x03:  # true
        return True, pos
    elif marker == 0x04:  # Integer (U29)
        val, pos = decode_u29(data, pos)
        return val, pos
    elif marker == 0x05:  # Double (8 bytes)
        val = struct.unpack(">d", data[pos:pos + 8])[0]
        return val, pos + 8
    elif marker == 0x06:  # String
        return _amf3_decode_string(data, pos, refs)
    elif marker == 0x09:  # Dynamic object
        return _amf3_decode_object(data, pos, refs, marker)
    elif marker == 0x0A:  # Array
        return _amf3_decode_array(data, pos, refs)
    elif marker == 0x0C:  # ByteArray
        ref, pos = decode_u29(data, pos)
        if ref & 1:
            return None, pos  # reference (skip)
        length = ref >> 1
        val = data[pos:pos + length]
        return val, pos + length
    else:
        # Unknown type — skip and return raw marker
        return f"<AMF3:0x{marker:02X}>", pos


def _amf3_decode_object(data, pos, refs, marker):
    """Decode an AMF3 object (0x09 = dynamic, 0x0A = typed)."""
    # Read traits (class definition)
    traits, pos = decode_u29(data, pos)
    is_dynamic = bool(traits & 1)
    traits >>= 1
    is_external = bool(traits & 1)
    traits >>= 1
    is_trait_ref = bool(traits & 1)
    count = traits >> 1

    if is_trait_ref:
        # Referenced class — we don't track class defs, return stub
        result = {"_class_ref": count}
    else:
        # Read class name
        class_name, pos = _amf3_decode_string(data, pos, refs)
        # Read sealed trait names
        trait_names = []
        for _ in range(count):
            name, pos = _amf3_decode_string(data, pos, refs)
            trait_names.append(name)
        # Read sealed values
        result = OrderedDict()
        for name in trait_names:
            val, pos = _amf3_decode_value(data, pos, refs)
            result[name] = val

    # Read dynamic members
    if is_dynamic:
        while True:
            if pos >= len(data):
                break
            name, pos = _amf3_decode_string(data, pos, refs)
            if name == "":
                break  # End of dynamic members
            val, pos = _amf3_decode_value(data, pos, refs)
            if isinstance(result, OrderedDict):
                result[name] = val

    return dict(result), pos


def _amf3_decode_array(data, pos, refs):
    """Decode an AMF3 array (0x0A)."""
    ref, pos = decode_u29(data, pos)
    if ref & 1:  # Reference
        return [], pos  # skip references
    count = ref >> 1
    # Read dense portion key (usually empty string = no associative keys)
    key, pos = _amf3_decode_string(data, pos, refs)
    items = []
    for _ in range(count):
        val, pos = _amf3_decode_value(data, pos, refs)
        items.append(val)
    return items, pos


def decode_amf3(data):
    """Decode an AMF3-encoded byte stream into Python objects."""
    refs = []
    pos = 0
    result = []
    while pos < len(data):
        val, pos = _amf3_decode_value(data, pos, refs)
        result.append(val)
    return result


# ─── AMF Remoting Request Builder ────────────────────────────────────────────
#
# AMF Remoting wraps AMF0/AMF3 data in a remote procedure call envelope:
#   [version:2] [header_count:2] [headers...] [body_count:2] [bodies...]
#
# Each body:
#   [target_uri_length:2] [target_uri:string]
#   [response_uri_length:2] [response_uri:string]
#   [data_length:4] [AMF0_value]

def build_amf_request(service, method, args):
    """
    Build an AMF Remoting request exactly matching the game client binary format.

    Uses PyAMF for the AMF3 body encoding (verified byte-identical to the game)
    and manually builds the AMF Remoting envelope.
    """
    import io
    from pyamf.amf3 import Encoder

    # ── Step 1: Encode the AMF3 body ─────────────────────────────────────
    # Game format: [method_name, [arg1, arg2, ...]]
    payload = [method, args]
    buf = io.BytesIO()
    enc = Encoder(buf)
    enc.writeElement(payload)
    enc.stream.flush()
    amf3_body = buf.getvalue()

    # ── Step 2: Wrap in AMF0 array with AMF3 switch ──────────────────────
    amf0_body = bytearray()
    amf0_body.append(0x0A)                   # AMF0 strict array
    amf0_body.extend(struct.pack(">I", 1))   # 1 element
    amf0_body.append(0x11)                   # AMF3 switch
    amf0_body.extend(amf3_body)

    # ── Step 3: Build AMF Remoting envelope ──────────────────────────────
    target_bytes = service.encode("utf-8")
    response_bytes = b"/1"

    body = bytearray()
    body.extend(struct.pack(">H", len(target_bytes)))
    body.extend(target_bytes)
    body.extend(struct.pack(">H", len(response_bytes)))
    body.extend(response_bytes)
    body.extend(struct.pack(">I", len(amf0_body)))
    body.extend(amf0_body)

    envelope = bytearray()
    envelope.extend(struct.pack(">H", 3))    # AMF3 version
    envelope.extend(struct.pack(">H", 0))    # 0 headers
    envelope.extend(struct.pack(">H", 1))    # 1 body
    envelope.extend(body)

    result = bytes(envelope)
    return result


def build_amf_request_direct(service, args):
    """
    Build AMF request WITHOUT the [method, args] wrapping.
    Used for SystemLogin.* calls where target URI IS the method.

    Payload format: args (flat list), not [args] (nested).
    PyAMF handles the encoding — verified byte-identical to the game client.
    """
    import io
    from pyamf.amf3 import Encoder

    # Encode args as flat AMF3 list (matching game client format)
    buf = io.BytesIO()
    enc = Encoder(buf)
    enc.writeElement(args)  # NOT [args] — flat list for SystemLogin calls
    enc.stream.flush()
    amf3_body = buf.getvalue()

    # ── Step 2: AMF0 wrapper ──────────────────────────────────────────────
    amf0_body = bytearray()
    amf0_body.append(0x0A)                  # AMF0 strict array
    amf0_body.extend(struct.pack(">I", 1))  # 1 element
    amf0_body.append(0x11)                  # AMF3 switch
    amf0_body.extend(amf3_body)

    # ── Step 3: AMF Remoting envelope ─────────────────────────────────────
    target_bytes = service.encode("utf-8")
    response_bytes = b"/1"

    body = bytearray()
    body.extend(struct.pack(">H", len(target_bytes)))
    body.extend(target_bytes)
    body.extend(struct.pack(">H", len(response_bytes)))
    body.extend(response_bytes)
    body.extend(struct.pack(">I", len(amf0_body)))
    body.extend(amf0_body)

    envelope = bytearray()
    envelope.extend(struct.pack(">H", 3))
    envelope.extend(struct.pack(">H", 0))
    envelope.extend(struct.pack(">H", 1))
    envelope.extend(body)

    return bytes(envelope)


def _encode_u29_direct(value):
    """Encode int as AMF3 U29 — handles multi-byte (unlike _encode_u29)."""
    if value < 0x80:
        return bytes([value])
    if value < 0x4000:
        return bytes([(value >> 7) | 0x80, value & 0x7F])
    if value < 0x200000:
        return bytes([(value >> 14) | 0x80,
                      ((value >> 7) & 0x7F) | 0x80,
                      value & 0x7F])
    return bytes([(value >> 22) | 0x80,
                  ((value >> 15) & 0x7F) | 0x80,
                  ((value >> 8) & 0x7F) | 0x80,
                  value & 0xFF])


def send_amf_direct(service, args=None):
    """Send AMF request without method wrapping. For SystemLogin.* calls."""
    if args is None:
        args = []
    request_body = build_amf_request_direct(service, args)

    req = urllib.request.Request(
        AMF_URL,
        data=request_body,
        headers=HEADERS,
        method="POST",
    )

    print(f"[*] Sending {service} ({len(request_body)} bytes)")

    with urllib.request.urlopen(req, timeout=15) as resp:
        raw = resp.read()
        print(f"    Response: {len(raw)} bytes, status={resp.status}")
        return raw


# ─── HTTP Transport ──────────────────────────────────────────────────────────

def send_amf(service, method, args=None):
    """Send an AMF request to the game server and return the raw response body."""
    if args is None:
        args = []
    request_body = build_amf_request(service, method, args)

    req = urllib.request.Request(
        AMF_URL,
        data=request_body,
        headers=HEADERS,
        method="POST",
    )

    print(f"[*] Sending {service} -> {method} ({len(request_body)} bytes)")

    with urllib.request.urlopen(req, timeout=15) as resp:
        raw = resp.read()
        print(f"    Response: {len(raw)} bytes, status={resp.status}")
        return raw


def strip_chunked_encoding(raw):
    """
    Strip HTTP chunked transfer encoding headers from raw AMF response.
    The server returns chunked data like: <hex_size>\r\n<data>\r\n0\r\n\r\n
    """
    pos = raw.find(b"\x00\x11")
    if pos < 0:
        pos = raw.find(b"\x00\x03")
    if pos >= 0:
        return raw[pos:]
    return raw


# ─── Castle Names (fixed order from the game) ────────────────────────────────
CASTLE_NAMES = [
    "Hiroshima", "Himeji", "Kumamoto", "Okazaki",
    "Inuyama", "Gifu", "Hikone",
]
# The game sends a character ID + device token with every request.
# From the Fiddler captures (2 players): char_id=532542 (exigency0527),
# char_id=591659 (clancy2807). Token changes per login session.
# No authentication required — these are just context identifiers.

DEFAULT_CHAR_ID = 532542
DEFAULT_TOKEN = "edbee744f71425e8ff231cb2c186a16a"


# ─── High-Level API Calls ────────────────────────────────────────────────────

# -- Clan Ranking & Info ------------------------------------------------------

def get_all_clans(char_id=DEFAULT_CHAR_ID, token=DEFAULT_TOKEN):
    """Fetch the full clan ranking list."""
    return send_amf("ClanService.executeService", "getClansForRequest",
                    [char_id, token])

def get_clan(char_id=DEFAULT_CHAR_ID, token=DEFAULT_TOKEN):
    """Fetch details for the player's own clan."""
    return send_amf("ClanService.executeService", "getClan",
                    [char_id, token])

def get_clan_status(char_id=DEFAULT_CHAR_ID, token=DEFAULT_TOKEN):
    """Fetch clan season info."""
    return send_amf("ClanService.executeService", "getClanStatus",
                    [char_id, token])

def get_clan_members(char_id=DEFAULT_CHAR_ID, token=DEFAULT_TOKEN):
    """Fetch clan member list with reputation, level, join date."""
    return send_amf("ClanService.executeService", "getMembersInfo",
                    [char_id, token])

def get_clan_history(char_id=DEFAULT_CHAR_ID, token=DEFAULT_TOKEN):
    """Fetch clan war/season history."""
    return send_amf("ClanService.executeService", "getClanHistory",
                    [char_id, token])

def get_clans_for_battle(char_id=DEFAULT_CHAR_ID, token=DEFAULT_TOKEN):
    """Fetch clans available for clan wars (the battle ranking list)."""
    return send_amf("ClanService.executeService", "getClansForBattle",
                    [char_id, token])

def get_clan_members_battle(char_id=DEFAULT_CHAR_ID, token=DEFAULT_TOKEN):
    """Fetch clan member battle stats for war."""
    return send_amf("ClanService.executeService", "getMembersInfoForBattle",
                    [char_id, token])

def get_stamina(char_id=DEFAULT_CHAR_ID, token=DEFAULT_TOKEN):
    """Fetch character stamina info."""
    return send_amf("ClanService.executeService", "getStamina",
                    [char_id, token])


# -- Crew Ranking & Info ------------------------------------------------------

def get_all_crews(char_id=DEFAULT_CHAR_ID, token=DEFAULT_TOKEN):
    """Fetch the full crew ranking list."""
    return send_amf("CrewService.executeService", "getCrewsForRequest",
                    [char_id, token])

def get_crew(char_id=DEFAULT_CHAR_ID, token=DEFAULT_TOKEN):
    """Fetch details for the player's own crew."""
    return send_amf("CrewService.executeService", "getCrew",
                    [char_id, token])

def get_crew_status(char_id=DEFAULT_CHAR_ID, token=DEFAULT_TOKEN):
    """Fetch crew season/phase info."""
    return send_amf("CrewService.executeService", "getCrewStatus",
                    [char_id, token])

def get_castles_info(char_id=DEFAULT_CHAR_ID, token=DEFAULT_TOKEN):
    """Fetch all castle ownership info (7 castles, which crew owns each)."""
    return send_amf("CrewService.executeService", "getCastlesInfo",
                    [char_id, token])

def get_castle_info(castle_index, char_id=DEFAULT_CHAR_ID, token=DEFAULT_TOKEN):
    """Fetch details for a specific castle (0-6)."""
    return send_amf("CrewService.executeService", "getCastleInfo",
                    [char_id, token, int(castle_index)])

def get_crew_members_battle(char_id=DEFAULT_CHAR_ID, token=DEFAULT_TOKEN):
    """Fetch crew member battle/boss stats."""
    return send_amf("CrewService.executeService", "getMembersInfoForBattle",
                    [char_id, token])


# -- Character & Daily --------------------------------------------------------

def get_daily_data(char_id=DEFAULT_CHAR_ID, token=DEFAULT_TOKEN):
    """Fetch daily tasks, rewards, and character summary."""
    return send_amf("DailyService.executeService", "getData",
                    [char_id, token, "normal"])

def get_skills(char_id=DEFAULT_CHAR_ID, token=DEFAULT_TOKEN):
    """Fetch character skills."""
    return send_amf("AdvancedAcademy.executeService", "getSkills",
                    [char_id, token])


# -- Authentication -------------------------------------------------------------

def _decode_amf_response(raw):
    """Decode AMF response, handling various envelope formats."""
    from pyamf import remoting, decode as amf_decode
    import io

    # Try direct decode first (response starts with envelope header)
    try:
        msg = remoting.decode(io.BytesIO(raw))
        for _, r in msg:
            body = r.body
            if isinstance(body, dict):
                return body
    except Exception:
        pass

    # Fallback: try at /1/onResult - 4
    pos = raw.find(b"/1/onResult")
    if pos >= 0 and pos >= 4:
        try:
            msg = remoting.decode(io.BytesIO(raw[pos - 4:]))
            for _, r in msg:
                body = r.body
                if isinstance(body, dict):
                    return body
        except Exception:
            pass

    return {}


# -- Authentication -------------------------------------------------------------

# Static AC verification hash (from game client binary)
_AC_HASH = (
    "8f86c474c8f46425c210n558k116dd0b62bc3797249e869658c481f622d686ecb3a6b9d0ca251c7daaf9"
    "6838e9ab044ce5115dda23b0c05fa88061486e02dbbbb437968d45741e33f198ac9ce708b3c33d4124ba6c"
    "12774b16112520ea31552c82a9653a0e77a872fd637403137382daa0d77b93447a7227a444f79134c6db86"
    "751464c93bd6d9d6251daad74f4c79b9c6df0449de33e986107201e56fdfa5ed3cec0b4f1915cff54734b0"
    "9fb3814d6879d82a8bf0d7c9d7730eea90f311ea21c7545d9d33bce15b968453819c34ff6f"
)


def verify_files(sessionkey):
    """Verify client files with the game server. Required before character queries."""
    raw = send_amf_direct("AC.verifyFiles", [sessionkey, _AC_HASH])
    body = _decode_amf_response(raw)
    return body is True or (isinstance(body, dict) and body.get("status") == 1)


def verify_code(action, acc_id, code, device_id="NR_K0SCYkkilPH8pZj20"):
    """
    Submit email verification code for a new device.
    Returns True on success.
    """
    raw = send_amf_direct("Account.verifyCode",
                          [action, str(acc_id), code, device_id])
    body = _decode_amf_response(raw)
    return body is True or bool(body.get("status"))


def set_cache_cleared(acc_id, areas=None):
    """Mark cache as cleared for a new device. Called after verifyFiles."""
    if areas is None:
        areas = ["all"]
    raw = send_amf_direct("FilesManager.setCacheCleared",
                          [str(acc_id), areas])
    body = _decode_amf_response(raw)
    return body is True or bool(body.get("status"))


def login_user(username, password, device_id="NR_K0SCYkkilPH8pZj20"):
    """
    Login to the game and return session info.
    Returns: {"acc_id": int, "sessionkey": str, ...} on success.
             {"needs_verification": True, "acc_id": int} if verification needed.
             None on failure.
    """
    raw = send_amf_direct("SystemLogin.loginUser",
                          [username, password, device_id])
    body = _decode_amf_response(raw)

    if not body.get("status"):
        return None

    # Check for new device — needs email verification
    if not body.get("verified"):
        return {"needs_verification": True, "acc_id": body.get("uid")}

    sessionkey = body.get("sessionkey")
    if not sessionkey:
        return None

    verify_files(sessionkey)
    set_cache_cleared(body.get("uid"))
    return {
        "acc_id": body.get("uid"),
        "sessionkey": sessionkey,
        "clan_season": body.get("clan_season"),
        "crew_season": body.get("crew_season"),
    }


def get_characters(acc_id, sessionkey):
    """
    Fetch character list for an account.
    Returns: [(char_id, name, level), ...] or [] on failure.
    """
    raw = send_amf_direct("SystemLogin.getAllCharacters",
                          [acc_id, sessionkey])
    body = _decode_amf_response(raw)

    if isinstance(body, dict) and "account_data" in body:
        result = []
        for char in body["account_data"]:
            if isinstance(char, dict):
                result.append((
                    char.get("char_id"),
                    char.get("character_name", "?"),
                    char.get("character_level", 0),
                ))
        return result
    return []


# ─── Response Display ────────────────────────────────────────────────────────

def dump_readable(raw, limit=40):
    """Extract and print readable strings from a raw AMF response.
    This is a fallback for when PyAMF is not available."""
    data = strip_chunked_encoding(raw)
    text = data.decode("ascii", errors="replace")
    # Find all runs of printable ASCII >= 3 chars
    import re
    strings = re.findall(rb'[\x20-\x7E]{3,}', data)
    results = []
    for s in strings:
        decoded = s.decode("ascii", errors="replace").strip()
        if decoded and not decoded.startswith(("HTTP", "Date:", "Serv", "Cont",
                                                "Tran", "Acce", "cf-", "CF-",
                                                "Nel:", "Conn")):
            results.append(decoded)

    print(f"\n{'='*70}")
    print("  Readable Strings in AMF Response")
    print(f"{'='*70}")
    # Group into "records" — each clan/crew record has ~12-20 strings
    record = []
    record_keys = {"clan_id", "clan_name", "crew_id", "crew_name",
                   "member_reputation", "character_name"}
    for s in results:
        if any(k == s for k in record_keys) and record:
            # Print previous record
            vals = {}
            for i in range(0, len(record) - 1, 2):
                if i + 1 < len(record):
                    vals[record[i]] = record[i + 1]
            if vals:
                print(f"  -- Record --")
                for k, v in vals.items():
                    print(f"    {k}: {v}")
            record = [s]
            if len(results) >= record_keys and any(vals for vals in [{}]):
                if len(results) - results.index(s) > limit * 20:
                    break
        else:
            record.append(s)

    # Print final record
    if record:
        vals = {}
        for i in range(0, len(record) - 1, 2):
            if i + 1 < len(record):
                vals[record[i]] = record[i + 1]
        if vals:
            print(f"  -- Record --")
            for k, v in vals.items():
                print(f"    {k}: {v}")

    print(f"\n  Total readable strings: {len(results)}")


def dump_with_pyamf(raw):
    """Decode and pretty-print using PyAMF."""
    from pyamf import remoting, decode, AMF3

    data = strip_chunked_encoding(raw)

    # PyAMF expects a stream — wrap in BytesIO
    import io
    stream = io.BytesIO(data)

    try:
        msg = remoting.decode(stream)
    except Exception as e:
        print(f"  PyAMF decode error: {e}")
        print(f"  Raw first 64 bytes hex: {data[:64].hex()}")
        return

    print(f"\n{'='*70}")
    print(f"  Decoded AMF Response (PyAMF)")
    print(f"{'='*70}")
    print(f"  Version: {msg.amfVersion}")
    print(f"  Bodies: {len(list(msg))}")

    for target, response in msg:
        payload = response.body
        print(f"\n  --- Body ---")
        print(f"    Target: {target}")
        print(f"    Status: {response.status}")
        print(f"    Payload type: {type(payload).__name__}")
        _print_payload(payload, indent="    ")


def _print_payload(obj, indent="", depth=0):
    """Recursively print decoded AMF payload."""
    if depth > 6:
        print(f"{indent}... (max depth)")
        return

    if isinstance(obj, dict):
        # Handle castle info
        if "castle_name" in obj or "castle_owners" in obj:
            if "castle_owners" in obj:
                print(f"{indent}Castle Owners:")
                for i, owner in enumerate(obj["castle_owners"]):
                    cname = CASTLE_NAMES[i] if i < len(CASTLE_NAMES) else f"#{i}"
                    print(f"{indent}  {cname:12s} -> {owner}")
            else:
                name = obj.get("castle_name", "?")
                owner = obj.get("crew_name", obj.get("clan_name", "None"))
                print(f"{indent}Castle: {name} — Owner: {owner}")
            for k, v in obj.items():
                if k not in ("castle_name", "crew_name", "clan_name",
                             "castle_owners"):
                    if isinstance(v, (list, dict)):
                        print(f"{indent}  {k}: [{len(v)} items]")
                    else:
                        print(f"{indent}  {k}: {v}")
            return
        # Handle clan/crew record
        elif "clan_id" in obj:
            name = obj.get("clan_name", "?")
            rep = obj.get("clan_reputation", 0)
            members = obj.get("clan_members", 0)
            print(f"{indent}Clan #{obj['clan_id']}: {name}")
            print(f"{indent}  Rep: {rep:,} | Members: {members}/{obj.get('clan_max_members', '?')}")
            for k, v in obj.items():
                if k not in ("clan_id", "clan_name", "clan_reputation",
                             "clan_members", "clan_max_members",
                             "clan_day_points", "clan_logo"):
                    if isinstance(v, (list, dict)):
                        print(f"{indent}  {k}: [{len(v)} items]")
                    else:
                        print(f"{indent}  {k}: {v}")
        elif "crew_id" in obj:
            name = obj.get("crew_name", "?")
            dmg = obj.get("crew_damage", 0)
            members = obj.get("crew_members", 0)
            print(f"{indent}Crew #{obj['crew_id']}: {name}")
            print(f"{indent}  Damage: {dmg:,} | Members: {members}/{obj.get('crew_max_members', '?')}")
        elif "character_name" in obj:
            name = obj.get("character_name", "?")
            rep = obj.get("member_reputation", 0)
            lvl = obj.get("character_level", 0)
            dmg = obj.get("member_damage", 0)
            info = f"Lv.{lvl}"
            if rep:
                info += f" Rep:{rep:,}"
            if dmg:
                info += f" Dmg:{dmg:,}"
            print(f"{indent}{name} ({info})")
        else:
            # Generic dict
            for k, v in list(obj.items())[:10]:
                if isinstance(v, (list, dict)):
                    print(f"{indent}{k}: [{len(v)} items]")
                elif isinstance(v, str) and len(v) > 80:
                    print(f"{indent}{k}: '{v[:80]}...'")
                else:
                    print(f"{indent}{k}: {v}")

    elif isinstance(obj, (list, tuple)):
        print(f"{indent}[{len(obj)} items]")
        for item in obj[:8]:  # Show first 8 items
            if isinstance(item, dict):
                _print_payload(item, indent + "  ", depth + 1)
            elif isinstance(item, (list, tuple)):
                if item and isinstance(item[0], dict):
                    for sub in item[:5]:
                        _print_payload(sub, indent + "  ", depth + 1)
                else:
                    print(f"{indent}  [...{len(item)} items]")
            else:
                print(f"{indent}  {item}")
        if len(obj) > 8:
            print(f"{indent}  ... and {len(obj) - 8} more items")

    elif isinstance(obj, bytes):
        print(f"{indent}<bytes: {len(obj)}>")
    else:
        print(f"{indent}{obj}")


# ─── Main / CLI ──────────────────────────────────────────────────────────────

def export_castles_json():
    """
    Login via env vars, fetch castle data once, write castle_data.json.
    For GitHub Actions — reads NR_USERNAME / NR_PASSWORD from environment.
    """
    import os
    username = os.environ.get("NR_USERNAME")
    password = os.environ.get("NR_PASSWORD")
    if not username or not password:
        print("ERROR: Set NR_USERNAME and NR_PASSWORD environment variables")
        sys.exit(1)

    print(f"[*] Logging in as {username}...")
    session = login_user(username, password)
    if session is None:
        print("ERROR: Login failed — check credentials")
        sys.exit(1)
    if session.get("needs_verification"):
        print(f"ERROR: Account '{username}' is not verified.")
        print("       Log in once via the TUI or game client to verify your email.")
        sys.exit(1)

    print(f"[*] Fetching castle data (acc_id={session['acc_id']})...")
    chars = get_characters(session["acc_id"], session["sessionkey"])
    if not chars:
        print("ERROR: No characters found on this account")
        sys.exit(1)
    char_id = chars[0][0]
    token = session["sessionkey"]
    print(f"[*] Using char_id={char_id} ({chars[0][1]})")

    castles = _fetch_castles(char_id, token)
    with open("castle_data.json", "w") as f:
        json.dump(castles, f, indent=2)
    print(f"[*] Wrote castle_data.json ({len(castles)} castles)")


def _fetch_castles(char_id, token):
    """Fetch all castle data via AMF. Returns list of 7 dicts."""
    from pyamf import remoting
    import io

    raw = get_castles_info(char_id, token)
    pos = raw.find(b"\x00\x03")
    if pos >= 0:
        raw = raw[pos:]
    msg = remoting.decode(io.BytesIO(raw))
    owners = list(msg)[0][1].body.get("castle_owners", [])

    castles = []
    for i in range(7):
        try:
            raw_d = get_castle_info(i, char_id, token)
            pos_d = raw_d.find(b"\x00\x03")
            if pos_d >= 0:
                raw_d = raw_d[pos_d:]
            msg_d = remoting.decode(io.BytesIO(raw_d))
            d = list(msg_d)[0][1].body if len(list(msg_d)) > 0 else {}
            castles.append({
                "castle": CASTLE_NAMES[i] if i < len(CASTLE_NAMES) else f"#{i}",
                "owner": owners[i] if i < len(owners) else "?",
                "wall_hp": round(d.get("wall_hp", 0), 1),
                "defender_hp": round(d.get("defender_hp", 0), 1),
            })
        except Exception:
            castles.append({
                "castle": CASTLE_NAMES[i] if i < len(CASTLE_NAMES) else f"#{i}",
                "owner": owners[i] if i < len(owners) else "?",
                "wall_hp": 0, "defender_hp": 0,
            })
    return castles


def _git_amend_push():
    """Commit castle_data.json with --amend and force-push."""
    import subprocess, os

    def _run(cmd, check=False):
        r = subprocess.run(cmd, capture_output=True, text=True)
        if r.returncode != 0:
            print(f"[git] {' '.join(cmd)} -> rc={r.returncode}")
            if r.stderr.strip():
                print(f"[git] stderr: {r.stderr.strip()}")
            if check:
                raise RuntimeError(f"git command failed: {' '.join(cmd)}")
        return r

    _run(["git", "config", "user.name", "github-actions"])
    _run(["git", "config", "user.email", "actions@github.com"])
    _run(["git", "add", "castle_data.json"])
    diff = _run(["git", "diff", "--staged", "HEAD"])
    if not diff.stdout.strip():
        return  # no changes

    _run(["git", "commit", "--amend", "--reset-author", "-m", "update castle data [daemon]"], check=True)

    branch = os.environ.get("GIT_BRANCH", "")
    if branch:
        _run(["git", "push", "origin", f"HEAD:{branch}", "--force"], check=True)
    else:
        _run(["git", "push", "origin", "HEAD", "--force"], check=True)


def serve_daemon():
    """
    Run indefinitely: fetch castle data every 30s and push to git.
    For GitHub Actions daemon mode — cron restarts every 6h.
    Reads NR_USERNAME / NR_PASSWORD from environment.
    """
    import os, subprocess, time as _time

    username = os.environ.get("NR_USERNAME")
    password = os.environ.get("NR_PASSWORD")
    if not username or not password:
        print("ERROR: Set NR_USERNAME and NR_PASSWORD environment variables")
        sys.exit(1)

    print(f"[daemon] Logging in as {username}...")
    session = login_user(username, password)
    if session is None:
        print("[daemon] ERROR: Login failed — check credentials")
        sys.exit(1)
    if session.get("needs_verification"):
        print(f"[daemon] ERROR: Account '{username}' is not verified.")
        sys.exit(1)

    chars = get_characters(session["acc_id"], session["sessionkey"])
    if not chars:
        print("[daemon] ERROR: No characters found")
        sys.exit(1)
    char_id = chars[0][0]
    token = session["sessionkey"]
    print(f"[daemon] Using char_id={char_id} ({chars[0][1]})")

    # Write initial data
    castles = _fetch_castles(char_id, token)
    with open("castle_data.json", "w") as f:
        json.dump(castles, f, indent=2)
    _git_amend_push()
    print(f"[daemon] Initial data written ({len(castles)} castles)")

    # Main loop
    count = 0
    while True:
        _time.sleep(30)
        count += 1
        try:
            castles = _fetch_castles(char_id, token)
            with open("castle_data.json", "w") as f:
                json.dump(castles, f, indent=2)
            _git_amend_push()
            print(f"[daemon] Loop #{count}: pushed {len(castles)} castles")
        except Exception as e:
            print(f"[daemon] Loop #{count} error: {e}")


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        print("\nQuick start:")
        print("  python game_bot.py clans              — all 2186 clans ranked by rep")
        print("  python game_bot.py crews              — all crews ranked by damage")
        print("  python game_bot.py clan_members       — clan member list + stats")
        print("  python game_bot.py castles            — castle ownership map")
        print("  python game_bot.py daily              — daily tasks + character summary")
        print("  python game_bot.py stamina            — character stamina info")
        print("  python game_bot.py skills             — character skills")
        print()
        return

    cmd = sys.argv[1].lower()

    if cmd == "clans":
        raw = get_all_clans()
    elif cmd == "crews":
        raw = get_all_crews()
    elif cmd == "clan":
        raw = get_clan()
    elif cmd == "crew":
        raw = get_crew()
    elif cmd == "clan_status":
        raw = get_clan_status()
    elif cmd == "crew_status":
        raw = get_crew_status()
    elif cmd == "clan_members":
        raw = get_clan_members()
    elif cmd == "clan_history":
        raw = get_clan_history()
    elif cmd == "clans_battle":
        raw = get_clans_for_battle()
    elif cmd == "clan_members_battle":
        raw = get_clan_members_battle()
    elif cmd == "stamina":
        raw = get_stamina()
    elif cmd == "castles":
        raw = get_castles_info()
    elif cmd == "castle":
        if len(sys.argv) < 3:
            print("Usage: python game_bot.py castle <0-6>")
            return
        raw = get_castle_info(sys.argv[2])
    elif cmd == "crew_members_battle":
        raw = get_crew_members_battle()
    elif cmd == "daily":
        raw = get_daily_data()
    elif cmd == "skills":
        raw = get_skills()
    elif cmd == "--export-castles":
        export_castles_json()
        return
    elif cmd == "--serve-daemon":
        serve_daemon()
        return
    else:
        print(f"Unknown command: {cmd}")
        print("Try: clans, crews, clan, clan_members, castles, daily, ...")

    # Try PyAMF first, fall back to raw string extraction
    try:
        import pyamf
        dump_with_pyamf(raw)
    except ImportError:
        print("[!] PyAMF not installed. Using raw string extraction.")
        print("[!] Install with: pip install pyamf")
        dump_readable(raw)


if __name__ == "__main__":
    main()

"""
Ninja Rift Game Bot — Terminal UI
==================================
Rich-powered terminal interface for exploring clan/crew data.

Run: python Internal/tui.py
"""

import sys, os, time, json, threading, math, io
from game_bot import *

try:
    from rich.console import Console
    from rich.table import Table as RichTable
    from rich.panel import Panel
    from rich.text import Text
    from rich import box
    from rich.markup import escape
except ImportError:
    print("[!] pip install rich")
    sys.exit(1)

import msvcrt

console = Console()
PAGE_SIZE = 20
CASTLE_NAMES = [
    "Hiroshima", "Himeji", "Kumamoto", "Okazaki",
    "Inuyama", "Gifu", "Hikone",
]

# ─── Shared State ────────────────────────────────────────────────────────────
_last_fetch: dict = {}       # cache: key -> (data, timestamp)
_auto_refresh = True
_refresh_interval = 60
_running = True

# ─── Header Cache (prevents API spam on every redraw) ────────────────────────
_header_cache = (0, 0, 0)     # (stamina, max_stamina, season)
_header_cache_time = 0         # timestamp of last fetch

# ─── Session State (set after login) ─────────────────────────────────────────
SESSION_CHAR_ID = None
SESSION_TOKEN = None
SESSION_USERNAME = None
SESSION_CHAR_NAME = None

SESSION_FILE = os.path.join(os.path.dirname(__file__), "session.json")


def session_context():
    """Return current (char_id, token) or raise if not logged in."""
    if not SESSION_CHAR_ID or not SESSION_TOKEN:
        raise RuntimeError("Not logged in")
    return SESSION_CHAR_ID, SESSION_TOKEN


def cached_fetch(key, fetcher):
    """Call fetcher() and cache result with timestamp."""
    global _last_fetch
    data = fetcher()
    _last_fetch[key] = (data, time.time())
    return data


def get_cached(key):
    """Return cached data if fresh, else None."""
    entry = _last_fetch.get(key)
    if entry is None:
        return None
    data, ts = entry
    if _auto_refresh and time.time() - ts > _refresh_interval:
        return None
    return data


# ─── Keyboard ────────────────────────────────────────────────────────────────

def get_key(timeout=0):
    """Non-blocking key read. Returns None if no key pressed."""
    import msvcrt
    if timeout:
        deadline = time.time() + timeout
        while time.time() < deadline:
            if msvcrt.kbhit():
                return msvcrt.getch().decode("utf-8", errors="replace").lower()
            time.sleep(0.05)
        return None
    else:
        while True:
            if msvcrt.kbhit():
                return msvcrt.getch().decode("utf-8", errors="replace").lower()
            time.sleep(0.05)


# ─── Header ──────────────────────────────────────────────────────────────────

def build_header():
    """Fetch and return header panel with char info. Cached to avoid API spam."""
    global _header_cache, _header_cache_time
    char_id, token = SESSION_CHAR_ID, SESSION_TOKEN
    now = time.time()

    # Use cache if fresh (< _refresh_interval seconds old)
    if char_id and token and now - _header_cache_time < _refresh_interval:
        st_cur, st_max, season = _header_cache
    else:
        st_cur, st_max, season = "?", "?", "?"
        if char_id and token:
            try:
                stam = get_stamina(char_id, token)
                pos = stam.find(b"\x00\x03")
                if pos >= 0:
                    stam = stam[pos:]
                from pyamf import remoting
                import io
                msg = remoting.decode(io.BytesIO(stam))
                for _, r in msg:
                    body = r.body
                    if isinstance(body, dict) and "char_data" in body:
                        cd = body["char_data"]
                        if isinstance(cd, dict):
                            st_cur = cd.get("stamina", "?")
                            st_max = cd.get("max_stamina", "?")
            except:
                pass

            try:
                status = get_crew_status(char_id, token)
                pos = status.find(b"\x00\x03")
                if pos >= 0:
                    status = status[pos:]
                msg = remoting.decode(io.BytesIO(status))
                for _, r in msg:
                    body = r.body
                    if isinstance(body, dict) and isinstance(body.get("result"), dict):
                        season = body["result"].get("crew_season", "?")
            except:
                pass

            _header_cache = (st_cur, st_max, season)
            _header_cache_time = now

    name = SESSION_CHAR_NAME or SESSION_USERNAME or "?"
    lines = [
        f"[bold cyan]Ninja Rift Game Bot[/bold cyan]",
        f"[dim]{name}  |  Stamina: {st_cur}/{st_max}  |  Season {season}[/dim]",
    ]
    return Panel("\n".join(lines), box=box.HEAVY, border_style="cyan")


# ─── Login Screen ─────────────────────────────────────────────────────────────

def load_saved_sessions():
    """Load saved accounts from session.json. Returns list of dicts."""
    if os.path.exists(SESSION_FILE):
        try:
            with open(SESSION_FILE, "r") as f:
                data = json.load(f)
                # Support both old single-account and new multi-account format
                if isinstance(data, dict):
                    return [data]
                if isinstance(data, list):
                    return data
        except:
            pass
    return []


def save_session(username, password, char_id):
    """Add or update an account in session.json."""
    accounts = load_saved_sessions()
    # Update existing or append
    for acc in accounts:
        if acc.get("username") == username:
            acc["password"] = password
            acc["last_char_id"] = char_id
            break
    else:
        accounts.append({"username": username, "password": password,
                         "last_char_id": char_id})
    os.makedirs(os.path.dirname(SESSION_FILE), exist_ok=True)
    with open(SESSION_FILE, "w") as f:
        json.dump(accounts, f, indent=2)


def remove_session(username):
    """Remove a single account from session.json."""
    accounts = load_saved_sessions()
    accounts = [a for a in accounts if a.get("username") != username]
    if accounts:
        with open(SESSION_FILE, "w") as f:
            json.dump(accounts, f, indent=2)
    else:
        if os.path.exists(SESSION_FILE):
            os.remove(SESSION_FILE)


def clear_sessions():
    """Delete all saved accounts."""
    if os.path.exists(SESSION_FILE):
        os.remove(SESSION_FILE)


def login_screen(saved_user=None, saved_pass=None, saved_char_id=None):
    """Prompt for username/password, login, pick character.
    If saved_user is provided, skips credential input."""
    global SESSION_CHAR_ID, SESSION_TOKEN, SESSION_USERNAME, SESSION_CHAR_NAME
    global _header_cache_time

    console.clear()
    console.print(Panel(
        "[bold cyan]Ninja Rift Game Bot[/bold cyan]\n[dim]Login[/dim]",
        box=box.HEAVY, border_style="cyan"
    ))
    console.print()

    if saved_user:
        console.print(f"[dim]Logging in as [cyan]{saved_user}[/cyan]...[/dim]")
        username, password = saved_user, saved_pass
    else:
        console.print("[bold]Username:[/bold] ", end="")
        username = input().strip()
        if not username:
            console.print("[red]Username required[/red]")
            time.sleep(1)
            return False

        console.print("[bold]Password:[/bold] ", end="")
        password = ""
        while True:
            ch = msvcrt.getch()
            if ch == b"\r" or ch == b"\n":
                break
            elif ch == b"\x08":
                if password:
                    password = password[:-1]
                    console.print("\b \b", end="")
            elif ch == b"\x1b":
                return False
            else:
                password += ch.decode("utf-8", errors="replace")
                console.print("*", end="")
        console.print()

    console.print("\n[dim]Logging in...[/dim]")
    try:
        session = login_user(username, password)
    except Exception as e:
        console.print(f"[red]Connection error: {e}[/red]")
        time.sleep(2)
        return False

    if not session:
        console.print("[red]Login failed — check credentials[/red]")
        time.sleep(2)
        return False

    # Handle new device verification
    if session.get("needs_verification"):
        console.print("\n[yellow][!] New device detected — verification required[/yellow]")
        console.print("[dim]A 6-digit code was sent to your email.[/dim]")
        console.print("\n[bold]Verification code: [/bold]", end="")
        code = input().strip()
        if not code.isdigit() or len(code) != 6:
            console.print("[red]Invalid code — must be 6 digits[/red]")
            time.sleep(2)
            return False

        console.print("[dim]Verifying...[/dim]")
        try:
            if not verify_code("login", session["acc_id"], code):
                console.print("[red]Invalid verification code[/red]")
                time.sleep(2)
                return False
        except Exception as e:
            console.print(f"[red]Verification error: {e}[/red]")
            time.sleep(2)
            return False

        # Re-login after verification
        console.print("[dim]Logging in again...[/dim]")
        try:
            session = login_user(username, password)
        except Exception as e:
            console.print(f"[red]Connection error: {e}[/red]")
            time.sleep(2)
            return False

        if not session or session.get("needs_verification"):
            console.print("[red]Login failed after verification[/red]")
            time.sleep(2)
            return False

    acc_id = session["acc_id"]
    token = session["sessionkey"]

    console.print(f"[dim]Fetching characters...[/dim]")
    chars = get_characters(acc_id, token)
    if not chars:
        console.print("[red]No characters found on this account[/red]")
        time.sleep(2)
        return False

    # Auto-select if saved_char_id matches
    auto_char = None
    if saved_char_id and len(chars) == 1:
        auto_char = chars[0]
    elif saved_char_id:
        for c in chars:
            if c[0] == saved_char_id:
                auto_char = c
                break

    if auto_char:
        char_id, char_name, char_level = auto_char
        console.print(f"[dim]Auto-selecting [cyan]{char_name}[/cyan] (Lv.{char_level})[/dim]")
    elif len(chars) == 1:
        char_id, char_name, char_level = chars[0]
    else:
        console.clear()
        console.print(Panel(
            f"[bold cyan]Select Character[/bold cyan]\n[dim]Account: {username}[/dim]",
            box=box.HEAVY, border_style="cyan"
        ))
        console.print()
        for i, (cid, cname, clvl) in enumerate(chars, 1):
            console.print(f"  [{i}] [cyan]{cname}[/cyan] [dim](Lv.{clvl}, ID:{cid})[/dim]")
        console.print()
        console.print("[bold]Choice: [/bold]", end="")
        choice = input().strip()
        if not choice.isdigit() or not (1 <= int(choice) <= len(chars)):
            console.print("[red]Invalid choice[/red]")
            time.sleep(1)
            return False
        char_id, char_name, char_level = chars[int(choice) - 1]

    global SESSION_CHAR_ID, SESSION_TOKEN, SESSION_USERNAME, SESSION_CHAR_NAME
    global _header_cache_time
    SESSION_CHAR_ID = char_id
    SESSION_TOKEN = token
    SESSION_USERNAME = username
    SESSION_CHAR_NAME = char_name
    _header_cache_time = 0

    # Save credentials if this was a manual login
    if not saved_user:
        existing = load_saved_sessions()
        already_saved = any(a.get("username") == username for a in existing)
        if already_saved:
            console.print(f"\n[dim]Update saved login for [cyan]{username}[/cyan]? (y/n): [/dim]", end="")
            save_choice = get_key()
            if save_choice and save_choice.lower() == "y":
                save_session(username, password, char_id)
                console.print("[green]Updated![/green]")
        else:
            console.print(f"\n[dim]Save login for next time? (y/n): [/dim]", end="")
            save_choice = get_key()
            if save_choice and save_choice.lower() == "y":
                save_session(username, password, char_id)
                console.print("[green]Saved![/green]")

    console.print(f"\n[green]Logged in as {char_name} (ID:{char_id})[/green]")
    time.sleep(1)
    return True


# ─── Paginated Table Views ───────────────────────────────────────────────────

def paginated_view(title, columns, rows, extra_info=None):
    """
    Display a sortable, searchable, paginated table.
    Returns when user presses 'q'.
    """
    page = 0
    sort_col = -1          # -1 = no sort (keep API order)
    sort_desc = True
    search = ""
    all_rows = list(rows)

    def _num(v):
        """Parse a number from string, handling commas like '74,236'."""
        try:
            return int(str(v).replace(",", ""))
        except:
            return 0

    while True:
        # Filter
        filtered = [r for r in all_rows if search.lower() in str(r).lower()]
        total_pages = max(1, math.ceil(len(filtered) / PAGE_SIZE))
        page = min(page, total_pages - 1)

        # Sort (only if user pressed 's')
        if sort_col >= 0:
            try:
                filtered.sort(
                    key=lambda x: _num(x[sort_col]) if isinstance(x, (list, tuple)) and sort_col < len(x) else 0,
                    reverse=sort_desc
                )
            except:
                pass

        # Slice
        start = page * PAGE_SIZE
        chunk = filtered[start:start + PAGE_SIZE]

        console.clear()
        console.print(build_header())
        console.print()

        # Build table
        table = RichTable(title=title, box=box.SIMPLE_HEAD, border_style="bright_black")
        for c in columns:
            table.add_column(
                c["header"],
                justify=c.get("justify", "left"),
                no_wrap=c.get("no_wrap", False),
                style=c.get("style", None),
            )
        for row in chunk:
            table.add_row(*[str(v) for v in row])

        console.print(table)

        # Pagination bar
        bar = Text()
        bar.append(f" Page {page + 1}/{total_pages} ", style="bold white")
        bar.append("| ")
        bar.append("n]ext ", style="green")
        bar.append("p]rev ", style="green")
        bar.append("/]search ", style="green")
        bar.append("s]ort ", style="green")
        bar.append("r]efresh ", style="green")
        bar.append("q]uit", style="red")
        if search:
            bar.append(f"  Filter: [yellow]{search}[/yellow]")
        if extra_info:
            bar.append(f"  {extra_info}")
        console.print(bar)

        key = get_key().lower() if not _auto_refresh else get_key(1.0)

        if key == "q":
            return
        elif key == "n":
            page = min(page + 1, total_pages - 1)
        elif key == "p":
            page = max(page - 1, 0)
        elif key == "s":
            sort_col = sort_col + 1
            if sort_col >= len(columns):
                sort_col = -1  # wrap back to no-sort
            sort_desc = not sort_desc if sort_col == 0 else True
        elif key == "r":
            return "refresh"
        elif key == "/":
            console.print("\n  Search: ", end="")
            search = input().strip()
            page = 0
        elif key == "\x1b":  # ESC
            return
        elif key and key.isdigit() and 1 <= int(key) <= len(chunk):
            # Quick-select row
            pass


# ─── View Functions ──────────────────────────────────────────────────────────

def view_clan_ranking():
    char_id, token = session_context()
    while True:
        raw = cached_fetch("clans", lambda: get_all_clans(char_id, token)) if get_cached("clans") is None else get_cached("clans")
        if isinstance(raw, tuple):
            raw = raw[0]  # cached fetcher returns raw bytes

        pos = raw.find(b"\x00\x03")
        if pos >= 0:
            raw = raw[pos:]
        from pyamf import remoting
        import io
        msg = remoting.decode(io.BytesIO(raw))
        clans = list(msg)[0][1].body if len(list(msg)) > 0 else []

        rows = []
        for i, c in enumerate(clans[:500]):  # top 500 for performance
            if isinstance(c, dict):
                rows.append((
                    str(i + 1),
                    c.get("clan_name", "?")[:25],
                    f"{c.get('clan_reputation', 0):,}",
                    f"{c.get('clan_members', 0)}/{c.get('clan_max_members', '?')}",
                    f"{c.get('clan_day_points', 0):,}",
                    str(c.get("clan_id", "")),
                ))

        result = paginated_view(
            "Clan Ranking (Top 500)",
            [
                {"header": "#", "justify": "right", "no_wrap": True},
                {"header": "Clan Name", "style": "cyan"},
                {"header": "Rep", "justify": "right", "style": "green"},
                {"header": "Members", "justify": "right"},
                {"header": "Day Pts", "justify": "right", "style": "yellow"},
                {"header": "ID", "justify": "right", "style": "dim"},
            ],
            rows,
        )
        if result != "refresh":
            return


def view_crew_ranking():
    char_id, token = session_context()
    while True:
        raw = cached_fetch("crews", lambda: get_all_crews(char_id, token)) if get_cached("crews") is None else get_cached("crews")
        if isinstance(raw, tuple):
            raw = raw[0]

        pos = raw.find(b"\x00\x03")
        if pos >= 0:
            raw = raw[pos:]
        from pyamf import remoting
        import io
        msg = remoting.decode(io.BytesIO(raw))
        crews = list(msg)[0][1].body if len(list(msg)) > 0 else []

        rows = []
        for i, c in enumerate(crews):
            if isinstance(c, dict):
                rows.append((
                    str(i + 1),
                    c.get("crew_name", "?")[:25],
                    f"{c.get('crew_damage', 0):,}",
                    f"{c.get('crew_members', 0)}/{c.get('crew_max_members', '?')}",
                    str(c.get("crew_id", "")),
                ))

        result = paginated_view(
            "Crew Ranking",
            [
                {"header": "#", "justify": "right", "no_wrap": True},
                {"header": "Crew Name", "style": "cyan"},
                {"header": "Damage", "justify": "right", "style": "green"},
                {"header": "Members", "justify": "right"},
                {"header": "ID", "justify": "right", "style": "dim"},
            ],
            rows,
        )
        if result != "refresh":
            return


def view_castles():
    char_id, token = session_context()
    while True:
        console.clear()
        console.print(build_header())
        console.print()

        # Use combined cache for owners + details
        cache_key = "castle_full"
        entry = _last_fetch.get(cache_key)
        if entry and _auto_refresh and time.time() - entry[1] > _refresh_interval:
            entry = None
        if entry:
            owners, details = entry[0]
        else:
            raw = get_castles_info(char_id, token)
            pos = raw.find(b"\x00\x03")
            if pos >= 0:
                raw = raw[pos:]
            from pyamf import remoting
            import io
            msg = remoting.decode(io.BytesIO(raw))
            owners = list(msg)[0][1].body.get("castle_owners", []) if len(list(msg)) > 0 else []

            details = []
            for i in range(7):
                try:
                    raw_d = get_castle_info(i, char_id, token)
                    pos_d = raw_d.find(b"\x00\x03")
                    if pos_d >= 0:
                        raw_d = raw_d[pos_d:]
                    msg_d = remoting.decode(io.BytesIO(raw_d))
                    details.append(list(msg_d)[0][1].body if len(list(msg_d)) > 0 else {})
                except Exception:
                    details.append({})
            _last_fetch[cache_key] = ((owners, details), time.time())

        table = RichTable(title="Castle Ownership", box=box.SIMPLE_HEAD, border_style="bright_black")
        table.add_column("Castle", style="cyan")
        table.add_column("Owner", style="yellow")
        table.add_column("Wall HP", justify="right", style="green")
        table.add_column("Def HP", justify="right", style="red")
        table.add_column("#", justify="right", style="dim")
        for i, owner in enumerate(owners):
            d = details[i] if i < len(details) else {}
            table.add_row(
                CASTLE_NAMES[i] if i < len(CASTLE_NAMES) else f"#{i}",
                owner,
                f"{d.get('wall_hp', 0):,.1f}",
                f"{d.get('defender_hp', 0):,.1f}",
                str(i),
            )

        console.print(table)

        bar = Text()
        bar.append(" 0-6] detail ", style="green")
        bar.append("r]efresh ", style="green")
        bar.append("q]uit", style="red")
        console.print(bar)

        key = get_key(1.0)
        if key == "q" or key == "\x1b":
            return
        elif key == "r":
            _last_fetch.pop("castle_full", None)
        elif key and key.isdigit():
            idx = int(key)
            if 0 <= idx < len(CASTLE_NAMES):
                _view_castle_detail(idx)


def _view_castle_detail(idx):
    char_id, token = SESSION_CHAR_ID, SESSION_TOKEN
    console.clear()
    console.print(build_header())
    console.print()

    try:
        raw = get_castle_info(idx, char_id, token)
        pos = raw.find(b"\x00\x03")
        if pos >= 0:
            raw = raw[pos:]
        from pyamf import remoting
        import io
        msg = remoting.decode(io.BytesIO(raw))
        data = list(msg)[0][1].body if len(list(msg)) > 0 else {}

        table = RichTable(title=f"Castle: {CASTLE_NAMES[idx]}", box=box.SIMPLE, border_style="cyan")
        table.add_column("Field")
        table.add_column("Value")
        for k, v in data.items():
            if isinstance(v, (list, dict)):
                table.add_row(k, f"[{len(v)} items]")
            elif isinstance(v, float):
                table.add_row(k, f"{v:,.1f}")
            else:
                table.add_row(k, str(v))
        console.print(table)
    except Exception as e:
        console.print(f"[red]Error: {e}[/red]")

    console.print("\n[dim]Press any key to return...[/dim]")
    get_key()


def view_daily():
    char_id, token = session_context()
    console.clear()
    console.print(build_header())
    console.print()

    try:
        raw = get_daily_data(char_id, token)
        pos = raw.find(b"\x00\x03")
        if pos >= 0:
            raw = raw[pos:]
        from pyamf import remoting
        import io
        msg = remoting.decode(io.BytesIO(raw))
        data = list(msg)[0][1].body if len(list(msg)) > 0 else {}

        table = RichTable(title="Daily Data", box=box.SIMPLE, border_style="cyan")
        table.add_column("Field")
        table.add_column("Value")
        for k, v in data.items():
            if isinstance(v, (list, dict)):
                table.add_row(k, f"[{len(v)} items]")
                if isinstance(v, list) and v and isinstance(v[0], dict):
                    for item in v[:5]:
                        for ik, iv in item.items():
                            table.add_row(f"  {ik}", str(iv)[:60])
            else:
                table.add_row(k, str(v)[:60])
        console.print(table)
    except Exception as e:
        console.print(f"[red]Error: {e}[/red]")

    console.print("\n[dim]Press any key to return...[/dim]")
    get_key()


def view_clan_war():
    char_id, token = session_context()
    while True:
        raw = cached_fetch("clans_battle", lambda: get_clans_for_battle(char_id, token)) if get_cached("clans_battle") is None else get_cached("clans_battle")
        if isinstance(raw, tuple):
            raw = raw[0]

        pos = raw.find(b"\x00\x03")
        if pos >= 0:
            raw = raw[pos:]
        from pyamf import remoting
        import io
        msg = remoting.decode(io.BytesIO(raw))
        clans = list(msg)[0][1].body if len(list(msg)) > 0 else []

        rows = []
        for i, c in enumerate(clans):
            if isinstance(c, dict):
                rows.append((
                    str(i + 1),
                    c.get("clan_name", "?")[:25],
                    f"{c.get('clan_reputation', 0):,}",
                    f"{c.get('clan_members', 0)}/{c.get('clan_max_members', '?')}",
                    str(c.get("clan_id", "")),
                ))

        result = paginated_view(
            "Clan War — Battle-Eligible Clans",
            [
                {"header": "#", "justify": "right", "no_wrap": True},
                {"header": "Clan Name", "style": "cyan"},
                {"header": "Rep", "justify": "right", "style": "green"},
                {"header": "Members", "justify": "right"},
                {"header": "ID", "justify": "right", "style": "dim"},
            ],
            rows,
        )
        if result != "refresh":
            return


def view_skills():
    char_id, token = session_context()
    console.clear()
    console.print(build_header())
    console.print()

    try:
        raw = get_skills(char_id, token)
        pos = raw.find(b"\x00\x03")
        if pos >= 0:
            raw = raw[pos:]
        from pyamf import remoting
        import io
        msg = remoting.decode(io.BytesIO(raw))
        data = list(msg)[0][1].body if len(list(msg)) > 0 else {}

        table = RichTable(title="Skills", box=box.SIMPLE, border_style="cyan")
        table.add_column("Element")
        table.add_column("Skills", style="yellow")
        for k, v in data.items():
            if isinstance(v, list):
                count = f"{len(v)} skills"
                names = []
                for s in v:
                    if isinstance(s, dict):
                        names.append(s.get("name", s.get("skill_name", "?")))
                if names:
                    count += " — " + ", ".join(names[:8])
                table.add_row(f"[cyan]{k}[/cyan]", count)
            else:
                table.add_row(f"[cyan]{k}[/cyan]", str(v)[:60])
        console.print(table)
    except Exception as e:
        console.print(f"[red]Error: {e}[/red]")

    console.print("\n[dim]Press any key to return...[/dim]")
    get_key()


def view_search_clan():
    console.clear()
    console.print(build_header())
    console.print()
    console.print("[bold]Search Clan[/bold]")
    console.print("Enter clan ID: ", end="")
    clan_id = input().strip()
    if not clan_id.isdigit():
        console.print("[red]Invalid ID[/red]")
        console.print("\n[dim]Press any key...[/dim]")
        get_key()
        return

    console.print(f"\n[dim]Fetching clan #{clan_id}...[/dim]")
    try:
        # Use the public PHP API for any clan lookup
        import urllib.request
        url = f"https://playninjarift.com/api/detail_clan_website.php?clan_id={clan_id}"
        req = urllib.request.Request(url, headers={"User-Agent": "clan-snapshot/1.0"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode())

        console.clear()
        console.print(build_header())
        console.print()

        panel = Panel(
            f"[bold cyan]{data.get('clan_name', '?')}[/bold cyan]\n"
            f"ID: {clan_id}  |  Logo: {data.get('clan_logo', '?')}",
            title="Clan Info", border_style="cyan"
        )
        console.print(panel)

        members = data.get("members", [])
        table = RichTable(title=f"Members ({len(members)})", box=box.SIMPLE_HEAD)
        table.add_column("#", justify="right", no_wrap=True)
        table.add_column("Name", style="cyan")
        table.add_column("Rep", justify="right", style="green")
        table.add_column("Level", justify="right")
        for i, m in enumerate(members[:30], 1):
            table.add_row(
                str(i),
                m.get("character_name", "?")[:25],
                f"{m.get('member_reputation', 0):,}",
                str(m.get("character_level", "")),
            )
        console.print(table)
    except Exception as e:
        console.print(f"[red]Error: {e}[/red]")

    console.print("\n[dim]Press any key to return...[/dim]")
    get_key()


# ─── Main Menu ───────────────────────────────────────────────────────────────

def main_menu():
    global _auto_refresh, _last_fetch

    # Login first — check for saved sessions
    if not SESSION_CHAR_ID:
        saved_accounts = load_saved_sessions()
        if saved_accounts:
            console.clear()
            console.print(Panel(
                "[bold cyan]Ninja Rift Game Bot[/bold cyan]\n"
                "[dim]Saved accounts[/dim]",
                box=box.HEAVY, border_style="cyan"
            ))
            console.print()
            for i, acc in enumerate(saved_accounts, 1):
                console.print(f"  [{i}] [cyan]{acc.get('username', '?')}[/cyan]")
            console.print()
            console.print("  [n] [dim]New login[/dim]")
            console.print("  [q] [red]Quit[/red]")
            console.print()
            console.print("[dim]Select: [/dim]", end="")
            key = get_key()
            if key == "q" or key == "\x1b":
                console.clear()
                console.print("[dim]Goodbye![/dim]")
                return
            elif key == "n":
                if not login_screen():
                    return
            elif key.isdigit() and 1 <= int(key) <= len(saved_accounts):
                acc = saved_accounts[int(key) - 1]
                if not login_screen(
                    saved_user=acc.get("username"),
                    saved_pass=acc.get("password"),
                    saved_char_id=acc.get("last_char_id"),
                ):
                    return
            else:
                # Default to first saved account
                acc = saved_accounts[0]
                if not login_screen(
                    saved_user=acc.get("username"),
                    saved_pass=acc.get("password"),
                    saved_char_id=acc.get("last_char_id"),
                ):
                    return
        else:
            if not login_screen():
                console.clear()
                console.print("[dim]Goodbye![/dim]")
                return

    while True:
        console.clear()
        console.print(build_header())
        console.print()

        menu = Panel(
            "\n".join([
                " [1] [cyan]Clan Ranking[/cyan]       [dim]All 2186 clans[/dim]",
                " [2] [cyan]Crew Ranking[/cyan]       [dim]All crews by damage[/dim]",
                " [3] [cyan]Castle Ownership[/cyan]   [dim]7 castles + owners[/dim]",
                " [4] [cyan]Daily Data[/cyan]         [dim]Calendar, tasks, rewards[/dim]",
                " [5] [cyan]Clan War[/cyan]           [dim]Battle-eligible clans[/dim]",
                " [6] [cyan]Skills[/cyan]             [dim]7 element trees[/dim]",
                " [7] [cyan]Search Clan[/cyan]        [dim]Lookup clan by ID[/dim]",
                "",
                f" {escape('[a]')} [dim]Toggle auto-refresh[/dim]",
                f" {escape('[q]')} [red]Quit[/red]",
            ]),
            title="Menu", border_style="cyan", box=box.HEAVY
        )
        console.print(menu)

        console.print(f"\n[dim]Auto-refresh: {'ON' if _auto_refresh else 'OFF'} ({_refresh_interval}s)[/dim]")
        console.print("[dim]Select: [/dim]", end="")
        key = get_key()

        if key == "1":
            view_clan_ranking()
        elif key == "2":
            view_crew_ranking()
        elif key == "3":
            view_castles()
        elif key == "4":
            view_daily()
        elif key == "5":
            view_clan_war()
        elif key == "6":
            view_skills()
        elif key == "7":
            view_search_clan()
        elif key == "q" or key == "\x1b":
            console.clear()
            console.print("[dim]Goodbye![/dim]")
            return
        elif key == "a":
            _auto_refresh = not _auto_refresh
            _last_fetch = {}


if __name__ == "__main__":
    try:
        main_menu()
    except KeyboardInterrupt:
        console.clear()
        console.print("[dim]Goodbye![/dim]")

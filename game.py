"""레벨 1 용사 — 게임 엔진.

events.xlsx(이벤트 / 선택지 / 후일담 시트)를 읽어 게임을 진행한다.
스토리를 고치려면 코드를 건드리지 말고 events.xlsx만 바꿔서 다시 올리면 된다.
"""
import random
import re
from pathlib import Path

from openpyxl import load_workbook

DATA_FILE = Path(__file__).parent / "events.xlsx"

MAX_HP = 100
START = {"hp": 100, "gold": 100, "power": 1}
POTION = "회복 물약"
POTION_HEAL = 30
COMPANION_BONUS = 2      # 전투력 판정에서 동료 1명당 보너스
MAX_COMPANIONS = 2
LAST_CHAPTER = 7
DEFEAT_ENDING = "E792"   # 패배 엔딩: 사망과 같게 처리하고 후일담을 붙이지 않는다

CMD_NEW = ("새 게임", "새게임", "시작", "처음부터", "다시 시작")
CMD_CONTINUE = ("이어하기", "계속")
CMD_STATUS = ("상태", "상태 보기")
CMD_POTION = ("물약 사용", "물약")


# ------------------------------------------------------------------ 데이터 읽기
def _clean(v):
    if v is None:
        return None
    if isinstance(v, str):
        v = v.strip()
        return v or None
    return v


def _num(v):
    try:
        return int(v)
    except (TypeError, ValueError):
        return 0


def load_data(path=DATA_FILE):
    wb = load_workbook(path, data_only=True)
    events, order = {}, []
    for row in wb["이벤트"].iter_rows(min_row=2, values_only=True):
        row = [_clean(c) for c in row] + [None] * 7
        if not row[0]:
            continue
        chapter = int(re.sub(r"\D", "", str(row[1])) or 0)
        events[row[0]] = {
            "id": row[0], "chapter": chapter, "type": row[2], "title": row[3] or "",
            "text": row[4] or "", "appear": row[6], "choices": [],
        }
        order.append(row[0])
    for row in wb["선택지"].iter_rows(min_row=2, values_only=True):
        row = [_clean(c) for c in row] + [None] * 17
        if not row[0] or row[0] not in events or not row[2]:
            continue
        events[row[0]]["choices"].append({
            "no": _num(row[1]), "label": str(row[2]), "need": row[3], "check": row[4],
            "text": row[5] or "", "hp": _num(row[6]), "gold": _num(row[7]), "power": _num(row[8]),
            "gain": row[9], "next": row[10] or "자동",
            "fail_text": row[11] or "", "fail_hp": _num(row[12]), "fail_gold": _num(row[13]),
            "fail_gain": row[14], "fail_next": row[15] or "자동",
        })
    for ev in events.values():
        ev["choices"].sort(key=lambda c: c["no"])
    epilogues = []
    if "후일담" in wb.sheetnames:
        for row in wb["후일담"].iter_rows(min_row=2, values_only=True):
            row = [_clean(c) for c in row] + [None] * 3
            if row[0] and row[1]:
                cond = re.sub(r"\s*\(.*?\)\s*$", "", str(row[0]))
                endings = [e.strip() for e in str(row[2] or "").split("/") if e.strip()]
                epilogues.append({"cond": cond, "text": row[1], "endings": endings})
    return {"events": events, "order": order, "epilogues": epilogues}


DATA = load_data()


# ------------------------------------------------------------------ 상태
def new_state(legacy=None, runs=0):
    s = {
        "hp": START["hp"], "gold": START["gold"], "power": START["power"],
        "abilities": [], "items": [], "clues": [], "companions": [], "lost": [],
        "attr": None, "chapter": 0, "queue": [], "current": None,
        "legacy": legacy, "runs": runs, "over": False,
    }
    if legacy:
        s["abilities"].append(legacy)
    return s


def status_line(s):
    return f"❤️ 체력 {s['hp']}  💰 골드 {s['gold']}  ⚔️ 전투력 {s['power']}"


def status_detail(s):
    def join(xs):
        return ", ".join(xs) if xs else "없음"
    items = {}
    for it in s["items"]:
        items[it] = items.get(it, 0) + 1
    item_txt = [f"{k} {v}개" if v > 1 else k for k, v in items.items()]
    lines = [
        status_line(s),
        "",
        f"능력: {join(s['abilities'])}",
        f"아이템: {join(item_txt)}",
        f"동료: {join(s['companions'])}",
        f"성검 속성: {s['attr'] or '없음'}",
        f"단서: {join(s['clues'])}",
    ]
    if s["companions"]:
        lines.append(f"(전투력 판정 시 동료 보너스 +{COMPANION_BONUS * len(s['companions'])})")
    return "\n".join(lines)


# ------------------------------------------------------------------ 조건·판정
def _tokens(spec):
    if not spec:
        return []
    return [t.strip() for t in str(spec).split(" / ") if t.strip()]


def meets(s, spec):
    """필요 조건 / 등장 조건(확률 제외)을 만족하는지."""
    for tok in _tokens(spec):
        m = re.match(r"^골드\s*>=\s*(\d+)$", tok)
        if m:
            if s["gold"] < int(m.group(1)):
                return False
            continue
        m = re.match(r"^전투력\s*>=\s*(\d+)$", tok)
        if m:
            if effective_power(s) < int(m.group(1)):
                return False
            continue
        if tok.startswith("확률:"):
            continue
        if ":" not in tok:
            continue
        kind, name = [x.strip() for x in tok.split(":", 1)]
        if kind == "능력" and name not in s["abilities"]:
            return False
        if kind == "아이템" and name not in s["items"]:
            return False
        if kind == "단서" and name not in s["clues"]:
            return False
        if kind == "속성" and s["attr"] != name:
            return False
        if kind == "동료":
            if name == "있음":
                if not s["companions"]:
                    return False
            elif name not in s["companions"]:
                return False
        if kind == "상실" and not s["lost"]:
            return False
    return True


def effective_power(s):
    return s["power"] + COMPANION_BONUS * len(s["companions"])


def roll_check(s, check, rng):
    """판정. 비어 있으면 항상 성공."""
    if not check:
        return True
    m = re.match(r"^전투력\s*>=\s*(\d+)$", check)
    if m:
        return effective_power(s) >= int(m.group(1))
    m = re.match(r"^확률\s*:\s*(\d+)$", check)
    if m:
        return rng.random() * 100 < int(m.group(1))
    return True


def gains_companion(spec):
    return any(t.startswith("동료:") for t in _tokens(spec))


def visible_choices(s, ev):
    out = []
    for c in ev["choices"]:
        if not meets(s, c["need"]):
            continue
        if gains_companion(c["gain"]) and len(s["companions"]) >= MAX_COMPANIONS:
            continue
        out.append(c)
    return out


def apply_gain(s, spec, notes, rng):
    for tok in _tokens(spec):
        if ":" not in tok:
            continue
        kind, name = [x.strip() for x in tok.split(":", 1)]
        if kind == "능력":
            if name not in s["abilities"]:
                s["abilities"].append(name)
                notes.append(f"능력 획득: {name}")
        elif kind == "아이템":
            s["items"].append(name)
            notes.append(f"아이템 획득: {name}")
        elif kind == "단서":
            if name not in s["clues"]:
                s["clues"].append(name)
                notes.append(f"단서 획득: {name}")
        elif kind == "동료":
            if name not in s["companions"] and len(s["companions"]) < MAX_COMPANIONS:
                s["companions"].append(name)
                notes.append(f"동료 합류: {name}")
        elif kind == "속성":
            s["attr"] = name
            notes.append(f"성검 속성: {name}")
        elif kind == "상실" and name == "동료" and s["companions"]:
            gone = s["companions"].pop(rng.randrange(len(s["companions"])))
            s["lost"].append(gone)
            notes.append(f"{gone}을(를) 잃었다")


def apply_delta(s, hp, gold, power, notes):
    if hp:
        before = s["hp"]
        s["hp"] = max(0, min(MAX_HP, s["hp"] + hp))
        if s["hp"] != before:
            notes.append(f"체력 {s['hp'] - before:+d}")
    if gold:
        before = s["gold"]
        s["gold"] = max(0, s["gold"] + gold)
        if s["gold"] != before:
            notes.append(f"골드 {s['gold'] - before:+d}")
    if power:
        before = s["power"]
        s["power"] = max(0, s["power"] + power)
        if s["power"] != before:
            notes.append(f"전투력 {s['power'] - before:+d}")


# ------------------------------------------------------------------ 진행
def build_queue(s, chapter, rng):
    evs = [DATA["events"][i] for i in DATA["order"] if DATA["events"][i]["chapter"] == chapter]
    intro = [e["id"] for e in evs if e["type"] in ("시작", "장 도입")]
    randoms = [e["id"] for e in evs if e["type"] == "무작위"]
    bosses = [e["id"] for e in evs if e["type"] == "보스"]
    rng.shuffle(randoms)
    for e in evs:
        if e["type"] != "숨김":
            continue
        appear = e["appear"] or ""
        m = re.search(r"확률\s*:\s*(\d+)", appear)
        if m and rng.random() * 100 >= int(m.group(1)):
            continue
        if not meets(s, appear):
            continue
        randoms.insert(rng.randrange(len(randoms) + 1), e["id"])
    return intro + randoms + bosses


def advance(s, target, rng):
    """다음 이벤트로 이동. target이 '자동'이면 장 진행 규칙대로."""
    if target and target != "자동" and target in DATA["events"]:
        if target in s["queue"]:
            s["queue"].remove(target)
        s["current"] = target
        s["chapter"] = DATA["events"][target]["chapter"] or s["chapter"]
        return
    while not s["queue"]:
        if s["chapter"] >= LAST_CHAPTER:
            s["current"] = None
            s["over"] = True
            return
        s["chapter"] += 1
        s["queue"] = build_queue(s, s["chapter"], rng)
    s["current"] = s["queue"].pop(0)


def pick_legacy(s, rng):
    return rng.choice(s["abilities"]) if s["abilities"] else None


# ------------------------------------------------------------------ 화면 만들기
def _reply(texts, buttons):
    return {"texts": [t for t in texts if t], "buttons": buttons}


def event_screen(s, pre=None):
    """현재 이벤트 화면. pre는 앞에 붙일 말풍선(직전 선택의 결과 등)."""
    ev = DATA["events"][s["current"]]
    head = f"[{ev['chapter']}장] {ev['title']}\n{status_line(s)}"
    buttons = [c["label"] for c in visible_choices(s, ev)]
    if POTION in s["items"] and s["hp"] < MAX_HP:
        buttons.append("물약 사용")
    buttons.append("상태")
    return _reply([pre, f"{head}\n\n{ev['text']}"], buttons)


def ending_screen(s, pre, rng):
    ev = DATA["events"][s["current"]]
    lines = [f"— {ev['title']} —", "", ev["text"]]
    defeat = ev["id"] == DEFEAT_ENDING
    if not defeat:
        extra = [e["text"] for e in DATA["epilogues"]
                 if (not e["endings"] or ev["id"] in e["endings"]) and meets(s, e["cond"])]
        if extra:
            lines += ["", "· · ·", ""] + extra
    legacy = pick_legacy(s, rng) if defeat else s.get("legacy")
    tail = "\n\n다시 시작하려면 '새 게임'을 눌러 주세요."
    if defeat and legacy:
        tail = f"\n\n다음 용사는 '{legacy}' 능력을 물려받습니다." + tail
    runs = s["runs"]
    s.clear()
    s.update(new_state(legacy=legacy if defeat else None, runs=runs))
    s["over"] = True
    return _reply([pre, "\n".join(lines) + tail], ["새 게임"])


def death_screen(s, pre, rng):
    legacy = pick_legacy(s, rng)
    msg = "체력이 다했다.\n\n당신의 검은 길 위에 남았다. 언젠가 누군가 그 앞에 잠시 설 것이다."
    if legacy:
        msg += f"\n\n다음 용사는 '{legacy}' 능력을 물려받습니다."
    msg += "\n\n다시 시작하려면 '새 게임'을 눌러 주세요."
    runs = s["runs"]
    s.clear()
    s.update(new_state(legacy=legacy, runs=runs))
    s["over"] = True
    return _reply([pre, msg], ["새 게임"])


def intro_screen():
    return _reply(
        ["레벨 1 용사\n\n성검이 하필 당신을 골랐다.\n선택에 따라 이야기가 달라지는 텍스트 모험입니다.\n\n"
         "· 아래 버튼을 눌러 선택하세요.\n· '상태'로 능력과 아이템을 확인할 수 있습니다.\n· 언제든 '새 게임'으로 처음부터 시작할 수 있습니다."],
        ["새 게임"],
    )


def start_game(s, rng):
    legacy, runs = s.get("legacy"), s.get("runs", 0) + 1
    s.clear()
    s.update(new_state(legacy=legacy, runs=runs))
    advance(s, "자동", rng)
    pre = f"선대 용사의 유품에서 '{legacy}'을(를) 물려받았다." if legacy else None
    return event_screen(s, pre)


# ------------------------------------------------------------------ 입력 처리
def handle(s, utterance, rng=random):
    """유저 발화 하나를 처리한다. s(상태 dict)는 제자리에서 바뀐다."""
    text = (utterance or "").strip()

    if text in CMD_NEW:
        return start_game(s, rng)
    if not s or s.get("current") is None:
        return intro_screen()
    if text in CMD_STATUS:
        scr = event_screen(s)
        return _reply([status_detail(s)], scr["buttons"])
    if text in CMD_POTION:
        if POTION in s["items"] and s["hp"] < MAX_HP:
            s["items"].remove(POTION)
            before = s["hp"]
            s["hp"] = min(MAX_HP, s["hp"] + POTION_HEAL)
            return event_screen(s, f"회복 물약을 마셨다. (체력 +{s['hp'] - before})")
        return event_screen(s, "지금은 물약을 쓸 수 없다.")

    ev = DATA["events"][s["current"]]
    choice = next((c for c in visible_choices(s, ev) if c["label"] == text), None)
    if choice is None:
        pre = None if text in CMD_CONTINUE else "아래 버튼 중에서 골라 주세요."
        return event_screen(s, pre)

    notes = []
    if roll_check(s, choice["check"], rng):
        result, target = choice["text"], choice["next"]
        apply_delta(s, choice["hp"], choice["gold"], choice["power"], notes)
        apply_gain(s, choice["gain"], notes, rng)
    else:
        result, target = choice["fail_text"] or choice["text"], choice["fail_next"]
        apply_delta(s, choice["fail_hp"], choice["fail_gold"], 0, notes)
        apply_gain(s, choice["fail_gain"], notes, rng)
    pre = result + (f"\n\n({', '.join(notes)})" if notes else "")

    if s["hp"] <= 0:
        return death_screen(s, pre, rng)

    advance(s, target, rng)
    if s["current"] is None:
        return death_screen(s, pre, rng)  # 이벤트가 더 없을 때의 안전장치
    if DATA["events"][s["current"]]["type"] == "엔딩":
        return ending_screen(s, pre, rng)
    return event_screen(s, pre)

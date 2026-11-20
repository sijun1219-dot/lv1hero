"""레벨 1 용사 — 게임 엔진 (v2: 검의 기억·광장·강화 포함).

events.xlsx를 읽어 게임을 진행한다.
스토리·가격·보상을 고치려면 코드를 건드리지 말고 events.xlsx만 바꿔서 다시 올리면 된다.

유저 한 명의 저장 데이터(state)는 두 부분이다.
  state["meta"] : 판이 끝나도 남는 것 (검의 기억, 강화, 본 엔딩, 플레이 횟수)
  state["run"]  : 지금 진행 중인 판 (없으면 None)
"""
import random
import re
from datetime import datetime, timedelta, timezone
from pathlib import Path

from openpyxl import load_workbook

DATA_FILE = Path(__file__).parent / "events.xlsx"
IMAGE_DIR = Path(__file__).parent / "images"   # 이벤트 이미지 폴더. 파일 이름은 이벤트ID (예: E290.jpg)

BASE_HP = 100
START = {"gold": 100, "power": 1}
POTION = "회복 물약"
POTION_HEAL = 30
COMPANION_BONUS = 2      # 전투력 판정에서 동료 1명당 보너스
MAX_COMPANIONS = 2
LAST_CHAPTER = 7
KING_CHAPTER = 6         # 이 장의 보스를 통과하면 '사천왕 통과'로 센다
DEFEAT_ENDING = "E792"   # 패배 엔딩: 사망과 같게 처리하고 후일담을 붙이지 않는다
BLOOD = "마족의 피"      # 체력이 낮을수록 전투력이 오르고, 이벤트마다 체력이 줄어드는 상태
MIASMA = "마왕령의 독기"  # 마왕령(6장)에서 이벤트를 넘길 때마다 체력이 줄어든다. 강화 '검들의 가호'로 줄인다

CMD_NEW = ("새 게임", "새게임", "게임 시작", "게임시작", "시작", "처음부터", "다시 시작")
CMD_FORFEIT = "포기하고 새 게임"
CMD_CONTINUE = ("이어하기", "계속")
CMD_STATUS = ("상태", "상태 보기")
CMD_POTION = ("물약 사용", "물약")
CMD_HUB = ("광장",)
CMD_RECORD = ("기록",)
NO_ATTR = "각인 없이 시작"


def today():
    """한국 시간 기준 오늘 날짜."""
    return datetime.now(timezone(timedelta(hours=9))).date().isoformat()


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
        row = [_clean(c) for c in row] + [None] * 9
        if not row[0]:
            continue
        chapter = int(re.sub(r"\D", "", str(row[1])) or 0)
        image = row[7] if isinstance(row[7], str) and row[7].startswith("http") else None
        if image is None:
            # 엑셀에 주소가 없으면 images 폴더에서 '이벤트ID.jpg'(또는 .png)를 찾아 쓴다
            for ext in ("jpg", "jpeg", "png"):
                if (IMAGE_DIR / f"{row[0]}.{ext}").exists():
                    image = f"/img/{row[0]}.{ext}"
                    break
        events[row[0]] = {
            "id": row[0], "chapter": chapter, "type": row[2], "title": row[3] or "",
            "text": row[4] or "", "appear": row[6], "image": image, "choices": [],
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

    upgrades = []
    if "강화" in wb.sheetnames:
        for row in wb["강화"].iter_rows(min_row=2, values_only=True):
            row = [_clean(c) for c in row] + [None] * 6
            if not row[0] or not row[1]:
                continue
            prices = [int(p) for p in re.findall(r"\d+", str(row[5] or ""))]
            upgrades.append({"key": str(row[0]), "name": str(row[1]), "desc": row[2] or "",
                             "kind": row[3], "value": row[4], "prices": prices})

    rewards = {}
    if "보상" in wb.sheetnames:
        for row in wb["보상"].iter_rows(min_row=2, values_only=True):
            row = [_clean(c) for c in row] + [None] * 3
            if row[0]:
                rewards[str(row[0])] = _num(row[2])

    params = {}
    if "상태 효과" in wb.sheetnames:
        for row in wb["상태 효과"].iter_rows(min_row=2, values_only=True):
            row = [_clean(c) for c in row] + [None] * 3
            if row[0]:
                params[str(row[0])] = _num(row[2])

    return {"events": events, "order": order, "epilogues": epilogues,
            "upgrades": upgrades, "rewards": rewards, "params": params}


DATA = load_data()


# ------------------------------------------------------------------ 영구 기록(meta)
def new_meta():
    return {"memory": 0, "upgrades": {}, "endings": [], "runs": 0, "deaths": 0,
            "best": 0, "daily": "", "legacy": []}


def ensure(state):
    """저장 데이터의 모양을 맞춘다(처음 온 유저, 예전 형식 포함)."""
    if not isinstance(state.get("meta"), dict):
        state.clear()
        state["meta"] = new_meta()
        state["run"] = None
    meta = new_meta()
    meta.update(state["meta"])
    state["meta"] = meta
    state.setdefault("run", None)
    state.setdefault("mode", None)
    return state


def level(meta, key):
    return int(meta["upgrades"].get(key, 0))


def bonus(meta, kind):
    """종류(hp/gold/power/legacy/hidden)별 강화 합계."""
    total = 0
    for up in DATA["upgrades"]:
        if up["kind"] == kind:
            total += _num(up["value"]) * level(meta, up["key"])
    return total


def owned_attrs(meta):
    return [str(up["value"]) for up in DATA["upgrades"]
            if up["kind"] == "attr" and level(meta, up["key"]) > 0]


def next_price(meta, up):
    lv = level(meta, up["key"])
    return up["prices"][lv] if lv < len(up["prices"]) else None


# ------------------------------------------------------------------ 한 판(run)
def new_run(meta, attr=None):
    hp_max = BASE_HP + bonus(meta, "hp")
    return {
        "hp": hp_max, "hp_max": hp_max,
        "gold": START["gold"] + bonus(meta, "gold"),
        "power": START["power"] + bonus(meta, "power"),
        "abilities": list(meta.get("legacy") or []), "items": [], "clues": [],
        "companions": [], "lost": [], "attr": attr, "status": [],
        "chapter": 0, "queue": [], "current": None, "kings": 0,
    }


def miasma_damage(meta):
    """마왕령의 독기 피해. '검들의 가호' 강화만큼 줄어든다."""
    return max(0, DATA["params"].get("miasma_damage", 0) - bonus(meta, "miasma"))


def potion_cap():
    return DATA["params"].get("potion_cap", 99)


def blood_bonus(r):
    """마족의 피: 잃은 체력에 비례한 전투력 판정 보너스."""
    if BLOOD not in r.get("status", []):
        return 0
    per = max(1, DATA["params"].get("blood_hp_per_point", 15))
    return min(DATA["params"].get("blood_max_bonus", 5), (r["hp_max"] - r["hp"]) // per)


def blood_help():
    p = DATA["params"]
    return (f"· 체력이 낮을수록 전투력 판정이 오른다 (잃은 체력 {p.get('blood_hp_per_point', 15)}마다 +1, "
            f"최대 +{p.get('blood_max_bonus', 5)})\n"
            f"· 이벤트를 넘길 때마다 체력 -{p.get('blood_drain', 5)}\n"
            f"· 체력을 가득 채우면 사라진다")


def status_line(r):
    line = f"❤️ 체력 {r['hp']}/{r['hp_max']}  💰 골드 {r['gold']}  ⚔️ 전투력 {r['power']}"
    if BLOOD in r.get("status", []):
        line += f"\n🩸 {BLOOD} (전투력 판정 +{blood_bonus(r)})"
    return line


def status_detail(r):
    def join(xs):
        return ", ".join(xs) if xs else "없음"
    items = {}
    for it in r["items"]:
        items[it] = items.get(it, 0) + 1
    item_txt = [f"{k} {v}개" if v > 1 else k for k, v in items.items()]
    lines = [
        status_line(r), "",
        f"능력: {join(r['abilities'])}",
        f"아이템: {join(item_txt)}",
        f"동료: {join(r['companions'])}",
        f"성검 속성: {r['attr'] or '없음'}",
        f"단서: {join(r['clues'])}",
    ]
    if r["companions"]:
        lines.append(f"(전투력 판정 시 동료 보너스 +{COMPANION_BONUS * len(r['companions'])})")
    if BLOOD in r.get("status", []):
        lines += ["", f"{BLOOD}", blood_help()]
    return "\n".join(lines)


def status_with_meta(state):
    r, meta = state["run"], state["meta"]
    text = status_detail(r)
    if miasma_damage(meta) > 0:
        text += f"\n\n{MIASMA}: 마왕령에서 이벤트마다 체력 -{miasma_damage(meta)}"
    return text


# ------------------------------------------------------------------ 조건·판정
def _tokens(spec):
    if not spec:
        return []
    return [t.strip() for t in str(spec).split(" / ") if t.strip()]


def effective_power(r):
    return r["power"] + COMPANION_BONUS * len(r["companions"]) + blood_bonus(r)


def meets(r, spec):
    """필요 조건 / 등장 조건(확률 제외)을 만족하는지. ' 또는 '으로 이으면 하나만 만족해도 된다."""
    if spec and " 또는 " in str(spec):
        return any(_meets_all(r, alt) for alt in str(spec).split(" 또는 "))
    return _meets_all(r, spec)


def _meets_all(r, spec):
    for tok in _tokens(spec):
        m = re.match(r"^골드\s*>=\s*(\d+)$", tok)
        if m:
            if r["gold"] < int(m.group(1)):
                return False
            continue
        m = re.match(r"^전투력\s*>=\s*(\d+)$", tok)
        if m:
            if effective_power(r) < int(m.group(1)):
                return False
            continue
        if tok.startswith("확률:") or ":" not in tok:
            continue
        kind, name = [x.strip() for x in tok.split(":", 1)]
        if kind == "능력" and name not in r["abilities"]:
            return False
        if kind == "아이템" and name not in r["items"]:
            return False
        if kind == "단서" and name not in r["clues"]:
            return False
        if kind == "속성" and r["attr"] != name:
            return False
        if kind == "동료":
            if name == "있음":
                if not r["companions"]:
                    return False
            elif name not in r["companions"]:
                return False
        if kind == "상실" and not r["lost"]:
            return False
        if kind == "상태" and name not in r.get("status", []):
            return False
    return True


def roll_check(r, check, rng):
    """판정. 비어 있으면 항상 성공."""
    if not check:
        return True
    m = re.match(r"^전투력\s*>=\s*(\d+)$", check)
    if m:
        return effective_power(r) >= int(m.group(1))
    m = re.match(r"^확률\s*:\s*(\d+)$", check)
    if m:
        return rng.random() * 100 < int(m.group(1))
    return True


def settle_blood(r):
    """체력이 가득 차면 마족의 피가 사라진다. 사라졌으면 True."""
    if BLOOD in r["status"] and r["hp"] >= r["hp_max"]:
        r["status"].remove(BLOOD)
        return True
    return False


def gains_companion(spec):
    return any(t.startswith("동료:") for t in _tokens(spec))


def visible_choices(r, ev):
    out = []
    for c in ev["choices"]:
        if not meets(r, c["need"]):
            continue
        if gains_companion(c["gain"]) and len(r["companions"]) >= MAX_COMPANIONS:
            continue
        out.append(c)
    return out


def apply_gain(r, spec, notes, rng):
    for tok in _tokens(spec):
        if ":" not in tok:
            continue
        m = re.match(r"^(.*)@\s*(\d+)\s*%?$", tok)   # 뒤에 @15% 가 붙으면 그 확률로만 적용
        if m:
            tok = m.group(1).strip()
            if rng.random() * 100 >= int(m.group(2)):
                continue
        kind, name = [x.strip() for x in tok.split(":", 1)]
        if kind == "상태":
            if name not in r["status"]:
                r["status"].append(name)
                notes.append(f"상태: {name}")
        elif kind == "해제":
            if name in r["status"]:
                r["status"].remove(name)
                notes.append(f"{name}가 가라앉았다")
            continue
        if kind == "능력":
            if name not in r["abilities"]:
                r["abilities"].append(name)
                notes.append(f"능력 획득: {name}")
        elif kind == "아이템":
            if name == POTION and r["items"].count(POTION) >= potion_cap():
                notes.append(f"{POTION}은 {potion_cap()}개까지만 들 수 있다")
                continue
            r["items"].append(name)
            notes.append(f"아이템 획득: {name}")
        elif kind == "단서":
            if name not in r["clues"]:
                r["clues"].append(name)
                notes.append(f"단서 획득: {name}")
        elif kind == "동료":
            if name not in r["companions"] and len(r["companions"]) < MAX_COMPANIONS:
                r["companions"].append(name)
                notes.append(f"동료 합류: {name}")
        elif kind == "속성":
            r["attr"] = name
            notes.append(f"성검 속성: {name}")
        elif kind == "상실" and name == "동료" and r["companions"]:
            gone = r["companions"].pop(rng.randrange(len(r["companions"])))
            r["lost"].append(gone)
            notes.append(f"{gone}을(를) 잃었다")


def apply_delta(r, hp, gold, power, notes):
    if hp:
        before = r["hp"]
        r["hp"] = max(0, min(r["hp_max"], r["hp"] + hp))
        if r["hp"] != before:
            notes.append(f"체력 {r['hp'] - before:+d}")
    if gold:
        before = r["gold"]
        r["gold"] = max(0, r["gold"] + gold)
        if r["gold"] != before:
            notes.append(f"골드 {r['gold'] - before:+d}")
    if power:
        before = r["power"]
        r["power"] = max(0, r["power"] + power)
        if r["power"] != before:
            notes.append(f"전투력 {r['power'] - before:+d}")


# ------------------------------------------------------------------ 장 진행
def build_queue(r, meta, chapter, rng):
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
        if m and rng.random() * 100 >= int(m.group(1)) + bonus(meta, "hidden"):
            continue
        if not meets(r, appear):
            continue
        randoms.insert(rng.randrange(len(randoms) + 1), e["id"])
    return intro + randoms + bosses


def advance(r, meta, target, rng):
    """다음 이벤트로 이동. target이 '자동'이면 장 진행 규칙대로."""
    if target and target != "자동" and target in DATA["events"]:
        if target in r["queue"]:
            r["queue"].remove(target)
        r["current"] = target
        r["chapter"] = DATA["events"][target]["chapter"] or r["chapter"]
        return
    while not r["queue"]:
        if r["chapter"] >= LAST_CHAPTER:
            r["current"] = None
            return
        r["chapter"] += 1
        r["queue"] = build_queue(r, meta, r["chapter"], rng)
    r["current"] = r["queue"].pop(0)


# ------------------------------------------------------------------ 판 종료와 보상
def finish_run(state, ending_id, rng):
    """판을 끝내고 검의 기억을 지급한다. (보상 설명 문자열을 돌려준다)"""
    meta, r = state["meta"], state["run"]
    rw = DATA["rewards"]
    lines, total = [], 0

    def add(label, amount):
        nonlocal total
        if amount > 0:
            total += amount
            lines.append(f"· {label} +{amount}")

    died = ending_id is None
    # 너무 이른 장에서 죽으면 기억이 남지 않는다 (일일 보너스도 쓰지 않고 남겨 둔다)
    too_early = died and r["chapter"] < rw.get("death_min_chapter", 1)
    if too_early:
        pass
    elif died:
        add(f"{r['chapter']}장까지 도달", rw.get("death_per_chapter", 0) * max(1, r["chapter"]))
    else:
        add(DATA["events"][ending_id]["title"], rw.get(ending_id, 0))
        if ending_id not in meta["endings"]:
            meta["endings"].append(ending_id)
            add("처음 본 엔딩", rw.get("first_ending", 0))
        if ending_id != DEFEAT_ENDING and r["companions"] and not r["lost"]:
            add("동료 전원 생존", rw.get("companions_alive", 0))
    add(f"사천왕 {r['kings']}명 통과", rw.get("king", 0) * r["kings"])
    if not too_early and meta["daily"] != today():
        meta["daily"] = today()
        add("오늘의 첫 판", rw.get("daily", 0))

    meta["memory"] += total
    meta["best"] = max(meta["best"], r["chapter"])
    if died or ending_id == DEFEAT_ENDING:
        meta["deaths"] += 1
        slots = 1 + bonus(meta, "legacy")
        pool = list(r["abilities"])
        rng.shuffle(pool)
        meta["legacy"] = pool[:slots]
    else:
        meta["legacy"] = []
    state["run"] = None
    state["mode"] = None

    if too_early:
        out = [f"검의 기억을 얻지 못했다.",
               f"({rw.get('death_min_chapter', 1)}장 이후에 쓰러져야 검에 기억이 남는다)"]
    else:
        out = [f"검의 기억 +{total}"] + lines
    out += ["", f"보유한 검의 기억: {meta['memory']}"]
    if meta["legacy"]:
        out.append(f"다음 용사가 물려받는 능력: {', '.join(meta['legacy'])}")
    return "\n".join(out)


# ------------------------------------------------------------------ 화면 만들기
def _reply(texts, buttons, card=None):
    """texts: 글 말풍선들 / card: 이벤트 카드(제목·본문·이미지·카드 안 버튼) / buttons: 화면 맨 아래 버튼"""
    return {"texts": [t for t in texts if t], "buttons": buttons, "card": card}


CONTINUE = "계속"
CARD_BUTTONS = 3        # 카드 안에 들어가는 버튼 수 한도
IMAGE_CARD_TEXT = 230   # 이미지 카드의 본문 한도
TEXT_CARD_TEXT = 400    # 글 카드의 제목+본문 한도


def event_image(r, ev):
    if ev.get("image"):
        return ev["image"]
    parent = DATA["events"].get(r.get("parent") or "")
    if ev["type"] == "연계" and parent:
        return parent.get("image")
    return None


END_BUTTONS = ["새 게임", "광장", "기록"]


def event_screen(state, pre=None):
    r = state["run"]
    ev = DATA["events"][r["current"]]
    title = f"[{ev['chapter']}장] {ev['title']}"
    labels = [c["label"] for c in visible_choices(r, ev)] or [CONTINUE]
    quick = []
    if POTION in r["items"] and r["hp"] < r["hp_max"]:
        quick.append("물약 사용")
    quick.append("상태")
    head = (pre + "\n\n" if pre else "") + status_line(r)
    image = event_image(r, ev)
    fits = (len(ev["text"]) <= IMAGE_CARD_TEXT) if image else (len(title) + len(ev["text"]) <= TEXT_CARD_TEXT)
    if fits and len(labels) <= CARD_BUTTONS and len(title) <= 50:
        card = {"title": title, "desc": ev["text"], "image": image, "buttons": labels}
        return _reply([head], quick, card)
    # 카드에 담기지 않으면 글 말풍선과 아래 버튼으로 보여 준다
    return _reply([head, f"{title}\n\n{ev['text']}"], labels + quick)


def ending_screen(state, pre, rng):
    r = state["run"]
    ev = DATA["events"][r["current"]]
    lines = [f"— {ev['title']} —", "", ev["text"]]
    if ev["id"] != DEFEAT_ENDING:
        extra = [e["text"] for e in DATA["epilogues"]
                 if (not e["endings"] or ev["id"] in e["endings"]) and meets(r, e["cond"])]
        if extra:
            lines += ["", "· · ·", ""] + extra
    reward = finish_run(state, ev["id"], rng)
    return _reply([pre, "\n".join(lines), reward], END_BUTTONS)


def death_screen(state, pre, rng):
    msg = "체력이 다했다.\n\n당신의 검은 길 위에 남았다. 언젠가 누군가 그 앞에 잠시 설 것이다."
    reward = finish_run(state, None, rng)
    return _reply([pre, msg, reward], END_BUTTONS)


def intro_screen():
    return _reply(
        ["레벨 1 용사\n\n성검이 하필 당신을 골랐다.\n선택에 따라 이야기가 달라지는 텍스트 모험입니다.\n\n"
         "· 아래 버튼을 눌러 선택하세요.\n· 판이 끝나면 '검의 기억'을 얻고, 광장에서 다음 용사를 강하게 만들 수 있습니다.\n"
         "· '상태'로 능력과 아이템을 확인할 수 있습니다."],
        ["새 게임", "광장", "기록"],
    )


def hub_screen(state, pre=None):
    meta = state["meta"]
    lines = ["— 광장 —", "돌바닥에 꽂힌 검들이 다음 용사를 기다린다.", "",
             f"보유한 검의 기억: {meta['memory']}", ""]
    buttons = []
    for up in DATA["upgrades"]:
        lv, price = level(meta, up["key"]), next_price(meta, up)
        stage = f" ({lv}/{len(up['prices'])})" if len(up["prices"]) > 1 else ""
        if price is None:
            lines.append(f"✔ {up['name']}{stage} — 완료")
        else:
            lines.append(f"· {up['name']}{stage} — {up['desc']} [기억 {price}]")
            buttons.append(up["name"])
    buttons = buttons[:9] + ["새 게임"]
    if len(buttons) < 10:
        buttons.append("기록")
    return _reply([pre, "\n".join(lines)], buttons)


def record_screen(state):
    meta = state["meta"]
    endings = [e for e in DATA["order"] if DATA["events"][e]["type"] == "엔딩"]
    lines = ["— 기록 —", "",
             f"플레이 횟수: {meta['runs']}판",
             f"쓰러진 횟수: {meta['deaths']}번",
             f"가장 멀리 간 곳: {meta['best']}장" if meta["best"] else "가장 멀리 간 곳: 아직 없음",
             f"보유한 검의 기억: {meta['memory']}", "",
             f"본 엔딩 ({len(meta['endings'])}/{len(endings)})"]
    for e in endings:
        seen = e in meta["endings"]
        lines.append(("✔ " + DATA["events"][e]["title"]) if seen else "· ???")
    buttons = ["이어하기"] if state["run"] else ["새 게임", "광장"]
    return _reply(["\n".join(lines)], buttons)


def buy(state, up):
    meta = state["meta"]
    price = next_price(meta, up)
    if price is None:
        return hub_screen(state, f"'{up['name']}'은(는) 이미 끝까지 강화했다.")
    if meta["memory"] < price:
        return hub_screen(state, f"검의 기억이 부족하다. ({meta['memory']}/{price})")
    meta["memory"] -= price
    meta["upgrades"][up["key"]] = level(meta, up["key"]) + 1
    return hub_screen(state, f"'{up['name']}'을(를) 얻었다. (검의 기억 -{price})")


def begin_run(state, attr, rng):
    meta = state["meta"]
    meta["runs"] += 1
    state["run"] = new_run(meta, attr)
    state["mode"] = None
    r = state["run"]
    advance(r, meta, "자동", rng)
    notes = []
    if r["abilities"]:
        notes.append(f"선대 용사의 유품에서 '{', '.join(r['abilities'])}'을(를) 물려받았다.")
    if attr:
        notes.append(f"성검에 {attr}의 각인이 깨어 있다.")
    return event_screen(state, "\n".join(notes) or None)


def start_game(state, rng):
    attrs = owned_attrs(state["meta"])
    if attrs:
        state["run"] = None
        state["mode"] = "pick_attr"
        return _reply(["성검에 새겨 둔 각인이 있다. 어떤 각인을 깨우고 떠날까?"],
                      [f"{a}으로 시작" for a in attrs] + [NO_ATTR])
    return begin_run(state, None, rng)


# ------------------------------------------------------------------ 입력 처리
def handle(state, utterance, rng=random):
    """유저 발화 하나를 처리한다. state는 제자리에서 바뀐다."""
    ensure(state)
    meta, text = state["meta"], (utterance or "").strip()

    # 각인 고르는 중
    if state.get("mode") == "pick_attr":
        attrs = owned_attrs(meta)
        for a in attrs:
            if text == f"{a}으로 시작":
                return begin_run(state, a, rng)
        if text == NO_ATTR:
            return begin_run(state, None, rng)
        if text not in CMD_HUB + CMD_RECORD:
            return _reply(["아래 버튼 중에서 골라 주세요."],
                          [f"{a}으로 시작" for a in attrs] + [NO_ATTR])
        state["mode"] = None

    if text in CMD_RECORD:
        return record_screen(state)

    r = state["run"]
    if r is not None:
        r.setdefault("status", [])
        r.setdefault("parent", None)

    # 판이 진행 중이 아닐 때: 광장
    if r is None:
        if text in CMD_NEW or text == CMD_FORFEIT:
            return start_game(state, rng)
        if text in CMD_HUB:
            return hub_screen(state)
        for up in DATA["upgrades"]:
            if text == up["name"]:
                return buy(state, up)
        if meta["runs"] == 0:
            return intro_screen()
        return hub_screen(state)

    # 판 진행 중
    if text == CMD_FORFEIT:
        state["run"] = None
        return start_game(state, rng)
    if text in CMD_NEW:
        return _reply(["진행 중인 모험이 있다. 포기하면 이번 판의 검의 기억은 얻지 못한다."],
                      [CMD_FORFEIT, "이어하기"])
    if text in CMD_HUB:
        return event_screen(state, "광장은 모험이 끝난 뒤에 들를 수 있다.")
    if text in CMD_STATUS:
        scr = event_screen(state)
        return _reply([status_with_meta(state)], scr["buttons"], scr["card"])
    if text in CMD_POTION:
        if POTION in r["items"] and r["hp"] < r["hp_max"]:
            r["items"].remove(POTION)
            before = r["hp"]
            r["hp"] = min(r["hp_max"], r["hp"] + POTION_HEAL)
            msg = f"회복 물약을 마셨다. (체력 +{r['hp'] - before})"
            if settle_blood(r):
                msg += f"\n\n몸이 온전해지자 끓던 피가 식었다. ({BLOOD}가 사라졌다)"
            return event_screen(state, msg)
        return event_screen(state, "지금은 물약을 쓸 수 없다.")

    ev = DATA["events"][r["current"]]
    visible = visible_choices(r, ev)
    choice = next((c for c in visible if c["label"] == text), None)
    if choice is None and not visible and text == CONTINUE:
        # 고를 수 있는 선택지가 하나도 없을 때의 안전장치: 그냥 다음으로 넘어간다
        choice = {"check": None, "text": "", "hp": 0, "gold": 0, "power": 0, "gain": None, "next": "자동"}
    if choice is None:
        pre = None if text in CMD_CONTINUE else "카드의 버튼 중에서 골라 주세요."
        return event_screen(state, pre)

    notes = []
    had_blood = BLOOD in r["status"]
    if roll_check(r, choice["check"], rng):
        result, target = choice["text"], choice["next"]
        apply_delta(r, choice["hp"], choice["gold"], choice["power"], notes)
        apply_gain(r, choice["gain"], notes, rng)
    else:
        result, target = choice["fail_text"] or choice["text"], choice["fail_next"]
        apply_delta(r, choice["fail_hp"], choice["fail_gold"], 0, notes)
        apply_gain(r, choice["fail_gain"], notes, rng)
    pre = result + (f"\n\n({', '.join(notes)})" if notes else "")

    if r["hp"] <= 0:
        return death_screen(state, pre, rng)

    # 연계 이벤트로 넘어가거나 되돌아오는 것은 '이벤트를 넘긴 것'으로 치지 않는다
    tgt = DATA["events"].get(target)
    if tgt and tgt["type"] == "연계":
        if ev["type"] != "연계":
            r["parent"] = ev["id"]
        r["current"] = target
        return event_screen(state, pre)
    if ev["type"] == "연계" and r.get("parent") and target == r["parent"]:
        r["current"], r["parent"] = r["parent"], None
        return event_screen(state, pre)
    if ev["type"] == "연계" and r.get("parent"):
        ev = DATA["events"][r["parent"]]   # 연계 이벤트에서 결판이 나면 원래 이벤트를 넘긴 것으로 친다
    r["parent"] = None

    # 마족의 피: 새로 걸렸으면 설명을 붙이고, 이미 걸려 있었으면 체력이 조금씩 줄어든다
    if BLOOD in r["status"] and not had_blood:
        pre += f"\n\n쓰러뜨린 것의 피가 입에 튀었다. 삼키고 말았다.\n[{BLOOD}]\n{blood_help()}"
    elif BLOOD in r["status"] and had_blood:
        drain = min(DATA["params"].get("blood_drain", 5), r["hp"] - 1)
        if drain > 0:
            r["hp"] -= drain
            pre += f"\n\n피가 끓는다. (체력 -{drain})"
    if settle_blood(r):
        pre += f"\n\n몸이 온전해지자 끓던 피가 식었다. ({BLOOD}가 사라졌다)"

    # 마왕령의 독기: 마왕령에서는 쉼터를 빼고 이벤트를 넘길 때마다 체력이 줄어든다 (죽을 수 있다)
    if ev["chapter"] == DATA["params"].get("miasma_chapter", 0) and ev["type"] != "쉼터":
        dmg = miasma_damage(meta)
        if dmg > 0:
            r["hp"] = max(0, r["hp"] - dmg)
            pre += f"\n\n{MIASMA}가 폐를 태운다. (체력 -{dmg})"
            if r["hp"] <= 0:
                return death_screen(state, pre, rng)

    if ev["type"] == "보스" and ev["chapter"] == KING_CHAPTER:
        r["kings"] += 1

    before_chapter = r["chapter"]
    advance(r, meta, target, rng)
    if r["current"] is None:
        return death_screen(state, pre, rng)  # 이벤트가 더 없을 때의 안전장치
    mc = DATA["params"].get("miasma_chapter", 0)
    if r["chapter"] == mc and before_chapter != mc and miasma_damage(meta) > 0:
        pre += (f"\n\n[{MIASMA}]\n여기서부터는 사람이 오래 버틸 수 없는 땅이다.\n"
                f"· 이벤트를 넘길 때마다 체력 -{miasma_damage(meta)}\n· 쉼터에서는 줄지 않는다")
    if DATA["events"][r["current"]]["type"] == "엔딩":
        return ending_screen(state, pre, rng)
    return event_screen(state, pre)

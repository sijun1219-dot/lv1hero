"""레벨 1 용사 — 카카오톡 챗봇 스킬 서버.

카카오 챗봇 관리자센터의 스킬 URL에  https://<서버 주소>/skill  을 등록한다.
유저 진행 상황은 Supabase의 players 테이블에 저장한다.
(SUPABASE_URL / SUPABASE_KEY 환경 변수가 없으면 서버 메모리에만 저장한다. 테스트용.)
"""
import logging
import os

import httpx
from fastapi import FastAPI, Request
from fastapi.concurrency import run_in_threadpool

import game

logging.basicConfig(level=logging.INFO)
log = logging.getLogger("level1-hero")

SUPABASE_URL = os.environ.get("SUPABASE_URL", "").rstrip("/")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY", "")
USE_DB = bool(SUPABASE_URL and SUPABASE_KEY)

app = FastAPI()
_memory = {}
_http = httpx.Client(timeout=3.0)


def _headers(extra=None):
    h = {"apikey": SUPABASE_KEY, "Authorization": f"Bearer {SUPABASE_KEY}",
         "Content-Type": "application/json"}
    h.update(extra or {})
    return h


def load_state(user_id):
    if not USE_DB:
        return dict(_memory.get(user_id, {}))
    r = _http.get(f"{SUPABASE_URL}/rest/v1/players",
                  params={"user_id": f"eq.{user_id}", "select": "state"}, headers=_headers())
    r.raise_for_status()
    rows = r.json()
    return rows[0]["state"] if rows else {}


def save_state(user_id, state):
    if not USE_DB:
        _memory[user_id] = dict(state)
        return
    r = _http.post(f"{SUPABASE_URL}/rest/v1/players",
                   params={"on_conflict": "user_id"},
                   headers=_headers({"Prefer": "resolution=merge-duplicates,return=minimal"}),
                   json={"user_id": user_id, "state": state})
    r.raise_for_status()


def kakao_response(reply):
    """게임 엔진의 결과를 카카오 스킬 응답(JSON)으로 바꾼다.
    말풍선은 최대 3개, 한 개는 1000자까지, 버튼은 최대 10개·14자까지."""
    texts = [t[:1000] for t in reply["texts"]][-3:] or ["…"]
    return {
        "version": "2.0",
        "template": {
            "outputs": [{"simpleText": {"text": t}} for t in texts],
            "quickReplies": [
                {"label": b[:14], "action": "message", "messageText": b}
                for b in reply["buttons"][:10]
            ],
        },
    }


@app.get("/")
def health():
    return {"ok": True, "events": len(game.DATA["events"]), "db": USE_DB}


@app.post("/skill")
async def skill(request: Request):
    try:
        body = await request.json()
        user = body.get("userRequest", {}).get("user", {})
        user_id = user.get("id") or "anonymous"
        utterance = body.get("userRequest", {}).get("utterance", "")
        state = await run_in_threadpool(load_state, user_id)
        reply = game.handle(state, utterance)
        await run_in_threadpool(save_state, user_id, state)
        return kakao_response(reply)
    except Exception:
        log.exception("skill error")
        return kakao_response({
            "texts": ["잠시 문제가 생겼어요. 조금 뒤에 '이어하기'를 눌러 주세요."],
            "buttons": ["이어하기", "새 게임"],
        })

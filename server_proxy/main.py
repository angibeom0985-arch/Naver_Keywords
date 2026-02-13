import os
import re
import time
import hmac
import hashlib
import base64
from datetime import datetime, timedelta

import requests
from fastapi import FastAPI, Header, HTTPException
from pydantic import BaseModel


SEARCHAD_BASE_URL = "https://api.searchad.naver.com"
SEARCHAD_KEYWORD_URI = "/keywordstool"
NAVER_BLOG_SEARCH_URL = "https://openapi.naver.com/v1/search/blog.json"
NAVER_DATALAB_SEARCH_URL = "https://openapi.naver.com/v1/datalab/search"


def _required_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(f"Missing environment variable: {name}")
    return value


SEARCHAD_ACCESS_KEY = _required_env("SEARCHAD_ACCESS_KEY")
SEARCHAD_SECRET_KEY = _required_env("SEARCHAD_SECRET_KEY")
SEARCHAD_CUSTOMER_ID = _required_env("SEARCHAD_CUSTOMER_ID")
NAVER_CLIENT_ID = _required_env("NAVER_CLIENT_ID")
NAVER_CLIENT_SECRET = _required_env("NAVER_CLIENT_SECRET")
PROXY_TOKEN = os.getenv("PROXY_TOKEN", "").strip()


app = FastAPI(title="Naver Keyword Proxy", version="1.0.0")


class KeywordRequest(BaseModel):
    keyword: str
    machine_id: str | None = None


def _verify_token(token: str | None):
    if PROXY_TOKEN and (token or "").strip() != PROXY_TOKEN:
        raise HTTPException(status_code=401, detail="Invalid proxy token")


def _parse_count(value):
    if value is None:
        return 0
    if isinstance(value, int):
        return value
    value_text = str(value).replace(",", "").strip()
    if value_text.startswith("<"):
        digits = re.sub(r"[^0-9]", "", value_text)
        return int(digits) if digits else 0
    digits = re.sub(r"[^0-9]", "", value_text)
    return int(digits) if digits else 0


def _searchad_signature(method: str, uri: str, timestamp: str) -> str:
    message = f"{timestamp}.{method}.{uri}"
    digest = hmac.new(
        SEARCHAD_SECRET_KEY.encode("utf-8"),
        message.encode("utf-8"),
        hashlib.sha256
    ).digest()
    return base64.b64encode(digest).decode("utf-8")


def _post_datalab(payload: dict):
    headers = {
        "X-Naver-Client-Id": NAVER_CLIENT_ID,
        "X-Naver-Client-Secret": NAVER_CLIENT_SECRET
    }
    resp = requests.post(NAVER_DATALAB_SEARCH_URL, headers=headers, json=payload, timeout=15)
    if resp.status_code != 200:
        raise HTTPException(status_code=502, detail=f"Datalab failed: {resp.text[:220]}")
    return resp.json()


@app.get("/health")
def health():
    return {"ok": True, "ts": int(time.time())}


@app.post("/related-keywords")
def related_keywords(req: KeywordRequest, x_proxy_token: str | None = Header(default=None)):
    _verify_token(x_proxy_token)
    keyword = req.keyword.strip()
    if not keyword:
        raise HTTPException(status_code=400, detail="keyword is required")

    timestamp = str(int(time.time() * 1000))
    headers = {
        "X-Timestamp": timestamp,
        "X-API-KEY": SEARCHAD_ACCESS_KEY,
        "X-Customer": SEARCHAD_CUSTOMER_ID,
        "X-Signature": _searchad_signature("GET", SEARCHAD_KEYWORD_URI, timestamp)
    }
    params = {"hintKeywords": keyword, "showDetail": 1}
    resp = requests.get(f"{SEARCHAD_BASE_URL}{SEARCHAD_KEYWORD_URI}", headers=headers, params=params, timeout=20)
    if resp.status_code != 200:
        raise HTTPException(status_code=502, detail=f"SearchAd failed: {resp.text[:220]}")

    rows = []
    for item in resp.json().get("keywordList", []):
        rel = str(item.get("relKeyword", "")).strip()
        if not rel:
            continue
        pc = _parse_count(item.get("monthlyPcQcCnt"))
        mob = _parse_count(item.get("monthlyMobileQcCnt"))
        rows.append({
            "keyword": rel,
            "monthly_pc_search": pc,
            "monthly_mobile_search": mob,
            "monthly_total_search": pc + mob
        })
    return {"ok": True, "rows": rows}


@app.post("/blog-count")
def blog_count(req: KeywordRequest, x_proxy_token: str | None = Header(default=None)):
    _verify_token(x_proxy_token)
    keyword = req.keyword.strip()
    if not keyword:
        raise HTTPException(status_code=400, detail="keyword is required")

    headers = {
        "X-Naver-Client-Id": NAVER_CLIENT_ID,
        "X-Naver-Client-Secret": NAVER_CLIENT_SECRET
    }
    params = {"query": keyword, "display": 1, "start": 1, "sort": "sim"}
    resp = requests.get(NAVER_BLOG_SEARCH_URL, headers=headers, params=params, timeout=12)
    if resp.status_code != 200:
        raise HTTPException(status_code=502, detail=f"Blog API failed: {resp.text[:220]}")
    total = int(resp.json().get("total", 0))
    return {"ok": True, "blog_document_count": total}


@app.post("/keyword-insight")
def keyword_insight(req: KeywordRequest, x_proxy_token: str | None = Header(default=None)):
    _verify_token(x_proxy_token)
    key = req.keyword.strip()
    if not key:
        raise HTTPException(status_code=400, detail="keyword is required")

    today = datetime.now().date()
    start_month = (today.replace(day=1) - timedelta(days=330)).replace(day=1)
    monthly_payload = {
        "startDate": start_month.strftime("%Y-%m-%d"),
        "endDate": today.strftime("%Y-%m-%d"),
        "timeUnit": "month",
        "keywordGroups": [{"groupName": key, "keywords": [key]}]
    }
    monthly_data = _post_datalab(monthly_payload).get("results", [{}])[0].get("data", [])
    month_sum = sum(float(x.get("ratio", 0.0)) for x in monthly_data) or 1.0
    month_ratio = []
    for row in monthly_data:
        period_raw = str(row.get("period", ""))[:7]
        value = float(row.get("ratio", 0.0))
        try:
            label = f"{int(period_raw.split('-')[1])}월"
        except Exception:
            label = period_raw
        month_ratio.append({"label": label, "value": (value / month_sum) * 100.0})

    daily_payload = {
        "startDate": (today - timedelta(days=89)).strftime("%Y-%m-%d"),
        "endDate": today.strftime("%Y-%m-%d"),
        "timeUnit": "date",
        "keywordGroups": [{"groupName": key, "keywords": [key]}]
    }
    daily_data = _post_datalab(daily_payload).get("results", [{}])[0].get("data", [])
    names = ["월", "화", "수", "목", "금", "토", "일"]
    wd_sum_map = {name: 0.0 for name in names}
    for row in daily_data:
        p = str(row.get("period", ""))
        try:
            wd = datetime.strptime(p, "%Y-%m-%d").weekday()
            wd_sum_map[names[wd]] += float(row.get("ratio", 0.0))
        except Exception:
            continue
    wd_total = sum(wd_sum_map.values()) or 1.0
    weekday_ratio = [{"label": k, "value": (v / wd_total) * 100.0} for k, v in wd_sum_map.items()]

    age_groups = [("10대", "1"), ("20대", "2"), ("30대", "3"), ("40대", "4"), ("50대 이상", "5")]
    age_values = []
    for _, age_code in age_groups:
        age_payload = {
            "startDate": (today - timedelta(days=89)).strftime("%Y-%m-%d"),
            "endDate": today.strftime("%Y-%m-%d"),
            "timeUnit": "date",
            "ages": [age_code],
            "keywordGroups": [{"groupName": key, "keywords": [key]}]
        }
        age_data = _post_datalab(age_payload).get("results", [{}])[0].get("data", [])
        age_total = sum(float(item.get("ratio", 0.0)) for item in age_data)
        age_values.append(age_total)
    age_sum = sum(age_values) or 1.0
    age_ratio = [{"label": age_groups[i][0], "value": (age_values[i] / age_sum) * 100.0} for i in range(len(age_groups))]

    return {"ok": True, "month_ratio": month_ratio, "weekday_ratio": weekday_ratio, "age_ratio": age_ratio}

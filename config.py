# =============================================
# API 키 / 설정 — 비밀값(secrets)에서 읽어온다
# =============================================
# 키는 코드에 적지 않는다. 다음 우선순위로 읽는다:
#   1) 환경변수 (예: CLI 실행 시  CLOVA_SECRET_KEY=... python translate.py)
#   2) Streamlit secrets (.streamlit/secrets.toml 또는 Streamlit Cloud > Settings > Secrets)
#
# 로컬에서 웹앱을 돌릴 때는 .streamlit/secrets.toml 에 키를 넣으면 된다.
# (secrets.toml.example 참고 — 실제 secrets.toml 은 .gitignore 로 커밋 제외)

import os

try:
    import streamlit as st
    _sec = dict(st.secrets)
except Exception:
    _sec = {}


def _get(key: str, default=None):
    """환경변수 → streamlit secrets 순으로 값을 찾는다."""
    val = os.environ.get(key)
    if val:
        return val
    return _sec.get(key, default)


# Naver CLOVA OCR
# console.ncloud.com → CLOVA OCR → 도메인 상세 → Secret Key / API URL
CLOVA_SECRET_KEY = _get("CLOVA_SECRET_KEY")
CLOVA_API_URL = _get("CLOVA_API_URL")

# DeepL API Free
# deepl.com → 계정 → Authentication Key for DeepL API
DEEPL_API_KEY = _get("DEEPL_API_KEY")

# 번역할 언어 목록 (DeepL 언어 코드) — CLI 배치 모드 기본값.
# 웹앱(app.py)에서는 요청마다 선택 언어를 함수 인자로 직접 전달한다.
# JA=일본어, ZH=중국어간체, ZH-HANT=중국어번체, EN-US=영어, DE=독일어, FR=프랑스어, ES=스페인어
TARGET_LANGS = ['JA']

#!/usr/bin/env python3
"""
다국어 상세페이지 번역 도구 — 웹앱 (Streamlit)

- 제품 한 개의 상세페이지 이미지들을 업로드하면
- CLOVA OCR로 한국어를 추출하고 DeepL로 선택 언어로 번역해
- 번역표 엑셀 + 번호 주석 이미지를 ZIP으로 다운로드한다.

로컬 실행:  streamlit run app.py
배포:       Streamlit Community Cloud (Settings > Secrets 에 API 키 입력)
"""

import io
import zipfile

import streamlit as st
from PIL import Image

import config
import translate
from translate import LANG_NAMES, filename_sort_key, run_pipeline

st.set_page_config(page_title="상세페이지 번역 도구", page_icon="🌐", layout="centered")

st.title("🌐 상세페이지 번역 도구")
st.caption("이미지 업로드 → 한국어 추출(CLOVA OCR) → 번역(DeepL) → 번역표 엑셀 다운로드")

# ── 로그인 / 도메인 제한 ─────────────────────
# 회사 구글 계정(@boosters.kr)만 사용 가능. Google OAuth([auth] secrets)가
# 설정돼 있을 때만 강제하며, 미설정 시에는 경고만 띄우고 동작(설정 전 임시).
ALLOWED_DOMAIN = "@boosters.kr"


def _auth_configured() -> bool:
    try:
        return "auth" in st.secrets
    except Exception:
        return False


if _auth_configured():
    if not st.user.is_logged_in:
        st.info("이 도구는 회사 구글 계정으로 로그인해야 사용할 수 있습니다.")
        st.button("🔐 Google로 로그인", on_click=st.login, type="primary")
        st.stop()

    email = (getattr(st.user, "email", "") or "").lower()
    verified = getattr(st.user, "email_verified", True)
    if not (email.endswith(ALLOWED_DOMAIN) and verified):
        st.error(
            f"`{email or '알 수 없음'}` 계정은 사용 권한이 없습니다.\n\n"
            f"**{ALLOWED_DOMAIN}** 로 끝나는 회사 구글 계정으로 로그인하세요."
        )
        st.button("다른 계정으로 로그인", on_click=st.logout)
        st.stop()

    with st.sidebar:
        st.write(f"👤 {email}")
        st.button("로그아웃", on_click=st.logout)
else:
    st.warning("⚠️ 로그인 보호가 아직 설정되지 않았습니다. (관리자가 Google OAuth 설정 후 활성화됩니다)")

# ── API 키 확인 ──────────────────────────────
if not (config.CLOVA_SECRET_KEY and config.CLOVA_API_URL and config.DEEPL_API_KEY):
    st.error(
        "API 키가 설정되지 않았습니다.\n\n"
        "- 로컬: `.streamlit/secrets.toml` 에 키 3개를 입력하세요 (`secrets.toml.example` 참고).\n"
        "- 배포: Streamlit Cloud **Settings → Secrets** 에 키를 입력하세요."
    )
    st.stop()

# ── 입력 폼 ──────────────────────────────────
st.subheader("1. 이미지 업로드")
st.write("제품 **한 개**의 상세페이지 이미지들을 한 번에 올려주세요. (파일명 숫자 순서로 위→아래 정렬됩니다)")

uploaded = st.file_uploader(
    "이미지 파일 선택 (여러 장 가능)",
    type=["jpg", "jpeg", "png", "tif", "tiff"],
    accept_multiple_files=True,
)

# 제품명 기본값: 첫 파일명에서 숫자/확장자 제외한 부분, 없으면 '제품'
default_name = "제품"
if uploaded:
    stem = uploaded[0].name.rsplit(".", 1)[0]
    cleaned = "".join(ch for ch in stem if not ch.isdigit()).strip("_-· ")
    default_name = cleaned or "제품"

st.subheader("2. 옵션")
product_name = st.text_input("제품명 (다운로드 파일 이름에 사용)", value=default_name)

lang_codes = list(LANG_NAMES.keys())
selected_langs = st.multiselect(
    "번역 언어 (여러 개 선택 가능)",
    options=lang_codes,
    default=["JA"],
    format_func=lambda code: f"{LANG_NAMES[code]} ({code})",
)

# ── 실행 ─────────────────────────────────────
st.subheader("3. 번역 실행")
run = st.button("번역 시작", type="primary", disabled=not (uploaded and selected_langs))

if run:
    if not uploaded:
        st.warning("이미지를 먼저 업로드해주세요.")
        st.stop()
    if not selected_langs:
        st.warning("번역 언어를 하나 이상 선택해주세요.")
        st.stop()

    images = [(f.name, Image.open(f)) for f in uploaded]
    images.sort(key=lambda t: filename_sort_key(t[0]))

    with st.status("처리 중...", expanded=True) as status:
        def progress(msg):
            st.write(msg)

        try:
            result = run_pipeline(images, product_name, selected_langs, progress=progress)
        except Exception as e:
            status.update(label="오류 발생", state="error")
            st.exception(e)
            st.stop()

        if not result:
            status.update(label="번역할 한국어 텍스트를 찾지 못했습니다.", state="error")
            st.stop()

        status.update(label=f"완료! 번역 항목 {result['count']}개", state="complete")

    # ZIP 묶기 (엑셀 + 주석 이미지)
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(f"{product_name}_translation.xlsx", result["excel"])
        zf.writestr(f"{product_name}_annotated.jpg", result["annotated"])

    st.success(f"번역 완료 — 총 {result['count']}개 항목, {len(selected_langs)}개 언어")
    st.download_button(
        "📥 결과 ZIP 다운로드 (엑셀 + 주석이미지)",
        data=zip_buf.getvalue(),
        file_name=f"{product_name}_번역.zip",
        mime="application/zip",
        type="primary",
    )

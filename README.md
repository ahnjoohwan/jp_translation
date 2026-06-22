# 상세페이지 번역 도구

제품 상세페이지 이미지에서 한국어를 추출(Naver CLOVA OCR)해 선택한 언어로
번역(DeepL)하고, **번역표 엑셀 + 번호 주석 이미지**를 만들어주는 웹 도구입니다.

- 입력: 제품 한 개의 상세페이지 이미지 여러 장 (웹 업로드)
- 출력: `엑셀(.xlsx) + 주석이미지(.jpg)` 를 묶은 ZIP 다운로드
- 번역 언어: 웹에서 선택 (일본어·중국어·영어 등 DeepL 지원 언어)

## 사용하는 외부 API

| API | 용도 | 비고 |
|---|---|---|
| Naver CLOVA OCR | 이미지 → 한국어 텍스트·좌표 추출 | 유료 (NCloud 종량제) |
| DeepL API Free | 한국어 → 선택 언어 번역 | 무료 플랜 월 50만 자 제한 |

## API 키 설정 (비밀값)

키는 코드에 적지 않고 비밀값으로 관리합니다.

**로컬에서 실행할 때**

`secrets.toml.example` 을 복사해 `.streamlit/secrets.toml` 을 만들고 키를 채웁니다.

```toml
CLOVA_SECRET_KEY = "..."
CLOVA_API_URL = "..."
DEEPL_API_KEY = "..."
```

`.streamlit/secrets.toml` 은 `.gitignore` 에 의해 커밋되지 않습니다.

## 로컬 실행

```bash
pip install -r requirements.txt
streamlit run app.py
```

브라우저가 열리면 이미지를 업로드하고 언어를 선택한 뒤 **번역 시작** → 완료되면
**결과 ZIP 다운로드**.

## 팀 배포 (Streamlit Community Cloud)

1. 이 폴더를 GitHub 저장소에 올립니다. (`.gitignore` 가 키·이미지를 제외하는지 확인)
2. https://share.streamlit.io 접속 → GitHub 연결 → 저장소와 `app.py` 선택 → **Deploy**.
3. 배포된 앱의 **Settings → Secrets** 에 위 toml 내용을 그대로 붙여넣습니다.
4. 회사 자료라면 앱을 **Private** 로 두고 열람 가능한 계정/이메일을 지정합니다.
   (무료 앱은 기본적으로 URL 공개)
5. 팀에 앱 URL 을 공유합니다.

## CLI 모드 (선택)

웹 없이 폴더 단위로 돌리던 기존 방식도 유지됩니다. 환경변수로 키를 지정하고
`input/<제품명>/` 구조를 만든 뒤:

```bash
CLOVA_SECRET_KEY=... CLOVA_API_URL=... DEEPL_API_KEY=... python translate.py
```

번역 언어는 `config.py` 의 `TARGET_LANGS` 로 지정합니다.

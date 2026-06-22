# jp_translation — 작업 가이드

상위 [`../CLAUDE.md`](../CLAUDE.md)의 규칙을 따른다.
특히 **Document Lake 개선 기록 하네스**:

이 프로젝트에서 의미있는 개선·변경·결정을 하면 `_records/`에 기록 md를 남긴다.
저장하면 PostToolUse 훅이 자동으로 옵시디언 vault `raw/`에 미러한다
(파일명 `{project}__{title}.md`).

```yaml
---
lake: true
project: jp_translation
title: <무엇을 개선했는지>
date: <YYYY-MM-DD>
summary: <한 줄 요약>
---
## 변경점
## 이유·맥락
## 영향 / 후속
```

- 새 주제 → 새 기록 md(insert), 같은 주제 연장 → 기존 기록 수정(update). 프로젝트 기록 ↔ raw는 1:1.
- 일괄/수동 동기화: `python3 ../tools/sync_md.py --all` (dry-run `--check`).
- 자세한 규칙·동작은 루트 `CLAUDE.md`의 "Document Lake" 섹션 참고.

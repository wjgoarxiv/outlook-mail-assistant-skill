![Outlook Mail Assistant cover](cover.png)

# outlook-mail-assistant

*Windows 데스크톱의 로컬 Outlook 메일 작업을 다루는 LLM 에이전트용 공개 skill directory입니다.*

<p align="center">
  <a href="#이-저장소의-성격">이 저장소의 성격</a> ·
  <a href="#skill-범위">Skill 범위</a> ·
  <a href="#저장소-구성">저장소 구성</a> ·
  <a href="#로컬-검증">로컬 검증</a> ·
  <a href="#안전-경계">안전 경계</a>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/type-LLM%20Agent%20Skill-111827?style=flat-square" alt="LLM Agent Skill" />
  <img src="https://img.shields.io/badge/platform-Windows-0078D4?style=flat-square&logo=windows" alt="Windows" />
  <img src="https://img.shields.io/badge/outlook-desktop%20COM-0A64AD?style=flat-square&logo=microsoftoutlook" alt="Outlook Desktop COM" />
  <img src="https://img.shields.io/badge/storage-JSONL%20%2B%20SQLite-2E7D32?style=flat-square" alt="JSONL and SQLite" />
</p>

---

## 이 저장소의 성격

이 저장소는 일반 사용자용 메일 애플리케이션이 아닙니다. Outlook 메일을 로컬 우선 방식으로 수집·정규화·추출·보고·후속조치할 수 있도록 돕는 **LLM 에이전트용 skill directory**입니다.

실제 skill 계약은 `SKILL.md`에 있고, `scripts/` 아래의 Python 코드는 에이전트가 검증이나 실행 단계에서 재사용할 수 있는 로컬 실행면을 제공합니다.

영문 문서는 [README.md](README.md)를 참고하세요.

## Skill 범위

이 skill은 다음과 같은 에이전트 워크플로를 대상으로 합니다.

- live Outlook desktop profile에서 메일 수집
- `.msg` 파일 및 `.pst` 아카이브 가져오기
- canonical schema로 메일 레코드 정규화
- 작업, 마감, 회의, 의사결정, 후속조치 추출
- Markdown, CSV, XLSX, DOCX 검토 산출물 생성
- dry-run 또는 확인된 Outlook 후속 액션 실행

공개 에이전트 엔트리 메타데이터는 `agents/openai.yaml`에 있습니다.

## 저장소 구성

```text
.
├─ SKILL.md
├─ README.md
├─ README.ko-KR.md
├─ agents/
│  └─ openai.yaml
├─ references/
│  ├─ canonical-schema.md
│  ├─ sqlite-schema.md
│  ├─ msg-gotchas.md
│  └─ pst-gotchas.md
├─ scripts/
│  ├─ export_outlook_json.py
│  ├─ import_pst.py
│  ├─ export_task_reports.py
│  ├─ execute_outlook_actions.py
│  ├─ convert_md_to_docx.py
│  └─ outlook_mail_assistant/
├─ tests/
├─ pyproject.toml
└─ .github/workflows/ci.yml
```

## 로컬 검증

editable 설치:

```powershell
python -m pip install -e .[dev]
```

선택적 런타임 백엔드:

- live Outlook desktop access: `python -m pip install pywin32`
- `.msg` 파싱: `python -m pip install extract-msg`
- `.pst` 파싱: `references/pst-gotchas.md` 기준으로 백엔드 선택

테스트 실행:

```powershell
python -m pytest -q tests
```

커버 이미지 재생성:

```powershell
python generate_cover.py
```

내장 DOCX 변환기 검증:

```powershell
python scripts/convert_md_to_docx.py README.md output.docx
```

## 안전 경계

이 저장소에는 **skill 정의와 보조 코드만** 있어야 하며, 실제 메일 데이터는 포함하면 안 됩니다.

- runtime workspace는 반드시 repo 밖에 둘 것
- 생성된 JSONL, SQLite, CSV, XLSX, DOCX, audit 출력은 민감 메일 산출물로 취급할 것
- 메일함/캘린더 변경 전에는 dry-run 경로를 우선 사용할 것
- 선택적 파서를 활성화하기 전 `references/`의 의존성/라이선스 제약을 검토할 것

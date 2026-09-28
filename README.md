# create-meeting-minutes

회의 중 나눈 대화 텍스트(STT 변환본, 채팅 로그, 메모 등)를 입력받아 표준 회의록 양식(`.docx`)에
맞춰 채워진 문서를 자동 생성하는 [Claude Code](https://claude.com/claude-code) 스킬입니다.

회의 내용을 분석·요약하는 것은 Claude가, 그 결과를 양식에 정확히 끼워 넣는 것은
`scripts/build_meeting_minutes.py`가 담당합니다. 폰트 크기, 굵기, 병합된 셀, 로마 숫자 목록
번호 매기기 등 문서 서식은 실제 회의록 파일들의 서식을 그대로 재현하도록 검증되어 있습니다.

## 설치

이 저장소를 클론한 뒤, 스킬 폴더를 회의록을 모아둘 프로젝트 루트의 `.claude/skills/`에 복사합니다.

```bash
git clone https://github.com/eunchae34/create-meeting-minutes.git
cp -r create-meeting-minutes <회의록-프로젝트-루트>/.claude/skills/create-meeting-minutes
```

`python-docx`가 필요합니다.

```bash
pip install python-docx
```

> **주의:** `--output`을 생략하면 스크립트는 자기 자신의 경로 기준(`<프로젝트루트>/.claude/skills/
> create-meeting-minutes/scripts/build_meeting_minutes.py`에서 4단계 위)으로 저장 위치를 추정합니다.
> 위 구조 그대로 프로젝트 단위(`.claude/skills/`)에 설치했다면 프로젝트 루트에 저장되지만, `~/.claude/
> skills/`처럼 사용자 전역 위치에 설치하면 홈 디렉터리에 저장되므로 이 경우엔 항상 `--output`을
> 명시적으로 지정하세요.

## 사용법

Claude Code 세션에서 회의 내용(녹취록, 대화 로그 등)을 붙여넣고 "회의록 작성해줘"라고 요청하면
자동으로 이 스킬이 트리거됩니다. 자세한 동작 방식과 데이터 스키마는 [SKILL.md](SKILL.md)를 참고하세요.

## 템플릿 커스터마이즈

`assets/meeting_template.docx`는 예시 회의록 양식입니다. 자신의 회사/팀 양식으로 바꾸려면 이
파일을 직접 쓰고 있는 실제 `.docx` 양식으로 교체하세요 — 표의 행/열 구조(회의 일시/부서/작성자/
회의 장소/참석자/회의 주제/회의 내용/To do list)만 유지되면 스크립트 로직은 그대로 동작합니다.

## 파일 구성

- `SKILL.md` — 스킬 정의 및 상세 동작 방식
- `scripts/build_meeting_minutes.py` — docx 생성 스크립트
- `scripts/sample_data.json` — 입력 데이터 스키마 예시
- `assets/meeting_template.docx` — 회의록 양식 예시

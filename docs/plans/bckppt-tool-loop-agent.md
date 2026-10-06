# BckPPT 도구 루프 에이전트와 일반 사용자 배포 계획

작성: 2026-10-06 · 상태: 진행 중(Phase A~E 구현, 카드 파이프라인·외부 작성) · 선행 문서: [독립형 PPT 에이전트 리팩토링 계획](bckppt-agent-first-refactoring.md)

## 1. 배경

2026-10-06 착수보고 PPT 작업에서 Claude Code가 BckPPT를 도구로 써서 다음을 했다.

- 기획보고서 5종(각 수백 쪽 Markdown)에서 개요·문제점·연구개발 내용을 뽑아 33쪽 카드 레이아웃을 복제한 2쪽에 채움
- 연계 수준 태그와 범례 추가, 1카드 쪽은 패널·띠 폭 축소
- 모든 텍스트 상자의 **실제 줄바꿈**을 측정하며 문구를 고쳐 "개발"만 넘어간 줄, 기준 줄 수 초과를 없앰
- 사용자 피드백("연구개발 내용 4줄 추가", "X 태그 삭제", "1줄씩 줄여줘")을 몇 분 안에 반영

결과 품질은 모델의 글솜씨보다 **측정 → 수정 → 재측정 → 렌더 확인** 반복에서 나왔다. BckPPT는
PowerPoint가 실제로 나눈 줄(`TextRange.Lines`)과 렌더 이미지를 돌려줄 수 있어 이 반복이 가능했다.
이 계획은 그 반복을 BckPPT 안에 넣어, Claude Code 없이 gitoky LLM만으로 일반 사용자가 같은 작업을 하게 만드는 것이다.

이번 작업 중 BckPPT에 반영한 기반(커밋 `feat: 도형 주소 규약과 배치·서식 API`):

- 도형 주소 규약 `이름`, `이름#n`, `그룹/이름`, `표 [RxCy]` (`Models/ShapeAddress.cs`)
- 그룹 안 도형 탐색: `GetTextLineBreaks`, `SetShapeText`, `GetShapeInfo`, `SetParagraphTexts`
- 새 API: `ListShapes`, `RenameShape`, `SetShapeGeometry`, `DeleteShape`, `AddTextShape`, `SetTextColor`, `SetTextFormat`
- 디버그 API `/api/shapes*`, `/api/line-breaks`
- `tmp/`, `artifacts/` 아래 obj가 섞여 빌드가 깨지던 문제 수정

## 2. 목표와 비목표

**목표**

1. 채팅창에서 "33쪽 레이아웃으로 이 보고서 5개를 카드로 정리해줘" 같은 요청을 받아, 사본에 적용하고 미리보기 후 승인받는다.
2. 결과가 레이아웃 규칙(줄 수, 고아 단어, overflow, 겹침)을 통과할 때까지 에이전트가 스스로 고친다.
3. 일반 사용자 PC(Windows + PowerPoint)에 설치해 쓸 수 있게 배포한다.

**비목표 (1차)**

- PowerPoint 없는 환경(웹, macOS) 지원
- 새 디자인 생성(빈 슬라이드에서 레이아웃 창작). 1차는 **기존 슬라이드를 복제해 채우는 작업**만 한다.
- 원본 파일 직접 덮어쓰기. 항상 사본에 적용한다(기존 `ApprovalWorkflow` 원칙 유지).

## 3. 현재 상태

| 영역 | 있는 것 | 부족한 것 |
| --- | --- | --- |
| LLM | `GitokyLlmClient`: OpenAI 호환 `/v1/chat/completions` 스트리밍, 재시도, `.env`의 `GITOKY_LLM_*` | `tools`(function calling), 이미지 입력, 비스트리밍 응답 |
| 에이전트 | `PptPromptBuilder` → `PptTaskPlan` 단발 JSON 계획, 명령 7종 | 결과를 보고 다시 계획하는 반복, 측정 도구 |
| 적용 | `ApprovalWorkflow`: 사본 생성, 사전상태 검사, 전후 썸네일 | 문단 단위 교체·도형 편집 명령 |
| 검증 | `RenderValidator`, `AutoFitAsync`, `GetTextLineBreaks` | 템플릿 대비 줄 수, 고아 단어 같은 규칙 판정 |
| 원문 | 문서 MD 내보내기(`ExportMarkdown`), 로컬 PPT RAG | PDF·장문 원문 검색 도구 |
| 도형 | 이번에 추가한 주소 규약·배치·서식 API | — |

## 4. 핵심 설계

### 4.1 도구 호출 루프

단발 계획 대신 **세션 안에서 도구를 여러 번 부르는 루프**를 만든다.

```text
사용자 요청
  → [AgentLoop] 사본 생성 (작업 파일)
  → 반복 (최대 N회, 토큰·시간 상한)
       LLM 호출 (대화 이력 + 도구 결과)
       ├─ 도구 호출 → ToolRegistry 실행 (작업 파일 대상) → 결과를 이력에 추가
       └─ finish(summary) → 루프 종료
  → LayoutLint 최종 판정 → 미리보기(전후 썸네일 + 변경 목록)
  → 사용자 승인 → 결과 사본 확정 / 거절 시 폐기
```

- 모든 쓰기 도구는 **작업 사본**에만 적용한다. 원본 경로를 도구 인자로 받지 않는다.
- 루프 상한: 기본 30턴, 도구 결과는 턴당 8,000자로 자른다. 상한 도달 시 지금까지 결과와 남은 문제를 보여준다.
- 사용자 후속 피드백("1줄씩 줄여줘")은 같은 세션·같은 작업 사본에서 루프를 이어간다.

### 4.2 도구 목록

`IPptComService`를 얇게 감싼 도구. 이름·인자는 JSON Schema로 정의해 LLM에 준다.

| 도구 | 대응 API | 비고 |
| --- | --- | --- |
| `list_slides` | `AnalyzePresentation` 요약 | 슬라이드 번호·제목 |
| `list_shapes(slide)` | `ListShapes` | 주소·위치·크기·텍스트 미리보기 |
| `read_text(slide, shape)` | `GetShapeInfo` | 문단 목록 포함 |
| `line_breaks(slide, shape)` | `GetTextLineBreaks` | **핵심 측정 도구** |
| `duplicate_slide(source, after)` | `DuplicateSlide` | |
| `set_paragraphs(slide, shape, paragraphs[])` | `SetParagraphTexts` | 문단 수 차이는 Insert/Delete로 맞춤. `\u000b`는 Shift+Enter |
| `replace_text(slide, shape, find, replace)` | `ReplaceText` | 런 서식 보존(제목처럼 글자 크기 섞인 곳) |
| `rename_shape`, `set_geometry`, `delete_shape` | 동명 API | |
| `add_text_shape`, `set_text_color`, `set_text_format` | 동명 API | 범례·배지 |
| `lint_slide(slide, baseline_slide?)` | `LayoutLint`(신규) | 4.3 참고 |
| `render_slide(slide)` | `GetSlideThumbnailAsync` | 이미지 입력 지원 시 모델에 전달 |
| `search_source(query)` / `read_source(doc, range)` | `SourceLibrary`(신규) | 4.4 참고 |
| `finish(summary)` | — | 루프 종료 |

`set_paragraphs`의 문단 수 맞춤 로직은 2026-10-06 프로토타입(스크래치 CLI)에서 검증했다.
같은 문단 수는 Replace, 늘면 앞에서부터 Insert, 줄면 뒤에서부터 Delete. 이를 `ParagraphEdit` 목록으로 바꾸는 순수 함수로 만들고 테스트한다.

### 4.3 LayoutLint: 결정적 레이아웃 판정

모델의 눈대중 대신 코드가 판정하고, 결과를 문장으로 모델에 돌려준다.

| 규칙 | 판정 | 예시 메시지 |
| --- | --- | --- |
| 기준 줄 수 | 복제 원본(baseline) 같은 역할 상자의 줄 수 이하 | `C2_rnd 12줄 > 기준 11줄` |
| 고아 단어 | 자동 줄바꿈된 문단의 마지막 줄이 4자 이하 | `C3_rnd 2문단 마지막 줄 '개발'만 남음` |
| overflow | `BoundHeight > Height + 허용오차`, 또는 상위 카드 아래 경계 초과 | `C4_rnd 카드 하단 3pt 초과` |
| 제목 균형 | 2줄 제목의 둘째 줄이 첫 줄의 30% 미만 | `C2_title 둘째 줄 '개발' — Shift+Enter로 균형` |
| 서식 이탈 | 같은 역할 상자끼리 글꼴·크기·자간 불일치 | `C1_overview 자간 -1.5 (다른 카드 0)` |
| 겹침 | 새로 추가한 도형이 기존 도형과 겹침 | `범례가 파란 띠와 12pt 겹침` |

baseline은 "복제 원본 슬라이드의 같은 주소 도형"으로 자동 연결한다(`duplicate_slide` 기록 이용).
고아 단어 기준(4자), 제목 비율(30%)은 `PptRules`에 설정으로 둔다.

### 4.4 원문 라이브러리

- 입력: pptx(`ExportMarkdown`), PDF(텍스트 레이어 추출 → Markdown), md/txt
- 쪽 단위로 쪼개 SQLite FTS5에 색인(기존 `TemplateRepository`와 같은 DB 방식)
- `search_source`는 상위 k개 조각(문서명, 쪽, 400자)을, `read_source`는 지정 쪽 범위 원문을 돌려준다.
- 이번 작업에서는 서브에이전트가 Grep으로 "요약문, 최종목표, 구성기술, RFP" 절을 찾아 읽었다. 같은 탐색 전략을 시스템 프롬프트에 적는다.

### 4.5 LLM 프로토콜

gitoky의 OpenAI 호환 API가 무엇을 지원하느냐에 따라 두 방식 중 하나를 쓴다. **Phase A에서 먼저 확인한다.**

1. **네이티브 function calling**: 요청에 `tools`, `tool_choice`를 넣고 응답의 `tool_calls`를 실행, 결과를 `role: tool` 메시지로 돌려준다.
2. **JSON 행동 프로토콜(대체안)**: 모델이 매 턴 `{"tool": "...", "args": {...}}` 또는 `{"finish": "..."}` 하나만 출력하게 하고, 기존 `PptTaskPlanDecoder`처럼 추출·검증한다.

`ILlmClient`에 비스트리밍 `CompleteAsync(LlmRequest)`와 `tools`·`tool_calls`를 담는 모델 확장을 추가한다.
모델은 `GITOKY_LLM_MODEL`(예: 품질 우선 라우트)로 바꿀 수 있게 유지하고, 루프용·요약용 모델을 따로 지정할 수 있게 한다.

## 5. 대표 시나리오: 템플릿 카드 채우기

입력: 대상 PPT, 템플릿 슬라이드 번호, 원문 문서들, 배치(쪽당 카드 수), 선택 사항(태그 기준).

1. `list_shapes(템플릿)`로 카드 구조를 파악하고, 이름이 겹치면 위치 기준으로 `C{열}_{역할}`로 이름을 바꿀지 결정한다.
2. 필요한 쪽 수만큼 `duplicate_slide`, 남는 카드는 `delete_shape`, 패널은 `set_geometry`.
3. 원문마다 `search_source`/`read_source`로 개요·문제점·연구개발 내용 초안 작성
4. `set_paragraphs`로 채우고 `lint_slide` → 위반 항목만 문구 수정 → 통과까지 반복
5. `render_slide`로 최종 확인 후 `finish`

완료 기준: 2026-10-06 착수보고 사례(5개 과제, 35~36쪽)를 사람 개입 없이 재현하고 LayoutLint 위반 0건.

## 6. 안전·비용·개인정보

- **원본 보호**: 작업 사본에서만 실행. 원본이 PowerPoint에 열려 있으면 열린 문서를 건드리지 않도록 사본을 별도로 연다.
- **승인**: 미리보기(전후 썸네일, 슬라이드별 변경 목록, Lint 결과) 후 적용. 거절하면 사본 삭제.
- **비용 상한**: 세션별 턴·토큰 상한, 사용량(`usage`) 누적 표시.
- **데이터 전송 고지**: 첫 실행 시 "문서 텍스트와 슬라이드 이미지가 gitoky LLM으로 전송됨" 동의. 보낼 문맥은 도구 결과로 필요한 만큼만(기존 데이터 최소화 원칙).
- **로그**: 세션 대화·도구 호출·결과를 `%LOCALAPPDATA%\BckSoft\BckPPT\sessions\`에 JSONL로 저장(재현·디버깅용). 원문 전체는 저장하지 않는다.

## 7. 일반 사용자 배포

- 요구사항 검사: 기존 `EnvironmentHealthService`에 PowerPoint 버전, COM 등록, 자동화 허용 검사 추가
- 설치: 기존 `scripts/package.ps1` 단일 폴더 배포 + `UpdateService` 자동 업데이트 유지. 첫 실행 마법사(키 입력·동의·연결 테스트).
- API 키: 1차는 사용자가 gitoky에서 발급받은 키를 입력(Windows DPAPI로 암호화 저장). `.env`는 개발용으로만 쓴다.
- 장애 대응: PowerPoint 응답 없음, 파일 잠김, 네트워크 오류에 사용자용 메시지와 재시도 버튼

## 8. 단계별 구현

### Phase A. LLM 기능 확인과 클라이언트 확장

- gitoky 품질 모델로 `tools` 요청, `tool_calls` 응답, 이미지 입력, 최대 문맥 길이를 실험해 결과를 이 문서 9장에 기록
- `ILlmClient.CompleteAsync`, `LlmRequest.Tools`, `LlmMessage`의 `tool_calls`/`tool_call_id`/이미지 파트 추가
- 지원하지 않으면 4.5의 JSON 행동 프로토콜 디코더 작성

완료 기준: 테스트 도구 1개(`echo`)를 모델이 호출하고 결과를 받아 답하는 통합 테스트 통과

### Phase B. ToolRegistry와 AgentLoop

- `Agent/Tools/`: 도구 정의(JSON Schema)와 실행기, 4.2 목록 중 읽기 도구 + `duplicate_slide`, `set_paragraphs`
- `Agent/AgentLoop.cs`: 턴 상한, 도구 결과 자르기, 취소, 세션 로그
- `set_paragraphs`의 문단 맞춤 순수 함수 + 단위 테스트
- `FakeLlmClient`로 녹화된 대화를 재생하는 테스트

완료 기준: 가짜 LLM 대본으로 "슬라이드 복제 → 문단 교체 → finish"가 사본에 적용됨

### Phase C. LayoutLint

- `PowerPoint/LayoutLint.cs`: 4.3 규칙, baseline 연결, `PptRules` 설정
- 규칙별 단위 테스트(줄 정보 입력 → 위반 목록)
- `lint_slide` 도구 연결, 루프 종료 전 자동 실행

완료 기준: 2026-10-06 사례의 초기 상태(12줄 상자, 고아 단어, 자간 -1.5)를 모두 잡아냄

### Phase D. 원문 라이브러리

- `Rag/SourceLibrary.cs`: pptx/PDF/md 수집, 쪽 단위 FTS5 색인, `search_source`/`read_source`
- 채팅창에 원문 파일·폴더 첨부 UI

완료 기준: 400쪽 보고서에서 "연구개발 내용" 관련 쪽을 상위 5개 안에 찾음

### Phase E. 채팅창 통합과 승인

- `AgentChatWindow`에 루프 진행 표시(현재 도구, 측정 결과), 중단 버튼
- 기존 `ApprovalWorkflow` 미리보기·승인 흐름에 루프 결과 연결, 후속 피드백으로 루프 재개
- 쓰기 도구 확대: 도형 편집·서식 도구

완료 기준: 5장 시나리오를 채팅창에서 처음부터 끝까지 수행

### Phase F. 배포 준비

- 첫 실행 마법사, 키 DPAPI 저장, 데이터 전송 동의, 환경 검사 강화
- 사용량 표시, 오류 메시지 정리, 베타 사용자 3~5명 테스트

완료 기준: 개발 환경이 없는 PC에 설치해 5장 시나리오 성공

## 9. 확인한 것 (Phase A, 2026-10-06 실측)

`GET /v1/models`와 직접 요청으로 확인했다. 결론: **4.5의 1번(네이티브 function calling)을 쓴다.** JSON 행동 프로토콜 디코더는 만들지 않았다.

| 항목 | 결과 |
| --- | --- |
| 품질 모델 이름/라우트 | `agent` (qwen3-32b, Bedrock 도쿄). "Agent loops" 전용이고 `auto`는 여기로 보내지 않으므로 명시 지정한다. 그 밖에 `general`(같은 모델), `code`(qwen3-coder-30b), `fast`(nova-micro), `vision`(nova-lite) |
| `tools` / `tool_calls` 지원 | 지원. 다섯 라우트 모두 capabilities에 tools. `role: tool` + `tool_call_id` 왕복 확인. 모델이 정수 인자를 `"1"`처럼 문자열로 보내기도 해서 `ToolArgs`가 둘 다 받는다 |
| 이미지 입력 지원 | `agent`는 거부(`403 model_not_allowed`). `vision`은 `image_url` data URL로 동작(색·글자를 정확히 읽음). → `render_slide`는 vision 라우트에 따로 묻고 **글 설명**을 루프에 돌려준다 |
| 최대 문맥·출력 길이 | `agent` context_window 32,768 토큰. 루프는 대화가 36,000자를 넘으면 오래된 도구 결과부터 한 줄로 줄인다(`AgentLoop.Compact`). 출력은 턴당 max_tokens 4,096 |
| 스트리밍 중 `tool_calls` 전달 방식 | 조각내지 않고 한 청크에 `id`·`name`·`arguments` 전체가 온다. 루프는 비스트리밍 `CompleteAsync`를 쓴다 |
| 요청당 비용·속도 | 평소 1턴 0.7초 안팎(입력 약 200토큰). 응답 헤더 `X-LLM-Cache: unsupported_model` — 프롬프트 캐시 없음. **같은 날 18시 40분경 게이트웨이 전체 지연이 8~15초로 늘었고(`/`, `/v1/models` 포함) 15초 근처에서 응답 없이 연결이 끊겼다**(HTTP 000, 스트리밍도 같음). 클라이언트는 재시도 후 `connection_lost`로 알린다. 비용 단가는 미확인 |

설정: `GITOKY_LLM_AGENT_MODEL`(기본 `agent`), `GITOKY_LLM_VISION_MODEL`(기본 `vision`).

## 9-1. 구현 현황

| Phase | 상태 | 주요 파일 |
| --- | --- | --- |
| A | 완료. 실서버 echo 왕복 테스트 통과(`GITOKY_LLM_API_KEY`가 있을 때만 실행) | `Llm/LlmModels.cs`, `GitokyLlmClient`의 `CompleteAsync`·`BuildPayload`·`ParseCompletion` |
| B | 완료. 가짜 LLM 대본 테스트에 더해, 실제 PowerPoint + gitoky `agent`로 2카드 슬라이드 복제·채우기를 끝까지 수행(19~20턴, 약 9만 토큰, Lint 위반 0). 첫 실행에서는 `list_shapes` 미리보기(문단을 공백으로 이은 것)를 원문으로 알고 `replace_text`가 3번 실패해, 미리보기에 ¶/↵ 경계를 넣고 프롬프트에 도구 선택 기준을 추가했다. 문구 품질(어색한 개요 문장)은 qwen3-32b의 한계로 보이며 원문 검색(Phase D) 후 다시 본다 | `Agent/AgentLoop.cs`, `AgentSession.cs`, `ParagraphFitPlanner.cs`, `SlideOrigins.cs`, `Agent/Tools/*` |
| C | 규칙·단위 테스트 완료. finish 때 건드린 슬라이드를 자동 검사해 위반이 있으면 2회까지 되돌려 보낸다 | `PowerPoint/LayoutLint.cs`, `PptRules`의 `OrphanMaxChars`·`TitleSecondLineMinRatio`·`TitleRoles` |
| D | 완료. pdf(PdfPig)·pptx(OpenXml)·md/txt를 쪽(구간) 단위로 FTS5 trigram 색인, 2글자 낱말은 LIKE 보완, tf-idf + 절 제목·붙여 쓴 구 가산점. 합성 400구간 보고서에서 "연구개발 내용" 절 1위, 실제 RFP PDF 27쪽 2.4초 색인. 채팅창 "원문 첨부" 버튼과 끌어 놓기 | `Rag/SourceLibrary.cs`, 도구 `source_outline`·`search_source`·`read_source` |
| E | 1차 완료. 채팅 "보내기"가 도구 루프로 바뀌었고(외부 요청 `create_job`은 기존 단발 계획 흐름 유지) 진행 로그·취소·전후 썸네일·"결과 저장/결과 열어 보기/작업 버리기", 후속 요청 이어가기. UI 자동화로 창에서 끝까지 실행 확인 | `Application/AgentWorkflow.cs`, `Views/AgentChatWindow.*` |
| F | 미착수 | |

실서버 대표 실행(2026-10-06, 카드 2개 + 합성 기획보고서 2종 첨부): 목차 → 요약문·연구개발 내용 쪽 읽기 → 6개 상자 채우기 → Lint 0 → finish, **17턴·약 10.8만 토큰·40초**.
여기까지 오며 실측으로 고친 것:

- 모델이 미리보기 표기 `¶`를 흉내 내 한 문단에 넣음 → `set_paragraphs`가 `¶`를 문단 경계로, `↵`를 Shift+Enter로 해석
- 검색 1건만 보고 부록을 읽은 뒤 내용을 지어냄 → `source_outline`(목차) 도구, 검색 최소 3건, 프롬프트에 "목차 먼저"
- 원문 문장을 그대로 옮겨 줄 수 초과, 문단 경계가 낀 `replace_text`를 반복 실패하며 30턴 소진 → 모든 줄 측정에 "(기준 N줄, 1줄 초과 — 약 28자 줄이세요)"를 붙이고, 문단을 넘는 `replace_text`는 즉시 `set_paragraphs`를 안내
- "이제 lint_slide 하겠습니다"라고 말만 하고 멈춤 → 슬라이드를 바꾼 뒤 도구 없이 답하면 한 번 이어가게 함
- 템플릿에 원래 있던 문제(손대지 않은 제목의 고아 단어)까지 지적 → 기준을 **원본 파일의 같은 슬라이드**(복제본은 복제 원본)로 바꾸고, 기준과 같은 상자는 판정하지 않음

남은 품질 문제: 한 문장을 두 문단(Enter)으로 쪼개 줄 수를 맞추는 경우가 있다.

## 9-2. 5장 시나리오 실측 (2026-10-06, 착수보고 PPT 31MB·38장 + 기획보고서 5종 1,370쪽)

정답: `Downloads\(착수보고) …_김병철35,36쪽추가.pptx`의 35쪽(카드 4: 1·6·7·9번)과 36쪽(카드 1: 11번, 패널 폭 236).
대상 PPT는 작업 폴더 사본에서 돌렸고 원본은 건드리지 않았다. 실행마다 막힌 곳과 고친 것:

| 실행 | 결과 | 막힌 곳 | 고친 것 |
| --- | --- | --- | --- |
| 1 | 1턴 중단 | `list_slides`가 `AnalyzePresentation`(전 도형 측정)으로 45초 COM 제한 초과 → COM 격리 | `ListSlides`(번호·제목만), 세션 동안 작업 사본·원본을 창 없이 열어 둠(`KeepOpen`) |
| 2 | 1턴 중단 | 모든 편집이 "Presentation cannot be modified" | 원인: **미리 보기·인쇄 전용(fsType 0x0004) 내장 글꼴**(G마켓 산스, KoPub돋움체 일부). 작업 사본에서 제한 글꼴만 제거(`RestrictedFontRemover`) |
| 3 | 약 30턴, 문맥 초과 | 32k 토큰 초과, 63개 도형 이름이 겹쳐 소제목을 본문으로 착각 | 실제 `prompt_tokens`로 예산 추정·초과 시 재시도, `list_shapes`를 카드별로 묶고 (소제목) 표시 |
| 4 | 80턴 상한, 188만 토큰 | 원문을 한 쪽씩 60턴 읽음, 긴 파일명을 한자로 틀리게 옮겨 씀 | 문서 별칭 D1…, `brief_source`(별도 요약 호출) |
| 5 | 43턴, 78만 토큰, 위반 15건 | `#n` 주소가 이름 바꿀 때마다 밀려 엉뚱한 상자에 이름, 태그 상자 폭 부족 | 겹치는 이름은 `@Id` 고정 주소, Id로 기준 짝짓기, 짧은 상자는 폭 넓히기 안내, 위반 목록 담은 재촉 |
| 6 | 55도구 호출 후 중단 | **gitoky 일일 토큰 한도(500만) 초과**. 그 전까지 카드 4장 24개 상자 이름 붙이기 오류 없음 | 카드 배경만 지우고 글 상자를 남김 → `delete_shape(with_contents)`, 계산식 인자("216*4") 허용 |
| 7 | 80턴, 127만 토큰 | 한 턴에 도구 하나씩 불러 이름 바꾸기에만 80턴 | 묶음 도구 `rename_shapes`·`set_texts` |
| 8 | 37초 중단 | 병렬 도구 호출 뒤 요청을 gitoky가 `upstream_rejected`로 거부 | 병렬 호출을 "호출→결과"로 풀어 보냄 |
| 9 | 80턴, 187만 토큰, 결과 사용 불가 | 번호·제목 상자 짝짓기가 계속 틀림. **qwen3-32b로는 구조 편집이 안 된다고 판단** | 구조는 코드가 하는 파이프라인으로 전환(아래) |
| 10 | 파이프라인 4.2만 + 루프 27.7만 | 루프 모델이 `fill_cards` 인자를 직접 쓰며 카드 5→2장 | 사양은 요청 원문에서 단발 호출로 뽑고 "N번" 누락 검사 |
| 11 | 파이프라인 6.5만 + 루프 35만 | 파이프라인은 5장 맞게 채움. 뒤이은 qwen 루프가 이름·패널·문구를 망침 | `fill_cards` 뒤 루프 종료, 카드 자리별 예산, 자간 통일, 범례 위치 |
| 12 | 중단 | 편집마다 27MB 저장 → 45초 COM 제한 초과 | 세션 문서는 닫을 때 한 번 저장 |
| 13 | **7.2만 토큰, 3분** | 원래 결과본과 같은 구조. qwen 문안 품질 문제(한자 섞임, 제목 바꿔 씀) | 사용자 고급 모델 외부 작성 흐름 |
| 14 | 외부 작성, gitoky 약 1.3만 | 문안은 원래 결과본 수준. qwen 길이 맞춤이 제목 어순·개요를 망침 | 제목은 과제명 그대로·균형은 코드, 길이 맞춤은 항목 수 같고 나아질 때만 받음, 요청문에 자리별 예산 |
| 15 | **외부 작성, gitoky 1.05만 토큰, 2분** | 원래 결과본과 구조·문안 거의 일치, 꼬리 단어 위반 3건 | — |

### 9-3. 바뀐 구조: 카드 채우기 파이프라인과 외부 작성

긴 도구 루프는 매 턴 대화 전체를 다시 보내 비싸고(프롬프트 캐시 없음), qwen3-32b는 63개 도형의 구조 편집을 해내지 못했다.
그래서 반복 작업인 "템플릿 카드 채우기"는 정해진 단계로 돈다(`Agent/Cards/CardPipeline.cs`, 도구 `fill_cards`).

1. 사양 추출: 사용자 원문 + 원문 목록 → `CardJob` JSON(qwen 1회, "N번" 누락 시 1회 재시도)
2. 템플릿 구조·예산: `CardLayoutAnalyzer`(위치만으로 카드·번호·태그·제목·소제목/본문·패널·머리말), 상자별 줄 수·한 줄 글자 수(코드)
3. 문안: **외부 작성**이면 원문 발췌(문서당 8천 자)·자리별 예산·예시·출력 형식을 담은 요청문을 사용자가 자기 고급 모델(ChatGPT·Claude 등)에 붙여넣고 답(JSON)을 붙여넣는다. 아니면 qwen이 요점 정리(문서당 1회)와 문안 작성(1회)
4. 배치(코드): 복제, 남는 카드 삭제, 패널·머리말 폭, 태그 폭, 제목(과제명 그대로·앞부분 서식 유지·2줄 균형), 범례, `C{n}_{역할}` 이름, 자간 통일
5. 길이 맞춤: 넘친 상자만 qwen 단발(항목 수 유지·나아질 때만 반영, 최대 2바퀴), 그래도 넘치면 마지막 항목부터 코드로 빼고 보고. 설정으로 LLM 없이 코드만 쓸 수 있다
6. 판정 후 루프를 끝내고 결과·남은 위반을 보여준다. 고칠 것은 사용자가 짚어 후속 요청

gitoky는 아직 비용 관리가 어려워 고급 모델(quality 라우트)을 쓰지 않는다. 비싼 글쓰기는 사용자 구독 모델에서, 측정·구조·짧은 교정은 BckPPT와 qwen이 맡는다.

비용 주의(실행 1~9, 루프 방식): 한 번에 78만~188만 토큰이 들었다. 일일 한도 500만이면 사용자 한 명이 하루 서너 번이다. 파이프라인과 외부 작성으로 1만 토큰대까지 줄였다(9-2·9-3).
턴마다 대화 전체(약 2만 토큰)를 다시 보내고 gitoky `agent` 라우트는 프롬프트 캐시가 없다(`X-LLM-Cache: unsupported_model`).
Phase F 전에 사용량 표시와 세션 토큰 상한 기본값, 캐시 지원 여부를 다시 정한다.

재실행: `scratchpad\real\run.sh <로그이름>`(이 PC의 실험 스크립트) 또는 스모크 테스트 환경변수 `BCKPPT_LIVE_PPTX`·`BCKPPT_LIVE_SOURCES`·`BCKPPT_LIVE_REQUEST`·`BCKPPT_LIVE_MAX_TURNS`·`BCKPPT_LIVE_KEEP=1`.

계획과 달라진 점:

- 4.2의 도구는 모두 만들었다(`search_source`/`read_source`는 원문 라이브러리가 있을 때만 노출). 쓰기 도구(`set_paragraphs`, `replace_text`, `set_text_format`)는 적용 직후 줄 측정 결과를 함께 돌려줘 측정 턴을 아낀다.
- `replace_text`는 기존 `ReplaceText`(최상위 도형 이름만, 런 단위)가 아니라 새 `ReplaceTextInShape`(주소 규약, `TextRange2.Replace`)를 쓴다.
- 서식 이탈 규칙을 위해 `GetTextFormat`(첫 런의 글꼴·크기·자간·굵기·색)을 추가했다. 같은 역할은 이름 규약 `C{n}_{역할}`로 묶고, 과반 서식과 다른 상자만 지적한다.
- 겹침은 새로 추가한 도형이 다른 도형에 **걸칠 때만** 잡는다. 패널 안에 완전히 들어간 배지는 의도한 배치로 본다.
- "상위 카드 아래 경계 초과" overflow는 아직 없다(텍스트 높이 > 상자 높이만 판정).
- 작업 사본은 `%LOCALAPPDATA%\BckSoft\BckPPT\work\`에 만들고, 승인하면 원본 옆 `이름_BckPPT_시각.pptx`로 복사, 거절하면 지운다(`AgentSession.Approve`/`Discard`).
- 실서버 스모크: `BCKPPT_LIVE_PPTX`(대상 파일)와 키를 주고 `dotnet test --filter 스모크`. `BCKPPT_LIVE_SOURCES`(원문, `|` 구분)·`BCKPPT_LIVE_REQUEST`로 바꿀 수 있고, 대화 로그 경로가 출력에 나온다. 원문 검색만 보려면 `BCKPPT_LIVE_SOURCE`·`BCKPPT_LIVE_QUERY`로 `--filter 실문서_검색`.

## 10. 다른 PC에서 시작하기

```powershell
git clone https://gitoky.com/bckim/bckppt.git
cd bckppt
copy .env.example .env   # GITOKY_LLM_API_KEY, GITOKY_LLM_MODEL 설정
.\run.ps1 --build-only   # 또는 dotnet build BckPPT.csproj -c Debug -p:Platform=x64
dotnet test BckPPT.Tests\BckPPT.Tests.csproj
```

- PowerPoint 데스크톱 앱이 있어야 COM 기능이 동작한다.
- 테스트는 `-p:Platform=x64` 없이 실행한다. 붙이면 출력 폴더가 한 단계 깊어져 `Assets/icon.pptx` 픽스처 경로가 어긋나 3개가 실패한다.
- 도형 API는 실행 중인 BckPPT의 디버그 API(`http://localhost:38700`)로도 바로 시험할 수 있다. 사용법은 README의 "도형 주소와 편집 API".
- 첫 작업은 Phase A(gitoky 모델 기능 확인)부터 한다. 결과에 따라 Phase B의 프로토콜이 정해진다.

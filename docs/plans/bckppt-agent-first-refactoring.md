# BckPPT 독립형 PPT 에이전트 리팩토링 계획

> 작성일: 2026-09-01  
> 대상 저장소: BckPPT  
> 상태: 구현 전 아키텍처 기준안

## 1. 결정 요약

BckPPT를 CellBox2의 부속 도구가 아닌 **독립형 Windows PPT 에이전트**로 개발한다.

- 프로그램을 실행하면 기존 편집창 대신 **에이전트 채팅창**이 먼저 열린다.
- 채팅창은 시작 화면이자 작업 허브이며, 기존 BckPPT 창은 필요할 때 여는 전문 편집창으로 유지한다.
- LLM은 `gitoky.com`의 `llm_api`를 사용한다.
- PPT 분석, 로컬 템플릿 검색, 작업계획 수립, PowerPoint COM 실행, 결과 검증까지 BckPPT가 소유한다.
- CellBox2 설치, 트레이, 동기화, 인증 및 대화 UI에 의존하지 않는다.
- 향후 외부 프로그램이 BckPPT를 호출할 수 있도록 안정적인 명령 계약만 별도로 제공한다.

## 2. 제품 목표

사용자는 PowerPoint 자동화 방법이나 BckPPT의 내부 기능을 몰라도 자연어로 다음과 같은 작업을 요청할 수 있어야 한다.

> 현재 PPT 3쪽의 목록 중 3번 항목을 구체화해서 한 장 추가해줘.

BckPPT는 요청을 이해하고 현재 문서와 로컬 PPT 라이브러리를 분석한 뒤 다음 중 적절한 방식을 스스로 선택한다.

1. 현재 문서의 기존 슬라이드 재사용
2. 개인 PPT 라이브러리의 슬라이드 재사용
3. 기존 레이아웃의 제한적 변형
4. 적합한 양식이 없을 때 신규 레이아웃 생성

결과는 적용 전에 변경 범위와 미리보기를 보여주고, 적용 후에는 줄바꿈·넘침·서식 유지 여부를 PowerPoint 렌더링 기준으로 검사한다.

## 3. 사용자 경험

### 3.1 시작 화면: AgentChatWindow

BckPPT 실행 시 채팅창을 기본 창으로 표시한다.

초기 화면의 최소 기능:

- PPT 파일 열기 및 드래그앤드롭
- 실행 중인 PowerPoint 문서 연결
- 최근 작업 3~5개 표시 및 대화 복원
- 현재 대상 파일과 선택 슬라이드 표시
- 자연어 요청 입력과 스트리밍 응답
- `미리보기`, `적용`, `다시 만들기`, `취소`
- `전문 편집창에서 열기`
- LLM 연결상태와 오류 안내

채팅창은 메시지만 나열하지 않고 작업상태를 카드로 표시한다.

```text
요청: 3쪽의 3번 항목을 구체화한 슬라이드 추가
대상: 사업기획안.pptx · 3쪽
상태: 레이아웃 후보 검색 → 문안 작성 → 줄 넘침 검사

[변경 미리보기] [적용] [다시 만들기] [전문 편집창]
```

### 3.2 전문 편집창: BckPptWindow

현재 창은 제거하지 않고 다음 작업을 위한 고급 도구로 유지한다.

- 슬라이드와 Shape 구조 확인
- 텍스트 역할 및 계층 수정
- 줄바꿈·문단·AutoFit 상세 조정
- 후보 템플릿 비교
- 변경 이력과 규칙 관리
- 인덱스 분류 결과의 수동 보정

채팅창과 편집창은 독립적으로 열리지만 동일한 `AgentSession`과 `DocumentSession`을 공유한다. 편집창을 닫아도 채팅과 작업은 유지하고, 채팅창을 닫을 때 전체 종료 여부를 처리한다.

## 4. 책임 경계

### BckPPT가 소유하는 것

- PPT 전용 시스템 프롬프트와 도구 정의
- 사용자 요청 해석 및 작업계획 수립
- `gitoky.com llm_api` 연동
- PowerPoint COM과 STA 실행 관리
- Open XML 및 COM 기반 PPT 분석
- 로컬 PPT 라이브러리 인덱싱과 검색
- 템플릿 재사용·변형·신규 생성 판단
- 구조화 명령 실행, 미리보기, 승인, 되돌리기
- 줄바꿈·overflow·서식·렌더링 검증 및 자동 보정
- 대화, 작업, 문서 세션의 로컬 저장

### 외부 프로그램에 맡길 수 있는 것

- BckPPT 실행
- 대상 파일 경로와 자연어 요청 전달
- 작업상태 조회
- 결과 파일 경로 수신

외부 프로그램이 Shape 단위 COM 명령이나 BckPPT 내부 프롬프트를 구성하게 하지 않는다. PPT 전문 판단은 항상 BckPPT 내부에 둔다.

## 5. 목표 아키텍처

```text
BckPPT.exe
├─ UI
│  ├─ AgentChatWindow          기본 시작창
│  └─ BckPptWindow             선택적 전문 편집창
├─ Application
│  ├─ AgentOrchestrator
│  ├─ SessionManager
│  ├─ PptTaskPlanner
│  └─ ApprovalWorkflow
├─ Agent
│  ├─ PromptBuilder
│  ├─ ToolRegistry
│  ├─ ContextAssembler
│  └─ ResponseDecoder
├─ Llm
│  ├─ ILlmClient
│  └─ GitokyLlmClient
├─ PowerPoint
│  ├─ IPptDocumentService
│  ├─ PptComService
│  ├─ StaDispatcher
│  └─ RenderValidator
├─ Rag
│  ├─ PptIndexer
│  ├─ TemplateRepository
│  ├─ TemplateRetriever
│  └─ LayoutDecisionRouter
└─ Infrastructure
   ├─ LocalStore
   ├─ CredentialStore
   ├─ Diagnostics
   └─ ExternalCommandEndpoint
```

초기에는 단일 EXE와 단일 프로젝트를 유지하되 네임스페이스와 인터페이스 경계를 먼저 만든다. COM 안정성이나 장기 작업 격리가 실제 문제가 될 때 `BckPPT.Worker.exe`를 분리한다. 리팩토링 시작부터 여러 프로세스로 나누지는 않는다.

## 6. LLM 연동 원칙

### 6.1 클라이언트 추상화

UI나 PPT 서비스가 HTTP API를 직접 호출하지 않는다.

```csharp
public interface ILlmClient
{
    IAsyncEnumerable<LlmStreamEvent> StreamAsync(
        LlmRequest request,
        CancellationToken cancellationToken);
}
```

첫 구현은 `GitokyLlmClient`이며, URL·모델·timeout·재시도 정책을 설정으로 분리한다. 이후 다른 API나 로컬 모델은 어댑터 추가로 대응한다.

### 6.2 구조화 작업계획

LLM이 C#이나 COM 코드를 생성하여 직접 실행하지 않도록 한다. 허용된 명령 DTO만 실행한다.

예:

- `analyze_presentation`
- `search_template_library`
- `duplicate_slide`
- `replace_shape_text`
- `insert_slide`
- `validate_layout`
- `save_copy`

모든 변경 명령에는 대상 문서, 슬라이드, Shape 식별자, 예상 사전상태를 포함하여 잘못된 문서에 적용되는 것을 방지한다.

### 6.3 데이터 최소화

- PPT 원본 파일을 기본적으로 LLM 서버에 전송하지 않는다.
- 요청 해결에 필요한 텍스트, 역할, 계층, 치수, 줄 수 등 구조화 문맥만 전송한다.
- 외부 이미지 전송이 필요하면 사용자에게 범위와 목적을 알리고 승인받는다.
- API 자격정보는 Windows Credential Manager 또는 DPAPI로 보호한다.
- 로그에는 API 키, 원문 전체, 민감한 파일경로를 남기지 않는다.

## 7. 로컬 PPT RAG

### 7.1 인덱싱 단위

프레젠테이션 전체가 아니라 슬라이드 단위로 검색할 수 있게 다음 정보를 저장한다.

- 출처 파일, 슬라이드 번호, 수정시각, 파일 해시
- 분류: 표지, 목차, 간지, 본문, 비교, 프로세스, 결론 등
- 텍스트량, 문단 수, 계층 깊이, 제목·본문·캡션 역할
- 카드·표·차트·이미지·서술형 영역의 개수
- 레이아웃 좌표, 정렬, 여백, 색상 및 글꼴 특징
- 완성도와 재사용 적합도
- 썸네일과 검색용 임베딩

### 7.2 단계적 분석

1. Open XML로 빠르게 구조와 텍스트를 추출한다.
2. 검색 후보 또는 현재 작업 문서만 PowerPoint COM으로 정밀 분석한다.
3. 파일 변경 시 변경된 슬라이드만 증분 갱신한다.

인덱싱 범위는 사용자가 선택한 폴더로 제한하며 제외 폴더, 일시중지, 전체 삭제 기능을 제공한다.

### 7.3 레이아웃 결정

`LayoutDecisionRouter`는 다음 순서로 판단한다.

1. 현재 문서에서 목적과 내용량이 맞는 슬라이드 검색
2. 개인 라이브러리에서 디자인 문법이 호환되는 슬라이드 검색
3. 높은 유사도의 레이아웃을 제한적으로 변형
4. 기준 점수 미달이면 신규 생성

사용자에게 매번 템플릿 선택을 요구하지 않는다. 다만 신뢰도가 낮거나 디자인 방향에 따라 결과가 크게 갈릴 때만 후보를 제시한다.

## 8. 세션과 안전한 적용

### 세션 모델

- `AgentSession`: 대화, 요약, 작업목록
- `DocumentSession`: 파일 경로, 열린 Presentation, 선택 슬라이드
- `PptJob`: 요청, 계획, 명령, 상태, 결과
- `ChangeSet`: 변경 전후 값과 되돌리기 정보
- `Artifact`: 미리보기, 생성본, 분석 리포트

### 적용 원칙

- 분석과 미리보기는 원본을 변경하지 않는다.
- 기본 저장은 원본 덮어쓰기보다 사본 생성을 우선한다.
- 원본 수정은 적용 직전에 명시적으로 확인한다.
- 적용 전 파일 해시나 수정시각을 확인하여 외부 변경 충돌을 감지한다.
- COM 작업은 STA에서 직렬 실행하고 취소·timeout·오류 복구 경계를 둔다.
- 적용 후 PowerPoint 렌더링 결과로 줄 넘침과 레이아웃을 다시 검사한다.

## 9. 외부 연계 계약

BckPPT의 독립성을 유지하면서 자동화에 사용할 최소 계약을 제공한다.

```json
{
  "protocolVersion": "1.0",
  "command": "create_job",
  "requestId": "uuid",
  "filePath": "C:\\...\\proposal.pptx",
  "instruction": "3쪽의 3번 항목을 한 장으로 구체화해줘",
  "saveMode": "copy"
}
```

초기 전달 방식은 명령행 인수 또는 localhost named pipe 중 하나를 선택한다. 기존 Debug HTTP/MCP API는 개발·진단 용도로 분리하고, 제품 연계 계약과 혼용하지 않는다.

## 10. 단계별 구현

### Phase 0. 기반 정리

- 루트 네임스페이스 `CellBox2` 의존을 `BckPPT`로 변경
- 현재 UI, ViewModel, COM 서비스의 직접 참조 관계 파악
- 핵심 서비스 인터페이스와 DI composition root 도입
- 기존 동작 회귀 테스트와 대표 PPT fixture 확보

완료 기준: 기존 편집 기능이 동일하게 동작하면서 CellBox2 이름과 런타임 의존이 제거된다.

### Phase 1. 채팅 우선 셸

- `AgentChatWindow` 추가 및 StartupUri 변경
- 파일 열기, 드래그앤드롭, 현재 PowerPoint 연결
- 전문 편집창 열기와 두 창의 `DocumentSession` 공유
- 최근 대화와 작업상태의 로컬 저장

완료 기준: BckPPT 실행 후 채팅창에서 문서를 선택하고 기존 편집창으로 같은 문서를 열 수 있다.

### Phase 2. gitoky LLM 연결

- `ILlmClient`와 `GitokyLlmClient` 구현
- 설정·자격정보 저장과 연결 테스트
- 스트리밍, 취소, timeout, 재시도, 오류 메시지
- PPT 전용 시스템 프롬프트와 구조화 응답 스키마

완료 기준: 채팅창에서 요청을 보내고 구조화된 작업계획을 받을 수 있다.

### Phase 3. 에이전트 도구화

- 기존 분석·치환·저장 기능을 명시적 도구로 래핑
- 허용 명령 검증과 사전조건 검사
- 계획 → 미리보기 → 승인 → 실행 상태 머신
- 변경 이력과 undo 연계

완료 기준: 단일 슬라이드 복제 및 텍스트 교체 작업을 자연어로 안전하게 완료한다.

### Phase 4. 품질 검증 루프

- 줄바꿈 유형과 서식 보존 검사
- PowerPoint 렌더링 기반 overflow 검사
- 제한된 자동 보정 규칙 적용
- 전후 썸네일과 검증결과 표시

완료 기준: 적용 결과가 기준을 벗어나면 자동 보정하거나 승인 전에 경고한다.

### Phase 5. 로컬 PPT RAG

- 사용자 동의와 인덱싱 폴더 설정
- Open XML 기반 슬라이드 인덱서
- SQLite 메타데이터와 벡터 저장소
- 분류·검색·재사용 점수 및 증분 갱신
- 현재 문서/개인 라이브러리/신규 생성 라우터

완료 기준: 자연어 요청에 맞는 기존 슬라이드를 검색하고 선택 근거를 작업 로그에 남긴다.

### Phase 6. 외부 호출과 배포

- 버전이 있는 외부 명령 계약
- 단일 인스턴스 처리와 기존 채팅 세션 전달
- 독립 인스톨러, 자동 업데이트, 코드 서명
- PowerPoint 미설치·버전 차이·COM hang 복구

완료 기준: 외부 앱이 BckPPT를 실행하여 작업을 요청할 수 있으나 BckPPT 없이도 외부 앱의 기능에는 영향이 없다.

## 11. 테스트 전략

- 단위 테스트: 작업계획 검증, 명령 DTO, 줄바꿈 보존, 검색 점수
- 계약 테스트: `GitokyLlmClient` 요청·스트림·오류 응답
- COM 통합 테스트: 실행 중/미실행 PowerPoint, 여러 문서, 저장 충돌
- 골든 파일 테스트: 대표 PPT의 텍스트, 줄 수, 글꼴, 위치, 도형 수 비교
- 시각 회귀 테스트: 적용 전후 슬라이드 렌더 이미지 비교
- 장애 테스트: 네트워크 중단, LLM timeout, COM 대화상자, 앱 강제종료 후 복구
- 보안 테스트: 자격정보 저장, 경로 검증, 외부 명령 인증 및 입력 제한

## 12. 초기 출시 범위

첫 버전에서는 다음 시나리오에 집중한다.

1. 채팅창에서 PPT 열기 또는 현재 문서 연결
2. 슬라이드 요약과 텍스트 구조 분석
3. 현재 문서 안의 기존 양식을 복제하여 문안 교체
4. 적용 전 미리보기와 승인
5. 줄바꿈·overflow 검증
6. 사본 저장과 전문 편집창 열기

개인 PPT 전체 인덱싱, 복합 슬라이드 재조립, 완전 신규 디자인 생성은 이 흐름이 안정된 뒤 추가한다.

## 13. 완료 판단 기준

- BckPPT가 CellBox2 없이 설치·실행·업데이트된다.
- 프로그램 실행 시 채팅창이 먼저 나타난다.
- 사용자가 자연어만으로 대표 PPT 수정 시나리오를 완료한다.
- LLM API 장애가 기존 수동 편집 기능을 막지 않는다.
- 모든 PPT 변경은 미리보기·승인·검증·되돌리기 경계를 갖는다.
- 외부 연계는 선택 사항이며 BckPPT 내부 구현을 노출하지 않는다.


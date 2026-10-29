# BckPPT

BckPPT는 기존 PowerPoint 슬라이드의 레이아웃과 서식을 유지하면서 텍스트를 분석·교체하는 Windows 애플리케이션입니다. PowerPoint가 실제로 해석한 문단, 강제 줄바꿈, 자동 줄바꿈과 텍스트 영역을 기준으로 결과를 검증하는 데 초점을 둡니다.

`gitoky.com`의 LLM API를 사용하는 독립형 PPT 에이전트이며, 채팅창에서 작업을 요청하고 필요할 때 전문 편집창을 열 수 있습니다. CellBox2에 포함되거나 의존하지 않습니다.

## 현재 기능

- PowerPoint 파일 열기와 슬라이드 탐색
- 슬라이드·Shape·텍스트 계층 구조 분석
- 텍스트 상자별 줄 수와 줄바꿈 유형 식별
  - 자동 줄바꿈
  - Enter 문단 나눔
  - Shift+Enter 강제 줄바꿈
- 문단별 텍스트 교체와 원본 서식 유지
- 슬라이드별 Markdown 내보내기
  - 그룹 도형·표·문단을 재귀적으로 추출
  - 슬라이드 PNG 미리보기를 함께 저장해 이미지·차트·도형도 원본 맥락으로 확인
- overflow 감지와 AutoFit 보정
- Shape 역할 분류와 구조 정보 저장
- 변경 이력 확인
- localhost 디버그 API를 통한 분석·편집 자동화
- 자연어 작업계획, 미리보기, 승인 및 사본 적용
- 적용 전후 렌더링과 overflow 자동 보정
- Open XML·SQLite 기반 로컬 PPT 템플릿 검색
- 버전이 있는 named pipe 외부 작업 계약과 단일 인스턴스 전달

## 개발 방향

- 에이전트 채팅창을 기본 시작 화면으로 제공
- `gitoky.com llm_api`를 추상화된 LLM 클라이언트로 연결
- 자연어 요청을 검증된 PPT 작업 명령으로 변환
- 적용 전 미리보기, 사용자 승인, 사본 저장과 되돌리기 제공
- 개인 PPT를 슬라이드 단위로 인덱싱하는 로컬 PPT RAG 구축
- 기존 양식 재사용, 제한적 변형, 신규 생성 중 적합한 방식을 자동 선택
- PowerPoint 렌더링 기준의 줄 넘침·서식·레이아웃 검증

상세 계획은 [독립형 PPT 에이전트 리팩토링 계획](docs/plans/bckppt-agent-first-refactoring.md)을 참고하십시오.

## 요구사항

- Windows 10 또는 11
- .NET 8 SDK 이상
- Microsoft PowerPoint 데스크톱 앱
- x64 환경 권장

PowerPoint COM 자동화를 사용하므로 현재 버전은 Windows 및 Microsoft PowerPoint 설치 환경을 전제로 합니다.

## LLM 키

채팅 요청(에이전트)에는 gitoky LLM 키가 필요합니다. 다음 줄을 `.env`에 둡니다.

```text
GITOKY_LLM_API_KEY=발급받은-키
```

`.env`는 실행 파일 폴더부터 위로 찾고, 없으면 `%LOCALAPPDATA%\BckSoft\BckPPT\.env`를 읽습니다.
플러그인 업데이트는 설치 폴더를 통째로 바꾸므로 키는 `%LOCALAPPDATA%` 쪽에 두는 편이 안전합니다.

키가 없으면 채팅 요청과 named pipe `create_job`만 꺼지고, 전문 편집창(분석·글꼴 찾기·텍스트 치환·
Markdown 내보내기)과 localhost 디버그 API는 그대로 동작합니다.

## 실행

PowerShell에서 다음 명령을 실행합니다.

```powershell
.\run.ps1
```

`run.ps1`은 소스가 실행파일보다 최신일 때만 다시 빌드한 후 BckPPT를 실행합니다.

빌드만 수행하려면 다음과 같이 실행합니다.

```powershell
.\run.ps1 -BuildOnly
```

BckPPT는 창 없이 작업 표시줄 알림 영역(트레이)에서 시작하고 MCP 서버를 켭니다. 트레이 아이콘을 두 번
누르면 채팅창이 열리고, 오른쪽 메뉴에서 전문 편집창·Claude Code 연결 명령 복사·종료를 고릅니다.
이미 실행 중일 때 다시 실행하면(CellBox 트레이 메뉴 포함) 채팅창을 띄웁니다.

## MCP 연결

실행 중인 BckPPT는 `http://localhost:38701/mcp`(Streamable HTTP)로 PPT 도구를 내보냅니다.
판단은 연결한 에이전트가 하고 BckPPT는 PowerPoint를 다루므로 **LLM 키가 없어도 됩니다.**

```powershell
claude mcp add --transport http bckppt http://localhost:38701/mcp
```

- 도구마다 `file` 인자로 대상을 고릅니다. 비우면 PowerPoint에서 보고 있는 문서, 절대 경로면 그 파일,
  아니면 열린 문서 이름 조각입니다.
- 열린 창의 문서는 그 자리에서 고치고 저장하지 않습니다(저장은 사용자가 합니다). 작업 중 사용자가
  도형·글을 선택하면 멈춥니다. 창 없는 파일은 고친 뒤 저장합니다.
- 내부 에이전트의 PPT 도구(슬라이드·도형 읽기, 줄 측정, 문단 교체, 도형·표·선 추가 등)에 더해
  `list_presentations`, `open_presentation`, `go_to_slide`, `list_fonts`(글꼴 찾기), `export_markdown`,
  `render_slide`(PNG 그림 반환)가 있습니다. LLM이 필요한 카드 채우기·원문 요약은 내보내지 않습니다.
- localhost에서만 열리고, localhost가 아닌 Origin(브라우저 페이지)의 요청은 거절합니다.

## 직접 빌드

```powershell
dotnet restore BckPPT.csproj
dotnet build BckPPT.csproj -c Debug -p:Platform=x64
```

## Markdown 내보내기

편집 창에서 대상 PPT를 선택한 뒤 상단의 `MD`를 누르고 저장 위치를 고릅니다. Markdown 파일과
`<파일명>_assets` 폴더가 함께 생성됩니다. 그룹 안에 들어간 텍스트, 표, 문단 들여쓰기를 추출하며,
텍스트로 바꿀 수 없는 이미지·차트·도형은 슬라이드 PNG 미리보기로 보존합니다.

빌드 결과는 `bin/x64/Debug/net8.0-windows/`에 생성됩니다.

## 도형 주소와 편집 API

도형을 이름으로 지정하는 모든 API(`SetShapeText`, `SetParagraphTexts`, `GetTextLineBreaks`,
`GetShapeInfo` 등)는 그룹 안 도형까지 찾고, 다음 주소 규약을 따릅니다.

| 주소 | 뜻 |
| --- | --- |
| `Text 19` | 그룹 안까지 펼친 순서에서 첫 번째 `Text 19` |
| `Text 19#3` | 같은 이름 중 세 번째 (카드를 복사해 이름이 겹칠 때) |
| `Group 171/Text 19` | 바로 위 그룹이 `Group 171`인 것 |
| `Table 1 [R2C3]` | 표 셀. 위 형식과 함께 쓸 수 있음 |

`ListShapes`가 슬라이드의 모든 도형과 그 주소를 돌려주므로, 이름이 겹치면 거기서 `#n` 주소를 확인합니다.
배치·서식 API로 `RenameShape`, `SetShapeGeometry`(뒤집힌 도형도 위치 유지), `DeleteShape`,
`AddTextShape`(범례·배지용, 글자 폭 자동 맞춤과 오른쪽 끝 정렬), `SetTextColor`(특정 글자만 색 지정),
`SetTextFormat`이 있습니다. 색은 `#RRGGBB`로 지정합니다.

디버그 API(`http://localhost:38700`)로도 같은 작업을 할 수 있습니다. 대상 문서는 `path`가 있으면 그 파일,
없으면 선택된 프레젠테이션입니다.

| 요청 | 본문/쿼리 |
| --- | --- |
| `GET /api/shapes` | `slide` |
| `GET /api/line-breaks` | `slide`, `shape` |
| `POST /api/shapes/rename` | `slide`, `shape`, `newName` |
| `POST /api/shapes/geometry` | `slide`, `shape`, `left?`, `top?`, `width?`, `height?` |
| `POST /api/shapes/delete` | `slide`, `shape` |
| `POST /api/shapes/add-text` | `slide`, `text`, `left`, `top`, `width`, `height`, `name?`, `kind?`(`roundRect`), `fillColor?`, `fontName?`, `fontSize?`, `bold?`, `fontColor?`, `autoWidth?`, `rightEdge?` |
| `POST /api/shapes/text-color` | `slide`, `shape`, `find`, `color` |
| `POST /api/shapes/text-format` | `slide`, `shape`, `size?`, `bold?`, `color?`, `spacing?`, `fontName?` |

## 외부 작업 요청

외부 요청은 항상 `saveMode: copy`만 허용하며 채팅창의 미리보기·승인 절차를 거칩니다.

```json
{
  "protocolVersion": "1.0",
  "command": "create_job",
  "requestId": "00000000-0000-0000-0000-000000000001",
  "filePath": "C:\\Presentations\\proposal.pptx",
  "instruction": "3쪽의 3번 항목을 한 장으로 구체화해줘",
  "saveMode": "copy"
}
```

요청 JSON을 파일로 저장한 뒤 실행 중인 BckPPT에 전달할 수 있습니다.

```powershell
BckPPT.exe --request-file C:\Requests\ppt-job.json
```

상태 조회는 같은 `requestId`와 `command: get_job_status`를 named pipe `BckPPT.Agent.v1`로 전달합니다.

## 배포 패키지

독립 실행 가능한 win-x64 ZIP과 SHA-256 파일을 생성합니다.

```powershell
.\scripts\package.ps1
```

코드 서명까지 수행하려면 `BCKPPT_SIGNTOOL`, `BCKPPT_CERT_SHA1` 환경변수를 설정합니다.

```powershell
.\scripts\package.ps1 -Sign
```

자동 업데이트 manifest 주소는 HTTPS `BCKPPT_UPDATE_MANIFEST_URL`로 주입합니다. manifest 형식은 다음과 같습니다.

```json
{
  "version": "2.1.1",
  "packageUrl": "https://example.com/BckPPT-2.1.1-win-x64.zip",
  "sha256": "64자리 SHA-256",
  "releaseNotes": "변경 내용"
}
```

현재 앱은 manifest 검증 경계까지 제공하며 무인 다운로드·교체는 활성화하지 않습니다.

## 프로젝트 구조

```text
BckPPT/
├─ Agent/        프롬프트, 구조화 계획 및 문맥 조립
├─ Application/  세션, 작업 상태 머신 및 승인 워크플로
├─ Infrastructure/ 외부 명령, 단일 인스턴스, 임시 artifact
├─ Llm/          GitOky 스트리밍 LLM 클라이언트
├─ Mcp/          외부 에이전트용 localhost MCP 서버와 도구 목록
├─ PowerPoint/   렌더링 및 품질 검증
├─ Rag/          Open XML 인덱싱, SQLite 검색 및 레이아웃 결정
├─ Models/       PPT, 슬라이드, Shape 및 분석 결과 모델
├─ Services/     PowerPoint COM, 구조 저장, 역할 분류, 변경 이력
├─ ViewModels/   메인 편집 작업 상태와 명령
└─ Views/        트레이, 채팅창과 분석·편집 보조창
Assets/          애플리케이션 아이콘과 원본 자산
docs/plans/      아키텍처 및 리팩토링 계획
```

## 안전 원칙

- PowerPoint COM 호출은 STA 실행 컨텍스트에서 직렬화합니다.
- COM 호출이 45초를 넘으면 해당 서비스를 격리하고 앱 재시작을 안내합니다.
- LLM이 임의의 C# 또는 COM 코드를 직접 실행하지 않도록 구조화 명령만 허용합니다.
- PPT 원본은 기본적으로 LLM 서버에 전송하지 않고 필요한 구조화 문맥만 사용합니다.
- 자동화 결과는 원본을 덮어쓰지 않고 사본으로 저장합니다.

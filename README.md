# BckPPT

BckPPT는 기존 PowerPoint 슬라이드의 레이아웃과 서식을 유지하면서 텍스트를 분석·교체하는 Windows 애플리케이션입니다. PowerPoint가 실제로 해석한 문단, 강제 줄바꿈, 자동 줄바꿈과 텍스트 영역을 기준으로 결과를 검증하는 데 초점을 둡니다.

현재는 전문 편집 도구를 제공하며, 앞으로 `gitoky.com`의 `llm_api`를 사용하는 독립형 PPT 에이전트로 확장합니다. 프로그램 실행 시 채팅창에서 작업을 요청하고 필요할 때 기존 전문 편집창을 여는 구조가 목표입니다. CellBox2에 포함되거나 의존하지 않습니다.

## 현재 기능

- PowerPoint 파일 열기와 슬라이드 탐색
- 슬라이드·Shape·텍스트 계층 구조 분석
- 텍스트 상자별 줄 수와 줄바꿈 유형 식별
  - 자동 줄바꿈
  - Enter 문단 나눔
  - Shift+Enter 강제 줄바꿈
- 문단별 텍스트 교체와 원본 서식 유지
- overflow 감지와 AutoFit 보정
- Shape 역할 분류와 구조 정보 저장
- 변경 이력 확인
- localhost 디버그 API를 통한 분석·편집 자동화

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

## 직접 빌드

```powershell
dotnet restore BckPPT.csproj
dotnet build BckPPT.csproj -c Debug -p:Platform=x64
```

빌드 결과는 `bin/x64/Debug/net8.0-windows/`에 생성됩니다.

## 프로젝트 구조

```text
BckPPT/
├─ Models/       PPT, 슬라이드, Shape 및 분석 결과 모델
├─ Services/     PowerPoint COM, 구조 저장, 역할 분류, 변경 이력
├─ ViewModels/   메인 편집 작업 상태와 명령
└─ Views/        메인창과 분석·편집 보조창
Assets/          애플리케이션 아이콘과 원본 자산
docs/plans/      아키텍처 및 리팩토링 계획
```

## 안전 원칙

- PowerPoint COM 호출은 STA 실행 컨텍스트에서 직렬화합니다.
- LLM이 임의의 C# 또는 COM 코드를 직접 실행하지 않도록 구조화 명령만 허용할 예정입니다.
- PPT 원본은 기본적으로 LLM 서버에 전송하지 않고 필요한 구조화 문맥만 사용합니다.
- 자동화 결과는 원본 덮어쓰기보다 사본 저장을 우선하는 방향으로 개발합니다.

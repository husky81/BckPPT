# BckPPT를 CellBox 플러그인으로 매립하는 계획

> 작성일: 2026-09-15
> 대상 저장소: BckPPT, cellbox2 (두 곳 모두 수정)
> 상태: **1단계(중복 제거) 완료** · 2단계부터 구현 대기
> 방향: CellBox 본체는 가볍게 두고, 플러그인은 **cellbox.ai 스토어**에서 내려받아 설치한다.

## 1. 결정 사항

| # | 결정 | 내용 |
|---|---|---|
| 1 | 소스 중복 해소 | **바이너리 아티팩트** — cellbox2는 BckPPT 소스를 갖지 않고 zip만 소비 |
| 2 | `installer/bckppt.iss` | **폐기** — 독립 Inno 설치관리자를 없애고 플러그인 체계로 일원화 |
| 3 | 배포 방식 | **스토어형** — CellBox만 설치하고, 플러그인은 cellbox.ai에서 필요할 때 받는다 |

3번 결정이 나머지를 자연스럽게 정리해 줍니다. 플러그인이 설치본에 들어가지 않으므로
**CellBox 설치관리자가 오히려 가벼워지고**, self-contained 용량 문제도 "받을 사람만 받는" 문제로 바뀝니다.

## 2. 1단계 — 소스 중복 제거 (완료)

### 왜 먼저 했나

BckPPT 소스가 두 벌이었고 이미 갈라져 있었습니다.

```
BckPPT/                       v2.1.0    ← 실제 개발하는 곳
cellbox2/apps/tray/BckPPT/    v2.0.44   ← 트레이가 빌드해 쓰던 곳
```

트레이에서 BckPPT를 누르면 2.0.44가 떴습니다. 이 상태로 메뉴에 버전을 붙이면
그 버전이 무엇을 가리키는지 모호해지므로, 플러그인화의 선행 조건이었습니다.

### 삭제해도 안전한지 확인한 방법

두 소스를 파일 단위·내용 단위로 대조했습니다.

- 트레이 쪽에만 있는 **파일은 하나도 없음** (독립형이 상위집합, 40여 개 파일이 더 많음)
- 내용 차이는 전부 **네임스페이스 재작성**(`CellBox2.BckPPT.*` → `BckPPT.*`) 아니면 **구버전 코드**
- 트레이 통합을 위한 고유 코드는 없었음

### 실제로 지운 것

| 대상 | 내용 |
|---|---|
| `apps/tray/BckPPT/` | 소스 55개 파일 (git 추적분) + bin/obj 잔여물, 합계 174MB |
| `installer/bckppt.iss` | 독립 Inno Setup 설치관리자 (결정 2) |
| `apps/tray/CellBox2.sln` | BckPPT 프로젝트 항목 + 빌드 구성 12줄 |
| `installer/build_client.ps1` | `[2d] BckPPT - Publish` 단계 35줄 + 빌드 요약 출력 |
| `.vscode/launch.json` | `BckPPT (C#)` 디버그 구성 |
| `.vscode/tasks.json` | `bckppt: run`, `bckppt: pre-launch` 태스크 |
| `.oky/scripts.json` | BckPPT 실행 스크립트 항목 |

검증: `dotnet sln list`로 솔루션 정상 파싱 확인, `CellBox2.csproj` 빌드 오류 0개,
수정한 JSON 3개 모두 파싱 정상.

### 건드리지 않은 것

- **`TrayManager.cs:83-104`** — BckPPT 메뉴 항목의 하드코딩.
  `File.Exists(BaseDirectory\BckPPT.exe)` 가드가 있어 지금은 **항목이 조용히 사라진 상태**입니다.
  다음 단계에서 이 자리를 플러그인 목록 기반으로 교체합니다. 교체 지점을 남겨두려고 일부러 두었습니다.
- **`apps/cellbox-mcp`, `apps/bckppt-mcp`** — BckPPT의 localhost 디버그 API(38700)를 부르는
  런타임 연동입니다. 소스 의존이 아니라 그대로 둡니다.

> **중간 상태 주의**: 지금 CellBox를 빌드하면 트레이에 BckPPT 항목이 없습니다.
> 3단계(플러그인 서비스)가 들어가야 다시 나타납니다.

## 3. 목표 구조

```
   BckPPT 레포
        │ 빌드 + 패키징
        ↓
   bckppt-2.1.0.zip + plugin.json          ──→  릴리스 버킷
                                                  plugins/bckppt/
                                                       ↑ 조회          ↓ 다운로드
   cellbox.ai/store  ──────────────────────────────────┤               │
   (카탈로그 웹)                                        │               │
                                                        │               │
   CellBox 트레이  ──── 설치 / 업데이트 / 제거 ─────────┴───────────────┘
        ↓
   %LOCALAPPDATA%\CellBox\Plugins\BckPPT\
```

핵심은 **CellBox 설치본에 플러그인이 들어가지 않는다**는 것입니다.
사용자는 CellBox만 설치하고, 필요한 플러그인을 나중에 고릅니다.

### 3.1 설치 위치

```
%LOCALAPPDATA%\CellBox\Plugins\BckPPT\
    plugin.json
    BckPPT.exe
    (self-contained 런타임 일체)
```

`Program Files`가 아니라 `%LOCALAPPDATA%`인 이유가 핵심입니다.
**관리자 권한 없이 설치·교체·제거가 됩니다.** 스토어에서 버튼 하나로 설치하고
버튼 하나로 지우려면 UAC 프롬프트가 끼면 안 됩니다.

self-contained로 담으므로 CellBox2의 .NET 런타임에 얹히지 않습니다.
**CellBox 업데이트가 플러그인을 깨뜨리지 않는다**는 뜻이고, 이게 플러그인 독립성의 실체입니다.

### 3.2 플러그인 규약

트레이가 BckPPT를 하드코딩하지 않도록 매니페스트를 정합니다.

```json
// plugin.json — 플러그인 폴더 루트
{
  "id": "bckppt",
  "name": "BckPPT",
  "version": "2.1.0",
  "description": "PPT 텍스트 분석·치환 도구",
  "executable": "BckPPT.exe",
  "arguments": "",
  "icon": "Assets/icon.ico",
  "processName": "BckPPT",
  "updateManifestUrl": "https://<release-bucket>/plugins/bckppt/latest.json",
  "dataDirs": [
    "%LOCALAPPDATA%\\BckSoft\\BckPPT"
  ],
  "minCellBoxVersion": "1.0.0"
}
```

두 필드가 특히 중요합니다.

**`dataDirs`** — 플러그인이 **자기 데이터 위치를 스스로 신고**합니다.
덕분에 트레이는 BckPPT 내부를 몰라도 완전 삭제를 할 수 있습니다.
BckPPT가 실제로 쓰는 곳은 `%LOCALAPPDATA%\BckSoft\BckPPT\` 로,
`history.jsonl`(치환 이력) · `marks.json`(표시 스냅샷) · `crash.log` · 템플릿 SQLite 색인이 들어갑니다.

**`processName`** — 제거·교체 전에 실행 중인 프로세스를 멈추는 데 씁니다.
BckPPT는 PowerPoint COM과 named pipe, localhost 디버그 API를 잡고 있어서
그냥 파일을 지우면 실패합니다.

## 4. 트레이 메뉴

```
CellBox v1.2.3 · 09/15 14:20
──────────────────────────────
Open Web Drive
Open Sync Folder
Settings
──────────────────────────────
BckPPT  v2.1.0                 ← 클릭 → Agent 창
플러그인...                     ← 관리 UI
──────────────────────────────
일시정지 ▸
종료
```

| 상태 | 메뉴 텍스트 | 클릭 동작 |
|---|---|---|
| 설치됨·최신 | `BckPPT  v2.1.0` | Agent 실행 |
| 설치됨·업데이트 있음 | `BckPPT  v2.1.0 → 2.2.0` | Agent 실행 (업데이트는 관리 UI) |
| 실행 중 | `BckPPT  v2.1.0 ●` | 기존 창 활성화 |
| 하나도 없음 | `플러그인 둘러보기...` | 스토어 열기 |

설치된 플러그인이 없으면 개별 항목 대신 스토어로 유도합니다.

버전은 `plugin.json`에서 읽습니다. 어셈블리를 로드하지 않으므로 메뉴 여는 속도에 영향이 없습니다.

"실행 중 → 기존 창 활성화"는 BckPPT의 `SingleInstanceCoordinator`가 이미 처리합니다.
두 번째 인스턴스가 뜨면 named pipe로 넘기고 스스로 종료합니다.

### 4.1 Agent 창으로 뜨는 문제

독립형 BckPPT는 이미 `AgentChatWindow`가 시작 화면입니다.
중복을 없앤 지금, 독립형 아티팩트를 쓰는 순간 **자동으로 해결됩니다.** 따로 할 일이 없습니다.

### 4.2 `--sync-folder` 인자

기존 트레이는 `--sync-folder "<경로>"` 를 넘겼지만 **독립형 BckPPT는 이 인자를 해석하지 않습니다**
(`App.OnStartup`은 JSON 외부 명령만 읽음). 지금까지 조용히 무시되고 있었습니다.

권장: BckPPT가 인자를 받아 **`열기` 대화상자의 기본 경로**로 쓰게 합니다.
동기화 폴더의 pptx를 바로 여는 흐름이 자연스러워집니다.

## 5. 플러그인 관리 UI

`SettingsWindow`의 NavigationView에 **`플러그인`** 페이지를 추가합니다.
이미 `일반 / 동기화 / 도구 / 고급` 구조가 있어 자리가 자연스럽습니다.

```
┌─ 플러그인 ──────────────────────────────────────────────┐
│                                        [스토어 열기 ↗]   │
│  설치됨                                                  │
│  ┌────────────────────────────────────────────────────┐  │
│  │ [아이콘]  BckPPT                       v2.1.0      │  │
│  │           PPT 텍스트 분석·치환 도구                  │  │
│  │           142MB · 설치 2026-09-15 · ● 실행 중       │  │
│  │                                                     │  │
│  │     [실행]  [업데이트 확인]  [재설치]  [제거]        │  │
│  └────────────────────────────────────────────────────┘  │
│                                                          │
│  설치 가능                                                │
│  ┌────────────────────────────────────────────────────┐  │
│  │ [아이콘]  BckHWP              v0.9.1   [설치]      │  │
│  └────────────────────────────────────────────────────┘  │
└──────────────────────────────────────────────────────────┘
```

### 5.1 제거

```
1. 실행 중이면 종료          processName으로 찾아 CloseMainWindow → 5초 대기 → Kill
2. 플러그인 폴더 삭제         %LOCALAPPDATA%\CellBox\Plugins\BckPPT\
3. (선택) 데이터 삭제         dataDirs
4. 트레이 메뉴 갱신
```

```
BckPPT를 제거할까요?

  ☐ 설정과 변경 이력도 함께 삭제
     (%LOCALAPPDATA%\BckSoft\BckPPT — 치환 이력, 표시 스냅샷, 규칙)

  체크를 해제하면 다시 설치했을 때 이전 설정이 복원됩니다.

                              [제거]  [취소]
```

**이 체크박스가 "깨끗한 제거"와 "재설치 편의"를 동시에 만족시키는 지점입니다.**
기본값은 해제 — 실수로 이력을 날리지 않게.

### 5.2 재설치

같은 버전을 다시 받아 덮어씁니다. 파일이 깨졌거나 업데이트가 반쯤 실패했을 때 쓰는 복구 버튼입니다.
데이터는 **건드리지 않습니다** — 그게 제거와의 차이입니다.

### 5.3 설치 · 업데이트

```
1. manifest 조회              {version, packageUrl, sha256}
2. zip 다운로드 → SHA-256 검증  검증 실패 시 중단
3. 실행 중이면 종료
4. 기존 폴더를 .old 로 이름 변경   (업데이트일 때)
5. 새 폴더 전개
6. 성공하면 .old 삭제, 실패하면 .old 를 되돌림
```

4~6단계의 **이름 변경** 방식이 중요합니다. 중간에 실패해도 이전 버전이 살아 있습니다.
덮어쓰기는 반쯤 망가진 상태를 만듭니다.

## 6. 스토어

### 6.1 배포 경로

```
https://<release-bucket>/
    releases/latest.json              ← CellBox 본체 (기존)
    plugins/index.json                ← 플러그인 카탈로그 (신규)
    plugins/bckppt/latest.json        ← 신규
    plugins/bckppt/bckppt-2.1.0.zip
```

```json
// plugins/bckppt/latest.json — BckPPT의 기존 UpdateManifest 형식 그대로
{
  "version": "2.1.0",
  "packageUrl": "https://<release-bucket>/plugins/bckppt/bckppt-2.1.0.zip",
  "sha256": "…64자…",
  "releaseNotes": "문단 단위 치환, 관용 파서, 치환 표시"
}
```

```json
// plugins/index.json — 카탈로그
{
  "plugins": [
    {
      "id": "bckppt",
      "name": "BckPPT",
      "description": "PPT 텍스트 분석·치환 도구",
      "icon": "https://<release-bucket>/plugins/bckppt/icon.png",
      "manifestUrl": "https://<release-bucket>/plugins/bckppt/latest.json",
      "sizeBytes": 148897792,
      "minCellBoxVersion": "1.0.0"
    }
  ]
}
```

### 6.2 서버 코드 변경 0

트레이 본체 업데이트는 `{apiUrl}/api/v1/version/latest` → S3 `releases/latest.json` 경로를 쓰는데,
**단일 제품 전용**이라 플러그인을 태울 수 없습니다.

그런데 BckPPT의 `UpdateService`가 이미 **임의의 manifest URL**을 읽도록 만들어져 있습니다
(`BCKPPT_UPDATE_MANIFEST_URL`, `{version, packageUrl, sha256, releaseNotes}` 형식).
릴리스 버킷에 정적 파일만 올리면 되고 **Go 서버는 손댈 필요가 없습니다.**

### 6.3 웹 스토어 페이지

`cellbox.ai/store` (또는 `/plugins`). `apps/web/src/app/` 에 이미 `download/` 가 있어 형제로 둡니다.

`plugins/index.json`을 읽어 카드로 렌더링하고, 각 카드에 두 경로를 둡니다.

- **CellBox가 설치돼 있으면** — `cellbox://plugin/install?id=bckppt` 커스텀 스킴으로
  트레이를 깨워 설치 UI를 띄운다
- **아니면** — CellBox 설치 페이지로 유도

커스텀 스킴 등록은 CellBox 설치관리자가 담당합니다. 이게 "스토어에서 클릭 → 바로 설치"를
만드는 연결 고리입니다.

## 7. 남은 작업

### BckPPT 레포

| 작업 | 내용 |
|---|---|
| 패키징 스크립트 | `scripts/package-plugin.ps1` — publish → `plugin.json` 생성 → zip → sha256 산출 |
| 인자 처리 | `--sync-folder` 수용 (`열기` 기본 경로로 사용) — 4.2 |
| 아이콘 내보내기 | 스토어 카드용 `icon.png` |

### cellbox2 레포

| 작업 | 내용 |
|---|---|
| 플러그인 모델 | `Models/PluginManifest.cs`, `Models/PluginState.cs` |
| 플러그인 서비스 | `Services/PluginService.cs` — 검색·설치·제거·재설치·업데이트·실행 |
| 트레이 메뉴 | `TrayManager.cs:83-104` 하드코딩을 플러그인 목록 기반으로 교체, 버전 표시 |
| 관리 UI | `Views/Pages/PluginsPage.xaml` + SettingsWindow NavigationView 항목 |
| 제거 확인 | 데이터 삭제 체크박스가 있는 ContentDialog |
| 커스텀 스킴 | `cellbox://` 등록 + `plugin/install` 처리 |
| 스토어 페이지 | `apps/web/src/app/store/page.tsx` |

### 릴리스 인프라

| 작업 | 내용 |
|---|---|
| 버킷 경로 | `plugins/bckppt/`, `plugins/index.json` 신설 |
| CI | BckPPT 태그 푸시 → zip + latest.json 업로드 + index.json 갱신 |

## 8. 권장 순서

**2단계 — 수동 설치로 동작 확인**
패키징 스크립트로 zip을 만들고 `%LOCALAPPDATA%\CellBox\Plugins\BckPPT\`에 손으로 풀어,
트레이 메뉴가 `plugin.json`을 읽어 `BckPPT v2.1.0`을 띄우고 Agent가 실행되는지 본다.
**스토어 없이 규약만 먼저 검증**하는 단계다.

**3단계 — 관리 UI**
설치·제거·재설치·업데이트를 UI에서. 이 시점에 버킷에 manifest를 올린다.

**4단계 — 스토어**
웹 카탈로그 + `cellbox://` 스킴 연결.

2단계가 가장 값어치가 큽니다. 여기서 `plugin.json` 규약과 프로세스 종료·폴더 교체가 맞는지
확인하고 나면 나머지는 그 위에 얹는 일입니다.

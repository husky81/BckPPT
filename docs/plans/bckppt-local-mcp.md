# BckPPT 로컬 MCP와 그림 칸·판 비교·렌더 도구 계획

작성일: 2026-10-30

## 배경

로컬에서 실행되는 에이전트(Claude Code, Codex 등)가 PPT 작업을 하며 겪은 불편과 사고를 정리하면 다음과 같다.

1. **그림 칸 채우기와 판 관리**
   - 매번 python-pptx 스크립트를 새로 짰다(`insert_images`, `finalize`, `finalize_codex`).
   - 에이전트가 만든 폴더 이름과 스크립트 속 폴더 이름이 달라 5쪽·9쪽이 비는 사고가 났다.
   - 사용자가 고친 판에서 무엇이 바뀌었는지는 도형 목록을 직접 뽑아 비교했다.
2. **확인용 렌더링**
   - PowerPoint로 쪽을 그림으로 내보낸 뒤 프로그램을 종료하지 않았다.
   - 숨은 PowerPoint가 파일을 잡아 판 이름이 보완2·3·4로 불어났고, 파일이 열리지 않거나 문서 복구 창이 뜨는 문제로 이어졌다.

## 현재 코드에서 확인한 점

- `PptComService.GetOrCreateInstance()`는 실행 중인 PowerPoint에 붙고, 없으면 보이는 인스턴스를 새로 만든 뒤 종료하지 않는다.
  썸네일·내보내기(`slide.Export`)도 원본 파일을 직접 연다. 렌더링 사고의 직접 원인이다.
- PowerPoint는 사용자당 단일 인스턴스 COM 서버라 "별도 PowerPoint"를 띄울 수 없다. `CreateInstance`도 실행 중인 인스턴스를 돌려준다.
  따라서 격리는 프로세스가 아니라 **사본 파일**로 한다.
- `Agent/Tools/PptAgentTools.cs`에 JSON Schema가 붙은 에이전트 도구가 약 40개 있다. 그대로 MCP 도구로 노출할 수 있다.
- `cellbox2\apps\bckppt-mcp`(Node, 2026-04)는 디버그 API(`localhost:38700`)를 감싼 8개 도구뿐이고, BckPPT 창이 떠 있어야 동작한다.
  BckPPT는 CellBox2에 의존하지 않는다는 원칙과도 맞지 않는다.

## 결정: `BckPPT.exe --mcp` (stdio)

- C# MCP SDK(`ModelContextProtocol` NuGet)로 stdio 서버를 띄운다. WPF 창 없이 COM 전용 STA 스레드 하나를 쓴다.
- `PptAgentTools`의 도구 목록을 MCP 도구로 노출한다. 에이전트 세션의 `WorkingPath` 대신 도구마다 `path` 인자를 받는다.
- GUI BckPPT와 MCP 프로세스가 동시에 떠도 COM 호출이 겹치지 않도록 프로세스 간 뮤텍스 `Global\BckPPT.PowerPoint`로 직렬화한다.
- 쓰기 도구는 기존 안전 원칙대로 사본 저장만 허용한다(`output_path` 필수, 원본 덮어쓰기 금지).
- 서버 안내문(instructions)은 BckHWP와 같은 형식으로 쓴다.
  - BckPPT가 PowerPoint를 다루는 유일한 통로다. python-pptx·PowerShell COM으로 직접 만들지 않는다.
  - 시간 초과 시 다시 보내지 말고 상태를 조회해 사용자에게 알린다.
  - 필요한 기능이 없으면 우회하지 말고 없는 기능을 사용자에게 알린다.
- 등록 예:
  - Claude Code: `claude mcp add bckppt -- "C:\...\BckPPT.exe" --mcp`
  - Codex: `config.toml`의 `[mcp_servers.bckppt]`에 같은 명령
- 정착 후 `cellbox2\apps\bckppt-mcp`는 폐기하거나 이 서버를 가리키도록 바꾼다.

## 새 도구

### `ppt_render` — PPT 몇 쪽 → 그림

- 입력: `path`, `slides`(번호 목록 또는 범위), `out_dir`, `width?`
- 출력: 쪽별 PNG 경로
- 규칙
  - 원본을 임시 폴더로 복사해 사본을 연다(`ReadOnly`, `WithWindow=false`). 원본은 잠기지 않는다.
  - BckPPT가 PowerPoint를 새로 띄웠으면 끝에 반드시 `Quit`한다. 사용자가 쓰던 인스턴스면 연 사본만 `Close`한다.
  - `finally`에서 Close·Release·Quit을 각각 따로 감싸 하나가 실패해도 나머지가 실행되게 한다.
  - 시간 제한을 넘기면 BckPPT가 연 사본만 정리하고 실패를 돌려준다.
  - 원본이 잠겨 있으면 Restart Manager(`RmGetList`)로 잡고 있는 프로세스 이름·PID를 알려 주고 멈춘다.
- 기존 `render_slide`, Markdown 내보내기 미리보기도 같은 경로를 쓰도록 바꾼다.

### `ppt_place_images` — 「쪽·칸 → 그림 파일」 표로 넣기·바꾸기

- 입력: `path`, `output_path`, `items: [{slide, cell, image, fit?}]`, `dry_run?`
  - `cell`: 도형 주소(README 「도형 주소」 규약) 또는 그림 자리 이름
  - `fit`: `contain` | `cover` | `stretch`(기본 `contain`)
- 2단계 실행
  1. **검사**: 없는 그림 파일, 없는 칸, 같은 칸 중복, 표에 빠진 그림 칸(빈 쪽)을 모두 나열한다. 하나라도 있으면 아무것도 바꾸지 않는다.
  2. **적용**: 사본에 넣고 쪽마다 결과(칸·파일·크기)를 돌려준다. 필요하면 `ppt_render`로 확인 그림까지 만든다.
- `output_path`가 이미 있거나 잠겨 있으면 실패로 돌려준다. 이름 뒤에 숫자를 자동으로 붙이지 않는다(보완2·3·4 방지).
- 보조 도구 `ppt_list_image_slots`: 쪽마다 그림 칸(그림 도형·그림 자리 표시자)의 주소·위치·크기.

### `ppt_diff` — 두 판 사이 도형 변경 비교

- 입력: `before`, `after`, `slides?`
- PowerPoint 없이 Open XML로 읽는다. 빠르고 파일을 잠그지 않는다.
- 도형 매칭: 같은 쪽 안에서 `cNvPr id` → 이름 → 위치 근접 순.
- 보고 항목: 추가·삭제·이동·크기 변경·글 변경(문단 단위)·그림 교체(blip 해시)·쪽 추가·삭제·순서 변경.

## 단계

1. **렌더 경로 정리**: 사본 렌더, 띄운 인스턴스만 `Quit`, 잠금 원인 보고. 기존 렌더 호출부 교체.
2. **MCP 뼈대**: `--mcp` 진입점, COM 뮤텍스, 기존 도구 노출, README에 등록 방법.
3. `ppt_render`, `ppt_diff`.
4. `ppt_place_images`, `ppt_list_image_slots`.

## 검증

- 렌더: PowerPoint를 미리 띄운 경우와 띄우지 않은 경우 모두, 렌더 뒤 `POWERPNT.EXE` 프로세스 수가 렌더 전과 같은지 확인. 원본을 즉시 다른 이름으로 저장·삭제할 수 있는지 확인.
- 잠금: 원본을 PowerPoint로 열어 둔 채 렌더·저장을 요청하면 잠근 프로세스가 보고되는지 확인.
- 그림 칸: 파일 하나를 빼고 실행하면 아무것도 바뀌지 않고 누락이 보고되는지 확인.
- 비교: 도형 하나를 옮기고, 글 하나를 바꾸고, 그림 하나를 교체한 판으로 세 변경이 모두 보고되는지 확인.

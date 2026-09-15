# 문단 단위 텍스트 치환 워크플로우 계획

> 작성일: 2026-09-15
> 대상 저장소: BckPPT
> 상태: **Phase 1 구현 완료** (설계와 달라진 점은 8장 참고)
> 배경: `로드맵_취합본_2608_정규산_김병철001_2-1분리.pptx` 보완 작업을 외부 스크립트로 처리한 사례를
> BckPPT UI 안에서 재현 가능하게 만든다.

## 1. 목표 워크플로우

```text
[1] run.ps1                → BckPPT 실행
[2] 파일 탑재              → "열기" 버튼 / 드래그&드롭 / PowerPoint 열린 파일 선택
[3] "텍스트 추출" 클릭     → 문단마다 ID가 붙은 텍스트 창이 열림
[4] "프롬프트+텍스트 복사" → 외부 에이전트에 붙여넣고 수정 요청
[5] 에이전트 답변 붙여넣기 → 설명문·표·코드블록이 섞여 있어도 그대로
[6] 자동 파싱 → 미리보기   → 변경표에서 항목별 체크
[7] "반영" 클릭            → PPT 적용 + 치환된 문단만 빨간색 표시
[8] 검토 후 "표시 해제"    → 원래 글자색 복원
```

핵심 설계 원칙은 **1:1 대칭 포맷을 강요하지 않는 것**이다. 사용자가 다른 에이전트의 답변을
가공 없이 붙여넣어도, 파서가 ID를 단서로 치환 대상만 추려낸다.

## 2. 현재 상태와 격차

| 단계 | 현재 구현 | 격차 |
|---|---|---|
| 파일 탑재 | `GetOpenPresentations()` — PowerPoint에 **이미 열린** 파일만 콤보박스에 표시 | 파일 열기 버튼·드래그&드롭 없음. 원본 보호(사본 작업) 없음 |
| 텍스트 추출 | `TextEditorWindow` — `GetTextsForCurrentView()`가 **shape 단위** `1: 전체텍스트` 생성, 문단 구분은 `\n` 이스케이프 | 문단 단위 ID 없음. 긴 도형은 한 줄에 뭉쳐 나옴 |
| 결과 반영 | `Apply_Click`의 `^(\d+):\s*(.*)$` — **완전 일치 형식만** 파싱 | 설명문·마크다운 표·코드펜스가 섞이면 전부 무시됨 |
| 미리보기 | `MessageBox` 문자열 diff | 항목별 취사선택 불가, 줄 넘침 사전 경고 없음 |
| 적용 | `SetShapeText()` — shape 전체 텍스트를 `\r`로 쪼개 문단별 적용, Bullet/Indent/Font 복원 | 문단 단위 진입점 없음 |
| 색 표시 | 없음 | 치환 흔적을 눈으로 확인할 방법 없음 |
| 되돌리기 | `_undoSnapshot` (shape 단위 1회) | 문단 단위 대응 필요 |

재사용할 수 있는 자산이 이미 많다. `SetShapeText()`의 문단 루프(백업 → `para.Text` 교체 →
Bullet·IndentLevel·Font 복원)와 `GetTextLineBreaks()`의 문단 순회 로직이 그대로 기반이 된다.

## 3. 구현 항목

### A. 문단 단위 ID 체계

**ID 형식: `S{슬라이드}-{도형순번}-{문단번호}`** — 예) `S1-12-3`

설계 제약에서 나온 형식이다.

- 산문 한가운데서도 정규식 `\bS(\d+)-(\d+)-(\d+)\b` 로 뽑을 수 있어야 한다.
- **`|` 를 쓰면 안 된다.** 마크다운 표 구분자와 충돌해 표 형태로 답한 에이전트의 출력을 못 읽는다.
- 도형순번은 창을 열 때 확정하고 창이 닫힐 때까지 고정한다. 도형 이름(`화살표: 오각형 69`)에는
  공백·콜론·한글이 섞여 있어 ID에 직접 넣기 부적합하므로, 창 내부 맵에서만 이름과 연결한다.

신규 모델:

```csharp
// BckPPT/Models/ParagraphEntry.cs
public record ParagraphEntry(
    string Id,              // "S1-12-3"
    int    SlideIndex,
    string ShapeName,       // 테이블 셀은 "Table 1 [R2C3]" 형식 그대로
    int    ParagraphIndex,  // 1-based (PowerPoint COM 규약)
    string Text,
    bool   IsTableCell);
```

신규 COM 메서드:

```csharp
IReadOnlyList<ParagraphEntry> GetParagraphs(string presentationPath, int? slideIndex = null);
```

`GetTextLineBreaks()`가 이미 `TextFrame2.TextRange.Paragraphs`를 순회하므로 그 순회 코드를
공용 헬퍼로 빼고 두 메서드가 공유한다. 빈 문단은 ID는 부여하되 목록에서는 `(빈 문단)`으로 표시한다.

### B. 관용 파서 — `PastedPlanParser`

**LLM 없이 결정적으로 동작한다.** 신규 클래스 + 단위 테스트로 만든다.

```csharp
// BckPPT/Services/PastedPlanParser.cs
public static ParseResult Parse(string pasted, IReadOnlyList<ParagraphEntry> known);

public record ParseResult(
    IReadOnlyList<ParsedChange> Changes,
    IReadOnlyList<string>       Warnings,      // 중복 ID, 목록에 없는 ID
    IReadOnlyList<string>       UnmatchedLines // ID처럼 보이지만 해석 실패한 줄
);
```

처리 규칙 (순서대로 적용):

1. 전체를 줄 단위로 스캔한다. 코드펜스(` ``` `)는 **경계로 쓰지 않고 그냥 제거**한다.
   에이전트마다 코드블록 안에 넣기도 하고 밖에 두기도 하기 때문이다.
2. 각 줄에서 첫 ID 토큰을 찾는다. 없으면 그 줄은 버린다 (설명 문장은 자연히 걸러진다).
3. 마크다운 표 행(`| S1-12-3 | 새 텍스트 | ... |`)이면 셀로 분해해 ID 다음 셀을 값으로 쓴다.
4. 일반 줄이면 ID 뒤의 구분자(`:` `：` 탭 · 2칸 이상 공백 · `→` · `=`) 다음을 값으로 쓴다.
5. 값에서 마크다운 장식을 제거한다 — 앞쪽 `-` `*` `>` `#`, 감싼 백틱과 `**`.
6. 값이 비었거나 `""` · `(빈 문단)` · `(삭제)` 이면 **빈 문자열 치환**으로 인정한다.
   (이번 작업의 `s1|96|p3` 케이스)
7. 같은 ID가 여러 번 나오면 마지막 것을 채택하고 경고에 남긴다.
   에이전트가 "변경 전/후"를 나란히 쓰는 경우가 흔하므로 필수 규칙이다.
8. 목록에 없는 ID는 무시하고 경고에 남긴다.
9. 원본과 동일한 값은 변경 목록에서 뺀다.

**문단 추가 지원**: `S1-12-3+` 처럼 ID 뒤 `+`는 "해당 문단 뒤에 새 문단 삽입"으로 해석한다.
이번 작업의 `s1|70` 케이스(원본이 1문단인데 2줄이 필요)가 여기 해당한다.
구현은 원본 문단을 복제한 뒤 텍스트만 바꾸는 방식이라 정렬·글꼴·크기가 자동으로 보존된다.

테스트는 이번 작업에서 실제로 주고받은 답변 형식을 그대로 픽스처로 쓴다 —
마크다운 표, 코드펜스 안의 `ID<탭>텍스트` 목록, 설명 문단이 섞인 형태.

### C. 변경 미리보기 창 — `ChangePreviewWindow` (신규)

`MessageBox` 문자열 diff를 DataGrid로 교체한다.

| 열 | 내용 |
|---|---|
| ☑ | 항목별 적용 여부 (기본 체크) |
| ID | `S1-12-3` |
| 위치 | 슬라이드 / 도형 이름 |
| 이전 → 이후 | 변경 내용 |
| ⚠ | **줄 넘침 예상** |

넘침 예상 열이 이 창의 존재 이유다. 이번 작업에서 반영 후에야 발견한 4건(`s1|96`, `s1|4`,
`s2|290`, `s2|145`)을 **반영 전에** 잡아낸다. 판정은 기존 자산을 그대로 쓴다 —
`PptSlideInfo.CalcHalfWidth()`로 새 텍스트의 반각 길이를 구하고, 도형의 `MaxLineHw`와
`BoundHeight / LineCount`(줄 높이)로 예상 줄 수를 계산해 `BoxHeight`와 비교한다.

하단에 경고 패널(파서 경고 + 해석 실패 줄)을 두고, `[전체 선택] [반영] [취소]` 를 배치한다.

### D. 빨간색 표시

신규 COM 메서드:

```csharp
// 적용 + 표시. 반환값은 문단별 원래 색 (복원용)
IReadOnlyList<ParagraphColorSnapshot> SetParagraphText(
    string presentationPath,
    IReadOnlyList<(int SlideIndex, string ShapeName, int ParagraphIndex, string NewText, bool Insert)> changes,
    bool markRed);

// 표시 해제 — 스냅샷의 원래 색으로 복원
void RestoreParagraphColors(string presentationPath, IReadOnlyList<ParagraphColorSnapshot> snapshots);
```

구현 시 주의점:

- 문단 접근은 기존 `SetShapeText()`와 동일하게 `tr.Paragraphs(i + 1, 1)`.
- **PowerPoint COM의 RGB는 BGR 순서다.** 빨강은 `0x0000FF`(십진 255)이지 `0xFF0000`이 아니다.
- 적용 **전에** `para.Font.Fill.ForeColor.RGB`와 `.ObjectThemeColor`를 함께 저장한다.
  테마색이 걸린 문단은 RGB만 복원하면 테마 연결이 끊긴다.
- Bullet → IndentLevel → Font 순서로 복원하는 기존 규칙(`PptComService.cs:772` 주석의
  "순서 필수")을 그대로 지킨다. 색 설정은 그 뒤에 한다.
- 도형 단위로 한 번만 `Presentations.Open` / `Save` 하도록 변경 목록을 슬라이드·도형 기준으로
  묶어서 처리한다. 문단마다 여는 방식은 느리다.

스냅샷은 `ChangeHistoryEntry`에 필드를 추가해 함께 저장한다. 그래야 BckPPT를 껐다 켜도
"표시 해제"가 동작한다.

### E. 파일 탑재

- 헤더에 **"열기"** 버튼 추가 → `OpenFileDialog` → 기존 `OpenPptxInPowerPoint(path)` 호출 →
  `Refresh()` → 해당 파일 자동 선택. 필요한 COM 메서드는 이미 있다.
- `BckPptWindow`에 `AllowDrop="True"` + `Drop` 핸들러. `.pptx`/`.ppt` 확장자만 수락.
- **"사본으로 작업" 체크박스** (기본 켜짐): 켜져 있으면 탑재 시
  `{원본명}_{yyyyMMdd-HHmm}.pptx`로 복사한 뒤 사본을 연다. 이번 작업에서 실제로 필요했던
  동작이고, 원본이 PowerPoint에서 열려 있어 잠겨 있을 때도 안전하다.

### F. LLM 연동 — 선택 사항

**핵심 흐름에는 필요 없다.** B의 파서는 정규식 기반이라 API 호출이 없고, 오프라인에서도 동작한다.

LLM을 붙일 지점은 두 곳뿐이며 둘 다 나중에 켜도 되는 부가기능이다.

1. **파서 실패 복구** — 에이전트가 ID 없이 "96번 도형 첫 줄을 X로 바꿔줘"처럼 산문으로만
   답한 경우. 미리보기 창에 변경이 0건이면 "LLM으로 해석" 버튼을 노출하고,
   붙여넣은 원문 + 문단 목록을 보내 `ID: 텍스트` 목록으로 변환시킨다.
2. **창 안에서 직접 수정 요청** — 외부 에이전트를 거치지 않고 BckPPT에서 바로 요청.

`GitokyLlmClient`(OpenAI 호환 `/v1/`, SSE 스트리밍, 재시도 2회)와 `PptPromptBuilder`가
이미 구현돼 있고 `.env`의 `GITOKY_LLM_API_KEY`만 채우면 동작한다. **신규 API 작업은 없다.**

## 4. 파일별 작업 목록

| 파일 | 작업 |
|---|---|
| `Models/ParagraphEntry.cs` | 신규 |
| `Models/ParagraphColorSnapshot.cs` | 신규 |
| `Models/ChangeHistoryEntry.cs` | 원래 색 필드 추가 |
| `Services/PastedPlanParser.cs` | 신규 — B |
| `Services/IPptComService.cs` | `GetParagraphs` / `SetParagraphText` / `RestoreParagraphColors` 추가 |
| `Services/PptComService.cs` | 위 3개 구현. 기존 문단 순회·복원 로직 재사용 |
| `Views/TextEditorWindow.xaml(.cs)` | 문단 ID 목록으로 전환, 프롬프트 문구 개정, 파싱을 파서에 위임 |
| `Views/ChangePreviewWindow.xaml(.cs)` | 신규 — C |
| `Views/BckPptWindow.xaml(.cs)` | "열기" 버튼, 드래그&드롭, "표시 해제" 버튼 |
| `ViewModels/BckPptViewModel.cs` | `GetParagraphsForCurrentView`, `ApplyParagraphChangesAsync`, `OpenFileAsync`, 사본 작업 |
| `BckPPT.Tests/PastedPlanParserTests.cs` | 신규 — 실제 에이전트 답변 픽스처 |

기존 shape 단위 `텍스트 치환` 버튼과 `ApplyTextChangesAsync`는 **남겨둔다.** 문단 단위가
안정화될 때까지 병행하고, 되돌리기 경로가 갈라지지 않도록 `_undoSnapshot`만 문단 단위를
함께 담을 수 있게 확장한다.

## 5. 단계

**Phase 1 — 이번 작업 재현 (필수)**
A → B → D → C → E 순. B와 D는 서로 독립이라 병행 가능하다.
완료 기준: `로드맵_취합본_2608_정규산_김병철001_2-1분리.pptx`를 탑재해
이번에 주고받은 답변을 그대로 붙여넣으면 13건 치환 + 1건 문단 추가가 빨간색으로 반영된다.

**Phase 2 — 품질 (선택)**
F(LLM 복구), 넘침 자동 보정(기존 `AutoFitAsync` 연계), 도형 높이 조정 제안.

## 6. 프롬프트 문구 개정안

`TextEditorWindow`의 `PromptBox`에 넣을 안내문. 현재 문구는 "텍스트 추가/삭제 불가"를
못박고 있는데, 문단 추가를 지원하므로 함께 고친다.

```text
# 아래는 PPT의 문단별 텍스트입니다. 요청에 맞게 수정해 주세요.
- 수정이 필요한 항목만 "ID: 새 텍스트" 형식으로 답해 주세요.
- ID는 반드시 원본 그대로 유지해 주세요. (예: S1-12-3)
- 문단을 지우려면 "S1-12-3: (빈 문단)" 으로 써 주세요.
- 문단을 추가하려면 "S1-12-3+: 새 문단" 처럼 ID 뒤에 +를 붙여 주세요.
- 설명을 함께 써 주셔도 됩니다. ID가 붙은 줄만 반영됩니다.
```

마지막 줄이 중요하다. 에이전트에게 설명을 금지하지 않고 허용하는 대신,
파서가 ID 있는 줄만 집어가는 구조임을 명시한다.

## 8. 구현하며 달라진 점

설계대로 되지 않아 바꾼 세 가지를 남긴다. 모두 실제 파일로 검증하다가 드러났다.

### 8.1 `(빈 문단)`은 비우기가 아니라 삭제다

문단 텍스트를 빈 문자열로 넣으면 PowerPoint가 **문단 기호까지 지워** 다음 문단과 합친다.
`s1|96`의 4문단이 3문단이 되는 식이다. "빈 문단으로 만든다"는 표현을 유지하면서
속으로는 문단이 사라지는 상태가 되므로, `ParagraphChangeKind.Clear`를
`ParagraphEditKind.Delete`로 명시적으로 옮겼다. 결과는 같지만 되돌리기가 정확해진다.

python-pptx로 같은 작업을 했을 때는 문단이 남아 빈 줄이 되었다. COM과 동작이 다르다.

### 8.2 되돌리기는 "적용 후" 인덱스를 써야 한다

되돌리기를 원본 인덱스로 만들면 문단이 한 칸씩 밀린다. 앞에서 문단이 지워졌으면 당겨지고,
삽입됐으면 밀리기 때문이다. `BuildUndoEdits`의 `Post()`가 그 보정이며,
순수 함수라 `ParagraphUndoTests`에서 단위 테스트로 고정했다.

COM 적용 순서도 함께 바꿨다. 삭제와 삽입을 따로 돌리면 "p2 삭제 + p3 뒤 삽입"처럼 섞였을 때
뒤 작업이 밀린 인덱스를 보게 되므로, **교체(앞→뒤) → 삭제·삽입 한 번에(뒤→앞)** 로 합쳤다.
큰 인덱스부터 내려오면 아직 처리하지 않은 작업의 위치는 항상 그대로다.

지웠던 첫 문단을 되살리려면 "맨 앞에 넣기"가 필요해 `ParagraphIndex = 0`을
`InsertBefore`로 정의했다.

### 8.3 `S1-27-1`과 `S1-27-1+`는 다른 작업이다

파서가 `+`를 뗀 ID로 중복 제거를 해서, 같은 문단에 대한 "치환"과 "뒤에 추가"가
서로를 덮어썼다. 이번 보완 작업이 정확히 그 형태(`s1|70` 1문단 → 2문단)라 바로 드러났다.
중복 제거 키에 `+`를 포함시켜 해결했다.

## 9. 검증 결과

실제 파일(`로드맵_취합본_2608_정규산_김병철001_2-1분리.pptx`)로 확인했다.

| 항목 | 결과 |
|---|---|
| 문단 추출 | 84개 — python-pptx로 뽑은 결과와 일치 |
| 관용 파싱 | 설명·표·코드펜스가 섞인 답변에서 **치환 13 + 추가 1 = 14건** 추출 |
| 산문 속 ID 언급 | 2줄을 올바르게 거부하고 경고에 표시 |
| 적용 | 실패 0건. 글꼴(`-윤고딕330`)·크기(7pt)·굵기·정렬 원본과 일치 |
| 삽입 문단 | 기준 문단의 서식을 그대로 상속 |
| 빨간색 표시 | 13건 적용 (삭제는 색 대상이 아님) |
| 되돌리기 | 84→84 문단, **텍스트 차이 0건**, 색도 원래 테마색(13)으로 복원 |
| UI | 편집창에 `열기` · `사본`(기본 켜짐) · `표시 해제` 렌더링 확인 |

단위 테스트 30개 통과.

### 알아둘 점

- COM이 보고하는 도형 이름은 XML에 저장된 이름과 다를 수 있다.
  `화살표: 오각형 69`가 COM에서는 `Pentagon 69`로 나온다.
  BckPPT는 읽기·쓰기 모두 COM을 쓰므로 문제되지 않지만, 외부 도구와 대조할 때 주의한다.
- 글자색은 RGB만 저장하면 안 된다. 이 파일의 문단들은 `RGB=0`이지만 실제로는
  `ObjectThemeColor=13`이라, RGB만 되돌리면 테마 연결이 끊긴다.

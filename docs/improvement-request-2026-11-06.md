# bckppt 개선 요청 (2026-11-06)

SAR 연차실적 발표자료(22쪽, 69MB) 취합 작업에서 bckppt로 안 되어 PowerPoint COM을 직접 쓴 것들. 11-03 작업에서 나온 항목도 함께 묶는다.

## 1. `open_presentation`이 읽기 전용으로 연다 (우선순위 높음)
- 증상: `open_presentation`으로 연 파일이 `Presentation.ReadOnly = -1`. 사용자가 Ctrl+S 하면 「다른 이름으로 저장」이 뜨고, 도구로 고친 내용을 원래 파일에 저장할 수 없다. 11-03에도 같은 문제.
- 우회: SaveAs(임시) → Close → 원본 위치로 복사 → `Presentations.Open(path, ReadOnly=0, Untitled=0, WithWindow=-1)`.
- 요청: 기본을 쓰기 모드로 열기. 읽기 전용이 필요하면 `read_only: true` 옵션. 결과 메시지에 읽기 전용 여부 표시.

## 2. 글자 형광(Highlight) 지우기·바꾸기
- 상황: 업체가 넣은 문단에 노란 형광(`Font.Highlight.RGB = 65535`)이 남음.
- 우회: `TextFrame2.TextRange.Runs(k).Font.Highlight.RGB = 16777215`(흰색). **진짜 "형광 없음"이 아니라 흰색**이라 배경이 흰색이 아니면 보인다.
- 요청: `set_text_format`에 `highlight: null|"#RRGGBB"` (null = 형광 없음, XML `a:highlight` 제거). 문단·글자 범위 지정 가능하게.

## 3. 문단·글자 범위 단위 서식
- 상황: 한 상자 안에서 소제목 한 줄(「◎ 2단계 … 확장안」)만 굵게, 일부 문단만 크기 바꾸기.
- 지금 `set_text_format`은 도형 전체만 바꾼다.
- 우회: `Paragraphs(i).Characters(1, n).Font.Bold = -1`.
- 요청: `set_text_format`에 `paragraph`(번호·범위)와 `find`(글자열) 대상 지정.

## 4. 빈 문단 하나 지우기 / 문단 단위 삭제
- 상황: 넘침을 줄이려 빈 문단 하나만 지움. `set_paragraphs`는 전체를 다시 쓰는 방식이라 겹치는 문단 서식이 뒤섞일 위험.
- 우회: `TextRange.Paragraphs(3).Delete()`.
- 요청: `delete_paragraphs(slide, shape, [번호])`.

## 5. 넘침 측정 결과를 바로 쓸 수 있게
- `read_text`가 "텍스트 높이 vs 상자 높이"를 주는 것은 좋음. 서식 변경 뒤에도 같은 값이 돌아오면 반복(14pt→13pt→빈 줄 삭제)을 도구 안에서 끝낼 수 있다.
- 요청: `set_text_format`·`set_paragraphs` 결과에 `textHeight / boxHeight / overflow(pt)` 포함. 선택 옵션 `fit: "shrink_font"`(지정 하한까지 줄여 맞춤).

## 6. 두 판 비교 (diff_presentations)
- 상황: 업체가 003 판을 고쳐 보냄 → 어느 쪽이 바뀌었는지 python-pptx로 쪽별 글·그림 수를 비교해서 찾음.
- 요청: `diff_presentations(a, b)` — 쪽마다 같음/바뀜, 바뀐 도형의 글 전/후, 그림 수 변화, 쪽 수 차이.

## 7. 여러 쪽 한 번에 PNG로 (render_slides)
- `render_slide`는 한 쪽씩. 검토 때 6쪽을 COM `Slide.Export`로 한꺼번에 뽑음.
- 요청: `render_slides(slides:[…], out_dir?)` — 파일 경로 목록과 미리보기.

## 8. 11-03에 이미 나온 항목 (다시 정리)
- 다른 pptx에서 쪽 가져오기 `Slides.InsertFromFile` (원본을 임시 폴더에 복사해야 성공했음)
- 쪽 복제·이동·삭제 (`Duplicate`·`MoveTo`·`Delete`)
- 표 칸 읽기·쓰기
- 그림 넣기 `AddPicture`(영역에 맞춤, 비율 유지)
- 다른 쪽 도형 복사·붙여넣기(묶음 도형은 PNG로 붙이기 옵션)
- 저장 `save` / `save_as` (1번과 함께)

## 참고: 이번 COM 우회 방식
- PowerShell 7에는 `Marshal.GetActiveObject`가 없음 → `New-Object -ComObject PowerPoint.Application`로 실행 중 인스턴스에 붙음(사용자 창이 있으므로 Quit 하지 않음).
- 대상 문서는 `Presentations | Where Name -like "*이름*"`으로 고름.

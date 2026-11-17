# bckppt 개선 요청 (2026-11-17 c) — 카드 PPT 3차 수정에서 걸린 점

2.2.17로 카드 PPT 3차 수정 9개 쪽을 넣으며(replace_picture·add_picture crop·inset·2×3 격자·image_preview·render_slides) 걸린 점.
(kiris-e0, 사용자 지시. 우선순위 1~4가 높음)

## 요청

1. **창 없는 문서**: PowerPoint 창이 닫혀 문서가 메모리에만 남으면(오늘 두 번) 다른 도구가 「열린 문서: 없음」만 말함 →
   창 없는 문서도 찾아 「창 없이 열려 있음 — open_presentation으로 창 띄우기」라고 알리거나, 그 자리에서 NewWindow로 창을 띄워 이어서 처리(옵션).
2. **`replace_pictures_in_box(slide, box, paths[], crops/keeps/focuses, rows/columns, gap, skip_user_modified?)`** — 칸 안(중심 기준) 그림을
   모두 지우고 새 목록으로 넣기를 한 번에. 지금은 delete_shape를 그림마다 부르고 add_pictures_in_box를 다시 불렀음.
3. **사용자가 고친 그림 알아보기**: bckppt가 넣은 그림에 넣을 때의 자리·크기·crop을 태그로 남기고, list_shapes가 지금 값과 다르면
   「사용자가 바꿈」 표시. 2번·replace_picture가 그런 그림을 덮어쓰기 전에 경고(force로 진행).
4. **`apply_picture_plan(file, plan_path | plan[])`** — 줄 형식 `쪽|칸 이름 또는 box|경로[@crop 또는 keep];경로…`
   (칸 이름은 예: 카드 PPT의 AS-IS/TO-BE 칸 정의를 인자로 받거나 「[그림 삽입]」 자리), 칸별로 넣고 결과 표 + 바뀐 쪽 축소판.
   참고 구현: CellBox\05 KB …\261117-시공안전 기술설명 카드 그림\_scripts\apply_open_ppt.py(COM).
5. **자르기를 남길 영역으로도**: `keep:{x0,y0,x1,y1}`(원본 기준 0~1) — crop(잘라 낼 양)과 둘 중 하나.
6. **`image_sheet(paths[], columns?)`** — 후보 사진 여러 장을 번호·파일 이름·픽셀 크기와 함께 한 장 축소판으로.
7. **해상도 경고**: 넣은 결과에 실제 인쇄 해상도(원본 px ÷ 넣은 크기 inch, crop 반영)를 적고 150dpi 미만이면 경고.
   487px·300px 원본이 흐렸던 것을 화면 보고서야 알았음.
8. **기준 도형 기준 배치**: add_picture에 `anchor:{shape, corner:"bottom-right"|…, size:0.45(기준 도형 너비 대비), margin:3}`.
9. **문서 크기 관리**: `document_size`(그림별 크기·자른 부분 포함 여부), `compress_pictures(dpi, delete_cropped=true, slides?)` — 27MB→62MB가 됨.
   사용자 문서에 적용은 사용자 확인 뒤에만.
10. **출처 기록**: add_picture 등에 `source`(URL·문서명), `license` 인자 → 태그 저장, `export_picture_credits(slides?)`로 쪽별 출처 목록(마크다운)
    또는 발표자 메모에 쓰기 옵션.
11. **`render_region(slide, box | shape, scale)`** — 칸 부분만 크게 렌더.

## 구현 메모

- 1: `ListWindowlessPresentations`(창 없는 문서 중 BckPPT가 작업용으로 연 것은 뺀다). 문서를 못 찾으면 오류 글에 그 문서들을 적고,
  모든 문서 도구의 `show_window=true`면 그 자리에서 창을 띄워(open_presentation과 같은 NewWindow) 이어서 처리한다. list_presentations에도 「창 없음」으로 보인다.
- 3: 태그 `BCKPPT_PLACED`(「l,t,w,h;L,T,R,B」)에 BckPPT가 마지막으로 정한 자리·자르기를 둔다(넣기·set_picture_crop·압축 때 갱신).
  자리·크기 0.75pt, 자르기 0.5%까지는 같다고 본다. 태그가 없는 그림(사용자가 넣었거나 옛 판)은 판정하지 않는다.
- 2·4: 칸 안 그림(중심 기준)과 「[그림…]」 글상자를 지우고 첫 그림의 Z순서 자리에 넣는다. 칸 도형 자신은 지우지 않는다.
  계획은 줄마다 따로 성공·실패하고(한 줄 오류가 나머지를 막지 않는다), 사용자가 바꾼 그림이 있는 칸은 기본으로 건너뛴다.
  계획 파일의 상대 경로는 계획 파일 폴더 기준, 거기 없으면 그 위 폴더(참고 구현처럼 _scripts 아래에 계획이 있을 때).
  그림마다 `@crop=l,t,r,b`, `@keep=x0,y0,x1,y1`, `@focus=x,y[,zoom]`.
- 7: dpi = 원본 px ÷ (도형 너비 ÷ 남긴 비율 ÷ 72). 넣기 결과·list_shapes·계획 결과에 보인다.
- 9: document_size는 디스크에 저장된 판의 zip을 읽는다(그림 파일은 zip 안 압축 크기, srcRect가 0.5% 넘으면 「자른 부분 포함」).
  compress_pictures는 confirm 없이 부르면 고치지 않고 줄일 그림·예상 크기만 보인다. 적용은 원본을 picture_preview와 같은 순서로 찾아
  (잘린 부분을 버리고) 쪽에 보이는 크기 기준 dpi로 줄인 JPEG(투명하면 PNG)로 같은 자리·Z순서·이름·테두리·그림자·출처 태그를 지켜 바꿔 넣는다.
  그 뒤에는 자르기를 넓힐 수 없으니 사용자 확인 뒤에만.
- 10: 태그 `BCKPPT_CREDIT`·`BCKPPT_LICENSE`. 발표자 메모에는 「그림 출처: …」 문단 하나를 쓰고, 다시 부르면 그 문단만 바꾼다.
- 11: 쪽을 pt당 scale 픽셀로 렌더해 영역만 잘라 준다(비우면 영역이 약 1200px 너비).

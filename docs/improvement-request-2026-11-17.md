# bckppt 개선 요청 (2026-11-17)

시공안전 기술설명 카드 PPT(36쪽, 26MB) 작업에서 나온 것. 11~36쪽의 AS-IS·TO-BE 그림 칸 52개(281.52×93.6pt, 3:1)에 사진을 넣고, 사용자 검토 뒤 칸마다 1~2장으로 바꿔 넣었다. 그림 넣기는 bckppt에 없어서 python-pptx(처음 넣기)와 PowerPoint COM(열린 창 다듬기)으로 했다.

## 1. 그림 넣기 `add_picture` (우선순위 높음)
- 상황: 칸 52개에 그림 넣기. 사용자가 창에서 몇 칸을 직접 고친 뒤라, 파일을 다시 만들지 않고 **열린 창에** 넣어야 했다.
- 우회: COM `Shapes.AddPicture(path, LinkToFile=False, SaveWithDocument=True, left, top, w, h)`. 비율은 미리 PIL로 칸 비율에 맞춰 잘라 둠.
- 요청: `add_picture(slide, path, left, top, width, height, fit)`
  - `fit: "cover"`(칸을 채우게 가운데 자름, `PictureFormat.Crop*`) / `"contain"`(비율 유지, 칸 안 가운데) / `"stretch"`
  - `crop: {left, top, right, bottom}`(원본 기준 비율 0~1) — 핵심 부분만 크게 보이게 지정
  - `name` 지정, 결과에 도형 주소·실제 크기·자른 양을 돌려줌
- 파일 경로는 CellBox·Downloads 아래로 한정해도 됨.

## 2. 칸(영역)에 여러 장 나란히 `add_pictures_in_box`
- 상황: 4:3 사진 한 장은 3:1 칸의 절반만 차서, 칸마다 2장을 3pt 간격으로 나란히 넣음.
- 요청: `add_pictures_in_box(slide, box:{left,top,width,height} 또는 기존 도형 주소, paths:[…], gap, fit)` — 칸을 n등분해 각각 cover로 넣기. 기존 도형 주소를 주면 그 자리·크기를 쓰고 그 도형(예: 「[그림 삽입]」 글상자)을 지우는 옵션 `replace_placeholder: true`.

## 3. 그림 바꾸기 `replace_picture`
- 상황: 사용자 검토 의견(「사진이 너무 작음」, 「다른 사진으로 교체」)으로 칸의 그림만 바꿈. 위치·크기·이름은 그대로 두고 파일만 바꾸고 싶음.
- 우회: 칸 범위 안의 그림을 찾아 `Delete` → `AddPicture`.
- 요청: `replace_picture(slide, shape, path, fit)` — 같은 자리·크기·Z순서 유지.

## 4. 영역 안 도형 찾기
- 상황: 「AS-IS 칸에 있는 그림」을 찾으려고 도형 목록에서 위치(Top 200~330, 중심 x가 칸 범위 안)로 걸렀다. 사용자가 그림을 칸보다 크게 끌어 놓거나 여러 조각으로 잘라 둔 경우도 있었다.
- 요청: `list_shapes`에 `in_box:{left,top,width,height}`·`type: "picture"` 필터, 그림이면 `crop`·원본 픽셀 크기도 표시.

## 5. 다운로드한 파일이 「제한된 보기」로 열리면 알려 주기
- 증상: 메일 첨부로 받은 pptx(Zone.Identifier 있음)를 `open_presentation`으로 열었더니 창은 `[Protected View]`로 뜨고, 이어지는 `list_slides`·`render_slide`가 「PowerPoint에서 보고 있는 문서가 없습니다」라고만 나옴.
- 우회: COM `ProtectedViewWindows(i).Close()` → `Unblock-File` → 다시 열기.
- 요청: `open_presentation` 결과에 제한된 보기 여부 표시, 옵션 `unblock: true`(사용자 파일 표시 제거 후 편집 모드로 열기) 또는 `ProtectedViewWindow.Edit()`.

## 6. 연 직후 「열린 문서 없음」
- 증상: 큰 파일(26MB)을 `open_presentation`으로 연 직후 `render_slide(file:"그림삽입")`이 「열린 문서: 없음」. 같은 시각 COM `Presentations`에는 문서가 있었고, 잠시 뒤 다시 부르니 됨.
- 요청: `open_presentation`이 문서가 `Presentations`에 올라올 때까지 기다린 뒤 돌려주거나, 다른 도구가 문서를 못 찾으면 잠깐 다시 찾아보기.

## 7. 여러 쪽 렌더 (11-06 요청 7번 다시)
- 26쪽 결과를 확인하느라 `render_slide`를 한 쪽씩 여러 번 불렀다. `render_slides(slides:[…])` 또는 여러 쪽을 한 장에 모은 축소판(contact sheet)이 있으면 검토가 빠르다.

## 참고: 이번 우회 스크립트
- 처음 넣기: `CellBox\05 KB …\261117-시공안전 기술설명 카드 그림\_scripts\insert_photos.py` (python-pptx, 「[그림 삽입]」 글상자 자리에 넣고 글상자 삭제)
- 열린 창 다듬기: 같은 폴더 `_scripts\apply_open_ppt.py` (COM, 칸의 기존 그림 지우고 1~2장 넣기, 사용자 수정 칸은 건너뜀, 저장은 사용자)

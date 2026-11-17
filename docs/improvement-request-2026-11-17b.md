# bckppt 개선 요청 (2026-11-17 b) — 그림 자르기를 PowerPoint 안에서

카드 PPT 그림 칸(281.52×93.6pt)에 사진을 넣을 때, 지금은 밖에서 PIL로 잘라 새 jpg를 만든다. 사용자는 단순 자르기는
PowerPoint 자르기(원본 보존, 나중에 다시 조정 가능)로 하길 원한다. 에이전트가 그림을 보며 편하게 자르고 다시 고칠 수 있게 해 달라.
(kiris-e0, 사용자 지시. 2.2.16 설치 뒤)

## 요청 (우선순위 순)

1. `set_picture_crop(slide, shape, crop?{left,top,right,bottom 원본 기준 0~1} | focus?{x,y 원본 기준 0~1} + zoom?(1=칸을 겨우 채움, 2=두 배 확대), keep_frame=true)`
   — 이미 넣은 그림의 자른 영역만 바꾸고 도형 자리·크기는 그대로. focus+zoom이면 칸 비율에 맞춰 그 점이 가운데 오게 자동 계산
   (가장자리 넘으면 안쪽으로 밀기). `reset=true`면 자르기 없음.
2. `picture_preview(slide, shape)` — 그 그림의 원본 전체를 축소 PNG로 돌려주고, 현재 자른 영역을 사각형으로, 0.1 간격 격자·눈금을
   그려서 crop/focus 값을 정할 수 있게. 결과에 원본 픽셀·현재 crop 비율.
3. 파일 미리보기도 같은 방식: `image_preview(path)` — 넣기 전 원본 파일에 격자·눈금(0~1)을 그려 돌려줌.
4. `add_picture`/`add_pictures_in_box`/`replace_picture`에 focus+zoom 옵션(1번과 같은 계산). crop과 같이 쓰면 crop 먼저.
5. 겹쳐 올리기 편의: `add_picture`에 `z?("front"|"back")`, `border?{color,width}`, `shadow?(bool)` — 큰 그림 위에 작은 inset 그림을 올릴 때.
6. 격자 배치: `add_pictures_in_box`에 `rows`/`columns`(예: 2×3) — 칸 오른쪽 절반에 센서 사진 6장을 2×3으로.
7. `list_shapes(type:"picture")`에서 그림마다 원본 파일명(넣을 때 이름을 기억할 수 있으면 대체텍스트/태그에 저장)·crop 표시.

덧붙여: 설치 스크립트(`package-plugin.ps1 -InstallLocal`)의 Plugins 폴더 이름 바꾸기가 CellBox2 트레이와 겹쳐 실패했으니 파일 단위 교체로.

## 구현 메모

- 계산은 `PictureFitter`(COM과 떼어 시험한다).
  - `FocusCrop`: crop 뒤 남은 부분에서 칸 비율의 가장 큰 창(zoom 1)을 잡고 zoom만큼 줄여 focus 가운데에 둔다. 가장자리를 넘으면 안쪽으로 민다.
  - `Recrop`: `keep_frame=true`면 도형 자리 그대로(crop은 cover로 맞춤, reset은 칸 안에 contain).
    `false`면 지금 배율·그림 자리를 지키고 도형이 새 자른 영역을 따라간다(PowerPoint 자르기 도구와 같다).
  - `Grid`: rows×columns, 왼쪽 위부터 줄 단위.
- 넣는 그림에는 원본 경로를 태그 `BCKPPT_SOURCE`에 적는다(대체 텍스트는 접근성용이라 건드리지 않는다).
  태그가 없으면 대체 텍스트가 그림 파일 이름처럼 생겼을 때 그것을 보여 준다.
- `picture_preview`의 원본은 ① 태그의 원본 파일(비율이 같을 때) → ② 저장된 pptx 안의 그림(SlideID·Shape.Id로 찾음) →
  ③ 복제본의 자르기를 풀어 `Shape.Export`로 뽑기(문서에 잠깐 도형이 생겼다 지워진다) 순서로 찾는다.
- 결과 글에 지금 자른 영역을 같은 뜻의 focus·zoom 값으로도 알려 준다(그대로 set_picture_crop에 넣어 조금씩 옮길 수 있게).
- 설치: 새 판을 임시 폴더에 풀고, 덮어쓸 파일만 임시 백업 → robocopy로 덮어쓰기(실패하면 백업에서 되돌림) →
  새 판에 없는 옛 파일은 지우되 잠겨 있으면 남기고 알린다.

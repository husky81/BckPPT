# BckPPT

PowerPoint 텍스트 진단 및 치환 도구

## 기능

- **텍스트 진단** — 열려있는 PowerPoint 파일의 모든 텍스트 상자/표 셀을 스캔하여 overflow(텍스트 넘침) 감지
- **텍스트 치환** — 슬라이드 텍스트를 번호 목록으로 추출 → LLM으로 수정 → 한번에 적용
- **글꼴 찾기** — PPT에서 사용된 모든 글꼴을 드롭다운으로 필터링, 누락 글꼴 탐색
- **레이아웃 미리보기** — 슬라이드의 텍스트/이미지 배치를 개략적으로 시각화

## 설치

[Releases](https://github.com/husky81/BckPPT/releases)에서 최신 설치 파일을 다운로드하세요.

## 요구 사항

- Windows 10/11 (x64)
- Microsoft PowerPoint (Office 2016 이상)

## 사용법

1. PowerPoint에서 PPT 파일을 열기
2. BckPPT 실행
3. 드롭다운에서 파일 선택 → 자동 분석
4. 슬라이드/shape 클릭 → PowerPoint에서 해당 위치로 이동

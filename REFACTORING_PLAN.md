\# 🛠️ Woondooran ERP 리팩토링 계획 (2026-01-16)



\## 1. 배경 (Background)

\* 현재 `index.html` (900 lines)과 `Code.gs` (300 lines) 두 파일에 모든 코드가 집중됨.

\* 추후 기능(반제품, 포장 등) 추가 시 코드량이 5,000줄 이상으로 늘어나 유지보수가 어려울 것으로 예상됨.



\## 2. 목표 (Goal)

\* \*\*모듈화 아키텍처 도입:\*\* Google Apps Script의 `include()` 패턴을 사용하여 파일을 기능별로 분리한다.

\* \*\*유지보수성 향상:\*\* 화면(View)과 로직(Logic)을 분리하여 관리 효율을 높인다.



\## 3. 변경할 파일 구조 (File Structure)



\### Server-Side (.gs)

\* `Code.gs`: `doGet`, `include` 함수만 유지 (진입점)

\* `Service.gs`: 비즈니스 로직 (데이터 CRUD 등)

\* `Config.gs`: 시트 이름 등 상수 관리



\### Client-Side (.html)

\* `index.html`: 메인 레이아웃 및 껍데기

\* `view-dashboard.html`: 대시보드 화면

\* `view-materials.html`: 원재료 관리 화면

\* `view-production.html`: 생산일지 화면

\* `modal-\*.html`: 각종 팝업창 분리

\* `stylesheet.html`: CSS 분리

\* `javascript.html`: JS 로직 분리



\## 4. 기대 효과

\* 파일당 코드 라인 수를 줄여 가독성 확보

\* 기능 추가 시 해당 파일만 수정하면 되므로 사이드 이펙트 최소화


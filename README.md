# 모의고사 운영 도구

정답·배점 기반 Google Forms 생성부터 응답 파일 분석, 개인 성적표와 주차별 비교 결과 생성까지 지원하는 웹 도구입니다.

## 주요 기능

- 응답 엑셀·CSV 검증 및 점수 재계산
- 학생별 성적표 엑셀·PDF 생성
- 수강반별 PDF ZIP 생성
- 정답·배점 기반 Google Forms 퀴즈와 시험 설정 Excel 생성
- 시험 설정 Excel을 불러와 정답·배점 자동 입력
- 학교·수강반·시간대와 성취도 문구 공통 설정
- 두 주차의 동일인 확인 및 성적 변화 비교
- 등수 상승 TOP10, 반별 TOP5, 주차별 점수 TOP10 제공
- 시스템 설정 연동 및 수동 전환이 가능한 다크 모드

업로드한 파일과 분석 데이터는 서버로 전송하지 않고 브라우저 안에서 처리합니다.

## 로컬 실행

```bash
npm ci
npm run dev
```

테스트와 프로덕션 빌드는 다음 명령으로 확인합니다.

```bash
npm test
npm run build
```

## Google Forms 설정

개인 Google 계정의 Google Cloud 프로젝트에서 Forms API와 Drive API를 활성화하고 웹 애플리케이션 OAuth 클라이언트를 만듭니다. 승인된 JavaScript 원본에는 로컬 주소와 실제 배포 주소를 등록합니다.

```bash
cp .env.example .env.local
```

`.env.local`의 `VITE_GOOGLE_CLIENT_ID`에 OAuth 클라이언트 ID를 입력합니다. OAuth 동의 화면이 테스트 상태라면 사용할 개인 Google 계정을 테스트 사용자로 추가해야 합니다. 앱은 `drive.file` 범위만 요청하며 액세스 토큰을 저장하지 않습니다.

GitHub Pages에서는 저장소의 **Settings → Secrets and variables → Actions → Variables**에 `VITE_GOOGLE_CLIENT_ID`를 등록합니다.

## GitHub Pages 배포

`main` 브랜치에 변경 사항을 푸시하면 GitHub Actions가 테스트와 빌드를 수행한 뒤 `dist` 산출물을 GitHub Pages에 배포합니다. 저장소의 **Settings → Pages → Build and deployment → Source**를 **GitHub Actions**로 설정해야 합니다.

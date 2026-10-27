# 모의고사 성적 분석

Google Forms 응답 엑셀 또는 Google Sheets에서 다운로드한 UTF-8 CSV를 브라우저에서 분석해 개인 성적표와 주차별 비교 결과를 만드는 웹 도구입니다.

## 주요 기능

- 응답 엑셀·CSV 검증 및 점수 재계산
- 학생별 성적표 엑셀·PDF 생성
- 수강반별 PDF ZIP 생성
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

## GitHub Pages 배포

`main` 브랜치에 변경 사항을 푸시하면 GitHub Actions가 테스트와 빌드를 수행한 뒤 `dist` 산출물을 GitHub Pages에 배포합니다. 저장소의 **Settings → Pages → Build and deployment → Source**를 **GitHub Actions**로 설정해야 합니다.

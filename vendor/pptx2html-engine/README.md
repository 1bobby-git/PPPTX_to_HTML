# 변환 엔진 패치와 재현 빌드

이 디렉터리는 MIT 라이선스의 [`@briank-dev/pptx-to-html` 2.1.0](https://github.com/brnyxx/pptx2html-turbo)을 기반으로 한 로컬 패치 버전 **2.1.0-ppptx.1**을 보관합니다. 브라우저와 Native CLI는 같은 패키지를 사용합니다.

## 변경 범위

`CHART-DEPTH.patch`는 차트 파서가 DrawingML 스타일 요소의 닫는 태그를 차트 계층 깊이에 잘못 반영하는 문제를 수정합니다. 정상적인 두 축 차트가 `incompatible-axes`로 잘못 분류되는 경우를 해결하며, 실제로 정의되지 않은 축은 계속 지원 불가로 처리합니다.

원본 차트의 모든 스타일을 재현하는 패치는 아닙니다. 확인한 예제에서 사용자 지정 계열 색상은 테마 색상으로 표시되며, 계열 선과 전체 차트 제목의 기존 지원 제한도 남아 있습니다.

`BUILD-INFO.json`에 원본 커밋, 다운로드 주소, 입력·출력 SHA256, 도구 버전, 빌드 명령과 검증 범위를 기록했습니다. `SHA256SUMS`는 배포 패키지와 관련 고지 파일의 해시입니다. 원본 MIT 고지는 `LICENSE`에 있습니다.

## 개발자용 재현 빌드

일반 사용자는 Rust, Cargo, wasm-bindgen 또는 명령어 실행이 필요하지 않습니다. 아래 절차는 **배포 엔진을 유지보수하는 개발자**를 위한 것으로, 저장소 루트에서 실행합니다.

재현 환경은 **Linux x64, Node.js v24.19.0, npm 11.9.0**입니다. `tar`, `patch`, `bash`와 고정된 공개 소스·도구 다운로드를 위한 네트워크 연결도 필요합니다.

```bash
node scripts/rebuild-engine.cjs
```

스크립트는 새 임시 디렉터리에 소스, 고정 버전 Rust 도구, wasm-bindgen 및 Cargo 캐시를 준비합니다. 전역 도구 설치나 셸 프로필 변경은 하지 않습니다. 다운로드 해시, 잠금 파일, 패치 전후 소스, 차트 회귀 검사, 생성된 WASM·바인딩과 최종 npm 패키지의 해시를 확인합니다. 다른 npm 버전으로 패키지 바이트가 달라지면 최종 해시 검사에서 실패합니다.

선택적으로 경로를 지정할 수 있습니다.

```bash
node scripts/rebuild-engine.cjs \
  --work-dir /tmp/pptx-engine-work \
  --cache-dir /tmp/pptx-engine-downloads \
  --output-dir /tmp/pptx-engine-artifacts
```

`--work-dir`는 새 디렉터리 또는 빈 디렉터리여야 합니다. 다운로드 캐시를 재사용해도 해시는 다시 확인합니다. 결과와 작업 경로는 완료 시 출력되며, 저장소의 배포 패키지를 자동으로 교체하지 않습니다. 이 버전은 별도의 작업 디렉터리에서 재빌드하여 최종 패키지 SHA256까지 일치함을 확인했습니다.

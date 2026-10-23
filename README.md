# PPTX → HTML

`.pptx` 프레젠테이션을 독립적인 HTML 파일로 변환합니다. 최신 버전은 **v1.6.0**입니다. **Windows 실행 파일에서 파일 선택 → 변환 → 미리보기 → HTML 저장까지 진행할 수 있습니다.** 일반 사용자는 Node.js 설치나 CMD 명령어 입력이 필요 없습니다. PowerPoint와 LibreOffice를 실행하지 않으며, 문서는 이 PC 안에서 변환합니다.

## 사용자: Windows 실행 파일로 변환하기

1. [v1.6.0 다운로드](https://github.com/1bobby-git/PPPTX_to_HTML/releases/tag/v1.6.0)에서 **`PPPTX_to_HTML_Desktop_v1.6.0_win32-x64.zip`**을 받습니다.
2. ZIP의 **압축을 모두 풀고** `PPPTX_to_HTML.exe`를 더블클릭합니다.
3. **파일 선택**으로 PPTX를 고릅니다. 필요한 경우 슬라이드 범위·숨김 슬라이드 포함·제목·문서 열기 암호를 설정합니다.
4. **HTML로 변환**을 누르고 결과를 **미리보기**로 확인합니다.
5. **HTML 저장**을 눌러 저장 위치를 선택합니다. 결과 HTML은 일반 브라우저에서 열 수 있습니다.

EXE 옆의 DLL·`resources`·`locales` 폴더도 필요합니다. **EXE 한 개만 분리해 복사하지 마세요.** 별도 설치 프로그램이나 자동 업데이트는 없습니다. 실행 파일이 원본 경로를 직접 읽고 내장 작업 스레드에서 기존 공통 변환 엔진을 실행합니다. 별도 `node.exe`나 PowerPoint 프로세스를 실행하지 않습니다. [자세한 실행 파일 안내](docs/DESKTOP.md).

회사 DRM이 이 프로그램에 평문 PPTX 읽기와 HTML 저장을 허용해야 보안 문서를 처리할 수 있습니다. EXE로 포장한 것만으로 NASCA 등의 암호문을 해독하지는 않습니다. **실제 NASCA 문서는 아직 검증되지 않았습니다.** 파일 형식을 읽지 못했다는 오류만으로 확장자 변경이나 DRM이라고 단정하지 않습니다.

## 웹 서비스와 오프라인 HTML

기존 웹 화면과 단일 `index.html`도 유지합니다. 웹 서비스에서는 파일 선택·변환·미리보기·저장을 브라우저로 진행하고 서버가 Native CLI를 실행합니다. 일반 사용자는 운영자가 제공한 웹 주소만 열면 됩니다.

CLI는 웹 서비스를 운영하는 컴퓨터에서 실행됩니다. 기본 설정은 `127.0.0.1`로, 같은 PC에서만 접속합니다. 여러 사람이 설치 없이 웹 주소 하나로 사용하려면 운영자가 사내 서버 등에 서비스를 한 번 배치해야 합니다. 이 경우 파일과 입력한 암호는 **그 변환 서버로 전달**됩니다. 현재 저장소 자체가 실행 중인 변환 서버를 제공하는 것은 아닙니다.

## 사용자: 웹 화면에서 변환하기

1. 운영자가 안내한 변환 웹 주소를 엽니다.
2. `.pptx` 파일을 선택하거나 끌어놓습니다.
3. 필요한 경우 제목, 슬라이드 범위, 숨긴 슬라이드 포함, 문서 열기 암호를 설정합니다.
4. **HTML로 변환**을 누릅니다. 기다리는 동안 **취소**할 수 있습니다.
5. 미리보기를 확인하고 **HTML 저장**을 누릅니다. 결과 HTML은 이후 서버 없이 열 수 있습니다.

웹 서비스는 원본 PPTX의 ZIP/XML에서 텍스트·표·이미지·SVG를 변환합니다. PDF나 슬라이드 전체 PNG를 중간 결과로 만들지 않습니다. 입력 파일과 결과를 작업별 임시 폴더에서 처리하고, 응답이 끝나거나 실패·취소되면 정리합니다. 동시 변환은 한 건이며, 다른 작업이 진행 중이면 다시 시도할 수 있는 안내를 표시합니다.

**운영자 설정:** [웹 서비스 실행 및 배포 안내](docs/WEB_SERVICE.md). 일반 사용자에게 이 설정 절차를 수행하도록 요구하지 않습니다. 서비스 실행 명령은 `npm start`이며 의존성을 먼저 준비해야 합니다. [GitHub Pages는 정적 호스팅](https://docs.github.com/en/pages/getting-started-with-github-pages/what-is-github-pages)이므로 Node.js CLI를 실행할 수 없습니다. 웹 서비스는 UI와 API를 같은 주소에서 제공합니다.

### 실행 방식 선택

| 방식 | 일반 사용자 준비 | 변환 위치와 파일 전달 |
| --- | --- | --- |
| **Windows 실행 파일** | Desktop ZIP을 모두 풀고 EXE 더블클릭 | EXE가 원본 경로를 직접 읽어 이 PC에서 변환, 서버 업로드 없음 |
| **웹 화면 + CLI 서비스** | 운영자가 제공한 웹 주소 접속 | 서비스가 실행되는 PC/서버로 파일 전달, 그곳에서 CLI 실행 |
| 기존 단일 `index.html` | HTML 한 파일 저장 후 열기 | 브라우저 안에서 변환, 서버 업로드 없음 |
| 직접 Native CLI | Node.js와 CLI 환경 준비 | 실행한 컴퓨터 안에서 변환 |

기존 **약 4.5 MiB의 단일 `index.html`**도 유지합니다. 이 오프라인 경로에는 Node.js, npm, 별도 서버, CDN, 인터넷 연결, PowerPoint 설치가 필요하지 않습니다. 회사 DRM을 해독하는 기능이 새로 추가된 것은 아니며, 자세한 범위는 아래 보안 문서 안내를 참고하세요.

## 참고 프로젝트를 검토해 반영한 개선

[`g21589/PPTX2HTML`](https://github.com/g21589/PPTX2HTML), [`javier-mora/pptx-to-html`](https://github.com/javier-mora/pptx-to-html), [Aspose 예제](https://gist.github.com/aspose-com-gists/223609d1fb2c102aac36f1c4707834bb), [`pptx-export` topic](https://github.com/topics/pptx-export)의 실제 소스·라이선스·변환 방향을 비교했습니다. 검토한 버전과 채택·제외 근거는 [참고자료 분석](docs/REFERENCE_REVIEW.md)에 기록했습니다.

- **하나의 변환 코어:** Windows 앱·브라우저·Native CLI가 범위 선택, 숨긴 슬라이드 포함, 엔진 결과 검증과 진단 처리를 공유합니다.
- **슬라이드 누락 방지:** 일부 차트 대체 표시 때문에 HTML 컨테이너 밖으로 밀려난 슬라이드를 복구하고, 엔진이 보고한 장수와 최종 HTML 장수를 대조합니다.
- **미지원 차트의 데이터 보존:** 정확한 슬라이드와 위치를 확인할 수 있는 차트에 한해 문서 내부 캐시의 원시 값을 표로 제공합니다. 정상적으로 그려지는 SVG 차트는 그대로 유지합니다. 캐시 표는 원본 차트의 모양이나 최신 외부 데이터의 재현을 의미하지 않습니다.
- **정상 차트의 오인식 수정:** 기존 엔진의 XML 깊이 계산을 수정해, DrawingML 스타일이 들어 있는 지원 차트를 잘못된 축으로 거부하던 문제를 해결했습니다. 정상 2축 막대 차트에서 값·범주·계열을 직접 SVG로 표시합니다. 축이 실제로 잘못된 차트와 미지원 구성은 기존 대체 표시를 유지합니다.
- **여러 테마의 색·글꼴 보정:** 관계와 상속이 명확한 경우 각 슬라이드의 테마를 적용하고 원래 글자 크기를 유지합니다. 선택적인 목록 스타일 요소가 생략된 정상 문서도 처리합니다. 모호한 override나 상속은 추정해서 바꾸지 않고 경고합니다.
- **슬라이드 내부 링크 보존:** 텍스트·도형·이미지에 설정된 페이지 이동과 다음·이전·처음·마지막 이동을 저장한 HTML에서도 사용할 수 있습니다. 숨김·범위 선택 이후의 결과 페이지에 맞춰 연결하고, 제외된 고정 목적지는 비활성화합니다. 키보드 Enter·Space 조작도 지원합니다.
- **환경별 자동 검증:** GitHub Actions에서 실제 Windows 배포 EXE, Windows·macOS·Linux의 Node CLI와 웹 흐름, 추가 Firefox·WebKit의 오프라인 흐름을 검사합니다. 배포 EXE는 저장소 밖으로 복사한 뒤 실행하여 개발 환경의 의존성에 기대지 않는지 확인합니다. 모든 검사가 통과한 새 버전은 GitHub Release에 배포합니다.

두 번째 JS 렌더러나 Aspose SDK를 중복 설치하지 않습니다. Aspose 평가판의 워터마크·텍스트 읽기 제한, HTML5의 외부 리소스 의존, topic 프로젝트의 반대 변환 방향 등을 검토한 결정입니다. 모든 PowerPoint 기능과 모든 환경의 시각적 일치를 보장하는 기능을 제공한다고 표시하지 않습니다.

## 운영자·자동화: CLI로 직접 변환하기 — Windows

[Node.js 24 LTS 최신 버전](https://nodejs.org/en/download)을 설치하고, 이 저장소를 **Code → Download ZIP**으로 내려받아 압축을 풉니다. 저장소 폴더에서 아래 명령을 실행하세요. Node.js 24 계열에서는 24.15.0 이상이 필요합니다.

```powershell
npm.cmd ci --omit=dev --ignore-scripts --no-audit --no-fund
.\tools\Convert-PptxNative.cmd --check
.\tools\Convert-PptxNative.cmd "C:\test.pptx" -o "C:\test.html"
```

의존성이 포함된 `PPPTX_to_HTML_Native_v1.6.0.zip`을 받았다면 `npm.cmd ci`는 생략합니다. 압축을 푼 폴더에서 바로 실행하세요. CMD를 사용하지 않고 다음처럼 실행해도 같습니다.

```powershell
node .\tools\convert-pptx.cjs "C:\test.pptx" -o "C:\test.html"
```

출력 파일이 이미 있으면 덮어쓰지 않습니다. 교체하려면 `--overwrite`를 추가하세요. PowerShell 실행 정책 변경은 필요하지 않습니다. **의존성 준비가 끝난 뒤에는 변환 서버나 인터넷 연결 없이 실행**합니다.

- 출력: 선택 가능한 텍스트, 표, 내장 이미지, SVG 도형과 탐색 UI가 들어 있는 HTML 한 파일.
- 지원: 숨김 슬라이드 포함, 원본 슬라이드 범위, 제목, JSON 결과, 시간 제한, 실행 환경 점검, 지원되는 Office 문서 열기 암호의 표준 입력.
- 보안 문서: **NASCA+SD 복호화 기능이 추가된 것은 아닙니다.** CLI가 읽은 바이트가 암호문이면 HTML 구조를 만들 수 없습니다. 알 수 없는 형식을 확장자 변경이나 DRM으로 단정하지 않고 오류 코드를 표시합니다.

설치 조건·옵션·오류 코드·한계는 [새 Native CLI 사용 설명](docs/NATIVE_CLI.md)에 정리했습니다. 기존 Python·LibreOffice 방식이 필요하면 [v1.3.0 CLI 사용 설명](docs/CLI.md)을 사용하세요. **`Convert-PptxNative.cmd`가 새 방식이며, `Convert-PptxCli.cmd`는 기존 LibreOffice 방식**입니다.

## 브라우저에서 바로 사용하기

1. 이 저장소의 [`index.html`](index.html)을 열고 **Download raw file**로 파일을 저장합니다. 저장한 파일을 더블클릭해 Chrome 또는 Edge의 `file://` 주소에서 엽니다.
2. 또는 GitHub의 **Code → Download ZIP**을 선택하고 압축을 푼 뒤, 그 안의 `index.html`을 더블클릭합니다.
3. `.pptx` 파일을 선택하거나 화면에 끌어놓습니다. 최대 크기는 **64 MiB**입니다.
4. 필요한 경우 출력 제목, 슬라이드 범위(예: `1,3-5`), **숨긴 슬라이드 포함**을 설정한 뒤 **HTML로 변환**을 누릅니다.
5. 결과를 미리보고 **HTML 저장**을 누릅니다. 새 탭에서 결과를 열 수도 있습니다.

기본적으로 숨긴 슬라이드는 제외합니다. 슬라이드 범위의 번호는 원본 파일의 슬라이드 순서에 따른 **1부터 시작하는 번호**입니다. `1,3-5`처럼 입력하면 선택한 슬라이드가 원본 순서대로 출력됩니다. 변환 중에는 진행 상태를 볼 수 있고 취소할 수 있습니다.

저장한 HTML은 이미지 등 필요한 리소스가 내장된 독립 파일입니다. 해당 파일에서 페이지 탐색, 전체 슬라이드 스크롤, 한 장씩 보기, 인쇄를 사용할 수 있습니다. 결과 파일을 보거나 공유할 때 변환기나 원본 PPTX를 다시 열 필요가 없습니다.

## 암호가 설정된 PPTX

1. 파일을 선택하고 **문서 암호·보안 안내**을 펼칩니다.
2. **문서 열기 암호**에 PowerPoint에서 입력하는 암호를 넣습니다.
3. **HTML로 변환**을 누릅니다. 암호를 입력하지 않고 시작한 경우에는 입력란을 자동으로 보여 줍니다.

단일 `index.html`에서는 암호 확인과 복호화가 브라우저의 작업 스레드 안에서 실행되어 암호를 서버로 보내지 않습니다. Windows 앱은 내장 작업 스레드에서 처리합니다. **웹 서비스에서는 파일과 암호를 접속한 변환 서버로 전달**하며, 서버가 암호를 CLI의 표준 입력으로 넘깁니다. 변환 후 암호 입력란을 비우고, 브라우저 저장소·결과 HTML·진단 정보에 암호를 보관하지 않습니다. 원본 파일은 바뀌지 않습니다.

지원 범위는 Office Agile AES-CBC(128/192/256비트, SHA-1/256/384/512, 일반적인 16바이트 솔트)와 Standard AES-128/SHA-1입니다. Agile의 데이터 무결성 정보가 있으면 HMAC을 검증합니다. 인증서만 사용하는 문서, 다른 암호화 방식, 암호 반복 횟수 1,000,000회 초과는 지원하지 않습니다. 회사 DRM의 로그인이나 사용 권한은 이 문서 암호와 다릅니다.

## 회사 DRM·IRM 문서와 NASCA+SD

**PowerPoint를 실행하지 않는 것과 회사 DRM 암호문을 읽는 것은 별개의 문제입니다.** 일반 PPTX와 지원되는 문서 열기 암호가 있는 PPTX는 로컬 브라우저에서 변환할 수 있습니다. NASCA+SD 등 회사 DRM으로 보호되어 브라우저에 암호화된 데이터만 전달되는 문서는 현재 변환할 수 없습니다.

NASCA는 사용자 인증과 문서 권한 확인을 거쳐 복호화하는 문서 보안 제품입니다. [삼성SDS 공식 제품 설명](https://www.samsungsds.com/kr/endpoint-security-platform/endpoint-security-platform.html)을 참고하세요. PowerPoint에서 문서가 열린다는 사실만으로 로컬 HTML에도 복호화된 내용이 전달된다고 가정할 수는 없습니다. 공개 자료에서 이 변환기에 바로 넣을 수 있는 NASCA 브라우저용 복호화 SDK는 확인하지 못했습니다. 회사 계약에 따른 별도 연동의 유무는 조직이나 공급사 확인이 필요합니다.

| 브라우저가 실제로 받는 데이터 | 현재 처리 |
| --- | --- |
| 일반 PPTX 또는 조직에서 허용한 일반 PPTX 사본 | 브라우저 안에서 변환 |
| 지원되는 Office 문서 열기 암호가 있는 PPTX | 사용자가 입력한 암호로 브라우저 안에서 복호화 후 변환 |
| 회사 DRM 암호문 또는 지원되지 않는 보호 형식 | 오류 코드와 형식 진단을 표시하고 중지 |

새 Native CLI와 Windows 앱은 브라우저와 같은 JavaScript/WASM 변환 엔진을 사용합니다. 파일을 읽는 프로그램에 평문 PPTX가 제공되는 경우에만 내용을 변환할 수 있습니다. 회사 보안 프로그램이 `node.exe` 또는 `PPPTX_to_HTML.exe`에 그 읽기와 결과 저장을 허용하는지는 이 도구가 결정하지 않습니다. 현재 버전에는 NASCA 전용 복호화가 구현되어 있지 않으며, 실제 NASCA 문서 성공 여부는 검증되지 않았습니다. 웹 서비스는 이 CLI를 실행하므로 같은 제한이 적용되며, 기존 LibreOffice CLI도 회사 DRM을 대신 인증하지 않습니다.

### 이전 PowerPoint 도우미와의 차이

v1.1.0 도우미는 PowerPoint의 슬라이드 PNG 내보내기 결과를 HTML에 넣는 이미지 기반 방식이었습니다. PowerPoint 자체의 외부 통신이나 회사 보안 제품의 제어까지 이 도우미가 제한하지는 못했습니다. 현재 Windows 앱·`index.html`·두 CLI는 해당 도우미를 호출하지 않습니다. PowerPoint·LibreOffice를 모두 사용하지 않으려면 **Windows 앱 또는 Native CLI**를 사용하세요.

저장소 `tools/Convert-PptxWithPowerPoint.*`는 이전 버전 참고 소스입니다. 새 CLI의 구성 요소가 아니며, 이전 도우미의 PNG 내보내기 오류를 고쳤거나 실제 NASCA+SD 문서 변환에 성공했다고 주장하지 않습니다.

## Native CLI와 브라우저 변환 범위

- 원본의 애니메이션과 슬라이드 전환은 정적 화면으로 표현됩니다.
- 차트는 종류와 축·캐시 구성에 따라 지원 범위가 다르며, 기본형 차트도 대체 표시로 나올 수 있습니다. SVG로 그려진 차트도 개별 계열의 채우기·선·전체 제목을 모두 재현하지 않으며 테마 팔레트를 사용할 수 있습니다. 미지원 차트는 대응 위치가 확실할 때 저장된 데이터를 표로 보완합니다. 외부 workbook은 열지 않으며, 캐시 표에는 숫자 서식을 적용하지 않습니다. 캐시가 없거나 한도를 넘거나 매핑이 모호한 경우에는 해당 안내 또는 기존 대체 표시를 유지합니다.
- SmartArt, OLE 개체, 수식은 대체 표현이 사용되거나 원본과 차이가 날 수 있습니다.
- 글꼴이 보는 기기에 설치되어 있는지에 따라 줄바꿈과 배치가 달라질 수 있습니다.
- 브라우저 변환은 구형 `.ppt`를 지원하지 않습니다. 인식한 구형 형식은 별도로 안내하며, 문서 열기 암호는 위 지원 범위에서 처리합니다.
- 외부 위치를 참조하는 이미지나 미디어는 오프라인에서 로드되지 않습니다.
- 일반 웹사이트 하이퍼링크는 클릭하면 브라우저의 새 탭에서 열립니다. 자동으로 외부 자료를 가져오지는 않습니다.
- 발표자 노트, 코멘트, 진단용 원시 메타데이터는 결과 HTML에 내보내지 않습니다.
- 과도한 메모리 사용을 방지하기 위해 ZIP 내부 파일 20,000개, 압축 해제 크기 합계 256 MiB, 관계 파일 한 개 4 MiB 한도를 둡니다. 변환은 2분 후 중지하거나 직접 취소할 수 있습니다.
- 오류가 나면 **파일 형식 진단 정보**에서 오류 코드, 크기, 인식한 형식 표식을 확인할 수 있습니다. 문서 내용이나 암호는 표시하지 않습니다.

중요한 발표 자료는 저장한 HTML을 열어 내용과 배치를 확인해 주세요. 원본 PPTX가 제공되지 않았으므로 특정 사용자 문서와의 시각적 일치 여부는 검증하지 않았습니다.

## 개발 및 검증

Native CLI 직접 실행 및 저장소 개발에는 Node.js가 필요합니다. 배포된 Windows 앱이나 단일 `index.html`을 사용하는 경우에는 설치할 필요가 없습니다.

```bash
npm ci
npm run build
npx playwright install chromium
npm run test:core
npm run test:charts
npm run test:engine
npm run test:fidelity
npm run test:links
npm test
npm run test:native
npm run test:web
npm run package:native
npm run package:desktop
```

`npm run build`는 의존성이 설치된 뒤 인터넷 연결 없이 실행할 수 있습니다. Playwright가 사용할 Chromium을 이미 설치했다면 `CHROME_BIN` 환경 변수에 실행 파일 경로를 지정할 수 있습니다. `npm test`는 먼저 테스트용 PPTX를 생성합니다. `npm run test:fixtures`로 테스트 파일만 다시 만들 수도 있습니다. Linux에서 한글 스크린샷을 검증하려면 로컬 한글 글꼴도 필요합니다.

### 확인한 동작

웹 서비스는 실제 Node CLI를 실행하여 일반·암호 PPTX, 숨김·범위, 미리보기·다운로드를 검증했습니다. 외부 요청을 차단한 상태에서 결과 HTML 탐색, 데스크톱·모바일 화면, 취소·시간 초과 후 프로세스와 임시 파일 정리, 같은 포트의 서비스 재시작 후 재연결을 확인했습니다. Host·Origin·요청 토큰과 업로드 한도 검사도 포함합니다. 이 테스트는 개발 환경에서 실행했으며 실제 회사 서버나 NASCA 문서의 성공을 의미하지 않습니다.

Chrome for Testing의 실제 `file://` 환경에서 HTTP/HTTPS 요청을 모두 차단한 상태로 아래 흐름을 검증했습니다. 파일 접근 보안을 완화하는 브라우저 옵션이나 개발 서버는 사용하지 않았습니다.

- 한글 텍스트, 표, 내장 PNG, SVG 도형, 차트가 있는 PPTX 변환
- 파일명 순서와 다른 실제 슬라이드 순서, 숨김 제외/포함, 슬라이드 범위 선택
- 결과 HTML 다운로드 후 별도 `file://` 페이지에서 다시 열기와 이미지 표시
- 저장 문서의 슬라이드 선택, 이전/다음, 한 장씩 보기와 모바일 가로 넘침 확인
- 잘못된 파일 처리, 설정 변경 후 결과 무효화, 변환 취소 후 늦은 응답 무시
- 문서 속 실행 코드가 텍스트로 유지되는지, 외부 이미지 요청이 없는지, 발표자 노트가 내보내지지 않는지 확인
- 독립적인 Python msoffcrypto-tool로 생성한 Agile 문서의 복호화, 암호 미입력·불일치, Standard 암호 문서와 무결성 오류 확인
- 정상 ZIP의 앞부분에 추가 바이트가 있는 경우, 불완전한 ZIP, IRM·구형 PPT·알 수 없는 복합 문서의 구분
- PowerPoint 도우미의 UI·내장 파일·다운로드 동작이 없는 브라우저 전용 배포 확인
- 일반 PPTX의 로컬 변환과 보호된 입력의 오류 안내에서 HTTP/HTTPS 요청이 없는지 확인

테스트는 생성한 PPTX를 기준으로 합니다. 모든 PowerPoint 기능이나 실제 사용자 문서의 시각적 일치를 보장하지 않습니다. [환경별 자동 검사](.github/workflows/compatibility.yml)는 Windows 배포 EXE, 세 OS의 Node 24·Chromium, Linux의 Firefox·WebKit을 대상으로 합니다. Playwright WebKit 검사는 실제 Safari·iOS 기기 검사를 대체하지 않습니다. 각 커밋의 실제 통과 여부는 [Actions](https://github.com/1bobby-git/PPPTX_to_HTML/actions)에서 확인합니다.

다른 브라우저 엔진을 로컬에서 검사하려면 해당 Playwright 브라우저를 설치한 뒤 `PPTX_TEST_BROWSER=firefox node tests/offline.cjs` 또는 `PPTX_TEST_BROWSER=webkit node tests/offline.cjs`를 실행합니다. Windows PowerShell에서는 `$env:PPTX_TEST_BROWSER="firefox"`처럼 환경 변수를 먼저 지정합니다.

### 주요 파일

| 경로 | 역할 |
| --- | --- |
| `desktop/`, `docs/DESKTOP.md` | 로컬 파일 선택·변환·미리보기·저장 화면, 제한된 IPC, 원본 직접 읽기와 Windows 사용 안내 |
| `scripts/package-desktop.cjs`, `tests/desktop.cjs` | 고정 Electron 런타임과 공통 엔진을 포함한 포터블 배포 생성 및 실제 배포 앱 검증 |
| `tools/serve-web.cjs`, `docs/WEB_SERVICE.md` | 웹 화면·API를 제공하고 Native CLI를 실행하는 서비스와 운영 안내 |
| `src/web-shell.html`, `src/server-client.js`, `tests/web-service.cjs` | 웹 서비스 화면, 요청·취소 처리, 실제 CLI/API/브라우저 검증 |
| `index.html` | 배포용 단일 HTML 파일. 다운로드 후 바로 실행합니다. |
| `tools/convert-pptx.cjs`, `tools/Convert-PptxNative.cmd` | Node.js에서 PPTX ZIP/XML을 HTML·SVG로 직접 변환하는 CLI |
| `src/native-cli-worker.cjs`, `docs/NATIVE_CLI.md` | Native CLI의 작업 스레드와 사용 설명 |
| `tests/native-cli.cjs`, `scripts/package-native-cli.cjs` | Native CLI 검증 및 의존성을 포함한 오프라인 ZIP 생성 |
| `tools/convert-pptx.py`, `tools/Convert-PptxCli.cmd` | PowerPoint 없이 LibreOffice headless → PDF → SVG → 독립 HTML을 만드는 CLI |
| `tools/requirements-cli.txt`, `docs/CLI.md` | CLI 의존성 버전과 사용 설명 |
| `tests/cli.py`, `tests/cli-browser.cjs` | 실제 CLI 변환과 결과 HTML의 로컬 브라우저 검증 |
| `src/shell.html`, `src/styles.css`, `src/app.js` | 변환기 화면과 사용자 흐름 |
| `src/worker.js` | 브라우저 내 PPTX 변환 작업 |
| `src/conversion-core.js`, `tests/conversion-core.cjs` | 브라우저·Node 공통 변환 계약과 장수·범위·정화 회귀 검사 |
| `src/chart-data.js`, `tests/chart-data.cjs` | 미지원 차트의 내부 캐시 데이터와 보수적인 위치 매핑 |
| `vendor/pptx2html-engine/`, `tests/chart-engine.cjs` | 동일 엔진의 국소 차트 파서 패치, 고정 빌드 출처·라이선스·배포 패키지와 실제 WASM 회귀 검사 |
| `src/theme-resolver.js` | 여러 master/theme의 명확한 색·글꼴 참조와 제한된 placeholder 상속 보정 |
| `tests/fidelity.cjs`, `scripts/create-fidelity-fixtures.cjs` | 병합 표·중첩 그룹·여러 테마를 사용하는 실제 OOXML 회귀 사례 |
| `docs/REFERENCE_REVIEW.md`, `.github/workflows/compatibility.yml` | 참고자료 분석·중복 제거 판단과 환경별 자동 검사 |
| `src/input-package.js` | 문서 형식·보호 방식 판별, 입력 암호를 이용한 Office 복호화 |
| `src/normalize-package.js` | PPTX 내부 절대 관계 경로 정규화와 압축 크기 검사 |
| `src/exporter.js` | 변환 결과의 독립 HTML 구성 |
| `tests/internal-links.cjs` | 실제 PPTX의 텍스트·도형·이미지 이동, 숨김·범위 선택, 키보드와 정화 회귀 검사 |
| `src/viewer.js`, `src/viewer.css` | 저장 결과의 슬라이드 탐색·보기·인쇄 UI |
| `scripts/build.mjs` | 소스와 의존성을 단일 `index.html`로 묶는 빌드 |
| `scripts/create-fixtures.cjs`, `tests/offline.cjs` | 생성 fixture 및 오프라인 검증 |
| `scripts/create-protected-fixtures.cjs`, `tests/protection.cjs`, `tests/reference/` | 암호·파일 형식 회귀 테스트와 독립 암호화 기준 파일 |
| `tools/Convert-PptxWithPowerPoint.ps1`, `.cmd` | v1.1.0의 참고 소스. 현재 HTML 배포·빌드·실행에는 사용하지 않음 |
| `licenses/` | 빌드 시 자동 생성되는 포함 의존성의 라이선스 고지 |

## 오픈 소스 고지

변환 엔진은 [`@briank-dev/pptx-to-html` 2.1.0](https://www.npmjs.com/package/@briank-dev/pptx-to-html)([상위 프로젝트 문서](https://github.com/brnyxx/pptx2html-turbo), MIT)에 차트 파서의 국소 수정을 적용한 **`2.1.0-ppptx.1`**입니다. 같은 이름의 로컬 tarball 패키지 하나를 설치하며, 별도 렌더러를 병렬로 추가하지 않습니다. 고정 원본 소스·수정 patch·빌드 도구·해시·MIT 원문은 [`vendor/pptx2html-engine/`](vendor/pptx2html-engine/)에 있습니다. 일반 사용자는 Rust나 빌드 도구를 설치할 필요가 없습니다.

HTML 정제에는 [`DOMPurify` 3.4.17](https://github.com/cure53/DOMPurify)(Apache-2.0 또는 MPL-2.0), 패키지 경로 호환 처리에는 [`JSZip` 3.10.1](https://github.com/Stuk/jszip)(MIT 선택)을 사용합니다. 빌드 과정에서 포함 의존성의 라이선스 문서가 `licenses/`에 생성되며, 엔진 버전·빌드 출처와 원문 고지는 배포용 `index.html` 안에도 보존됩니다.

Office 암호 처리에는 [`officecrypto-tool` 0.0.19](https://github.com/zurmokeeper/officecrypto-tool)의 암호화 모듈, `cfb` 1.2.2, `xml2js` 0.6.2 및 Browserify 브라우저 구현을 포함합니다. 실제 번들에 포함된 패키지의 고지는 `licenses/OfficeCrypto-BUNDLED-LICENSES.txt`와 HTML 안에 함께 들어갑니다. 독립 기준 파일의 생성 정보와 해시는 `tests/reference/manifest.json`에 있습니다.

Native CLI는 같은 변환·입력·정화 코드를 재사용하며 DOM 처리에 [`jsdom` 30.1.3](https://github.com/jsdom/jsdom)(MIT)을 사용합니다. 의존성을 포함한 Native ZIP에는 해당 패키지의 라이선스 파일도 보존합니다. 브라우저나 Office 프로세스를 실행하지 않고 Node.js의 작업 스레드 안에서 처리합니다.

Windows 앱은 [Electron](https://www.electronjs.org/) 45.0.0의 공식 런타임을 포함합니다. 내장 Node.js의 작업 스레드에서 동일한 변환 엔진을 실행하며 UI와 문서 미리보기는 격리합니다. Electron의 `LICENSE`·`LICENSES.chromium.html`과 포함된 패키지의 라이선스를 배포 폴더에 보존합니다. 공식 런타임 URL·크기·SHA256은 `desktop/runtime.json`에 고정합니다.

사용자가 제시한 [`g21589/PPTX2HTML`](https://github.com/g21589/PPTX2HTML)의 ZIP/XML → HTML·SVG 변환 구조를 참고했습니다. 해당 프로젝트는 브라우저용 JavaScript 구현입니다. 이 저장소의 실제 변환은 위에 명시한 `pptx2html-turbo` 엔진을 재사용합니다.

웹 서비스의 multipart 업로드 처리에는 [`busboy` 1.6.0](https://github.com/mscdex/busboy)(MIT)을 사용합니다. 웹 UI는 CDN 없이 저장소의 정적 파일을 같은 서버에서 받습니다.

저장소 자체 소스의 라이선스는 [MIT](LICENSE)입니다. 기존 Python CLI가 별도로 사용하는 [LibreOffice](https://www.libreoffice.org/about-us/licenses/)와 [PyMuPDF](https://pymupdf.readthedocs.io/en/latest/about.html)의 라이선스는 각각의 프로젝트를 따릅니다. PyMuPDF와 MuPDF는 AGPL-3.0 또는 Artifex 상용 라이선스로 제공되며, MIT 고지가 이 외부 의존성의 라이선스를 대체하지 않습니다. Native CLI에는 LibreOffice와 PyMuPDF가 필요하지 않으며 Native ZIP에도 포함하지 않습니다.

GitHub Pages에 게시하지 않아도 로컬에서 실행할 수 있습니다.

# PPTX → HTML Native CLI · v1.5.1

이 CLI는 **PowerPoint, LibreOffice, 브라우저 프로그램을 실행하지 않습니다.** Node.js 안에서 PPTX의 ZIP/XML을 읽고, 기존 JavaScript/WASM 변환 엔진으로 HTML·표·이미지·SVG 도형을 만듭니다. PDF나 슬라이드 전체 PNG를 중간 결과로 저장하지 않습니다.

결과물은 이미지와 탐색 UI가 내장된 **HTML 한 파일**입니다. 변환이 끝나면 해당 HTML을 더블클릭하여 로컬 브라우저에서 볼 수 있습니다. 결과를 보는 컴퓨터에는 Node.js나 원본 PPTX가 필요하지 않습니다.

사용자가 CLI를 직접 다루지 않게 하려면 [웹 서비스 안내](WEB_SERVICE.md)를 따라 운영자가 서비스를 준비하세요. v1.5.1 의존성 포함 ZIP에도 웹 서비스 파일이 들어 있습니다. 일반 사용자는 제공된 웹 주소만 열면 됩니다.

## Windows에서 실행

### 1. Node.js 준비

[Node.js 공식 다운로드](https://nodejs.org/en/download)에서 **24 LTS 최신 버전**을 설치합니다. 지원 버전은 22.x의 22.22.2 이상, 24.x의 24.15.0 이상, 또는 26 이상입니다. 설치 후 새 PowerShell이나 명령 프롬프트를 열어 확인합니다.

```powershell
node --version
```

Node.js를 설치할 수 없는 환경에서 회사가 허용한 휴대용 Node.js를 사용한다면, 프로젝트 최상위 폴더에 `node.exe`를 둘 수 있습니다. `tools/Convert-PptxNative.cmd`는 그 파일을 먼저 찾고, 없으면 `PATH`의 `node`를 사용합니다. 제공하는 ZIP에 Node.js 실행 파일은 포함하지 않습니다.

### 2. 변환기 준비

**의존성 포함 배포본** `PPPTX_to_HTML_Native_v1.5.1.zip`을 받았다면 압축만 풉니다. `node_modules`도 함께 들어 있으므로 npm 설치 단계 없이 오프라인에서 사용할 수 있습니다.

**GitHub의 Code → Download ZIP**으로 소스를 받았다면, 압축을 푼 프로젝트 폴더에서 최초 한 번 다음 명령을 실행합니다.

```powershell
npm.cmd ci --omit=dev --ignore-scripts --no-audit --no-fund
```

이 설치 단계에는 인터넷 또는 사내 npm 저장소가 필요합니다. 설치 후 변환 과정에는 서버나 인터넷 연결이 필요하지 않습니다. `npm.cmd`를 명시하므로 `npm.ps1`의 PowerShell 실행 정책 문제를 피할 수 있습니다. 실행 정책을 변경하지 않습니다.

### 3. 점검과 변환

압축을 푼 프로젝트 폴더에서 실행합니다.

```powershell
.\tools\Convert-PptxNative.cmd --check
.\tools\Convert-PptxNative.cmd "C:\test.pptx" -o "C:\test.html"
```

Node.js를 직접 호출해도 같습니다.

```powershell
node .\tools\convert-pptx.cjs "C:\test.pptx" -o "C:\test.html"
```

출력 폴더는 미리 만들어져 있어야 합니다. `-o`를 생략하면 원본 옆에 같은 이름의 `.html`을 저장합니다. 원본 파일은 수정하지 않습니다.

이미 존재하는 HTML을 교체하려면 `--overwrite`를 붙입니다. 변환이 끝난 뒤에만 결과를 쓰며, 변환 중 실패하면 기존 HTML을 보존합니다.

```powershell
.\tools\Convert-PptxNative.cmd "C:\test.pptx" -o "C:\test.html" --overwrite
```

**이름을 구분하세요:** `Convert-PptxNative.cmd`가 Node.js 방식이고, 기존 `Convert-PptxCli.cmd`는 Python·LibreOffice 방식입니다. 예전 `Convert-PptxWithPowerPoint.*`는 사용하지 않습니다.

## 옵션

| 옵션 | 동작 |
| --- | --- |
| `-o`, `--output FILE` | 출력 HTML 경로 |
| `--slides "1,3-5"` | 원본 슬라이드 번호로 범위 선택 |
| `--include-hidden` | 숨긴 슬라이드 포함. 기본은 제외 |
| `--title "검토 자료"` | HTML 제목 지정, 최대 300자 |
| `--timeout 120` | 변환 작업 스레드의 제한 시간(초). 암호 표준 입력 대기에도 별도로 적용. 기본 120, 범위 1~3600 |
| `--overwrite` | 성공한 결과로 기존 출력 교체 |
| `--password-stdin` | 지원되는 Office 문서 열기 암호를 표준 입력에서 읽기 |
| `--json` | 표준 출력에 결과 JSON 하나를 기록. 진행 상태는 표준 오류에 기록 |
| `--check` | 입력 파일 없이 실행 환경 검사 |
| `--help` | 사용법 표시 |

```powershell
.\tools\Convert-PptxNative.cmd "C:\자료\보고서.pptx" -o "C:\자료\보고서.html" --slides "1,3-5" --include-hidden --title "검토 자료" --json
```

번호는 PPTX 내부 XML 파일명이 아니라 **원본 프레젠테이션 순서의 1부터 시작하는 번호**입니다. 숨긴 슬라이드를 선택하려면 `--include-hidden`을 함께 지정합니다. 선택한 범위에 표시할 슬라이드가 없으면 실패하며 빈 HTML을 만들지 않습니다.

입력 최대 64 MiB, 압축 내부 항목 최대 20,000개, 선언된 압축 해제 크기 합계 256 MiB, 개별 관계 파일 4 MiB, 원본 슬라이드 최대 1,000장, 결과 HTML 최대 256 MiB 한도가 있습니다. 시간 제한은 문서 복호화·정규화·렌더링·HTML 정화를 수행하는 작업 스레드에 적용됩니다.

## 문서 열기 암호

지원되는 일반 Office 문서 열기 암호는 `--password-stdin`으로 전달할 수 있습니다. 암호를 명령 인수로 받는 `--password` 옵션은 없습니다. 암호가 있는 파일에 입력이 없으면 `PASSWORD_REQUIRED`, 암호가 다르면 `INCORRECT_PASSWORD`를 반환합니다.

자동화 프로그램에서 UTF-8로 인코딩한 암호를 CLI의 표준 입력에 쓰고 입력 스트림을 닫습니다. 마지막 줄바꿈 하나는 제거하며 암호 길이는 최대 255자, 입력 바이트는 최대 1,024바이트입니다. 일반적인 한글 암호도 UTF-8로 전달해야 합니다. 암호를 JSON·진행 로그·결과 HTML에 기록하지 않습니다. 입력 스트림이 닫히지 않으면 `--timeout` 후 중지하며, Ctrl+C도 처리합니다.

직접 입력할 때는 Windows PowerShell의 숨김 입력을 다음처럼 사용할 수 있습니다. 실제 암호를 명령 문자열에 적지 않습니다.

```powershell
$pptxPassword = Read-Host "문서 열기 암호" -AsSecureString
$pptxPasswordPtr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($pptxPassword)
$pptxPreviousEncoding = $OutputEncoding
try {
    $OutputEncoding = [System.Text.UTF8Encoding]::new($false)
    [Runtime.InteropServices.Marshal]::PtrToStringBSTR($pptxPasswordPtr) | node .\tools\convert-pptx.cjs "C:\test.pptx" -o "C:\test.html" --password-stdin
} finally {
    $OutputEncoding = $pptxPreviousEncoding
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pptxPasswordPtr)
    $pptxPassword.Dispose()
}
```

지원하는 형식은 Office Standard AES-128/SHA-1 및 Agile AES-CBC 128/192/256, SHA-1/256/384/512, 일반적인 16바이트 솔트입니다. Agile 무결성 정보가 있으면 HMAC을 검사합니다. 지원하지 않는 암호화·인증서 전용 문서는 오류로 중지합니다. 이는 회사 DRM 인증과 다른 기능입니다.

## NASCA+SD 문서

`g21589/PPTX2HTML`처럼 ZIP/XML을 해석하는 방식은 **프로그램이 실제로 읽은 바이트가 PPTX 구조일 때** 동작합니다. NASCA가 Node.js에 복호화된 읽기를 허용한다면 그 내용으로 변환을 시도할 수 있지만, 암호문이 전달되면 XML을 읽을 수 없습니다. 이 프로젝트가 NASCA의 허용 프로그램이나 내보내기 권한을 바꾸지는 않습니다.

이 CLI에는 NASCA 복호화 SDK나 DRM 해제 기능이 없습니다. 인식되는 IRM 형식은 `DRM_PROTECTED`, 인식하지 못한 구조는 `UNKNOWN_CONTAINER` 등으로 구분합니다. 알 수 없는 입력을 단순히 "확장자만 바꾼 파일"이라고 판단하지 않습니다. 실제 NASCA+SD 보호 파일과 사용자 Windows PC에서의 성공 여부는 검증하지 못했습니다.

PowerPoint와 Office의 클라우드 연동을 호출하지 않습니다. 변환 코드에는 HTTP 업로드나 원격 변환 요청이 없으며, jsdom에는 외부 리소스 로딩과 문서 속 스크립트 실행을 활성화하지 않습니다. 이 설명은 해당 변환 코드에 대한 것이며, PC의 다른 프로그램이나 보안 에이전트의 통신까지 제어한다는 뜻은 아닙니다.

## 결과 표현과 제한

- 텍스트는 HTML 텍스트, 표는 셀, 도형은 HTML/SVG, 포함 이미지는 data URI로 저장합니다. HTML에서 텍스트 선택·복사를 할 수 있습니다.
- 슬라이드 선택, 이전/다음, 전체 보기·한 장 보기, 인쇄를 지원합니다. 뷰어의 번호는 선택된 출력 안에서의 순서입니다.
- 발표자 노트·코멘트·원시 진단 메타데이터와 엔진 실행 코드는 결과에서 제거합니다.
- 자동으로 불러오는 외부 리소스는 제거하고, 포함된 이미지와 미디어만 유지합니다. 일반 웹 하이퍼링크는 사용자 클릭으로 열 수 있습니다.
- 애니메이션·전환은 정적인 화면입니다. SmartArt, OLE, 수식, 일부 차트·효과는 대체 표시 또는 경고가 생길 수 있습니다.
- 미지원 차트는 원래 순서·관계·위치가 유일하게 일치할 때 내부 캐시의 원시 값을 표로 표시합니다. 외부 workbook은 열지 않으며 숫자 서식은 적용하지 않습니다. 최대 100개 차트, 차트당 100행·16계열, 셀당 256자, XML당 2 MiB·누적 16 MiB, 메타데이터 1 MiB 한도가 적용됩니다.
- v1.5.1의 `2.1.0-ppptx.1` 엔진은 DrawingML 스타일 때문에 정상 차트를 잘못된 축으로 오인하던 파서 오류를 수정합니다. SVG 차트의 개별 계열 채우기·선·전체 제목까지 완전하게 재현하는 기능은 아니며, 잘못된 축과 미지원 차트는 기존 경고·대체 표시를 유지합니다. `--check`의 `enginePackage`와 `engineVersion`으로 실제 설치 엔진을 확인할 수 있습니다.
- 여러 master/theme는 관계가 명확한 색·명시 글꼴과 제한된 placeholder 상속을 보완합니다. 불확실한 상속·override는 원본을 유지하고 `THEME_INHERITANCE_UNRESOLVED` 경고를 반환합니다. 구현 범위와 실제 회귀 사례는 [참고자료 분석](REFERENCE_REVIEW.md)에 있습니다.
- 글꼴은 결과를 보는 기기에 설치된 글꼴에 영향을 받습니다. 모든 문서의 배치와 줄바꿈이 PowerPoint와 같지는 않습니다.

## 오류와 자동화

`--json`의 성공 결과는 `ok: true`, `engine: "native"`, `version`, `output`, `slideCount`, `width`, `height`, `warnings`, `sourceSlides`를 포함합니다. `sourceSlides`는 실제로 출력한 원본 슬라이드 번호이며, 엔진의 번호를 확인할 수 없으면 `null`입니다. 실패 결과는 `ok: false`, `code`, `error`를 포함합니다. `warnings`가 있으면 지원되지 않는 기능이 있으므로 결과를 확인하세요.

| 주요 코드 | 의미 |
| --- | --- |
| `UNSUPPORTED_NODE` | Node.js 버전이 지원 범위보다 낮음 |
| `DEPENDENCY_MISSING` | 변환 의존성 또는 WASM 파일이 누락됨 |
| `OUTPUT_EXISTS` | 출력이 이미 존재함. 교체하려면 `--overwrite` 사용 |
| `PATH_COLLISION` | 입력과 출력이 같은 파일 또는 같은 파일을 가리키는 경로임 |
| `INPUT_LIMIT`, `OUTPUT_LIMIT`, `SLIDE_LIMIT` | 크기 또는 슬라이드 수 한도 초과 |
| `PASSWORD_REQUIRED`, `INCORRECT_PASSWORD` | 문서 열기 암호 미입력 또는 불일치 |
| `DRM_PROTECTED` | 인식되는 DRM/IRM 보호 구조 |
| `UNKNOWN_CONTAINER` | 다른 문서 형식 또는 알 수 없는 보호 구조 등, 원인을 단정할 수 없음 |
| `INVALID_CONTAINER`, `INVALID_PACKAGE` | 읽을 수 없는 구조·무결성 오류 |
| `UNSUPPORTED_ENCRYPTION` | 지원하지 않는 암호화 방식 |
| `EMPTY_SELECTION` | 선택한 범위에 표시할 슬라이드가 없음 |
| `SLIDE_COUNT_MISMATCH` | 엔진이 만든 슬라이드 수와 정화 후 출력 수가 달라 중지 |
| `TIMEOUT`, `CANCELLED` | 시간 초과 또는 Ctrl+C 취소 |
| `CONVERSION_FAILED` | 그 외 엔진 처리 오류 |

정상 완료는 종료 코드 `0`, 처리 실패는 `1`, 잘못된 명령 인수는 `2`, 취소는 `130`입니다. 파일 경로·권한 문제는 운영체제 오류 코드가 함께 표시될 수 있습니다. 출력은 같은 폴더의 임시 파일을 완성한 뒤 교체합니다. 기본 덮어쓰기 방지는 파일시스템의 하드 링크를 사용하므로, 이를 지원하지 않는 일부 외장·네트워크 파일시스템에서는 저장 오류가 날 수 있습니다.

## Linux / macOS

같은 지원 버전의 Node.js와 의존성을 준비하고 실행합니다.

```bash
npm ci --omit=dev --ignore-scripts --no-audit --no-fund
node tools/convert-pptx.cjs --check
node tools/convert-pptx.cjs ./presentation.pptx -o ./presentation.html
```

의존성 포함 Native ZIP을 사용하는 경우 npm 단계는 생략합니다.

## 개발 및 배포본 생성

다음 명령은 GitHub의 전체 소스를 받은 개발 환경에서 사용합니다. 의존성 포함 Native ZIP에는 실행에 필요한 파일만 들어 있습니다.

```bash
npm ci --ignore-scripts --no-audit --no-fund
npm run test:native
npm run package:native
```

테스트는 CLI 자체의 외부 프로그램 실행·네트워크 호출을 금지한 상태에서 변환을 확인하고, 생성 HTML은 별도로 Chromium의 `file://`에서 검증합니다. 브라우저 검증에는 테스트용 Chromium이 필요하며 `CHROME_BIN`으로 경로를 지정할 수 있습니다. CLI 실행에 Chromium이 필요한 것은 아닙니다.

배포본 생성은 설치된 production 의존성과 필요한 소스·라이선스를 ZIP으로 묶습니다. 사용자 문서, 테스트 입력, Node.js 실행 파일, Python·LibreOffice·PyMuPDF는 포함하지 않습니다.

## 참고 프로젝트와 라이선스

- [`g21589/PPTX2HTML`](https://github.com/g21589/PPTX2HTML): 사용자가 제시한 JavaScript ZIP/XML → HTML·SVG 구현을 구조 참고로 검토했습니다. 이 CLI의 렌더러로 해당 저장소를 직접 가져오지는 않았습니다.
- [`pptx2html-turbo` / `@briank-dev/pptx-to-html`](https://github.com/brnyxx/pptx2html-turbo): 실제 변환 엔진, MIT. 2.1.0 기반의 로컬 패치 버전 **2.1.0-ppptx.1**을 사용합니다. [패치와 재현 빌드 안내](../vendor/pptx2html-engine/README.md)를 참고하세요.
- [`DOMPurify` 3.4.17](https://github.com/cure53/DOMPurify): HTML·SVG 정화, Apache-2.0 또는 MPL-2.0.
- [`jsdom` 30.1.3](https://github.com/jsdom/jsdom): Node.js 안의 DOM 처리, MIT.
- [`JSZip` 3.10.1](https://github.com/Stuk/jszip): ZIP 읽기 및 호환 처리, MIT 선택.
- [`officecrypto-tool` 0.0.19](https://github.com/zurmokeeper/officecrypto-tool), `cfb`, `xml2js`: 지원되는 일반 Office 암호 문서 처리.

프로젝트 소스는 [MIT](../LICENSE)이며 각 의존성은 해당 패키지의 라이선스를 따릅니다. 배포 ZIP에 포함된 `node_modules`의 라이선스 고지를 보존합니다. 별도 Python CLI의 PyMuPDF·LibreOffice 라이선스는 그 경로에 적용되며 이 Native CLI에는 해당 의존성이 없습니다.

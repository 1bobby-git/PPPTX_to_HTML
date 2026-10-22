# PPTX → HTML

`.pptx` 프레젠테이션을 독립적인 HTML 파일로 변환하는 로컬 도구입니다. 변환할 파일을 서버로 업로드하지 않습니다. 일반 PPTX와 지원되는 문서 열기 암호가 설정된 PPTX는 **약 4.1 MiB의 `index.html` 한 파일**로 브라우저 안에서 처리합니다. 이 경로에는 Node.js, npm, 별도 서버, CDN, 인터넷 연결 또는 PowerPoint 설치가 필요하지 않습니다.

**v1.3.0에는 PowerPoint를 실행하지 않는 CLI 변환기가 추가되었습니다.** LibreOffice를 `--headless`로 호출하고, 변환한 각 페이지를 SVG로 만들어 한 개의 HTML에 넣습니다. 변환 중 PowerPoint·LibreOffice 창이나 브라우저를 열지 않습니다. 기존 단일 `index.html` 방식도 그대로 사용할 수 있습니다.

## CLI로 바로 변환하기 — Windows

[Python](https://www.python.org/downloads/windows/)과 [LibreOffice](https://www.libreoffice.org/download/download-libreoffice/)를 설치하고, 이 저장소를 **Code → Download ZIP**으로 내려받아 압축을 풉니다. 저장소 폴더에서 아래 명령을 실행하세요.

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r .\tools\requirements-cli.txt
.\.venv\Scripts\python.exe .\tools\convert-pptx.py --check
.\.venv\Scripts\python.exe .\tools\convert-pptx.py "C:\test.pptx" -o "C:\test.html"
```

이후에는 다음 명령으로 변환할 수 있습니다.

```powershell
.\tools\Convert-PptxCli.cmd "C:\test.pptx" -o "C:\test.html"
```

출력 파일이 이미 있으면 덮어쓰지 않습니다. 교체하려면 `--overwrite`를 추가하세요. 가상환경 활성화나 PowerShell 실행 정책 변경은 필요하지 않습니다. **의존성 설치가 끝난 뒤에는 변환 서버나 인터넷 연결 없이 실행**합니다.

- 출력: SVG 슬라이드·탐색 UI·읽기용 텍스트가 내장된 HTML 한 파일. 슬라이드 전체를 PNG로 캡처하지 않습니다.
- 지원: 숨김 슬라이드 포함, 원본 슬라이드 범위, 제목, JSON 결과, 시간 제한, 실행 환경 점검.
- 보안 문서: 파일의 ZIP 표식만으로 먼저 거절하지 않고 LibreOffice에 원본 파일 읽기를 요청합니다. **NASCA+SD 복호화 기능이 추가된 것은 아닙니다.** 해당 PC의 보안 정책과 LibreOffice의 읽기 권한에 따라 실패할 수 있으며, 실제 NASCA 문서 성공 여부는 검증되지 않았습니다.

설치 조건·옵션·오류 코드·한계는 [CLI 사용 설명](docs/CLI.md)에 정리했습니다.

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

암호 확인과 복호화는 브라우저의 작업 스레드 안에서 실행합니다. 암호를 서버로 보내거나 브라우저 저장소에 보관하지 않으며, 변환 요청 후 입력란을 비웁니다. 결과 HTML과 진단 정보에도 암호를 포함하지 않습니다. 원본 파일은 바뀌지 않습니다.

지원 범위는 Office Agile AES-CBC(128/192/256비트, SHA-1/256/384/512, 일반적인 16바이트 솔트)와 Standard AES-128/SHA-1입니다. Agile의 데이터 무결성 정보가 있으면 HMAC을 검증합니다. 인증서만 사용하는 문서, 다른 암호화 방식, 암호 반복 횟수 1,000,000회 초과는 지원하지 않습니다. 회사 DRM의 로그인이나 사용 권한은 이 문서 암호와 다릅니다.

## 회사 DRM·IRM 문서와 NASCA+SD

**PowerPoint를 실행하지 않는 것과 회사 DRM 암호문을 읽는 것은 별개의 문제입니다.** 일반 PPTX와 지원되는 문서 열기 암호가 있는 PPTX는 로컬 브라우저에서 변환할 수 있습니다. NASCA+SD 등 회사 DRM으로 보호되어 브라우저에 암호화된 데이터만 전달되는 문서는 현재 변환할 수 없습니다.

NASCA는 사용자 인증과 문서 권한 확인을 거쳐 복호화하는 문서 보안 제품입니다. [삼성SDS 공식 제품 설명](https://www.samsungsds.com/kr/endpoint-security-platform/endpoint-security-platform.html)을 참고하세요. PowerPoint에서 문서가 열린다는 사실만으로 로컬 HTML에도 복호화된 내용이 전달된다고 가정할 수는 없습니다. 공개 자료에서 이 변환기에 바로 넣을 수 있는 NASCA 브라우저용 복호화 SDK는 확인하지 못했습니다. 회사 계약에 따른 별도 연동의 유무는 조직이나 공급사 확인이 필요합니다.

| 브라우저가 실제로 받는 데이터 | 현재 처리 |
| --- | --- |
| 일반 PPTX 또는 조직에서 허용한 일반 PPTX 사본 | 브라우저 안에서 변환 |
| 지원되는 Office 문서 열기 암호가 있는 PPTX | 사용자가 입력한 암호로 브라우저 안에서 복호화 후 변환 |
| 회사 DRM 암호문 또는 지원되지 않는 보호 형식 | 오류 코드와 형식 진단을 표시하고 중지 |

CLI는 브라우저와 별도로 설치된 LibreOffice 엔진을 사용합니다. 원본 파일을 정상적으로 읽고 내보낼 수 있다면 HTML을 만들지만, LibreOffice가 암호화된 내용만 받거나 읽기가 차단되면 변환할 수 없습니다. 실패한 문서를 성공한 HTML로 저장하지 않습니다. 현재 버전에는 NASCA 전용 복호화가 구현되어 있지 않으며, CLI로 실행한다는 것만으로 DRM 권한이 달라지지는 않습니다.

### 이전 PowerPoint 도우미와의 차이

v1.1.0 도우미는 PowerPoint의 슬라이드 PNG 내보내기 결과를 HTML에 넣는 이미지 기반 방식이었습니다. PowerPoint 자체의 외부 통신이나 회사 보안 제품의 제어까지 이 도우미가 제한하지는 못했습니다. 현재 `index.html`과 새 CLI는 해당 도우미를 호출하지 않습니다. 명령행에서는 **`Convert-PptxCli.cmd` 또는 `convert-pptx.py`**를 사용하세요.

저장소 `tools/Convert-PptxWithPowerPoint.*`는 이전 버전 참고 소스입니다. 새 CLI의 구성 요소가 아니며, 이전 도우미의 PNG 내보내기 오류를 고쳤거나 실제 NASCA+SD 문서 변환에 성공했다고 주장하지 않습니다.

## 브라우저 변환 범위와 주의사항

- 원본의 애니메이션과 슬라이드 전환은 정적 화면으로 표현됩니다.
- 차트는 종류와 축·캐시 구성에 따라 지원 범위가 다르며, 기본형 차트도 대체 표시로 나올 수 있습니다. 테스트에서는 원형 차트의 실제 렌더링과 지원되지 않는 막대 차트 구성의 경고·대체 표시를 각각 확인합니다.
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

저장소에서 빌드하려는 개발자에게만 Node.js와 npm이 필요합니다.

```bash
npm ci
npm run build
npx playwright install chromium
npm test
```

`npm run build`는 의존성이 설치된 뒤 인터넷 연결 없이 실행할 수 있습니다. Playwright가 사용할 Chromium을 이미 설치했다면 `CHROME_BIN` 환경 변수에 실행 파일 경로를 지정할 수 있습니다. `npm test`는 먼저 테스트용 PPTX를 생성합니다. `npm run test:fixtures`로 테스트 파일만 다시 만들 수도 있습니다. Linux에서 한글 스크린샷을 검증하려면 로컬 한글 글꼴도 필요합니다.

### 확인한 동작

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

테스트는 생성한 PPTX를 기준으로 합니다. 모든 PowerPoint 기능이나 실제 사용자 문서의 시각적 일치를 보장하지 않습니다. Firefox와 Safari는 별도로 검증하지 않았습니다.

### 주요 파일

| 경로 | 역할 |
| --- | --- |
| `index.html` | 배포용 단일 HTML 파일. 다운로드 후 바로 실행합니다. |
| `tools/convert-pptx.py`, `tools/Convert-PptxCli.cmd` | PowerPoint 없이 LibreOffice headless → PDF → SVG → 독립 HTML을 만드는 CLI |
| `tools/requirements-cli.txt`, `docs/CLI.md` | CLI 의존성 버전과 사용 설명 |
| `tests/cli.py`, `tests/cli-browser.cjs` | 실제 CLI 변환과 결과 HTML의 로컬 브라우저 검증 |
| `src/shell.html`, `src/styles.css`, `src/app.js` | 변환기 화면과 사용자 흐름 |
| `src/worker.js` | 브라우저 내 PPTX 변환 작업 |
| `src/input-package.js` | 문서 형식·보호 방식 판별, 입력 암호를 이용한 Office 복호화 |
| `src/normalize-package.js` | PPTX 내부 절대 관계 경로 정규화와 압축 크기 검사 |
| `src/exporter.js` | 변환 결과의 독립 HTML 구성 |
| `src/viewer.js`, `src/viewer.css` | 저장 결과의 슬라이드 탐색·보기·인쇄 UI |
| `scripts/build.mjs` | 소스와 의존성을 단일 `index.html`로 묶는 빌드 |
| `scripts/create-fixtures.cjs`, `tests/offline.cjs` | 생성 fixture 및 오프라인 검증 |
| `scripts/create-protected-fixtures.cjs`, `tests/protection.cjs`, `tests/reference/` | 암호·파일 형식 회귀 테스트와 독립 암호화 기준 파일 |
| `tools/Convert-PptxWithPowerPoint.ps1`, `.cmd` | v1.1.0의 참고 소스. 현재 HTML 배포·빌드·실행에는 사용하지 않음 |
| `licenses/` | 빌드 시 자동 생성되는 포함 의존성의 라이선스 고지 |

## 오픈 소스 고지

변환 엔진은 [`@briank-dev/pptx-to-html` 2.1.0](https://www.npmjs.com/package/@briank-dev/pptx-to-html)([상위 프로젝트 문서](https://github.com/brnyxx/pptx2html-turbo), MIT)을 사용합니다. HTML 정제에는 [`DOMPurify` 3.4.17](https://github.com/cure53/DOMPurify)(Apache-2.0 또는 MPL-2.0), 패키지 경로 호환 처리에는 [`JSZip` 3.10.1](https://github.com/Stuk/jszip)(MIT 선택)을 사용합니다. 빌드 과정에서 포함 의존성의 라이선스 문서가 `licenses/`에 생성되며, 원문 고지는 배포용 `index.html` 안에도 보존됩니다.

Office 암호 처리에는 [`officecrypto-tool` 0.0.19](https://github.com/zurmokeeper/officecrypto-tool)의 암호화 모듈, `cfb` 1.2.2, `xml2js` 0.6.2 및 Browserify 브라우저 구현을 포함합니다. 실제 번들에 포함된 패키지의 고지는 `licenses/OfficeCrypto-BUNDLED-LICENSES.txt`와 HTML 안에 함께 들어갑니다. 독립 기준 파일의 생성 정보와 해시는 `tests/reference/manifest.json`에 있습니다.

저장소 자체 소스의 라이선스는 [MIT](LICENSE)입니다. CLI가 별도로 사용하는 [LibreOffice](https://www.libreoffice.org/about-us/licenses/)와 [PyMuPDF](https://pymupdf.readthedocs.io/en/latest/about.html)의 라이선스는 각각의 프로젝트를 따릅니다. PyMuPDF와 MuPDF는 AGPL-3.0 또는 Artifex 상용 라이선스로 제공되며, MIT 고지가 이 외부 의존성의 라이선스를 대체하지 않습니다. 이 저장소는 해당 실행 파일이나 Python 패키지를 함께 배포하지 않습니다.

GitHub Pages에 게시하지 않아도 로컬에서 실행할 수 있습니다.

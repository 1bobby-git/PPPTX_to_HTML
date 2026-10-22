# PowerPoint 없이 CLI로 PPTX → HTML

> 이 문서는 v1.3.0에서 추가한 **Python·LibreOffice 방식**입니다. PowerPoint와 LibreOffice를 모두 실행하지 않는 v1.4.0의 **Node.js 방식**은 [Native CLI 설명](NATIVE_CLI.md)을 사용하세요. 기존 `Convert-PptxCli.cmd`의 동작은 유지합니다.

이 CLI는 PowerPoint를 실행하지 않습니다. 설치된 LibreOffice를 화면 없는 `--headless` 모드로 호출하고, 그 결과를 한 개의 HTML 파일로 저장합니다. 변환할 때 브라우저를 열거나 로컬 웹 서버를 실행할 필요가 없습니다.

## Windows 설치와 실행

필요한 프로그램은 **Python 3.10 이상**과 **LibreOffice**입니다. Python은 64비트 3.12 이상을 권장합니다. LibreOffice에는 Impress 구성 요소가 필요합니다.

- [Python 공식 다운로드](https://www.python.org/downloads/windows/)
- [LibreOffice 공식 다운로드](https://www.libreoffice.org/download/download-libreoffice/)
- [프로젝트 ZIP 다운로드](https://github.com/1bobby-git/PPPTX_to_HTML/archive/refs/heads/main.zip)

프로젝트 ZIP을 풀고 해당 폴더에서 PowerShell 또는 명령 프롬프트를 엽니다. 아래 명령의 `C:\test.pptx`를 실제 입력 파일 경로로 바꾸세요.

### 1. 최초 한 번만 의존성 설치

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r .\tools\requirements-cli.txt
```

PyMuPDF 1.28.2를 프로젝트 안의 `.venv`에 설치합니다. `Activate.ps1`은 실행하지 않으므로 PowerShell 실행 정책을 바꾸지 않습니다. 이미 설치된 Python 패키지도 변경하지 않습니다. 이 설치 단계에는 인터넷 또는 사내 패키지 저장소가 필요하며, 변환 단계에는 온라인 변환 서비스를 사용하지 않습니다.

### 2. 설치 상태 확인

```powershell
.\.venv\Scripts\python.exe .\tools\convert-pptx.py --check
```

Python·LibreOffice·PyMuPDF의 버전과 실행 경로를 확인합니다. LibreOffice를 자동으로 찾지 못하면 경로를 지정할 수 있습니다.

```powershell
.\.venv\Scripts\python.exe .\tools\convert-pptx.py --check --soffice "C:\Program Files\LibreOffice\program\soffice.com"
```

### 3. 변환

```powershell
.\tools\Convert-PptxCli.cmd "C:\test.pptx" -o "C:\test.html"
```

CMD 파일은 프로젝트 `.venv`의 Python을 우선 사용합니다. CMD를 거치지 않고 다음처럼 직접 실행해도 같습니다.

```powershell
.\.venv\Scripts\python.exe .\tools\convert-pptx.py "C:\test.pptx" -o "C:\test.html"
```

성공하면 출력 경로와 슬라이드 수를 표시합니다. 그때 `C:\test.html`을 더블클릭하여 브라우저로 열면 됩니다. 원본 PPTX는 수정하지 않습니다. 출력 폴더는 미리 존재해야 하며 해당 위치에 저장 권한이 있어야 합니다.

동일한 HTML이 있으면 기본적으로 중지합니다. 다시 변환하여 교체하려면 다음처럼 실행합니다.

```powershell
.\tools\Convert-PptxCli.cmd "C:\test.pptx" -o "C:\test.html" --overwrite
```

변환이나 검증에 실패하면 `--overwrite`를 지정했더라도 기존 HTML을 보존합니다.

## 옵션

| 옵션 | 동작 |
| --- | --- |
| `-o`, `--output` | 출력 HTML 경로. 생략하면 입력 파일 옆에 같은 이름의 `.html`로 저장 |
| `--soffice` | LibreOffice 실행 파일의 절대 경로. Windows에서는 `soffice.com` 권장 |
| `--slides 1,3-5` | 선택한 슬라이드만 원본 순서로 출력 |
| `--include-hidden` | 숨긴 슬라이드를 포함. 기본값은 제외 |
| `--title "보고서"` | 결과 HTML 제목 |
| `--timeout 120` | LibreOffice 변환 제한 시간(초). 기본 120, 허용 1~3600 |
| `--overwrite` | 성공한 결과로 기존 출력 교체 |
| `--json` | 성공 또는 오류를 JSON 한 개로 출력 |
| `--check` | 입력 파일 없이 실행 환경 검사 |
| `--help` | 전체 사용법 |

```powershell
.\tools\Convert-PptxCli.cmd "C:\자료\보고서.pptx" -o "C:\자료\보고서.html" --slides "1,3-5" --include-hidden --title "검토 자료"
```

일반 PPTX에서 `--slides`는 **원본 프레젠테이션의 1부터 시작하는 슬라이드 번호**입니다. 파일 내부의 `slide1.xml` 같은 파일명 순서와 다를 수 있습니다. 숨긴 슬라이드를 선택하려면 `--include-hidden`을 함께 지정하세요.

CLI가 원본 슬라이드 목록을 직접 읽을 수 없지만 LibreOffice는 정상적으로 변환한 경우, 내보낸 PDF의 페이지 번호를 사용하고 경고를 표시합니다. JSON의 `numbering`은 이때 `exported_page`이며 일반적인 경우에는 `original`입니다. 원본 목록을 읽었는데 실제 출력 페이지 수가 맞지 않으면 성공으로 처리하지 않습니다.

## 결과 HTML

변환 과정은 **PPTX → 임시 PDF → 페이지별 SVG → 독립 HTML**입니다. 임시 PDF와 변환 전용 LibreOffice 프로필은 작업 종료 시 정리합니다.

- 슬라이드 전체를 PNG로 캡처하지 않습니다. 도형·윤곽선 글자는 SVG 벡터이며, 원본 사진 등 비트맵 이미지는 포함된 이미지로 유지됩니다.
- HTML에 필요한 리소스를 모두 내장합니다. 생성된 HTML을 보는 데 Python·LibreOffice·원본 PPTX가 필요하지 않습니다.
- 이전/다음, 슬라이드 선택, 전체/한 장 보기, 키보드 이동, 인쇄를 지원합니다.
- 그림 속 글자는 글꼴 윤곽선으로 표시됩니다. 텍스트 선택·복사는 각 슬라이드의 **슬라이드 텍스트**를 펼쳐 사용하세요. 이 텍스트는 PDF에서 추출하므로 읽기 순서나 특수문자가 원본과 다를 수 있습니다.
- 발표자 노트와 노트 페이지는 내보내지 않습니다. PDF 내부 속성이나 첨부 파일도 결과 HTML로 복사하지 않습니다.
- 애니메이션·영상·슬라이드 전환·클릭 동작은 재생하지 않는 정적 결과입니다.

글꼴은 **변환하는 PC**에 설치되어 있어야 원본 배치와 가깝게 변환됩니다. LibreOffice와 PowerPoint의 표현 차이로 차트·수식·SmartArt·복잡한 레이아웃은 달라질 수 있습니다. 모든 문서의 시각적 일치를 보장하지 않습니다.

## NASCA+SD와 문서 암호

CLI는 원본 경로를 LibreOffice에 전달합니다. 파일 앞부분이 일반 ZIP이 아니라는 이유만으로 먼저 거절하거나 "확장자만 바꾼 파일"이라고 판단하지 않습니다. NASCA가 해당 프로그램의 문서 읽기·내보내기를 허용하고 LibreOffice가 내용을 읽을 수 있는 경우에만 변환할 수 있습니다.

**CLI 실행 자체가 NASCA 암호를 해제하지는 않습니다.** 이 도구에는 NASCA SDK·키 추출·PowerPoint 자동화·별도의 DRM 해제 기능이 없습니다. 실제 NASCA+SD 보호 문서는 제공받지 않았으므로 그 환경에서의 성공을 검증하지 못했습니다. 읽을 수 없는 문서는 `CONVERSION_FAILED`로 중지합니다.

CLI에는 일반 Office 문서의 암호 입력 옵션도 없습니다. 문서 열기 암호가 설정된 일반 PPTX는 저장소의 `index.html` 방식과 지원되는 암호 입력란을 사용하세요. NASCA 인증은 그 암호 입력과 다릅니다.

PowerPoint와 그 Office 클라우드 연동을 호출하지 않습니다. 다만 `--headless`는 화면을 표시하지 않는 실행 옵션이며 **운영체제의 모든 외부 통신을 차단하는 방화벽 기능은 아닙니다.** 변환기의 외부 업로드 기능은 없고, 외부 자료가 필요한 문서와 보안 에이전트의 통신은 PC 정책을 따릅니다. 생성된 HTML은 외부 리소스 URL을 사용하지 않으며 CSP로 자동 네트워크 요청을 제한합니다.

## 오류와 종료 코드

| 오류 코드 | 의미 |
| --- | --- |
| `LIBREOFFICE_NOT_FOUND` | LibreOffice 미설치 또는 실행 경로를 찾지 못함 |
| `PYMUPDF_NOT_FOUND` | 현재 명령이 사용하는 Python 환경에 PyMuPDF가 없음 |
| `INPUT_NOT_FOUND` | 입력 파일 경로를 찾지 못함 |
| `OUTPUT_EXISTS` | 출력 파일이 이미 있고 `--overwrite`를 지정하지 않음 |
| `CONVERSION_FAILED` | 엔진이 읽기 또는 PDF 내보내기에 실패. DRM·문서 암호·손상·형식 호환성 등을 확인해야 함 |
| `TIMEOUT` | LibreOffice가 지정된 시간 안에 변환하지 못함 |
| `SLIDES_UNAVAILABLE` | 없는 슬라이드 또는 제외된 숨김 슬라이드를 선택함 |
| `SLIDE_COUNT_MISMATCH` | 원본 목록과 출력 페이지 수가 달라 결과 신뢰성을 확인할 수 없음 |
| `IO_ERROR` | 파일 접근·저장·권한 문제 |

정상 완료는 종료 코드 `0`, 처리 실패는 `1`, 잘못된 명령 인수는 `2`, 사용자의 취소는 `130`입니다. JSON을 사용하는 자동화에서는 `ok`와 `error.code`를 검사하세요. PDF 내보내기의 프로세스 종료 코드만으로 성공을 판정하지 않고, 결과 파일 형식과 페이지 수까지 확인합니다.

입력 64 MiB, HTML 256 MiB, 최대 1,000장으로 제한합니다. 일반 PPTX 압축 내부 크기와 메타데이터에도 별도 한도를 적용합니다. 시간 제한은 LibreOffice 단계에 적용됩니다.

## Linux / macOS

Python 3.10 이상과 LibreOffice를 설치한 뒤 다음처럼 실행합니다.

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r tools/requirements-cli.txt
.venv/bin/python tools/convert-pptx.py --check
.venv/bin/python tools/convert-pptx.py ./presentation.pptx -o ./presentation.html
```

macOS 기본 설치 경로와 Linux `PATH`에서 LibreOffice를 찾습니다. 별도 위치에 설치했다면 `--soffice`를 지정합니다.

## 개발 검증

브라우저 앱과 CLI 테스트는 구분합니다. CLI 테스트에는 Python 의존성, LibreOffice, Node 개발 의존성, Chromium이 필요합니다.

```bash
npm ci
npm run test:fixtures
npx playwright install chromium
.venv/bin/python -m unittest discover -s tests -p cli.py -v
```

실제 LibreOffice로 생성한 PPTX를 변환하고, 결과를 Chromium `file://`로 다시 열어 검사합니다. 테스트용 보호 표식·잘못된 입력의 실패 검사는 실제 NASCA 복호화 성공 검사가 아닙니다. Windows/NASCA PC에서의 실행은 별도 확인이 필요합니다.

## 엔진 문서와 라이선스

- [LibreOffice 명령행 옵션](https://help.libreoffice.org/latest/en-US/text/shared/guide/start_parameters.html)
- [LibreOffice PDF 내보내기 옵션](https://help.libreoffice.org/latest/en-US/text/shared/guide/pdf_params.html)
- [PyMuPDF SVG 출력](https://pymupdf.readthedocs.io/en/latest/page.html#Page.get_svg_image)
- [PyMuPDF 라이선스](https://pymupdf.readthedocs.io/en/latest/about.html)

이 저장소는 LibreOffice 실행 파일과 PyMuPDF 패키지를 포함하지 않습니다. LibreOffice와 PyMuPDF 각각의 라이선스가 적용되며, PyMuPDF/MuPDF는 AGPL-3.0 또는 Artifex 상용 라이선스 제품입니다. 저장소의 MIT 고지가 외부 엔진 라이선스를 대체하지 않습니다.

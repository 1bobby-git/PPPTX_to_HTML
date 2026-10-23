# PPTX → HTML 참고자료 검토와 적용 판단

검토일: 2026-10-23. 이 문서는 네 참고자료의 실제 소스와 공식 문서를 비교하고, 현재 구현에 도움이 되는 부분과 채택하지 않은 부분을 기록한다. 외부 프로젝트의 기능 소개만으로 변환 품질을 보장하지 않는다.

## 결론과 이번 적용 범위

현재의 `@briank-dev/pptx-to-html` **2.1.0** Rust/WASM 엔진과 오프라인 HTML exporter를 유지한다. 별도 변환 엔진을 나란히 추가하면 ZIP/XML 해석, 테마·레이아웃 처리, 글꼴·이미지·차트 출력과 보안 처리가 중복된다. 참고자료에서는 재현 가능한 결함과 작은 개선 아이디어를 골라 현재 경로에 반영한다.

| 항목 | 상태 | 적용 내용과 의미 |
| --- | --- | --- |
| 브라우저·Node 공통 변환 계약 | 반영 | `src/conversion-core.js`에서 범위 해석, 엔진 호출, 치수·장수·출력 크기 검증, 진단 요약과 WASM 객체 해제를 공통 처리한다. 두 worker가 같은 코어를 사용한다. |
| 숨긴 슬라이드 포함 동작 | 반영 | 고정 엔진에서 빈 범위가 표시 슬라이드만 선택하는 특성을 고려한다. `includeHidden`이 켜져 있고 범위가 비었으면 모든 원본 번호를 명시 전달하여 브라우저와 Node의 차이를 제거한다. |
| 컨테이너 밖으로 밀려난 슬라이드 복구 | 반영 | 일부 차트 대체 HTML이 컨테이너를 조기에 닫는 경우, 공통 exporter에서 유효한 `.slide-shell`을 순서대로 회수한다. 직접 `.slide` 자식이 없는 shell은 정상 슬라이드로 세지 않는다. |
| 최종 슬라이드 장수·원본 번호 확인 | 반영 | 엔진 metadata 장수와 sanitize/export 이후 장수를 대조한다. 슬라이드가 사라졌는데 성공했다고 반환하지 않는다. 원본 번호가 유일하고 유효할 때만 `sourceSlides`를 제공하여 링크 매핑에 사용한다. |
| 미지원 차트의 캐시 데이터 접근성 대체 표시 | 반영 | `src/chart-data.js`가 내부 캐시를 읽고 실제 프레젠테이션 순서·관계·위치가 유일하게 맞는 placeholder에 표를 제공한다. 원시 값임을 명시하고 기존 경고를 유지한다. 정상 SVG 차트와 외부 workbook은 건드리지 않는다. 원래 차트의 모양 재현을 의미하지 않는다. |
| 여러 테마의 색상·글꼴 상속 | 반영, 범위 제한 | `src/theme-resolver.js`가 연결이 명확한 여러 master/theme의 색상·명시 글꼴을 해소하고, 제한된 placeholder 조건에서 전체 문단 기본값을 보존한다. 확인하지 못한 override나 상속은 원본을 유지하며 경고한다. |
| 그룹·색상 상속 진단 | 반영 | `chOff`, `chExt`, `masterClrMapping` 관련 진단을 무조건 숨기지 않고 도형·효과 경고로 요약한다. 실제 기하·상속 검증 없이 완전 지원으로 판단하지 않는다. |
| 추가 변환 dependency·외부 엔진 코드 | 추가 없음 | 네 자료의 원저작 코드를 복제하지 않았다. Aspose SDK나 두 번째 JS 렌더러도 설치·추가하지 않았다. |

공통 코어와 exporter 변경에 대한 회귀 항목은 `tests/conversion-core.cjs`에 있다. 범위·숨김 처리, 장수 불일치, 객체 해제, UTF-8 출력 제한, orphan shell 복구, 원본 번호와 안전한 링크를 검사한다. `tests/chart-data.cjs`는 실제 WASM·native worker·브라우저를 통해 캐시의 빈칸, 숫자, namespace, 악성처럼 보이는 텍스트, 한도, 순서·숨김·범위·모호한 위치를 검사한다.

차트 테스트에서 엔진의 raw diagnostic `slide_index`가 실제 프레젠테이션 순서와 다를 수 있음을 확인했다. 차트는 `presentation.xml` 관계로 다시 매핑하며, 일반 경고에서는 이 불확실한 번호를 사용자 슬라이드 번호로 표시하지 않는다. 원본 번호가 확인되는 최종 출력의 `sourceSlides`는 유지한다.

## 네 참고자료의 비교

| 참고자료 | 확인한 버전 또는 불변 revision | 라이선스 | 현재 구현과 겹치는 기능 | 판단 |
| --- | --- | --- | --- | --- |
| [g21589/PPTX2HTML](https://github.com/g21589/PPTX2HTML) | `05824403f375d028280c8cd9bb2b58ac4f8d7571` | MIT, 포함된 제삼자 자산은 별도 확인 필요 | 브라우저 worker, ZIP/XML, 텍스트·도형·이미지·표·차트, HTML 다운로드 | 구형 엔진 전체는 제외. 미지원 차트의 캐시 데이터 추출이라는 작은 관심사만 독립적으로 참고 |
| [javier-mora/pptx-to-html](https://github.com/javier-mora/pptx-to-html) | `a67cda43c5e7e43659e8c906dc1d1d1a2d922627`; `@jvmr/pptx-to-html` 1.1.2 | MIT | OOXML 순서·상속·도형·텍스트·이미지 추출, HTML fragment 생성 | 두 번째 렌더러는 제외. 비교 과정에서 확인한 현재 wrapper의 브라우저/Node 차이와 장수 계약을 개선 |
| [Aspose Python 예제 Gist](https://gist.github.com/aspose-com-gists/223609d1fb2c102aac36f1c4707834bb) | `49500592c8930e6086e5733bd20e77996b51eaff` | Gist에 LICENSE 파일 없음; 사용하는 SDK는 상용·독점 라이선스 | PPT/PPTX 로드, HTML 저장, notes와 responsive 옵션 | SDK 사용 예제이며 독립 무료 변환기가 아님. 설치·라이선스·품질 실험 없이 자동 fallback에 추가하지 않음 |
| [GitHub `pptx-export` topic](https://github.com/topics/pptx-export) | 검토일에 표시된 다섯 저장소를 아래 SHA로 확인 | 프로젝트별 상이 | 대부분 HTML/내용 → PPTX 생성. DeckMason은 PPTX → JSON 내용 추출도 있음 | 시각 충실도 변환 엔진으로 채택할 대상 없음. 방향·라이선스와 중복을 구분하여 제외 |

이 결과는 공개 소스의 정적 분석에 기초한다. 구형 프로젝트는 핵심 변환·파서·라이선스 파일을, 현대 프로젝트는 `src/`, `test/`, package와 README·라이선스를 확인했다. 전체 소스가 제공되는 자료는 일반적인 소스 분석으로 읽었으며 **REA 엔진을 실행했다고 주장하지 않는다**. 참고 프로젝트를 설치하거나 원본 PowerPoint로 비교 렌더링한 결과도 아니다.

## 1. 구형 PPTX2HTML: 데이터 접근성 아이디어만 참고

고정 소스: [js/worker.js](https://github.com/g21589/PPTX2HTML/blob/05824403f375d028280c8cd9bb2b58ac4f8d7571/js/worker.js), [js/pptx2html.js](https://github.com/g21589/PPTX2HTML/blob/05824403f375d028280c8cd9bb2b58ac4f8d7571/js/pptx2html.js), [LICENSE](https://github.com/g21589/PPTX2HTML/blob/05824403f375d028280c8cd9bb2b58ac4f8d7571/LICENSE).

이 구현은 브라우저에서 FileReader와 Web Worker로 PPTX를 읽고, HTML/CSS와 차트 작업을 전달한다. 차트는 D3/NVD3로 렌더링한다. 현재 엔진과 기능 범위가 크게 겹치지만 더 충실한 fallback이라고 보기 어려운 구체적 이유가 있다.

- 슬라이드 순서를 `presentation.xml`의 `sldIdLst`가 아닌 Content Types 나열에서 찾는다. 원본에서 슬라이드 순서를 바꾼 문서는 파일 번호 순서와 다를 수 있다.
- 그룹의 `ext/chExt`를 좌표계 배율로 합성하지 않고 차이로 다루는 부분이 있다. 중첩·비균등 배율 그룹에 재사용할 안전한 알고리즘이 아니다.
- 테마는 전역 하나에 의존하고 `clrMap` 처리는 TODO다. 다중 master와 동아시아 글꼴 상속을 보강할 근거가 되지 않는다.
- 표의 행·열 병합, EMF/WMF의 실제 브라우저 변환, SmartArt와 이미지 자르기에는 명확한 한계가 있다.
- 원문 텍스트·링크를 문자열 HTML로 합치는 경로를 그대로 가져오지 않는다. 현재 exporter의 sanitize·URL 제한·CSP를 유지한다.

유용한 부분은 `extractChartData`가 OOXML의 계열·범주·값 캐시를 읽는다는 점이다. 이번에 구현한 대체 표시는 이 **데이터 추출과 렌더링의 분리**를 참고하며 NVD3나 원저작 코드를 가져오지 않는다. 캐시가 없거나 오래되었으면 실제 최신 차트 데이터를 보장할 수 없고, 외부 workbook을 다운로드해서 보완해서도 안 된다. 경고를 유지하면서 복구 가능한 데이터만 명확한 대체 표시로 제공하는 범위다.

## 2. 현대 TypeScript 구현: 대체 엔진보다 공통 계약 개선

고정 소스: [SlideExtractor.ts](https://github.com/javier-mora/pptx-to-html/blob/a67cda43c5e7e43659e8c906dc1d1d1a2d922627/src/core/SlideExtractor.ts), [저장소 트리](https://github.com/javier-mora/pptx-to-html/tree/a67cda43c5e7e43659e8c906dc1d1d1a2d922627), [LICENSE](https://github.com/javier-mora/pptx-to-html/blob/a67cda43c5e7e43659e8c906dc1d1d1a2d922627/LICENSE).

`getSlidePathsInPresentationOrder`가 presentation 관계를 따라가는 방향은 적절하지만, 현재 구현도 이미 원본 순서와 관계 경로 정규화를 다룬다. 이 기능을 위해 다른 파서를 추가할 필요는 없다. 반대로 이 참고 구현은 숨긴 슬라이드의 `show` 속성을 검사하지 않고, `theme1.xml`과 일부 기본 스타일에 의존한다. 혼합 run 서식·master `txStyles`·테마별 색상 매핑·그룹 회전과 flip을 완전하게 해결하는 근거도 없다. 폰트 임베딩이나 전체 오프라인 viewer를 제공하는 대체재도 아니다.

비교에서 확인한 실제 개선 대상은 현재 wrapper의 차이였다. 브라우저와 Node가 같은 엔진을 쓰더라도 빈 범위·숨김 옵션, 출력 shell 복구, 장수·원본 번호 검증이 다르면 결과가 달라질 수 있었다. 이를 `conversion-core.js`와 공통 exporter로 정리했다. 전체 엔진 교체보다 변화 범위가 작고, 기존 sanitizer와 viewer를 계속 사용할 수 있다.

### 실제 렌더링을 비교해 수정한 여러 테마 문제

`scripts/create-fidelity-fixtures.cjs`와 `tests/fidelity.cjs`로 실제 WASM 출력을 검사했다. 행·열 동시 2×2 병합 표는 셀 구조·행 높이를 보존했고, 비영점 child offset과 서로 다른 X/Y 배율을 가진 중첩 그룹은 기대 좌표·크기와 일치했다. 회전·flip을 포함한 모든 그룹 조합을 확인한 것은 아니다.

서로 다른 두 master/theme를 사용하는 사례에서는 둘째 슬라이드가 첫 테마의 색과 동아시아 글꼴을 사용했다. 기대값은 둘째 테마의 `#BB3311`·`Microsoft YaHei`였지만 원래 엔진은 `#1122AA`·`Malgun Gothic`를 반환했다. `src/theme-resolver.js`는 관계가 명확한 여러 테마에 한해 다음을 보완한다.

- theme의 명시적인 RGB 색과 `+mj/+mn` 글꼴 참조를 해소하고 tint·alpha 등 기존 자식 변환을 유지한다.
- `body`·`title`·`ctrTitle` placeholder의 slide list style이 비어 있고 layout/master의 지역 텍스트 서식이 없는 경우, 해당 master의 문단 기본값 전체를 복사한다. 색·글꼴만 부분 삽입할 때 생겼던 글자 크기 변경을 방지한다.
- XML 전체를 다시 직렬화하지 않고 검증한 원문 구간만 수정한다. 도형·이미지·그룹의 혼합 순서와 알 수 없는 namespace 내용을 보존한다.
- 단일 테마에는 원래 bytes를 반환한다. 상속·공유 관계가 모호하거나 override가 있거나 검사 한도를 넘으면 해당 원본을 유지하고 `THEME_INHERITANCE_UNRESOLVED`로 안내한다. 차트·다이어그램의 모든 테마 상속과 언어별 빈 글꼴 기본값은 해결했다고 주장하지 않는다.

수정한 사례는 두 슬라이드 모두 원래 24pt 크기를 유지하고 각자의 색·동아시아 글꼴 이름을 사용한다. 이것은 출력 CSS와 OOXML 상속의 검증이며, 해당 글꼴이 모든 기기에 설치되어 있거나 동일한 글리프가 렌더링된다는 보장은 아니다. `--raw-theme-probe` 옵션은 보정 전 고정 엔진의 결함을 의도적으로 재현하는 별도 진단용이다.

## 3. Aspose: 명시적으로 구성할 수 있는 향후 후보

Gist의 기본 HTML, notes, responsive 예제는 `aspose.slides` SDK 호출이다. 예제의 `notes_comments_layouting`과 달리 현재 공식 안내는 `NotesCommentsLayoutingOptions`를 `HtmlOptions.slides_layout_options`에 지정한다. [현재 HTML 변환 안내](https://docs.aspose.com/slides/python-net/convert-powerpoint-to-html/)와 [HtmlOptions API](https://reference.aspose.com/slides/python-net/aspose.slides.export/htmloptions/)를 기준으로 검토해야 한다.

| 검토 항목 | 공식 근거와 통합 시 의미 |
| --- | --- |
| 로컬 실행·PowerPoint 필요 여부 | [시스템 요구사항](https://docs.aspose.com/slides/python-net/system-requirements/)은 PowerPoint 없이 작동하는 독립 엔진이라고 설명한다. 온라인 Aspose 변환기나 Cloud API를 쓸 필요는 없다. |
| HTML·SVG와 HTML5 차이 | [HTML5 안내](https://docs.aspose.com/slides/python-net/export-to-html5/)에 따르면 표준 HTML은 SVG 기반이며, HTML5는 도형 애니메이션·전환용 CSS/JS와 public CDN의 jQuery·Anime.js를 사용한다. HTML5를 그대로 오프라인 단일 HTML로 취급하면 안 된다. |
| 글꼴·이미지 | [HTML 안내](https://docs.aspose.com/slides/python-net/convert-powerpoint-to-html/)는 대부분의 리소스를 포함하는 기본 출력과 font embedding 옵션을 설명한다. 렌더링 장치의 원본 폰트 확보와 웹폰트 배포 권한은 별도로 필요하다. |
| Python·CPU·운영체제 | [설치 안내](https://docs.aspose.com/slides/python-net/installation/)와 [PyPI 26.10.0 파일 목록](https://pypi.org/project/aspose-slides/26.10.0/#files)은 Python 3.5–3.14, Windows x86/x64, Linux x64/ARM64, macOS Intel/Apple silicon을 제시한다. Linux/macOS에는 native 라이브러리와 shared libpython 등의 조건이 있다. |
| 운영체제 검증 한계 | 시스템 요구사항 표는 macOS 12를 나열하지만 wheel tags에는 Intel 10.14+, ARM64 11+가 있다. 설치 문서에는 macOS 설치 미검증 문구도 있다. 배포 파일 존재와 실제 지원 환경 검증을 구분해야 한다. |
| 평가판 | [라이선스 안내](https://docs.aspose.com/slides/python-net/licensing/)는 모든 저장 슬라이드의 평가 watermark와 읽은 텍스트의 잘림을 명시한다. 무라이선스 결과를 정상 fallback 성공으로 반환하면 안 된다. |
| 오프라인 라이선스 | 같은 안내는 파일 라이선스를 로컬에서 검증한다고 설명한다. 반면 [metered licensing](https://docs.aspose.com/slides/python-net/metered-licensing/)은 지속적인 인터넷 연결을 필요로 한다. |
| 배포 권한 | [SDK EULA](https://about.aspose.com/legal/eula/)와 구매한 계약에 따른다. 프로젝트의 MIT 라이선스가 SDK나 폰트의 재배포 권한까지 부여하지 않는다. |

이번에는 Aspose SDK를 설치하거나 상용 라이선스를 확보하지 않았으며, 실제 문서 변환이나 기존 엔진과의 충실도 비교도 하지 않았다. 따라서 기본 또는 자동 fallback으로 추가하지 않았다. 평가 watermark·텍스트 잘림, native dependency, 플랫폼별 확인 부족, HTML5의 외부 리소스 의존을 숨긴 채 엔진을 전환하는 것은 성공 조건을 바꾸는 일이기 때문이다.

향후 검증하려면 관리자가 로컬 Python·고정 SDK 버전·유효한 파일 라이선스·폰트를 명시적으로 구성하고, 격리된 worker에서 실행해야 한다. 사용자 흐름은 기존 웹의 파일 선택·변환·다운로드를 유지할 수 있다. 출력 리소스와 오프라인 동작을 실제 확인하고 engine/version을 기록한 뒤 선택 기능으로 제공하는 순서가 적절하다. 외부 서비스 문서 업로드나 PowerPoint 실행은 이 후보의 전제가 아니다.

## 4. `pptx-export` topic의 다섯 저장소

Topic은 변할 수 있으므로 아래는 검토일의 고정 SHA 기준이다. 이름에 `pptx-export`가 있다는 이유만으로 PPTX 입력을 HTML로 충실하게 변환한다고 판단하지 않는다.

| 저장소·확인 소스 | Commit SHA | 실제 방향·라이선스 | 채택하지 않은 이유 |
| --- | --- | --- | --- |
| [JunZhaoNathan/Chiselo](https://github.com/JunZhaoNathan/Chiselo/tree/877b7773b2fa2d2b04a6f80f1dd42e859657ade7), `Chiselo/HTMLRenderExporter.swift`, `LICENSE` | `877b7773b2fa2d2b04a6f80f1dd42e859657ade7` | macOS WKWebView로 HTML을 편집·렌더링하고 PPTX를 작성. 현재 CC-BY-NC 4.0 | 반대 방향이며 현재 소스는 비상업적 제한이 있다. LICENSE가 명시한 `v0.1.23` 이하의 과거 Apache-2.0 이력을 현재 main에 적용하지 않는다. |
| [Danamagentapink178/Chiselo](https://github.com/Danamagentapink178/Chiselo/tree/891fe67d35617f0196a8fe1f3ebd75d1c60c1ce2), 같은 exporter와 `LICENSE` | `891fe67d35617f0196a8fe1f3ebd75d1c60c1ce2` | HTML → PPTX. Personal Non-Commercial License 1.0 | 반대 방향. 상업·회사 내부 업무 사용 등의 제한을 포함하므로 MIT 코드로 가져오지 않는다. |
| [AI-PPT-Slides-Builder](https://github.com/sirimallachandana4/AI-PPT-Slides-Builder/tree/66b7f5c88f10bc6f86709aad3d2ddf275797681f), `src/lib/pptxExport.ts`, `server.ts` | `66b7f5c88f10bc6f86709aad3d2ddf275797681f` | Gemini로 내용을 생성하고 `pptxgenjs`로 PPTX 작성. LICENSE 없음 | PPTX 입력 렌더러가 아니며 코드 재사용 허가도 확인되지 않는다. 서버의 외부 생성 API는 현재 로컬 변환 목적과 다르다. |
| [a692570/deckforge](https://github.com/a692570/deckforge/tree/40b59944cc5bb5a26b577a2070a5409a7e0d5f0e), `scripts/pptx_export.py`, README·LICENSE | `40b59944cc5bb5a26b577a2070a5409a7e0d5f0e` | HTML deck 제작·PPTX export. MIT | archived이며 DeckMason으로 이름을 바꾼 전신이다. 별개 변환 엔진으로 중복 추가할 근거가 없다. |
| [a692570/deckmason](https://github.com/a692570/deckmason/tree/27b0d9da7673f174acfa4a80dc2318e5eda935b9), `scripts/pptx_import.py`, `scripts/pptx_export.py`, LICENSE | `27b0d9da7673f174acfa4a80dc2318e5eda935b9` | PPTX → JSON 내용 추출 후 agent가 HTML을 재구성하는 기능도 있음. MIT | 내용 재디자인 도구이며 원래 화면을 보존하는 renderer가 아니다. 현재 파서·dependency와 중복된다. |

DeckMason의 import는 실제 존재한다. 다섯 저장소를 모두 단순히 HTML → PPTX라고 분류하는 것은 부정확하다. 다만 소스에서 본문은 텍스트·문단 수준만 기록하고 폰트 run·본문 좌표·원래 도형을 보존하지 않는다. 이미지는 파일명과 위치를 JSON에 적지만 실제 image blob을 파일로 추출하지 않는다. 첫 테마의 Latin major/minor 폰트 처리와 최상위 shape 순회만으로는 차트·표·그룹·테마 상속을 포함한 원본 재현을 해결할 수 없다. 별도 재디자인 모드가 필요한 경우에는 참고할 수 있지만 현재 충실도 fallback으로는 제외한다.

## 호환성과 ‘완벽한 변환’의 검증 한계

로컬에서 실제 검증한 환경은 **Linux, Node 24, Chromium·Firefox**다. `.github/workflows/compatibility.yml`에는 Windows·macOS·Linux의 Node 24·Chromium, 추가 Linux Firefox·WebKit 검사를 구성했다. 각 커밋의 통과 상태는 GitHub Actions 결과를 기준으로 판단한다. WebKit은 로컬 환경의 native 라이브러리 부재로 실행하지 못했으며, Playwright WebKit 검사도 실제 Safari·iOS 기기 검증을 대신하지 않는다. responsive viewer가 있다는 사실은 모든 화면·브라우저·입력 장치에서 실제 사용성을 확인했다는 뜻이 아니다.

첫 [GitHub Actions 실행](https://github.com/1bobby-git/PPPTX_to_HTML/actions/runs/39879591262)에서는 Linux·macOS의 전체 경로와 Firefox가 통과했고, Windows·WebKit에서 추가 문제가 드러났다. Windows식 CRLF로 읽은 viewer의 해시는 HTML 파서가 LF로 정규화한 실제 스크립트와 달라 CSP가 실행을 차단한다. 실제 Chromium으로 이를 재현하여 빌드·Native exporter·웹 응답의 세 경로 모두 동일한 LF 소스로 해시를 계산하도록 수정했다. viewer 검사는 슬라이드 마크업의 존재뿐 아니라 선택 목록과 현재 슬라이드 초기화를 기다린 뒤 기존 조작 검증을 수행한다.

WebKit의 Blob Worker는 생성자 반환 후 스크립트를 비동기로 읽을 수 있다. 생성 직후 URL을 폐기하던 코드를 첫 작업 메시지 또는 종료 시점에 정리하도록 바꿨다. [Playwright 공식 저장소의 재현과 maintainer 설명](https://github.com/microsoft/playwright/issues/33794)도 같은 자원 수명 문제를 확인한다. CSP나 네트워크 차단 조건을 완화하지 않았으며, 수정 후 플랫폼별 성공 여부는 해당 커밋의 Actions 결과로 확인한다.

Firefox에서 생성 fixture의 PNG CRC 오류를 확인해 정상 이미지로 교체했다. 독립 암호 검증용 `tests/reference` 파일은 동일 bytes를 검증하는 고정 벡터로 유지했다.

후속 검사에서는 Playwright 1.62.1 / Firefox r1538의 iframe 조회가 간헐적으로 멈췄다. 실제 iframe의 `src`는 Blob으로 바뀌었지만 검사 도구의 frame URL은 `about:blank`에 머무르고 DOM 조회가 끝나지 않았다. 공식 [Playwright #43124](https://github.com/microsoft/playwright/issues/43124)는 동일 버전의 실행 컨텍스트 회귀를 기록하며 [#43494](https://github.com/microsoft/playwright/pull/43494)의 r1564에서 수정했다. 앱의 iframe 교체나 검사 생략으로 우회하지 않았다.

검토 시점의 안정판 1.64.1에는 이 수정이 포함되지 않아 **Firefox CI에만 공식 `1.65.0-alpha-2026-10-22`와 그에 대응하는 Firefox r1564를 정확히 고정**했다. 운영 의존성과 일반 Playwright 설치는 변경하지 않고 별도 임시 경로에서 검사 도구를 설치한다. 로컬에서도 이 조합으로 동일한 전체 오프라인 검사를 통과했다. Windows·macOS·Linux의 Chromium과 WebKit은 기존 고정 도구를 사용한다. 이 제한적인 사전 배포판 사용은 Firefox 검사 도구의 회귀를 해결하기 위한 것으로, 안정판 1.65 이상 채택 시 제거할 수 있다.

현재 입력 처리·패키지 정규화·WASM 변환·공통 exporter·오프라인 viewer는 PowerPoint를 실행하지 않는 경로다. 저장소의 별도 legacy Python/LibreOffice 및 PowerPoint 도구의 존재를 이번 브라우저·Node 경로의 실행 의존성으로 해석하지 않는다. 참고 검토 때문에 사용자의 웹 흐름에 명령 입력을 추가하지 않는다.

원본 폰트가 없으면 줄바꿈과 위치가 달라질 수 있다. 복잡한 그룹 변환, 테마·master 상속, SmartArt·OLE·수식, 미지원 이미지·브라우저 코덱, 애니메이션·전환과 일부 링크 동작도 개별 확인이 필요하다. 그림으로 대체하거나 캐시 데이터 표를 제공하는 것은 유용한 손실 완화이지만 원래 편집 가능한 개체나 원래 애니메이션을 재현한 것은 아니다. 미지원 경고를 유지하고 미리보기와 원본을 비교해야 한다.

제공된 실제 원본문서와 회사 DRM 표본은 없었다. 테스트용 보호 문서의 성공을 모든 기업 DRM·IRM 지원으로 확장하지 않는다. 합법적으로 복호화된 PPTX와 지원하는 암호 형식의 처리를, 권한이 없거나 접근 제어된 문서의 처리와 구분한다. 이 검토에서는 DRM 우회, 원본 문서의 외부 업로드, PowerPoint 실행을 하지 않았다.

장수 검증은 **조용한 누락을 방지하는 계약**이며 픽셀 단위 시각 일치를 증명하지 않는다. 모든 환경의 완벽성을 약속하는 대신, 확인한 환경과 fixture, 복구한 내용, 남은 경고를 함께 기록한다. 추후 개선도 실제로 실패하는 fixture를 먼저 확보한 뒤 현재 경로의 작은 수정 또는 검증된 선택 백엔드로 해결한다.

#!/usr/bin/env python3
"""Convert a presentation locally with LibreOffice and PyMuPDF.

This tool never starts PowerPoint or a browser and does not upload documents.
LibreOffice must be installed separately. PyMuPDF is licensed under AGPL or a
commercial license; review its license before distributing a bundled copy.
"""

from __future__ import annotations

import argparse
import base64
import html
import json
import os
from pathlib import Path
import re
import shutil
import signal
import subprocess
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET

MAX_INPUT = 64 * 1024 * 1024
MAX_HTML = 256 * 1024 * 1024
MAX_PAGES = 1000
MAX_MESSAGE = 600
VERSION = "1.3.0"


class ConversionError(Exception):
    def __init__(self, code: str, message: str):
        super().__init__(message)
        self.code = code


def fail(code: str, message: str) -> None:
    raise ConversionError(code, message)


def check_python() -> None:
    if sys.version_info < (3, 10):
        fail("PYTHON_VERSION", "Python 3.10 이상이 필요합니다.")


def find_soffice(explicit: str | None) -> Path:
    if explicit:
        candidate = Path(explicit).expanduser()
        if candidate.is_file():
            return candidate.resolve()
        fail("LIBREOFFICE_NOT_FOUND", f"지정한 LibreOffice 실행 파일을 찾을 수 없습니다: {candidate}")
    names = ("soffice.com", "soffice.exe", "soffice", "libreoffice") if os.name == "nt" else ("soffice", "libreoffice")
    for name in names:
        found = shutil.which(name)
        if found:
            return Path(found).resolve()
    extras = [
        Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
        Path("/Applications/LibreOfficeDev.app/Contents/MacOS/soffice"),
    ]
    if os.name == "nt":
        for root_name in ("PROGRAMFILES", "PROGRAMFILES(X86)", "ProgramW6432"):
            root = os.environ.get(root_name)
            if root:
                extras.extend(Path(root, product, "program", executable)
                              for product in ("LibreOffice", "LibreOfficeDev")
                              for executable in ("soffice.com", "soffice.exe"))
    for candidate in extras:
        if candidate.is_file():
            return candidate.resolve()
    fail("LIBREOFFICE_NOT_FOUND", "LibreOffice를 찾을 수 없습니다. 설치하거나 --soffice 실행파일 경로를 지정하세요.")


def get_fitz():
    try:
        import pymupdf
    except ImportError:
        fail("PYMUPDF_NOT_FOUND", "PyMuPDF가 필요합니다: py -3 -m pip install -r tools/requirements-cli.txt")
    return pymupdf


def safe_version(command: list[str]) -> str:
    try:
        result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                text=True, errors="replace", timeout=10, check=False)
        lines = (result.stdout or result.stderr).strip().splitlines()
        if result.returncode == 0 and lines:
            return lines[0][:160]
    except (OSError, subprocess.TimeoutExpired):
        pass
    fail("LIBREOFFICE_CHECK_FAILED", "LibreOffice 버전을 확인할 수 없습니다. --soffice 경로와 설치 상태를 확인하세요.")


def slide_metadata(source: Path) -> list[bool] | None:
    """Return each original slide's visibility when plain OOXML is readable."""
    try:
        with zipfile.ZipFile(source) as archive:
            members = archive.infolist()
            if len(members) > 20_000 or sum(item.file_size for item in members) > 256 * 1024 * 1024:
                fail("PACKAGE_LIMIT", "PPTX 압축 파일의 항목 수 또는 압축 해제 크기가 허용 범위를 초과했습니다.")
            for member in ("ppt/presentation.xml", "ppt/_rels/presentation.xml.rels"):
                if archive.getinfo(member).file_size > 4 * 1024 * 1024:
                    fail("PACKAGE_LIMIT", "PPTX 메타데이터가 4 MiB 제한을 초과했습니다.")
            root = ET.fromstring(archive.read("ppt/presentation.xml"))
            relroot = ET.fromstring(archive.read("ppt/_rels/presentation.xml.rels"))
            rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            p_ns = "http://schemas.openxmlformats.org/presentationml/2006/main"
            package_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
            targets = {rel.get("Id"): rel.get("Target") for rel in relroot.findall(f"{{{package_ns}}}Relationship")
                       if rel.get("Type", "").endswith("/slide")}
            slides = root.find(f"{{{p_ns}}}sldIdLst")
            if slides is None:
                return None
            if len(slides) > MAX_PAGES:
                fail("PACKAGE_LIMIT", "PPTX 슬라이드 수가 1000장 제한을 초과했습니다.")
            result: list[bool] = []
            for item in slides.findall(f"{{{p_ns}}}sldId"):
                target = targets.get(item.get(f"{{{rel_ns}}}id"))
                if not target:
                    return None
                member = target.lstrip("/") if target.startswith("/") else "ppt/" + target
                # Presentation slide relationship targets normally live in ppt/slides/.
                parts: list[str] = []
                for part in member.split("/"):
                    if part == "..":
                        if not parts:
                            return None
                        parts.pop()
                    elif part and part != ".":
                        parts.append(part)
                member = "/".join(parts)
                if archive.getinfo(member).file_size > 4 * 1024 * 1024:
                    fail("PACKAGE_LIMIT", "PPTX 슬라이드 메타데이터가 4 MiB 제한을 초과했습니다.")
                slide = ET.fromstring(archive.read(member))
                result.append(slide.get("show", "1").lower() not in ("0", "false"))
            return result or None
    except (OSError, zipfile.BadZipFile, KeyError, ET.ParseError, ValueError):
        return None


def select_numbers(spec: str | None, available: list[int]) -> list[int]:
    if spec is None:
        return available
    selected: set[int] = set()
    for chunk in spec.split(","):
        match = re.fullmatch(r"\s*(\d+)(?:\s*-\s*(\d+))?\s*", chunk)
        if not match:
            fail("INVALID_SLIDES", "--slides 형식은 1,3-5처럼 지정하세요.")
        start, end = int(match[1]), int(match[2] or match[1])
        if start < 1 or end < start or end > MAX_PAGES:
            fail("INVALID_SLIDES", "--slides 범위가 올바르지 않습니다.")
        selected.update(range(start, end + 1))
    missing = selected.difference(available)
    if missing:
        fail("SLIDES_UNAVAILABLE", f"선택한 슬라이드를 출력할 수 없습니다: {', '.join(map(str, sorted(missing)[:20]))}. 숨김 슬라이드에는 --include-hidden을 사용하세요.")
    return [n for n in available if n in selected]


def isolated_profile(path: Path) -> Path:
    profile = path / "profile"
    (profile / "user").mkdir(parents=True)
    # Set the LibreOffice Common/Security/Scripting preference in this private
    # profile before the process starts. No existing user profile is touched.
    (profile / "user" / "registrymodifications.xcu").write_text(
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<oor:items xmlns:oor="http://openoffice.org/2001/registry">'
        '<item oor:path="/org.openoffice.Office.Common/Security/Scripting">'
        '<prop oor:name="DisableMacrosExecution" oor:op="fuse">'
        '<value>true</value></prop></item></oor:items>', encoding="utf-8")
    return profile


def run_office(command: list[str], timeout: int) -> subprocess.CompletedProcess[str]:
    kwargs: dict = {"stdout": subprocess.PIPE, "stderr": subprocess.PIPE,
                    "text": True, "errors": "replace", "stdin": subprocess.DEVNULL}
    if os.name == "nt":
        kwargs["creationflags"] = subprocess.CREATE_NEW_PROCESS_GROUP
    else:
        kwargs["start_new_session"] = True
    try:
        process = subprocess.Popen(command, **kwargs)
    except OSError as exc:
        fail("LIBREOFFICE_START", f"LibreOffice 실행에 실패했습니다: {exc}")
    try:
        stdout, stderr = process.communicate(timeout=timeout)
    except subprocess.TimeoutExpired:
        stop_office(process)
        fail("TIMEOUT", f"LibreOffice 변환이 {timeout}초 제한을 초과했습니다.")
    except KeyboardInterrupt:
        stop_office(process)
        fail("CANCELLED", "사용자가 변환을 중단했습니다.")
    return subprocess.CompletedProcess(command, process.returncode, stdout, stderr)


def stop_office(process: subprocess.Popen) -> None:
    """Stop only the subprocess tree started for this conversion."""
    if os.name == "nt":
        try:
            subprocess.run(["taskkill", "/PID", str(process.pid), "/T", "/F"],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                           timeout=5, check=False)
        except (OSError, subprocess.TimeoutExpired):
            pass
    else:
        try:
            os.killpg(process.pid, signal.SIGKILL)
        except ProcessLookupError:
            pass
    try:
        process.communicate(timeout=5)
    except subprocess.TimeoutExpired:
        try:
            process.kill()
        except ProcessLookupError:
            pass
        try:
            process.communicate(timeout=2)
        except subprocess.TimeoutExpired:
            pass


CSS = """*{box-sizing:border-box}body{margin:0;background:#161a22;color:#f0f2f5;font:15px system-ui,'Malgun Gothic',sans-serif}header{position:sticky;top:0;z-index:2;background:#232936;border-bottom:1px solid #3c4657;padding:.7rem 1rem;display:flex;align-items:center;gap:.7rem;flex-wrap:wrap}h1{font-size:1rem;margin:0 auto 0 0;max-width:55vw;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}button,select{background:#3a465a;color:white;border:1px solid #69778e;border-radius:.4rem;padding:.4rem .65rem;font:inherit;cursor:pointer}button:disabled{opacity:.4;cursor:default}main{max-width:1500px;margin:auto;padding:1rem}figure{margin:0 auto 1.5rem;background:white;box-shadow:0 6px 22px #0007;max-width:100%;scroll-margin-top:5rem}figure img{display:block;width:100%;height:auto}figcaption{background:#232936;color:#f0f2f5;padding:.35rem .6rem;font-size:.85rem}details{padding:.45rem .7rem;color:#20242c}details summary{cursor:pointer}details pre{white-space:pre-wrap;overflow-wrap:anywhere;font:inherit}.single figure{display:none}.single figure.active{display:block}@media print{body{background:white}header,figcaption,details{display:none}main{padding:0;max-width:none}figure,.single figure{display:block!important;box-shadow:none;break-after:page;margin:0}figure img{width:100%;max-height:100vh;object-fit:contain}}"""
JS = """(()=>{'use strict';const root=document.querySelector('main'),slides=[...document.querySelectorAll('figure')],select=document.querySelector('#slide-select'),prev=document.querySelector('#prev'),next=document.querySelector('#next'),mode=document.querySelector('#mode'),printButton=document.querySelector('#print');let current=0;function show(index){current=Math.max(0,Math.min(slides.length-1,index));slides.forEach((s,i)=>s.classList.toggle('active',i===current));select.value=String(current);prev.disabled=current===0;next.disabled=current===slides.length-1;if(!root.classList.contains('single'))slides[current].scrollIntoView({block:'start'})}prev.addEventListener('click',()=>show(current-1));next.addEventListener('click',()=>show(current+1));select.addEventListener('change',()=>show(Number(select.value)));mode.addEventListener('click',()=>{root.classList.toggle('single');mode.setAttribute('aria-pressed',String(root.classList.contains('single')));mode.textContent=root.classList.contains('single')?'전체 보기':'한 장씩 보기';show(current)});printButton.addEventListener('click',()=>window.print());document.addEventListener('keydown',e=>{if(['SELECT','INPUT','BUTTON'].includes(document.activeElement.tagName))return;let index=null;if(e.key==='ArrowRight'||e.key==='ArrowDown')index=current+1;else if(e.key==='ArrowLeft'||e.key==='ArrowUp')index=current-1;else if(e.key==='Home')index=0;else if(e.key==='End')index=slides.length-1;if(index!==null){e.preventDefault();show(index)}});show(0)})();"""


def render_html(pdf: Path, destination: Path, title: str, numbers: list[int], fitz) -> tuple[int, int]:
    import hashlib
    js_hash = base64.b64encode(hashlib.sha256(JS.encode()).digest()).decode()
    css_hash = base64.b64encode(hashlib.sha256(CSS.encode()).digest()).decode()
    csp = ("default-src 'none'; img-src data:; style-src 'sha256-" + css_hash +
           "'; script-src 'sha256-" + js_hash + "'; connect-src 'none'; base-uri 'none'; form-action 'none'")
    escaped_title = html.escape(title, quote=True)
    total = 0

    def write(out, fragment: str) -> None:
        nonlocal total
        data = fragment.encode("utf-8")
        total += len(data)
        if total > MAX_HTML:
            fail("OUTPUT_TOO_LARGE", "HTML 크기가 256 MiB 제한을 초과했습니다.")
        out.write(data)

    try:
        doc = fitz.open(pdf)
        if doc.page_count < 1 or doc.page_count > MAX_PAGES:
            fail("INVALID_PDF", "변환된 PDF의 페이지 수가 허용 범위를 벗어났습니다.")
        if len(numbers) != doc.page_count:
            fail("INTERNAL_PAGE_MAPPING", "슬라이드 번호 대응에 실패했습니다.")
        with destination.open("wb") as out:
            write(out, '<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">')
            write(out, '<meta http-equiv="Content-Security-Policy" content="' + html.escape(csp, quote=True) + '">')
            write(out, "<title>" + escaped_title + "</title><style>" + CSS + "</style></head><body>")
            write(out, '<header><h1>' + escaped_title + '</h1><button id="prev" type="button" aria-label="이전 슬라이드">이전</button>')
            write(out, '<select id="slide-select" aria-label="슬라이드 선택">')
            for index, number in enumerate(numbers):
                write(out, f'<option value="{index}">슬라이드 {number}</option>')
            write(out, '</select><button id="next" type="button" aria-label="다음 슬라이드">다음</button><button id="mode" type="button" aria-pressed="false">한 장씩 보기</button><button id="print" type="button">인쇄</button></header><main>')
            # Page SVG is inert when loaded through <img>; SVG scripts and links
            # cannot execute or navigate in the parent document.
            for page_index, number in enumerate(numbers):
                page = doc[page_index]
                svg = page.get_svg_image(text_as_path=True)
                encoded = base64.b64encode(svg.encode("utf-8")).decode("ascii")
                accessible_text = html.escape(page.get_text("text")[:100_000], quote=True)
                write(out, f'<figure><img alt="슬라이드 {number}" src="data:image/svg+xml;base64,{encoded}"><figcaption>슬라이드 {number}</figcaption><details><summary>슬라이드 텍스트</summary><pre>{accessible_text}</pre></details></figure>')
            write(out, '</main><script>' + JS + '</script></body></html>')
        return doc.page_count, total
    except ConversionError:
        raise
    except (OSError, RuntimeError, ValueError) as exc:
        fail("RENDER_FAILED", f"PDF 페이지를 HTML로 만드는 데 실패했습니다: {str(exc)[:160]}")
    finally:
        if "doc" in locals():
            doc.close()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="PowerPoint 없이 로컬에서 PPTX를 단일 HTML로 변환합니다.")
    parser.add_argument("input", nargs="?", type=Path, help="입력 .pptx 파일")
    parser.add_argument("-o", "--output", type=Path, help="출력 .html 경로 (기본: 입력 파일과 같은 위치)")
    parser.add_argument("--soffice", help="LibreOffice soffice 실행 파일 경로")
    parser.add_argument("--include-hidden", action="store_true", help="숨김 슬라이드 포함")
    parser.add_argument("--slides", help="슬라이드 선택, 예: 1,3-5")
    parser.add_argument("--title", help="HTML 제목")
    parser.add_argument("--timeout", type=int, default=120, help="LibreOffice 제한 시간(초)")
    parser.add_argument("--overwrite", action="store_true", help="기존 출력 덮어쓰기")
    parser.add_argument("--json", action="store_true", help="결과를 JSON으로 출력")
    parser.add_argument("--check", action="store_true", help="의존성만 검사")
    return parser.parse_args()


def convert(args: argparse.Namespace) -> dict:
    check_python()
    office = find_soffice(args.soffice)
    fitz = get_fitz()
    if args.check:
        return {"ok": True, "version": VERSION, "python": sys.version.split()[0], "libreoffice": str(office),
                "libreoffice_version": safe_version([str(office), "--version"]),
                "pymupdf_version": fitz.VersionBind, "pymupdf_path": str(Path(fitz.__file__).resolve())}
    if args.input is None:
        fail("INPUT_REQUIRED", "입력 PPTX 파일을 지정하세요.")
    if args.timeout < 1 or args.timeout > 3600:
        fail("INVALID_TIMEOUT", "--timeout은 1~3600초여야 합니다.")
    source = args.input.expanduser().resolve()
    output = (args.output or source.with_suffix(".html")).expanduser().resolve()
    if source.suffix.lower() != ".pptx":
        fail("INPUT_EXTENSION", "입력 파일은 .pptx여야 합니다.")
    if not source.is_file():
        fail("INPUT_NOT_FOUND", f"입력 파일을 찾을 수 없습니다: {source}")
    input_size = source.stat().st_size
    if input_size == 0:
        fail("INPUT_EMPTY", "입력 파일이 비어 있습니다.")
    if input_size > MAX_INPUT:
        fail("INPUT_TOO_LARGE", "입력 파일이 64 MiB 제한을 초과했습니다.")
    if output == source:
        fail("OUTPUT_EQUALS_INPUT", "출력 경로가 입력 파일과 같습니다.")
    if not output.parent.is_dir():
        fail("OUTPUT_DIRECTORY", f"출력 폴더가 없습니다: {output.parent}")
    if output.exists() and not args.overwrite:
        fail("OUTPUT_EXISTS", f"출력 파일이 이미 있습니다: {output} (--overwrite 사용 가능)")
    metadata = slide_metadata(source)
    warnings: list[str] = []
    with tempfile.TemporaryDirectory(prefix="pptx-html-") as work_string:
        work = Path(work_string)
        profile = isolated_profile(work)
        pdf_dir = work / "pdf"
        pdf_dir.mkdir()
        options = {"ExportHiddenSlides": {"type": "boolean", "value": "true" if args.include_hidden else "false"},
                   "ExportNotesPages": {"type": "boolean", "value": "false"},
                   "ExportNotes": {"type": "boolean", "value": "false"},
                   "ExportOnlyNotesPages": {"type": "boolean", "value": "false"},
                   "ExportBookmarks": {"type": "boolean", "value": "false"},
                   "UseLosslessCompression": {"type": "boolean", "value": "true"},
                   "ReduceImageResolution": {"type": "boolean", "value": "false"}}
        command = [str(office), "-env:UserInstallation=" + profile.as_uri(), "--headless", "--nologo",
                   "--nodefault", "--norestore", "--infilter=Impress MS PowerPoint 2007 XML",
                   "--convert-to", "pdf:impress_pdf_Export:" + json.dumps(options, separators=(",", ":")),
                   "--outdir", str(pdf_dir), str(source)]
        result = run_office(command, args.timeout)
        pdf = pdf_dir / (source.stem + ".pdf")
        if result.returncode != 0 or not pdf.is_file() or pdf.stat().st_size < 8:
            detail = (result.stderr or result.stdout).strip().replace("\r", " ").replace("\n", " ")[:MAX_MESSAGE]
            fail("CONVERSION_FAILED", "LibreOffice가 입력을 읽거나 PDF로 변환하지 못했습니다. 암호화/DRM, 파일 형식, LibreOffice 호환성 또는 손상 가능성을 확인하세요." + (" 세부 정보: " + detail if detail else ""))
        with pdf.open("rb") as stream:
            if stream.read(5) != b"%PDF-":
                fail("INVALID_PDF", "LibreOffice 출력이 유효한 PDF가 아닙니다.")
        try:
            doc = fitz.open(pdf)
            count = doc.page_count
            doc.close()
        except (RuntimeError, ValueError) as exc:
            fail("INVALID_PDF", f"변환된 PDF를 읽을 수 없습니다: {str(exc)[:160]}")
        if count < 1 or count > MAX_PAGES:
            fail("INVALID_PDF", "변환된 PDF의 페이지 수가 허용 범위를 벗어났습니다.")
        if metadata is not None:
            expected = [i for i, visible in enumerate(metadata, 1) if visible or args.include_hidden]
            if len(expected) != count:
                fail("SLIDE_COUNT_MISMATCH", f"원본에서 예상한 슬라이드 {len(expected)}장과 PDF 출력 {count}장이 다릅니다. 숨김 슬라이드 또는 PDF 내보내기 설정을 확인하세요.")
            available = expected
            numbering = "original"
        else:
            available = list(range(1, count + 1))
            numbering = "exported_page"
            warnings.append("원본 슬라이드 목록을 읽을 수 없어 출력 PDF 페이지 번호를 사용했습니다.")
        selected = select_numbers(args.slides, available)
        selected_indexes = [available.index(n) for n in selected]
        # A small selected-page PDF prevents rendering unselected slides into HTML.
        if len(selected) != count:
            subset = pdf_dir / "selected.pdf"
            source_doc = fitz.open(pdf)
            selected_doc = fitz.open()
            try:
                for index in selected_indexes:
                    selected_doc.insert_pdf(source_doc, from_page=index, to_page=index)
                selected_doc.save(subset)
            finally:
                selected_doc.close()
                source_doc.close()
            pdf = subset
        html_temp = tempfile.NamedTemporaryFile(prefix=".pptx-html-", suffix=".tmp", dir=output.parent, delete=False)
        temp_path = Path(html_temp.name)
        html_temp.close()
        try:
            pages, size = render_html(pdf, temp_path, args.title or source.stem, selected, fitz)
            if args.overwrite:
                os.replace(temp_path, output)
            else:
                try:
                    os.link(temp_path, output)
                except FileExistsError:
                    fail("OUTPUT_EXISTS", f"출력 파일이 이미 있습니다: {output}")
                temp_path.unlink()
        finally:
            temp_path.unlink(missing_ok=True)
    return {"ok": True, "input": str(source), "output": str(output), "slides": pages,
            "slide_numbers": selected, "numbering": numbering, "bytes": size, "warnings": warnings}


def main() -> int:
    args = parse_args()
    try:
        result = convert(args)
    except KeyboardInterrupt:
        result = {"ok": False, "error": {"code": "CANCELLED", "message": "사용자가 변환을 중단했습니다."}}
    except ConversionError as exc:
        result = {"ok": False, "error": {"code": exc.code, "message": str(exc)[:MAX_MESSAGE]}}
    except (OSError, PermissionError) as exc:
        result = {"ok": False, "error": {"code": "IO_ERROR", "message": str(exc)[:MAX_MESSAGE]}}
    if args.json:
        print(json.dumps(result, ensure_ascii=False))
    elif result["ok"]:
        if args.check:
            print(f"Python {result['python']}; LibreOffice {result['libreoffice_version']} ({result['libreoffice']}); PyMuPDF {result['pymupdf_version']}")
        else:
            print(f"변환 완료: {result['output']} ({result['slides']}장, {result['bytes']}바이트)")
            for warning in result["warnings"]:
                print("주의: " + warning, file=sys.stderr)
    else:
        print(f"오류 [{result['error']['code']}]: {result['error']['message']}", file=sys.stderr)
    return 0 if result["ok"] else (130 if result["error"]["code"] == "CANCELLED" else 1)


if __name__ == "__main__":
    sys.exit(main())

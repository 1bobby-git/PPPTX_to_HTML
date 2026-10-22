#requires -Version 5.1
<#
.SYNOPSIS
설치된 Windows PowerPoint의 정상 PNG 내보내기로 독립 HTML을 만듭니다.
.DESCRIPTION
Windows PowerShell 5.1과 데스크톱 PowerPoint가 필요합니다. 인수가 없으면
파일 열기/저장 대화상자를 표시합니다. 이미 열린 같은 파일은 현재 상태를
읽기만 하며 닫지 않습니다. 새 파일은 읽기 전용의 보이는 창으로 엽니다.
PowerPoint/DRM 제품이 열기와 이미지 내보내기를 허용해야 합니다.
암호/DRM을 해제하거나 보안 정책을 우회하지 않습니다. 도우미 자체에는
네트워크 요청이 없으며, PowerPoint와 보안 제품의 인증은 기존 정책을 따릅니다.
결과는 PNG 이미지 기반입니다. 텍스트 선택/검색, 애니메이션, 영상/음성,
클릭 가능한 링크, 발표자 노트 및 편집 가능한 개체를 보존하지 않습니다.
.PARAMETER InputPath
로컬 PPTX 파일의 literal 경로입니다. 공백/한글 경로는 따옴표로 감싸세요.
.PARAMETER OutputPath
로컬 HTML 파일 경로입니다. 생략하면 저장 대화상자를 표시합니다.
.PARAMETER IncludeHiddenSlides
숨긴 슬라이드를 포함합니다. 기본값은 제외이며 원본 순서 번호를 유지합니다.
.PARAMETER MaxImageSize
이미지의 긴 변의 최대 픽셀 수입니다. 기본값 1920, 허용 범위 320~3840.
.PARAMETER Overwrite
명시한 OutputPath의 기존 파일을 모든 내보내기가 성공한 뒤 교체합니다.
저장 대화상자에서는 Windows의 덮어쓰기 확인을 사용합니다.
.EXAMPLE
.\Convert-PptxWithPowerPoint.ps1
.EXAMPLE
.\Convert-PptxWithPowerPoint.ps1 -InputPath 'C:\자료\보고서.pptx' -OutputPath 'C:\자료\보고서.html' -IncludeHiddenSlides
.NOTES
PowerPoint API: Presentations.Open(FileName, ReadOnly, Untitled, WithWindow)
https://learn.microsoft.com/office/vba/api/powerpoint.presentations.open
https://learn.microsoft.com/office/vba/api/powerpoint.slide.export
이 스크립트의 한글을 Windows PowerShell 5.1에서 읽도록 UTF-8 BOM을 유지하세요.
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string] $InputPath,
    [Parameter(Position = 1)]
    [string] $OutputPath,
    [switch] $IncludeHiddenSlides,
    [ValidateRange(320, 3840)]
    [int] $MaxImageSize = 1920,
    [switch] $Overwrite
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Release-ComReference {
    param([object] $Reference)
    if ($null -ne $Reference -and [Runtime.InteropServices.Marshal]::IsComObject($Reference)) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Reference) }
        catch { Write-Warning 'PowerPoint COM 참조를 해제하지 못했습니다. PowerPoint 창은 종료하지 않습니다.' }
    }
}

function Get-ComIdentity {
    param([object] $Reference)
    $pointer = [Runtime.InteropServices.Marshal]::GetIUnknownForObject($Reference)
    try { return $pointer.ToInt64() }
    finally { [void][Runtime.InteropServices.Marshal]::Release($pointer) }
}

function Get-LocalLiteralPath {
    param([string] $Path)
    # This resolves a literal path without wildcard expansion or a shell command.
    $provider = $null
    $drive = $null
    $resolved = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path, [ref]$provider, [ref]$drive)
    if ($provider.Name -ne 'FileSystem') { throw '파일 시스템의 로컬 경로를 선택해 주세요.' }
    $resolved = [IO.Path]::GetFullPath($resolved)
    if ($resolved.StartsWith('\\', [StringComparison]::Ordinal)) {
        throw '네트워크/UNC 경로 대신 이 PC의 로컬 파일과 저장 폴더를 선택해 주세요.'
    }
    $volume = [IO.DriveInfo]::new([IO.Path]::GetPathRoot($resolved))
    if ($volume.DriveType -eq [IO.DriveType]::Network) {
        throw '네트워크 드라이브 대신 이 PC의 로컬 파일과 저장 폴더를 선택해 주세요.'
    }
    return $resolved
}

function Assert-ExportedPng {
    param([string] $Path, [int] $Width, [int] $Height)
    $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    try {
        if ($stream.Length -lt 45) { throw 'PNG 파일이 비어 있거나 불완전합니다.' }
        $header = [byte[]]::new(24)
        if ($stream.Read($header, 0, $header.Length) -ne $header.Length) { throw 'PNG 헤더를 읽지 못했습니다.' }
        $signature = [byte[]](137, 80, 78, 71, 13, 10, 26, 10)
        for ($i = 0; $i -lt $signature.Length; $i++) {
            if ($header[$i] -ne $signature[$i]) { throw 'PowerPoint가 유효한 PNG를 만들지 못했습니다.' }
        }
        if ([Text.Encoding]::ASCII.GetString($header, 12, 4) -ne 'IHDR') { throw 'PNG 크기 정보가 없습니다.' }
        $actualWidth = $header[16] * 16777216L + $header[17] * 65536L + $header[18] * 256L + $header[19]
        $actualHeight = $header[20] * 16777216L + $header[21] * 65536L + $header[22] * 256L + $header[23]
        if ($actualWidth -ne $Width -or $actualHeight -ne $Height) { throw '내보낸 PNG 크기가 요청한 슬라이드 크기와 다릅니다.' }
        [void]$stream.Seek(-12, [IO.SeekOrigin]::End)
        $tail = [byte[]]::new(12)
        if ($stream.Read($tail, 0, $tail.Length) -ne $tail.Length) { throw 'PNG 끝부분을 읽지 못했습니다.' }
        $endChunk = [byte[]](0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130)
        for ($i = 0; $i -lt $endChunk.Length; $i++) {
            if ($tail[$i] -ne $endChunk[$i]) { throw 'PNG 내보내기가 완전히 끝나지 않았습니다.' }
        }
    }
    finally { $stream.Dispose() }
}

function Write-StandaloneHtml {
    param(
        [string] $Path, [string] $Title, [object[]] $Items,
        [int] $SourceCount, [int] $Width, [int] $Height
    )
    $safeTitle = [Net.WebUtility]::HtmlEncode($Title)
    $orientation = if ($Width -ge $Height) { 'landscape' } else { 'portrait' }
    $printHeight = if ($Width -ge $Height) { '177mm' } else { '264mm' }
    $stream = [IO.File]::Open($Path, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
    $writer = $null
    try {
        $writer = [IO.StreamWriter]::new($stream, [Text.UTF8Encoding]::new($false))
        $writer.WriteLine('<!doctype html>')
        $writer.WriteLine('<html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">')
        $writer.WriteLine('<meta http-equiv="Content-Security-Policy" content="default-src ''none''; img-src data:; style-src ''unsafe-inline''; script-src ''unsafe-inline''; connect-src ''none''; object-src ''none''; base-uri ''none''; form-action ''none''">')
        $writer.WriteLine('<title>' + $safeTitle + '</title>')
        $writer.WriteLine(@'
<style>
:root{color-scheme:light;font-family:system-ui,-apple-system,"Segoe UI",sans-serif;color:#172437;background:#eef2f7}
*{box-sizing:border-box}body{margin:0}header,main{width:min(100% - 32px,1440px);margin:0 auto}
header{padding:20px 0 14px}h1{font-size:1.2rem;line-height:1.5;overflow-wrap:anywhere;margin:0 0 12px}
nav{display:flex;gap:8px;align-items:center;flex-wrap:wrap}nav[hidden]{display:none}
button,select{font:inherit;color:inherit;background:white;border:1px solid #bac6d5;border-radius:8px;padding:9px 12px;max-width:100%}
button{cursor:pointer}button:disabled{opacity:.45;cursor:default}button:focus-visible,select:focus-visible{outline:3px solid #1967d2;outline-offset:2px}
#status{font-size:.9rem}p.note{font-size:.875rem;line-height:1.6;color:#4c5e73;margin:12px 0 0}
.slide{margin:0 0 24px;break-inside:avoid}.js .slide{display:none}.js .slide.active{display:block}
figure{margin:0}img{display:block;width:100%;height:auto;margin:0 auto;background:white;box-shadow:0 2px 12px #10243b18}
figcaption{text-align:center;padding:10px;font-size:.875rem;color:#4c5e73}
@media(max-width:480px){header,main{width:calc(100% - 16px)}button,select{padding:8px}#status{flex-basis:100%}}
'@)
        $writer.WriteLine('img{max-width:' + $Width + 'px}@page{size:A4 ' + $orientation + ';margin:10mm}')
        $writer.WriteLine('@media print{body{background:white}header{display:none}main{width:100%;margin:0}.js .slide,.slide{display:block!important;margin:0;break-after:page;page-break-after:always;break-inside:avoid;page-break-inside:avoid}.slide:last-child{break-after:auto;page-break-after:auto}img{width:100%;max-width:100%;max-height:' + $printHeight + ';object-fit:contain;box-shadow:none}figcaption{font-size:8pt;padding:2mm 0 0}}')
        $writer.WriteLine('</style></head><body><header><h1>' + $safeTitle + '</h1>')
        $writer.WriteLine('<nav id="controls" aria-label="슬라이드 탐색" hidden><button id="previous" type="button">이전</button><label for="slide-picker">슬라이드</label><select id="slide-picker">')
        for ($i = 0; $i -lt $Items.Count; $i++) {
            $label = '원본 ' + $Items[$i].Number
            if ($Items[$i].Hidden) { $label += ' (숨김)' }
            $writer.WriteLine('<option value="' + $i + '">' + [Net.WebUtility]::HtmlEncode($label) + '</option>')
        }
        $writer.WriteLine('</select><button id="next" type="button">다음</button><button id="print" type="button">전체 인쇄</button><span id="status" role="status" aria-live="polite"></span></nav>')
        $writer.WriteLine('<p class="note">PowerPoint에서 내보낸 PNG 이미지로 만든 정적 문서입니다. 텍스트 선택·검색, 애니메이션, 영상·음성, 클릭 가능한 링크 및 발표자 노트는 포함하지 않습니다. 방향키·Page Up/Down·Home/End로 이동할 수 있습니다. 인쇄 시 모든 슬라이드를 한 장씩 출력합니다.</p></header>')
        $writer.WriteLine('<main id="slides" data-source-count="' + $SourceCount + '">')
        foreach ($item in $Items) {
            $label = '원본 슬라이드 ' + $item.Number + ' / ' + $SourceCount
            if ($item.Hidden) { $label += ' (숨김)' }
            $safeLabel = [Net.WebUtility]::HtmlEncode($label)
            $writer.Write('<section class="slide" data-original="' + $item.Number + '" aria-label="' + $safeLabel + '"><figure><img width="' + $Width + '" height="' + $Height + '" alt="' + $safeLabel + '" src="data:image/png;base64,')
            $writer.Write([Convert]::ToBase64String([IO.File]::ReadAllBytes($item.Path)))
            $writer.WriteLine('"><figcaption>' + $safeLabel + '</figcaption></figure></section>')
        }
        $writer.WriteLine(@'
</main><script>
(() => {
  'use strict';
  const slides = Array.from(document.querySelectorAll('.slide'));
  if (!slides.length) return;
  const picker = document.getElementById('slide-picker');
  const previous = document.getElementById('previous');
  const next = document.getElementById('next');
  const status = document.getElementById('status');
  const sourceCount = document.getElementById('slides').dataset.sourceCount;
  let current = 0;
  function show(index) {
    current = Math.max(0, Math.min(slides.length - 1, index));
    slides.forEach((slide, i) => slide.classList.toggle('active', i === current));
    picker.value = String(current);
    previous.disabled = current === 0;
    next.disabled = current === slides.length - 1;
    status.textContent = `원본 ${slides[current].dataset.original} / ${sourceCount} · 출력 ${current + 1} / ${slides.length}`;
  }
  previous.addEventListener('click', () => show(current - 1));
  next.addEventListener('click', () => show(current + 1));
  picker.addEventListener('change', () => show(Number(picker.value)));
  document.getElementById('print').addEventListener('click', () => window.print());
  document.addEventListener('keydown', event => {
    if (event.altKey || event.ctrlKey || event.metaKey || event.target.closest('select,input,textarea,[contenteditable]')) return;
    const moves = {ArrowRight: current + 1, ArrowDown: current + 1, PageDown: current + 1,
      ArrowLeft: current - 1, ArrowUp: current - 1, PageUp: current - 1, Home: 0, End: slides.length - 1};
    if (Object.prototype.hasOwnProperty.call(moves, event.key)) {
      event.preventDefault();
      show(moves[event.key]);
    }
  });
  show(0);
  document.body.classList.add('js');
  document.getElementById('controls').hidden = false;
})();
</script></body></html>
'@)
        $writer.Flush()
        $stream.Flush($true)
    }
    finally {
        if ($null -ne $writer) { $writer.Dispose() }
        else { $stream.Dispose() }
    }
}

$powerPoint = $null
$presentations = $null
$presentation = $null
$slides = $null
$pageSetup = $null
$openedByHelper = $false
$workingDirectory = $null
$stagedHtml = $null
$exitCode = 0

try {
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
        throw '이 도우미는 Windows의 데스크톱 PowerPoint와 Windows PowerShell 5.1에서 실행하세요.'
    }
    if ([Threading.Thread]::CurrentThread.ApartmentState -ne [Threading.ApartmentState]::STA) {
        throw '동봉된 CMD를 실행하거나 powershell.exe -NoProfile -STA -File 옵션으로 실행하세요.'
    }
    if ([string]::IsNullOrWhiteSpace($InputPath) -or [string]::IsNullOrWhiteSpace($OutputPath)) {
        Add-Type -AssemblyName System.Windows.Forms
    }
    if ([string]::IsNullOrWhiteSpace($InputPath)) {
        $dialog = [Windows.Forms.OpenFileDialog]::new()
        try {
            $dialog.Title = 'PowerPoint에서 정상적으로 열 수 있는 PPTX 선택'
            $dialog.Filter = 'PowerPoint 프레젠테이션 (*.pptx)|*.pptx'
            $dialog.CheckFileExists = $true
            $dialog.Multiselect = $false
            $dialog.RestoreDirectory = $true
            if ($dialog.ShowDialog() -ne [Windows.Forms.DialogResult]::OK) { return }
            $InputPath = $dialog.FileName
        }
        finally { $dialog.Dispose() }
    }
    $sourcePath = Get-LocalLiteralPath $InputPath
    if (-not [IO.File]::Exists($sourcePath)) { throw '선택한 원본 파일을 찾을 수 없습니다.' }
    if (-not [IO.Path]::GetExtension($sourcePath).Equals('.pptx', [StringComparison]::OrdinalIgnoreCase)) {
        throw '이 도우미에는 .pptx 파일을 선택해 주세요.'
    }
    $overwriteApproved = $Overwrite.IsPresent
    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        $dialog = [Windows.Forms.SaveFileDialog]::new()
        try {
            $dialog.Title = '독립 HTML 저장 위치'
            $dialog.Filter = 'HTML 문서 (*.html)|*.html'
            $dialog.DefaultExt = 'html'
            $dialog.AddExtension = $true
            $dialog.CheckPathExists = $true
            $dialog.OverwritePrompt = $true
            $dialog.RestoreDirectory = $true
            $dialog.InitialDirectory = [IO.Path]::GetDirectoryName($sourcePath)
            $dialog.FileName = [IO.Path]::GetFileNameWithoutExtension($sourcePath) + '.html'
            if ($dialog.ShowDialog() -ne [Windows.Forms.DialogResult]::OK) { return }
            $OutputPath = $dialog.FileName
            $overwriteApproved = [IO.File]::Exists($OutputPath)
        }
        finally { $dialog.Dispose() }
    }
    $destinationPath = Get-LocalLiteralPath $OutputPath
    if ($sourcePath.Equals($destinationPath, [StringComparison]::OrdinalIgnoreCase)) { throw '원본과 다른 HTML 저장 경로를 선택해 주세요.' }
    if ([IO.Path]::GetExtension($destinationPath).ToLowerInvariant() -notin @('.html', '.htm')) { throw '출력 파일 확장자는 .html 또는 .htm이어야 합니다.' }
    $destinationDirectory = [IO.Path]::GetDirectoryName($destinationPath)
    if (-not [IO.Directory]::Exists($destinationDirectory)) { throw '출력 폴더가 없습니다. 기존 로컬 폴더를 선택해 주세요.' }
    if ([IO.File]::Exists($destinationPath) -and -not $overwriteApproved) {
        throw '출력 파일이 이미 있습니다. 새 이름을 지정하거나 -Overwrite를 명시하세요.'
    }
    Write-Host 'PowerPoint에서 파일을 읽습니다. 필요한 정상 인증/암호 입력은 PowerPoint 창에서 진행하세요.'
    Write-Host '문서 보안 정책이 이미지 내보내기를 허용해야 하며, 변환 중에는 해당 프레젠테이션을 편집하지 마세요.'
    try { $powerPoint = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') }
    catch { $powerPoint = New-Object -ComObject PowerPoint.Application }
    $powerPoint.Visible = -1
    $presentations = $powerPoint.Presentations
    $existingIdentities = [Collections.Generic.HashSet[long]]::new()
    $existingCount = [int]$presentations.Count
    for ($i = 1; $i -le $existingCount; $i++) {
        $candidate = $null
        try {
            $candidate = $presentations.Item($i)
            [void]$existingIdentities.Add((Get-ComIdentity $candidate))
            if ($null -eq $presentation -and $sourcePath.Equals([string]$candidate.FullName, [StringComparison]::OrdinalIgnoreCase)) {
                $presentation = $candidate
                $candidate = $null
            }
        }
        finally { Release-ComReference $candidate }
    }
    if ($null -eq $presentation) {
        # Honor the Trust Center (2), or retain a pre-existing ForceDisable (3).
        # The application's original runtime setting is restored immediately.
        $previousSecurity = $powerPoint.AutomationSecurity
        try {
            $powerPoint.AutomationSecurity = [Math]::Max(2, [int]$previousSecurity)
            $presentation = $presentations.Open($sourcePath, -1, 0, -1)
            # Open can reuse an existing object. Never claim ownership of that object.
            $openedByHelper = -not $existingIdentities.Contains((Get-ComIdentity $presentation))
        }
        finally { $powerPoint.AutomationSecurity = $previousSecurity }
    }
    if (-not $openedByHelper) { Write-Host '이미 열린 해당 프레젠테이션의 현재 상태를 내보냅니다. 이 창은 닫지 않습니다.' }
    $slides = $presentation.Slides
    $sourceCount = [int]$slides.Count
    if ($sourceCount -lt 1) { throw '내보낼 슬라이드가 없습니다.' }
    $pageSetup = $presentation.PageSetup
    $sourceWidth = [double]$pageSetup.SlideWidth
    $sourceHeight = [double]$pageSetup.SlideHeight
    if ($sourceWidth -le 0 -or $sourceHeight -le 0 -or [double]::IsNaN($sourceWidth) -or [double]::IsNaN($sourceHeight) -or [double]::IsInfinity($sourceWidth) -or [double]::IsInfinity($sourceHeight)) {
        throw '슬라이드 크기를 읽지 못했습니다.'
    }
    $scale = $MaxImageSize / [Math]::Max($sourceWidth, $sourceHeight)
    $imageWidth = [int][Math]::Max(1, [Math]::Round($sourceWidth * $scale))
    $imageHeight = [int][Math]::Max(1, [Math]::Round($sourceHeight * $scale))
    $localTempDirectory = Get-LocalLiteralPath ([IO.Path]::GetTempPath())
    $workingDirectory = [IO.Path]::Combine($localTempDirectory, 'PptxToHtml-' + [Guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($workingDirectory)
    $exported = [Collections.Generic.List[object]]::new()
    for ($i = 1; $i -le $sourceCount; $i++) {
        $slide = $null
        $transition = $null
        try {
            $slide = $slides.Item($i)
            $transition = $slide.SlideShowTransition
            $isHidden = [int]$transition.Hidden -eq -1
            if ($isHidden -and -not $IncludeHiddenSlides) { continue }
            Write-Progress -Activity 'PowerPoint PNG → HTML' -Status ('원본 슬라이드 ' + $i + ' / ' + $sourceCount) -PercentComplete ([int](100 * ($i - 1) / $sourceCount))
            $pngPath = [IO.Path]::Combine($workingDirectory, ('slide-{0:D5}.png' -f $i))
            try {
                # Use only PowerPoint's own export. There is no capture or fallback bypass.
                $slide.Export($pngPath, 'PNG', $imageWidth, $imageHeight)
                Assert-ExportedPng -Path $pngPath -Width $imageWidth -Height $imageHeight
            }
            catch {
                throw [InvalidOperationException]::new(('원본 ' + $i + '번 슬라이드의 PNG 내보내기에 실패했습니다. PowerPoint/DRM 정책이 내보내기를 허용하는지 확인하세요. ' + $_.Exception.Message), $_.Exception)
            }
            $exported.Add([pscustomobject]@{ Number = $i; Hidden = $isHidden; Path = $pngPath })
        }
        finally {
            Release-ComReference $transition
            Release-ComReference $slide
        }
    }
    if ($exported.Count -eq 0) { throw '내보낼 슬라이드가 없습니다. 모두 숨김 상태라면 -IncludeHiddenSlides를 사용하세요.' }
    $stagedHtml = [IO.Path]::Combine($destinationDirectory, ('.pptx-html-' + [Guid]::NewGuid().ToString('N') + '.tmp'))
    Write-StandaloneHtml -Path $stagedHtml -Title ([IO.Path]::GetFileNameWithoutExtension($sourcePath)) -Items $exported.ToArray() -SourceCount $sourceCount -Width $imageWidth -Height $imageHeight
    # Same-directory staging permits an atomic commit. Never delete the old output first.
    if ([IO.File]::Exists($destinationPath)) {
        if (-not $overwriteApproved) { throw '변환 중 같은 이름의 출력 파일이 생성되었습니다. 기존 파일은 보존했습니다.' }
        [IO.File]::Replace($stagedHtml, $destinationPath, $null)
    }
    else { [IO.File]::Move($stagedHtml, $destinationPath) }
    $stagedHtml = $null
    Write-Host ('완료: 원본 ' + $sourceCount + '장 중 ' + $exported.Count + '장, ' + $imageWidth + ' × ' + $imageHeight + ' PNG 기반 정적 HTML')
    Write-Host $destinationPath
}
catch {
    $exitCode = 1
    [Console]::Error.WriteLine('변환 실패: ' + $_.Exception.Message)
    [Console]::Error.WriteLine('PowerPoint에서 정상적으로 열고 PNG로 내보낼 권한이 필요합니다. 정책이 차단하면 문서 소유자나 관리자에게 허용된 배포본을 요청하세요.')
}
finally {
    Write-Progress -Activity 'PowerPoint PNG → HTML' -Completed
    Release-ComReference $pageSetup
    Release-ComReference $slides
    if ($null -ne $presentation -and $openedByHelper) {
        try {
            # If a user edited the helper's window, preserve that work and leave it open.
            if ([int]$presentation.Saved -eq -1) { $presentation.Close() }
            else { Write-Warning '변환 중 변경된 프레젠테이션이 있어 창을 열어 두었습니다. 필요한 변경을 직접 저장하세요.' }
        }
        catch { Write-Warning '도우미가 연 프레젠테이션 창을 닫지 못했습니다. 필요하면 직접 닫아 주세요.' }
    }
    Release-ComReference $presentation
    Release-ComReference $presentations
    # PowerPoint may share a running instance. Never call Application.Quit().
    Release-ComReference $powerPoint
    if ($null -ne $stagedHtml -and [IO.File]::Exists($stagedHtml)) {
        try { [IO.File]::Delete($stagedHtml) }
        catch { Write-Warning ('임시 HTML을 삭제하지 못했습니다: ' + $stagedHtml) }
    }
    if ($null -ne $workingDirectory -and [IO.Directory]::Exists($workingDirectory)) {
        try { [IO.Directory]::Delete($workingDirectory, $true) }
        catch { Write-Warning ('임시 PNG 폴더를 삭제하지 못했습니다: ' + $workingDirectory) }
    }
}
exit $exitCode

# LIRA 1차 멘토링 점검표 자동 채우기 v2
# HWPX 안의 placeholder 텍스트를 LIRA 프로젝트 실제 내용으로 치환

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$srcPath  = 'C:\Users\User\Documents\경기도청 메신저\Talk 받은 파일\[경기도의회] 1차 멘토링 프로젝트 점검표(법제과).hwpx'
$outPath  = 'C:\Users\User\Documents\이지영\9. AI agent\클로드\LIRA\docs\[법제과-LIRA] 1차 멘토링 프로젝트 점검표.hwpx'
$workDir  = Join-Path $env:TEMP ("hwpx_fill_" + (Get-Random))

if (-not (Test-Path -LiteralPath $srcPath)) {
    Write-Error "원본 없음: $srcPath"
    exit 1
}
Write-Output "원본: $srcPath"

# 1) 원본 ZIP으로 펼치기 (System.IO.Compression 직접 사용 — 확장자 무관)
New-Item -ItemType Directory -Path $workDir -Force | Out-Null
[System.IO.Compression.ZipFile]::ExtractToDirectory($srcPath, $workDir)
Write-Output "압축 풀기 완료: $workDir"

# 검증
$xmlPath = Join-Path $workDir 'Contents\section0.xml'
if (-not (Test-Path $xmlPath)) {
    Write-Error "section0.xml 없음: $xmlPath"
    Get-ChildItem $workDir -Recurse | Select-Object FullName | Format-Table
    exit 1
}

# 2) section0.xml 로드
$xml = [System.IO.File]::ReadAllText($xmlPath, [System.Text.Encoding]::UTF8)
Write-Output "원본 section0.xml 크기: $($xml.Length) chars"

# 3) placeholder → 실제 내용 치환
$replacements = @{
    '어떤 업무를 개선하기 위한 프로젝트인지 한 문장으로 작성해주세요.' = '의원 발의안 조례안을 5개 영역(인용 무결성, 법체계 적법성, 자치법규 정합성, 입법기술, 양성평등 차별 영향)에서 자동 점검하고 검토의견서 ODT 초안을 생성하는 단일 HTML AI 도구 (LIRA, Legislative Intelligent Review Assistant)'

    '현재 실제로 구현한 기능을 작성해주세요.' = '① HWPX TXT 발의안 파일 업로드 및 텍스트 자동 추출 / ② 5개 영역 AI 자동 검토 / ③ 법제처 행안부 100개 chunk RAG (자치법규 입안 길라잡이, 알기쉬운 법령정비기준, 법령입안심사기준, 행안부 자치법규 업무매뉴얼, 지방의회 운영가이드북 등 8개 자료) / ④ 국가법령정보센터 OpenAPI 실시간 연동 (법령, 자치법규, 판례, 헌재결정, 법령해석례, 행정심판재결례, 자치법규해석례) / ⑤ Multi-AI (Claude Sonnet 4.6 / Gemini 2.5 Flash) 선택 가능 / ⑥ 검토의견서 ODT 종합 보고서 다운로드 (검토의견 + 발송문 + 의안 첨부 + 신구조문대비표 + 관계법령 발췌서) / ⑦ 신구조문대비표 자동 생성 (단어 단위 LCS diff, 빨강 취소선 파랑 굵게) / ⑧ 발의안 본문 인라인 수정 표시 / ⑨ 약칭 자동 인식 (제2조 정의 조항 파싱) / ⑩ Rate limit 자동 재시도 (429/503/504 백오프)'

    '계획은 있으나 아직 구현하지 못한 기능을 작성해주세요.' = '① 재정부담 자동 점검 (지방자치법 제148조) / ② 강행 임의 규정 자동 구분 점검 / ③ 분야별 chunks 보강 (보안, AI, 교육 등 전문 분야) / ④ PDF 직접 첨부 (현재 HWPX 변환 필요) / ⑤ 결과 시각화 (수정율 통계 등) / ⑥ 사용자 계정 로그 이력 관리 (서버화 단계)'

    '현재 팀이 판단하는 전체 구현률을 선택해주세요.' = '5개 핵심 영역 검토 + ODT 종합 보고서 + 신구조문대비표까지 모두 작동. 시연 가능 수준. (구현률 80% 이상)'

    '프로젝트에 투입되는 자료 또는 파일을 작성해주세요.' = '① 의원 발의안 조례안 (HWPX, TXT 파일 또는 텍스트 직접 입력) / ② AI API 키 (Claude 또는 Gemini) / ③ 국가법령정보센터 사용자 ID (선택, 미입력해도 작동) / ④ 보고서 메타정보 (발의자, 의안번호, 검토 담당자, 의뢰일자, 회신일자)'

    '프로젝트를 통해 최종적으로 생성되는 결과물을 작성해주세요.' = '① 화면: 5영역 검토 결과 카드 + 사전조회 결과 + AI 수정안 + 전체 종합 통계 / ② 다운로드 파일: 검토의견서 ODT 종합 보고서 (자치법규안 검토의견서 표 + 발송문 + 의안 첨부 본문 + 인라인 수정 표시 + 관계법령 발췌서 + 신구조문대비표 — 경기천년체 적용)'

    '자료 입력부터 결과물 생성까지의 흐름을 순서대로 작성해주세요.' = '① 발의안 HWPX 첨부 또는 텍스트 붙여넣기 → ② AI 제공자 선택 + API 키 입력 → ③ 보고서 메타 입력 (발의자 등, 선택) → ④ 검토 시작 클릭 → ⑤ 사전 조회 (약칭 추출, 인용법령 조문 조회, 경기조례 검색, 판례 헌재 해석례 행정심판 자치법규해석례 자동 수집) → ⑥ 5영역 검토 병렬 호출 (동시성 2 배치) → ⑦ AI 수정안 생성 → ⑧ 결과 화면 표시 → ⑨ 검토의견서 ODT 다운로드'

    '실행 과정에서 발생하는 오류나 문제가 있다면 작성해주세요.' = '① Claude API 분당 입력 토큰 30K 한도 초과 (rate_limit_error 429) → 동시성 2 배치 + 자동 재시도(35초 백오프)로 해결 / ② Gemini 503 high demand (UNAVAILABLE) → 자동 재시도(15초 35초 75초)로 해결 / ③ Gemini 1.5 Flash deprecated (Google 2024년 9월 종료) → 2.5 Flash로 교체 / ④ AI 응답 토큰 한도(4096) 초과로 수정안 JSON 잘림 → max_tokens 8192 상향 + corrections 배열 부분 복구 로직 추가 / ⑤ 의회 PC 보안망에서 외부 API CORS 차단 가능성 → allorigins 프록시 fallback 적용 / ⑥ Claude API 크레딧 소진 → Gemini 무료 전환 안내'

    '멘토링에서 가장 우선적으로 도움받고 싶은 부분을 선택해주세요.' = '주요 관심: Claude Code 사용 워크플로 최적화, 데이터 구조 설계 (RAG chunks 효율화 및 카테고리 매핑 정교화), 의회 사내 보안망 배포 (단일 HTML 더블클릭 방식)'

    '최종적으로 완성하고 싶은 결과물의 형태를 구체적으로 작성해주세요.' = '의회 PC에서 단일 HTML 파일을 더블클릭만으로 작동 → 발의안 HWPX 1개 첨부 → 약 60~80초 후 입법조사관 검토의견서 표준 양식의 종합 보고서 ODT 다운로드 → 그대로 인쇄 가능. 누구나 (법제과 자문관, 의원실 보좌관, 타 광역의회) 본인 API 키로 즉시 사용. 별도 설치 서버 불필요 (의회 보안망 충돌 회피).'

    '강사님께 가장 먼저 확인받고 싶은 내용을 작성해주세요.' = '① LIRA의 출품 차별화 포인트가 분명한지 (양성평등 차별 자동 점검 + 100개 chunk RAG + 단일 HTML 보급성 + Multi-AI 교체 가능 구조) / ② 실제 입법조사관(김 홍 6급) 검토의견 vs LIRA 검토의견 비교 시 LIRA의 현재 수준 평가 / ③ 출품 후 의회 전체 배포 전략 (API 키 발급, 폰트 설치, 사용 매뉴얼)'

    '현재 기능 중 강사님 점검이 필요한 부분을 작성해주세요.' = '① ODT 종합 보고서 생성 결과 — 한글 LibreOffice에서 정상 렌더링 여부 (경기천년체 폰트 적용 확인) / ② 인라인 수정 표시(취소선+색상) 정확도 — 단어 단위 LCS diff 결과 / ③ 5개 영역 검토 의견 품질 — 특히 분야 전문성 (보안 AI 등) 부족 여부 / ④ 신구조문대비표 2칸 표 자동 생성 정확도 / ⑤ 국가법령정보센터 OpenAPI 응답 파싱 (lawService.do MST 조회 결과)'

    '추가로 구현하고 싶은 기능이 있다면 작성해주세요.' = '① 재정부담 자동 점검 (지방자치법 제148조 — 비용 예산 지원 키워드 자동 검출) / ② 강행 임의 규정 자동 구분 점검 / ③ 분야별 chunks 보강 (보안 AI 교육 환경 등 전문 분야 발췌 추가) / ④ 다중 AI 결과 비교 도구 (Claude vs Gemini 동시 호출 후 차이 표시) / ⑤ 검토 이력 관리 (의회 출품 후 서버화 단계) / ⑥ PDF 발의안 직접 첨부 (현재 HWPX 변환 필요)'

    '실제 업무 자료로 테스트했는지 선택해주세요.' = '예. 경기도 외국인 공공보건 접근성 향상 조례안, 경기도 생성형 인공지능 플랫폼 운영 조례안 등 실 발의안 2건 이상 테스트 완료. 입법조사관(김 홍 6급) 작성 검토의견서와 비교 분석 보고서 별도 작성.'

    '현재 결과물을 실제 업무에 활용할 수 있는 수준인지 선택해주세요.' = '일부 수정 필요. 5개 핵심 기능 + ODT 종합 보고서 작동하나, 분야 전문성(보안 AI 등) 보강 및 N/A 항목 추가 정리 필요. 입법조사관 검토와 비교 시 LIRA는 표준화 양성평등 자동 점검에서 우수, 분야 전문지식 구체성에서 부족.'

    '2차 멘토링 전까지 보완하거나 고도화하고 싶은 부분을 작성해주세요.' = '① 분야별 chunks 보강 (전문 분야 30~50개 추가) / ② 재정부담 + 강행/임의 규정 점검 영역 추가 / ③ N/A 출력 추가 정리 및 결과 검증 (v0.10.3에서 1차 완료, 추가 다듬기) / ④ 실 발의안 5건 이상 검증 / ⑤ 시연 영상 제작 (1분짜리 데모) / ⑥ README v0.10 동기화 및 사용 매뉴얼 작성'
}

$count = 0
foreach ($key in $replacements.Keys) {
    $val = $replacements[$key]
    if ($xml.Contains($key)) {
        $xml = $xml.Replace($key, $val)
        $count++
        Write-Output ("  OK : " + $key.Substring(0, [Math]::Min(35, $key.Length)) + "...")
    } else {
        Write-Output ("  MISS: " + $key.Substring(0, [Math]::Min(35, $key.Length)) + "...")
    }
}
Write-Output "치환 $count / $($replacements.Count) 건"

# 4) 체크박스 처리
$xml = $xml.Replace('□ 80% 이상', '■ 80% 이상')
$xml = $xml.Replace('□ Claude Code 사용', '■ Claude Code 사용')
$xml = $xml.Replace('□ 데이터구조설계', '■ 데이터구조설계')
$xml = $xml.Replace('□ 배포', '■ 배포')
$xml = $xml.Replace('□ 예  □ 아니오', '■ 예  □ 아니오')
$xml = $xml.Replace('□ 일부수정필요', '■ 일부수정필요')

# 5) 저장 (UTF-8 BOM 없음)
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($xmlPath, $xml, $utf8NoBom)
Write-Output "section0.xml 저장: $($xml.Length) chars"

# 6) HWPX 재패키지 — mimetype 첫 파일 STORED
$tmpZip = Join-Path $env:TEMP ("hwpx_filled_" + (Get-Random) + ".zip")
if (Test-Path $tmpZip) { Remove-Item $tmpZip -Force }

$zipStream = [System.IO.File]::Open($tmpZip, [System.IO.FileMode]::Create)
$archive = New-Object System.IO.Compression.ZipArchive($zipStream, [System.IO.Compression.ZipArchiveMode]::Create)

# (1) mimetype — STORED (무압축)
$mimePath = Join-Path $workDir 'mimetype'
if (Test-Path $mimePath) {
    $mimeEntry = $archive.CreateEntry('mimetype', [System.IO.Compression.CompressionLevel]::NoCompression)
    $mimeStream = $mimeEntry.Open()
    $mimeBytes = [System.IO.File]::ReadAllBytes($mimePath)
    $mimeStream.Write($mimeBytes, 0, $mimeBytes.Length)
    $mimeStream.Close()
}

# (2) 나머지 파일 — DEFLATE
Get-ChildItem $workDir -Recurse -File | Where-Object { $_.Name -ne 'mimetype' } | ForEach-Object {
    $relativePath = $_.FullName.Substring($workDir.Length + 1).Replace('\', '/')
    $entry = $archive.CreateEntry($relativePath, [System.IO.Compression.CompressionLevel]::Optimal)
    $entryStream = $entry.Open()
    $fileBytes = [System.IO.File]::ReadAllBytes($_.FullName)
    $entryStream.Write($fileBytes, 0, $fileBytes.Length)
    $entryStream.Close()
}

$archive.Dispose()
$zipStream.Close()

# 7) 최종 위치로 이동
if (Test-Path -LiteralPath $outPath) { Remove-Item -LiteralPath $outPath -Force }
Move-Item -LiteralPath $tmpZip -Destination $outPath
Remove-Item -LiteralPath $workDir -Recurse -Force

$outFile = Get-Item -LiteralPath $outPath
Write-Output ""
Write-Output "✓ 생성 완료: $($outFile.FullName)"
Write-Output "  크기: $($outFile.Length) bytes"

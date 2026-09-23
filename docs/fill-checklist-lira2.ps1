# LIRA2 1차 멘토링 점검표 자동 채우기
# 경기도의회 빈 템플릿의 placeholder를 LIRA2(조례안 자동검토 어시스턴트) 실제 사양으로 치환
# LIRA2 핵심: 단일 HTML / 검토 전용(수정안·신구조문대비표 생성 없음) / 출력=검토의견서 HWPX(양식 A)
#             / 5대 입안원칙 등 9개 항목 / LLM 2콜(구조화+검토) / 스트리밍(SSE) 호출 / 키 메모리 전용

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$srcPath  = 'C:\Users\User\Documents\경기도청 메신저\Talk 받은 파일\[경기도의회] 1차 멘토링 프로젝트 점검표(법제과).hwpx'
$outPath  = 'C:\Users\User\Documents\이지영\9. AI agent\클로드\LIRA\docs\[법제과-LIRA] 1차 멘토링 프로젝트 점검표.hwpx'
$workDir  = Join-Path $env:TEMP ("hwpx_fill2_" + (Get-Random))

if (-not (Test-Path -LiteralPath $srcPath)) {
    Write-Error "원본 없음: $srcPath"
    exit 1
}
Write-Output "원본: $srcPath"

# 1) 원본 ZIP으로 펼치기
New-Item -ItemType Directory -Path $workDir -Force | Out-Null
[System.IO.Compression.ZipFile]::ExtractToDirectory($srcPath, $workDir)
Write-Output "압축 풀기 완료: $workDir"

$xmlPath = Join-Path $workDir 'Contents\section0.xml'
if (-not (Test-Path $xmlPath)) {
    Write-Error "section0.xml 없음: $xmlPath"
    exit 1
}

# 2) section0.xml 로드
$xml = [System.IO.File]::ReadAllText($xmlPath, [System.Text.Encoding]::UTF8)
Write-Output "원본 section0.xml 크기: $($xml.Length) chars"

# 3) placeholder -> LIRA2 실제 내용 치환
$replacements = [ordered]@{
    # 프로젝트명 (블랭크 템플릿에 구 명칭이 들어 있음)
    '조례안 자동 검토 시스템' = '조례안 자동검토 어시스턴트 (LIRA2)'

    # 프로젝트 한 줄 소개
    '어떤 업무를 개선하기 위한 프로젝트인지 한 문장으로 작성해주세요.' = '의원 발의 조례안의 1차 검토를 보조하는 단일 HTML AI 도구. 조례안 HWPX를 올리면 추출·구조화 -> 국가법령정보센터 법령 조회 -> 5대 입안원칙 등 9개 항목 자동 점검을 거쳐, 검토의견서(HWPX 양식 A)를 한글에서 바로 열 수 있게 생성한다. (LIRA2, Legislative Intelligent Review Assistant — 입법, 다시 한 번 점검하는 AI 동료)'

    # 현재까지 구현한 기능
    '현재 실제로 구현한 기능을 작성해주세요.' = '① 조례안 HWPX 업로드 -> 텍스트·표 자동 추출 / ② LLM 1콜 구조화(제명·발의자·제안이유·주요내용·개정문·인용법령·후보법령 등 9필드) / ③ 국가법령정보센터 OpenAPI 실시간 연동(법령·자치법규·판례·헌재결정·법령해석례·행정심판재결례) — 0.3초 간격 레이트리미터·세션 캐시 / ④ 100개 chunk RAG(자치법규 입안 길라잡이, 알기쉬운 법령정비기준, 법령입안심사기준, 행안부 자치법규 업무매뉴얼, 지방의회 운영가이드북 등 8개 자료) / ⑤ 5대 입안원칙 등 9개 항목 일괄 검토(LLM 1콜) — 소관사무·법령우위·법률유보·권한경계·보조금·인용무결성·용어(양성평등·차별)·정합성·입법기술 / ⑥ 검토 결과 화면 표시(양식 B) + 검토의견서 HWPX 다운로드(양식 A) — 한글에서 바로 열림 / ⑦ Multi-AI(Claude Sonnet 4.6 / Gemini 2.5 Flash) 선택 / ⑧ AI 호출 스트리밍(SSE) 처리 — 의회 보안망 idle 연결 끊김 회피 / ⑨ 인용 법령 최신성·무결성 자동 경고(폐지·개정·조문이동) / ⑩ 키 메모리 전용(localStorage·sessionStorage 미사용)·단일 HTML 더블클릭 실행'

    # 아직 구현되지 않은 기능
    '계획은 있으나 아직 구현하지 못한 기능을 작성해주세요.' = '① 재정부담 자동 점검 고도화 (지방자치법 제148조 — 비용·예산·지원 키워드 자동 검출) / ② 강행·임의 규정 자동 구분 점검 / ③ 분야별 chunks 보강 (보안·AI·교육·환경 등 전문 분야) / ④ 분야 전문지식 기반 구체적 통합·중복 제거 제안 (입법조사관 수준) / ⑤ 결과 시각화·검토 이력 관리 (서버화 단계)'

    # 현재 진행률
    '현재 팀이 판단하는 전체 구현률을 선택해주세요.' = '추출 -> 법령조회 -> 9개 항목 검토 -> 검토의견서 HWPX 생성까지 직선형 파이프라인 전 구간 작동. 실 발의안으로 시연 가능. (구현률 80% 이상)'

    # 입력 데이터
    '프로젝트에 투입되는 자료 또는 파일을 작성해주세요.' = '① 의원 발의 조례안 (HWPX 권장, TXT 허용 — 구버전 HWP는 한글에서 HWPX로 저장 후 첨부) / ② AI API 키 (Claude 또는 Gemini, 메모리에만 보관·미저장) / ③ 국가법령정보센터 OpenAPI 사용자 ID(OC) — 선택, 미입력 시 "AI 단독" 모드로 작동'

    # 출력 결과
    '프로젝트를 통해 최종적으로 생성되는 결과물을 작성해주세요.' = '① 화면(양식 B): 구조화 9필드 요약 + 법제처 조회 결과 + 9개 항목별 검토의견(적합/검토필요/부적합 + 지적·제안·근거) / ② 다운로드 파일: 검토의견서 HWPX (양식 A) — 한글에서 바로 열리는 자치법규안 검토의견서 양식. 모든 지적에 근거(법령 조문·판례·해석례·RAG 청크 출처) 명시, 말미에 면책 고지 3종 포함'

    # 업무 흐름도
    '자료 입력부터 결과물 생성까지의 흐름을 순서대로 작성해주세요.' = '① 발의안 HWPX 첨부 -> ② AI 제공자 선택 + API 키 입력 + OC 입력(선택) -> ③ 검토 시작 -> ④ extract: 텍스트·표 추출 후 LLM 1콜로 9필드 구조화(후보법령 추론 포함) -> ⑤ lawfinder: 법제처 2단계 탐색(인용법령 유효성 -> 후보법령 본문·판례·해석례·헌재·행정심판·경기 자치법규), 0.3초 간격 -> ⑥ review: RAG 청크 검색 + 9개 항목 일괄 검토(LLM 1콜) -> ⑦ report: 화면 양식 B 렌더링 + 검토의견서 HWPX(양식 A) 생성·다운로드'

    # 발생하는 오류
    '실행 과정에서 발생하는 오류나 문제가 있다면 작성해주세요.' = '① 의회·관공서 보안망이 idle 외부 연결을 끊어 AI 검토 호출이 ERR_CONNECTION_RESET/타임아웃 -> AI 호출을 스트리밍(SSE)으로 전환, 데이터가 계속 흐르게 하여 해결 / ② 외부 API CORS 차단 -> allorigins 프록시 fallback / ③ 자치법규해석례(cgmExpc)는 법제처 lawSearch.do가 지원하지 않아 404 -> 해당 소스 제거(나머지 6종 정상) / ④ 무한 대기 -> fetch 타임아웃(법령 10초·AI idle 30초) 적용 / ⑤ OC 미입력 처리 -> 입력란 기본값(실제 등록 OC) 지정 + "AI 단독" 모드 분기 / ⑥ AI 응답 토큰 한도 -> max_tokens 8192 상향'

    # 가장 도움이 필요한 부분
    '멘토링에서 가장 우선적으로 도움받고 싶은 부분을 선택해주세요.' = '주요 관심: Claude Code 워크플로 최적화, 데이터 구조 설계(RAG chunks 효율화·카테고리 매핑), 의회 보안망 배포(단일 HTML 더블클릭 + 스트리밍 호출 안정성)'

    # 원하는 결과물
    '최종적으로 완성하고 싶은 결과물의 형태를 구체적으로 작성해주세요.' = '의회 PC에서 단일 HTML 더블클릭만으로 작동 -> 발의안 HWPX 1건 첨부 -> 추출·법령조회·검토를 거쳐 검토의견서 HWPX(양식 A)를 한글에서 바로 열어 그대로 인쇄·편집. 법제과 자문관·의원실 보좌관·타 광역의회 누구나 본인 API 키로 즉시 사용. 설치형 서버 불필요(의회 보안망 충돌 회피).'

    # 가장 먼저 확인받고 싶은 부분
    '강사님께 가장 먼저 확인받고 싶은 내용을 작성해주세요.' = '① LIRA2의 출품 차별화 포인트가 분명한지(9개 항목 표준 점검 + 양성평등·차별 자동 + 100개 chunk RAG + 단일 HTML 보급성 + Multi-AI 교체 가능 구조) / ② 실제 입법조사관 검토의견 vs LIRA2 비교 시 현재 수준 평가 / ③ 출품 후 의회 배포 전략(API 키 발급, 사용 매뉴얼)'

    # 점검이 필요한 기능
    '현재 기능 중 강사님 점검이 필요한 부분을 작성해주세요.' = '① 검토의견서 HWPX(양식 A) — 한글에서 정상 렌더링 여부 / ② 9개 항목 검토의견 품질 — 특히 분야 전문성(보안·AI 등) / ③ 모든 지적의 근거 인용 정확도(법령 조문·판례·해석례·RAG 청크 출처) / ④ 국가법령정보센터 OpenAPI 응답 파싱(lawSearch.do 목록·lawService.do MST 본문) / ⑤ 의회 보안망에서 스트리밍 AI 호출 안정성'

    # 추가 구현 희망
    '추가로 구현하고 싶은 기능이 있다면 작성해주세요.' = '① 재정부담 자동 점검 (지방자치법 제148조) / ② 강행·임의 규정 자동 구분 점검 / ③ 분야별 chunks 보강 (보안·AI·교육·환경) / ④ 분야 전문지식 기반 통합·중복 제거 제안 / ⑤ 검토 이력 관리 (서버화 단계) / ⑥ PDF 발의안 직접 첨부 (현재 HWPX 변환 필요)'

    # 실제 의회 업무 데이터 테스트 여부
    '실제 업무 자료로 테스트했는지 선택해주세요.' = '예. 경기도 외국인 공공보건 접근성 향상 조례안, 경기도 생성형 인공지능 플랫폼 운영 조례안 등 실 발의안 2건 이상 테스트 완료. 입법조사관(김 홍 6급) 작성 검토의견서와 비교 분석 보고서 별도 작성.'

    # 실제 업무 활용 가능성
    '현재 결과물을 실제 업무에 활용할 수 있는 수준인지 선택해주세요.' = '일부 수정 필요. 직선형 파이프라인 + 검토의견서 HWPX 생성 작동하나, 분야 전문성(보안·AI 등) 보강 필요. 입법조사관 검토와 비교 시 LIRA2는 표준화·양성평등 자동 점검·근거 검증 가능성에서 우수, 분야 전문지식 구체성에서 부족.'

    # 2차 멘토링 전 고도화 희망
    '2차 멘토링 전까지 보완하거나 고도화하고 싶은 부분을 작성해주세요.' = '① 분야별 chunks 보강(전문 분야 30~50개 추가) / ② 재정부담 + 강행·임의 규정 점검 영역 추가 / ③ 의회 보안망 실환경 검증(스트리밍 호출 안정성) / ④ 실 발의안 5건 이상 검증 / ⑤ 시연 영상(1분 데모) 제작 / ⑥ README·사용 매뉴얼 v0.11 동기화'
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

# 4) 체크박스 처리 (구현률 80% 이상 / 도움 항목 / 테스트 예 / 일부수정필요)
$xml = $xml.Replace('□ 80% 이상', '■ 80% 이상')
$xml = $xml.Replace('□ Claude Code 사용', '■ Claude Code 사용')
$xml = $xml.Replace('□ 데이터구조설계', '■ 데이터구조설계')
$xml = $xml.Replace('□ 배포', '■ 배포')
$xml = $xml.Replace('□ 예', '■ 예')
$xml = $xml.Replace('□ 일부수정필요', '■ 일부수정필요')

# 5) 저장 (UTF-8 BOM 없음)
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($xmlPath, $xml, $utf8NoBom)
Write-Output "section0.xml 저장: $($xml.Length) chars"

# 6) HWPX 재패키지 — mimetype 첫 파일 STORED
$tmpZip = Join-Path $env:TEMP ("hwpx_filled2_" + (Get-Random) + ".zip")
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
Write-Output "생성 완료: $($outFile.FullName)"
Write-Output "  크기: $($outFile.Length) bytes"

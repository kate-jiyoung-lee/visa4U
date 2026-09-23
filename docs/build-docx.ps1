# LIRA vs 입법조사관 검토의견 비교 분석 — Word 문서 생성 스크립트
# Run: powershell -File build-docx.ps1

$outPath = 'C:\Users\User\Documents\이지영\9. AI agent\클로드\LIRA\docs\LIRA-vs-입법조사관-비교분석.docx'

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0  # wdAlertsNone

$doc = $word.Documents.Add()

# 페이지 여백 (2cm)
$doc.PageSetup.TopMargin    = $word.CentimetersToPoints(2.0)
$doc.PageSetup.BottomMargin = $word.CentimetersToPoints(2.0)
$doc.PageSetup.LeftMargin   = $word.CentimetersToPoints(2.5)
$doc.PageSetup.RightMargin  = $word.CentimetersToPoints(2.5)

$selection = $word.Selection

# 기본 폰트 — 경기천년바탕 + 맑은 고딕 fallback (한국어 정상 표시 보장)
function Set-DefaultFont {
    param($sel)
    $sel.Font.Name = '맑은 고딕'
    $sel.Font.NameAscii = 'Malgun Gothic'
    $sel.Font.NameOther = 'Malgun Gothic'
    $sel.Font.NameFarEast = '맑은 고딕'
    $sel.Font.Size = 11
    $sel.Font.Bold = $false
    $sel.Font.Color = 0  # black
}

function Insert-Title {
    param($sel, $text)
    $sel.Font.NameFarEast = '맑은 고딕'
    $sel.Font.Name = 'Malgun Gothic'
    $sel.Font.Size = 18
    $sel.Font.Bold = $true
    $sel.ParagraphFormat.Alignment = 1  # center
    $sel.ParagraphFormat.SpaceAfter = 12
    $sel.TypeText($text)
    $sel.TypeParagraph()
    $sel.ParagraphFormat.Alignment = 0  # left
    Set-DefaultFont $sel
}

function Insert-Subtitle {
    param($sel, $text)
    $sel.Font.Size = 11
    $sel.Font.Bold = $false
    $sel.Font.Italic = $true
    $sel.Font.Color = 7368816  # gray
    $sel.ParagraphFormat.Alignment = 1
    $sel.ParagraphFormat.SpaceAfter = 18
    $sel.TypeText($text)
    $sel.TypeParagraph()
    Set-DefaultFont $sel
    $sel.Font.Italic = $false
    $sel.Font.Color = 0
    $sel.ParagraphFormat.Alignment = 0
}

function Insert-Heading1 {
    param($sel, $text)
    $sel.Font.NameFarEast = '맑은 고딕'
    $sel.Font.Name = 'Malgun Gothic'
    $sel.Font.Size = 14
    $sel.Font.Bold = $true
    $sel.Font.Color = 2236962  # navy-ish
    $sel.ParagraphFormat.SpaceBefore = 14
    $sel.ParagraphFormat.SpaceAfter = 8
    $sel.TypeText($text)
    $sel.TypeParagraph()
    Set-DefaultFont $sel
}

function Insert-Heading2 {
    param($sel, $text)
    $sel.Font.Size = 12
    $sel.Font.Bold = $true
    $sel.Font.Color = 4202515  # darker
    $sel.ParagraphFormat.SpaceBefore = 8
    $sel.ParagraphFormat.SpaceAfter = 4
    $sel.TypeText($text)
    $sel.TypeParagraph()
    Set-DefaultFont $sel
}

function Insert-Para {
    param($sel, $text)
    Set-DefaultFont $sel
    $sel.ParagraphFormat.LineSpacingRule = 5  # multiple
    $sel.ParagraphFormat.LineSpacing = 16
    $sel.ParagraphFormat.SpaceAfter = 4
    $sel.TypeText($text)
    $sel.TypeParagraph()
}

function Insert-Bullet {
    param($sel, $text)
    Set-DefaultFont $sel
    $sel.ParagraphFormat.LeftIndent = $word.CentimetersToPoints(0.5)
    $sel.TypeText("• $text")
    $sel.TypeParagraph()
    $sel.ParagraphFormat.LeftIndent = 0
}

function Insert-Table {
    param($sel, $headers, $rows)
    $rowCount = $rows.Count + 1
    $colCount = $headers.Count

    $range = $sel.Range
    $table = $doc.Tables.Add($range, $rowCount, $colCount)
    $table.Borders.Enable = $true
    $table.Borders.OutsideLineStyle = 1
    $table.Borders.InsideLineStyle = 1
    $table.Borders.OutsideColor = 0
    $table.Borders.InsideColor = 8421504  # gray

    # 헤더 행
    for ($c = 0; $c -lt $colCount; $c++) {
        $cell = $table.Cell(1, $c + 1)
        $cell.Range.Font.Name = 'Malgun Gothic'
        $cell.Range.Font.NameFarEast = '맑은 고딕'
        $cell.Range.Font.Size = 10.5
        $cell.Range.Font.Bold = $true
        $cell.Shading.BackgroundPatternColor = 15263976  # light gray
        $cell.Range.Text = $headers[$c]
    }
    # 데이터 행
    for ($r = 0; $r -lt $rows.Count; $r++) {
        for ($c = 0; $c -lt $colCount; $c++) {
            $cell = $table.Cell($r + 2, $c + 1)
            $cell.Range.Font.Name = 'Malgun Gothic'
            $cell.Range.Font.NameFarEast = '맑은 고딕'
            $cell.Range.Font.Size = 10
            $cell.Range.Font.Bold = $false
            $cell.Range.Text = [string]$rows[$r][$c]
        }
    }

    # 표 다음 줄로
    $sel.EndKey(6, 0) | Out-Null  # wdStory
    $sel.TypeParagraph()
    Set-DefaultFont $sel
}

# =====================================================
# 문서 본문 작성
# =====================================================

Insert-Title $selection 'LIRA vs 입법조사관 검토의견 비교 분석'
Insert-Subtitle $selection '경기도 생성형 인공지능 플랫폼 운영에 관한 조례안 검토 결과 비교  ·  LIRA v0.10.2 / 2026-06-01'

# ===== 섹션 1: 구조 비교 =====
Insert-Heading1 $selection '1. 구조 비교'

$structureHeaders = @('측면', 'LIRA (AI)', '입법조사관 (인간, 김 홍)')
$structureRows = @(
    @('구조', '5영역 고정 (인용·법체계·정합성·입법기술·양성평등)', '조문 단위 자유 구조'),
    @('분량', '매우 많음 (각 영역 평균 3-4건, N/A 다수 포함)', '압축적 (5개 ○ 항목, 핵심만)'),
    @('조문 식별 정확도', '일반론 위주 (예: "제2조제1호 및 제2호")', '매우 구체적 (예: "제6조제1항제1호의 [기밀성]")'),
    @('출처 표기', '풍부 (법제처 길라잡이 페이지·국가법령정보센터·발췌번호)', '본문 내 자연 인용'),
    @('수정안 표시', '별도 신구조문대비표 (이번엔 Gemini 503으로 실패)', '본문 내 인라인 (붉은색 자구·푸른색 제안)')
)
Insert-Table $selection $structureHeaders $structureRows

# ===== 섹션 2: 공통점 =====
Insert-Heading1 $selection '2. 공통점'
Insert-Para $selection '1. 법리적 결론 일치 — 두 의견서 모두 "본 조례안은 「지방자치법」 제13조에 따른 자치사무, 법리적으로 특별한 문제 없음" 일치'
Insert-Para $selection '2. 상위법 인용 — 둘 다 「인공지능 발전과 신뢰 기반 조성 등에 관한 기본법」 인용'
Insert-Para $selection '3. 표준 양식 — 자치법규 제명 / 관계 법령 / 검토 의견 / 검토 기간 / 담당자 5행 구조 동일'

# ===== 섹션 3: 차이점 — 입법조사관에만 있는 것 =====
Insert-Heading1 $selection '3. 차이점 (1) — 입법조사관에만 있는 것 (LIRA가 보강 필요)'

Insert-Heading2 $selection '3.1 분야 전문지식 활용'
Insert-Bullet $selection '"기밀성은 정보보안의 3원칙(기밀성·무결성·가용성) 요소" → 보안성과 묶어 제6조 통합 제안'
Insert-Bullet $selection '"공정성·투명성·책임성"은 「경기도 인공지능윤리 기반 조성에 관한 조례」 제5조제2항제2호의 기본계획 사항 → 중복 제거 제안'

Insert-Heading2 $selection '3.2 재정부담 검토 (지방자치법 §148)'
Insert-Bullet $selection '"안 제7조 내지 제8조 → 새로운 재정 부담 → 지자체장 의견 청취 필요"'
Insert-Bullet $selection 'LIRA의 5영역 체크리스트에는 이 항목이 누락되어 있음'

Insert-Heading2 $selection '3.3 행정 수단 선택 자유 보호'
Insert-Bullet $selection '"국산 인공지능 모델 및 반도체 활용하여" → 강행규정으로 행정 수단 과도 제한 → "적극 활용할 수 있다"로 임의규정화'

Insert-Heading2 $selection '3.4 임의규정 vs 강행규정 구분'
Insert-Bullet $selection '"행정망 보안 강화 등"을 임의적 규정으로 — 집행기관 재량 인정'

# ===== 섹션 4: 차이점 — LIRA에만 있는 것 =====
Insert-Heading1 $selection '4. 차이점 (2) — LIRA에만 있는 것 (LIRA의 강점)'

Insert-Bullet $selection '양성평등·차별 점검 — 입법조사관 의견서엔 없음 (별도 절차일 듯)'
Insert-Bullet $selection '5영역 표준 체크리스트 — 누구나 같은 기준 적용 가능'
Insert-Bullet $selection '출처 페이지·발췌번호 명시 — 검증 가능성 확보'
Insert-Bullet $selection '인용 법령 광범위 표기 — 입법조사관은 "「인공지능…기본법」 등"으로 처리, LIRA는 9개 법령 풀 인용'

# ===== 섹션 5: 영역별 평가 =====
Insert-Heading1 $selection '5. 누가 더 우수한가 — 영역별 평가'

$evalHeaders = @('평가 영역', '우수', '이유')
$evalRows = @(
    @('법리적 정확성',     '입법조사관', '분야 전문지식 (보안 3원칙·인공지능윤리 체계)'),
    @('구체성·조문 식별',  '입법조사관', '"제6조제1항제1호의 [기밀성]" 수준'),
    @('행정 수단 보호',    '입법조사관', '강행/임의 구분 + 행정재량 보호 관점'),
    @('재정부담 점검',     '입법조사관', '지방자치법 §148 명시, LIRA 누락'),
    @('체계성·표준화',     'LIRA',       '5영역 일관 적용'),
    @('양성평등 점검',     'LIRA',       '자동 점검 (입법조사관 의견서엔 없음)'),
    @('출처 검증 가능성',  'LIRA',       '페이지·발췌번호까지'),
    @('분량 효율성',       '입법조사관', 'LIRA는 N/A 항목 너무 많음 (전체 검토의견의 약 40%)'),
    @('신구조문대비표',    '입법조사관', '본문 인라인 자구수정 (LIRA는 이번엔 Gemini 503으로 실패)')
)
Insert-Table $selection $evalHeaders $evalRows

# ===== 섹션 6: 종합 평가 =====
Insert-Heading1 $selection '6. 종합 평가'
Insert-Para $selection '현재(v0.10.2) 입법조사관 의견이 LIRA 의견보다 종합적으로 우수합니다. 격차의 핵심은 다음과 같습니다.'
Insert-Bullet $selection '① 분야 전문지식 — 입법조사관은 보안·AI·재정·행정재량 등 여러 분야 지식을 자유롭게 조합'
Insert-Bullet $selection '② 구체성 — LIRA는 "제2조 정의 삭제 검토" 같은 일반 권고, 입법조사관은 "제6조제1항제1호의 [기밀성]은 보안 3원칙이므로 통합" 수준'
Insert-Bullet $selection '③ 분량 효율 — LIRA는 N/A 항목·중복 종합이 검토 의견 글자 수의 약 40%'
Insert-Para $selection ''
Insert-Para $selection '그러나 LIRA의 가치는 분명합니다.'
Insert-Bullet $selection '1차 점검의 누락 방지 (양성평등·차별·인용 무결성 자동)'
Insert-Bullet $selection '누구나 일관된 표준 적용'
Insert-Bullet $selection '출처 명시로 검증 가능'

# ===== 섹션 7: LIRA 개선 방향 =====
Insert-Heading1 $selection '7. LIRA가 입법조사관 수준에 근접하려면 (출품 후 개선 방향)'
Insert-Bullet $selection '① N/A 제거 + 문제 있는 것만 출력 (v0.10.3에서 처리 완료)'
Insert-Bullet $selection '② 재정부담 자동 점검 영역 추가 — 지방자치법 §148 자동 검출 ("비용", "예산", "지원" 키워드)'
Insert-Bullet $selection '③ 강행/임의 규정 구분 점검'
Insert-Bullet $selection '④ 분야별 chunks 보강 — 보안·AI·교육 등'
Insert-Bullet $selection '⑤ AI 모델 업그레이드 — Sonnet 4.6 vs Gemini 2.5 Flash 격차'

# ===== 푸터 =====
Insert-Para $selection ''
Insert-Para $selection ''
Set-DefaultFont $selection
$selection.Font.Size = 9
$selection.Font.Color = 7368816
$selection.Font.Italic = $true
$selection.ParagraphFormat.Alignment = 1
$selection.TypeText('LIRA — 의안 자동 검토 어시스턴트 · 경기도의회 법제과 의회 AI 바이브 코딩 워크숍 출품작')
$selection.TypeParagraph()
$selection.TypeText('작성: 이지영 법률자문관 · 2026-06-01')

# 저장
$doc.SaveAs([ref] $outPath, [ref] 16)  # 16 = wdFormatDocumentDefault (.docx)
$doc.Close()
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null

Write-Output "✓ 생성 완료: $outPath"
Get-Item $outPath | Select-Object Name, Length, LastWriteTime

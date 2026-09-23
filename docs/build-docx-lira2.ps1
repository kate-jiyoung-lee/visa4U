# LIRA2 vs 입법조사관 검토의견 비교 분석 — Word 문서 생성 (LIRA2 v0.11.0 기준 갱신·정리)
# Run: powershell -File build-docx-lira2.ps1
# 변경점(vs 구버전): 산출물 ODT->HWPX / 수정안·신구조문대비표 생성 제거(검토 전용으로 범위 한정)
#                    / 5영역->5대 입안원칙 등 9개 항목(양성평등은 용어점검에 포함) / v0.11.0·2026-06-02

$outPath = 'C:\Users\User\Documents\이지영\9. AI agent\클로드\LIRA\docs\LIRA-vs-입법조사관-비교분석.docx'

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

$doc = $word.Documents.Add()

$doc.PageSetup.TopMargin    = $word.CentimetersToPoints(2.0)
$doc.PageSetup.BottomMargin = $word.CentimetersToPoints(2.0)
$doc.PageSetup.LeftMargin   = $word.CentimetersToPoints(2.5)
$doc.PageSetup.RightMargin  = $word.CentimetersToPoints(2.5)

$selection = $word.Selection

function Set-DefaultFont {
    param($sel)
    $sel.Font.Name = '맑은 고딕'
    $sel.Font.NameAscii = 'Malgun Gothic'
    $sel.Font.NameOther = 'Malgun Gothic'
    $sel.Font.NameFarEast = '맑은 고딕'
    $sel.Font.Size = 11
    $sel.Font.Bold = $false
    $sel.Font.Italic = $false
    $sel.Font.Color = 0
}

function Insert-Title {
    param($sel, $text)
    $sel.Font.NameFarEast = '맑은 고딕'
    $sel.Font.Name = 'Malgun Gothic'
    $sel.Font.Size = 18
    $sel.Font.Bold = $true
    $sel.ParagraphFormat.Alignment = 1
    $sel.ParagraphFormat.SpaceAfter = 12
    $sel.TypeText($text)
    $sel.TypeParagraph()
    $sel.ParagraphFormat.Alignment = 0
    Set-DefaultFont $sel
}

function Insert-Subtitle {
    param($sel, $text)
    $sel.Font.Size = 11
    $sel.Font.Bold = $false
    $sel.Font.Italic = $true
    $sel.Font.Color = 7368816
    $sel.ParagraphFormat.Alignment = 1
    $sel.ParagraphFormat.SpaceAfter = 18
    $sel.TypeText($text)
    $sel.TypeParagraph()
    Set-DefaultFont $sel
    $sel.ParagraphFormat.Alignment = 0
}

function Insert-Heading1 {
    param($sel, $text)
    $sel.Font.NameFarEast = '맑은 고딕'
    $sel.Font.Name = 'Malgun Gothic'
    $sel.Font.Size = 14
    $sel.Font.Bold = $true
    $sel.Font.Color = 2236962
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
    $sel.Font.Color = 4202515
    $sel.ParagraphFormat.SpaceBefore = 8
    $sel.ParagraphFormat.SpaceAfter = 4
    $sel.TypeText($text)
    $sel.TypeParagraph()
    Set-DefaultFont $sel
}

function Insert-Para {
    param($sel, $text)
    Set-DefaultFont $sel
    $sel.ParagraphFormat.LineSpacingRule = 5
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
    $table.Borders.InsideColor = 8421504

    for ($c = 0; $c -lt $colCount; $c++) {
        $cell = $table.Cell(1, $c + 1)
        $cell.Range.Font.Name = 'Malgun Gothic'
        $cell.Range.Font.NameFarEast = '맑은 고딕'
        $cell.Range.Font.Size = 10.5
        $cell.Range.Font.Bold = $true
        $cell.Shading.BackgroundPatternColor = 15263976
        $cell.Range.Text = $headers[$c]
    }
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

    $sel.EndKey(6, 0) | Out-Null
    $sel.TypeParagraph()
    Set-DefaultFont $sel
}

# =====================================================
# 문서 본문
# =====================================================

Insert-Title $selection 'LIRA2 vs 입법조사관 검토의견 비교 분석'
Insert-Subtitle $selection '경기도 생성형 인공지능 플랫폼 운영에 관한 조례안 검토 결과 비교  ·  LIRA2 v0.11.0 / 2026-06-02'

# ===== 0. 한눈에 보기 =====
Insert-Heading1 $selection '0. 한눈에 보기'
Insert-Para $selection 'LIRA2는 의원 발의 조례안의 1차 검토를 보조하는 단일 HTML AI 도구다. 입법조사관(사람)의 최종 검토를 대체하는 것이 아니라, 누락 방지·표준 점검·근거 제시를 자동화해 검토의 출발점을 만든다. 같은 조례안을 두고 비교한 결과, 최종 법리 판단의 깊이는 입법조사관이 앞서지만, 표준화·양성평등 자동 점검·근거 검증 가능성은 LIRA2가 보완적 가치를 분명히 갖는다.'
Insert-Bullet $selection 'LIRA2 강점 — 9개 항목 표준 점검 / 양성평등·차별 자동 점검 / 모든 지적에 근거 명시 / 단일 HTML 보급성'
Insert-Bullet $selection '입법조사관 강점 — 분야 전문지식(보안·재정·행정재량) / 조문 단위 구체성 / 재정부담·강행임의 구분'
Insert-Bullet $selection '핵심 포지셔닝 — LIRA2 = 1차 검토 보조(범위 한정), 입법조사관 = 최종 검토. 경합이 아니라 단계 분담.'

# ===== 1. 구조 비교 =====
Insert-Heading1 $selection '1. 구조 비교'
$structureHeaders = @('측면', 'LIRA2 (AI)', '입법조사관 (인간, 김 홍)')
$structureRows = @(
    @('검토 구조', '5대 입안원칙 등 9개 표준 항목 (소관사무·법령우위·법률유보·권한경계·보조금·인용무결성·용어[양성평등·차별]·정합성·입법기술)', '조문 단위 자유 구조'),
    @('분량', '항목별 적합/검토필요/부적합 + 근거. 해당 없는 항목은 생략 (v0.11에서 N/A 정리)', '압축적 (핵심 ○ 항목 위주)'),
    @('조문 식별', '일반론~조문 단위 (예: "안 제2조제1호")', '매우 구체적 (예: "제6조제1항제1호의 [기밀성]")'),
    @('근거 표기', '모든 지적에 근거 명시 (법령 조문·판례·해석례·RAG 청크 출처)', '본문 내 자연 인용'),
    @('산출물', '검토의견서 HWPX (양식 A) — 한글에서 바로 열림', '검토의견서 (한글 문서)'),
    @('수정안', '생성하지 않음 — 1차 검토(누락 방지·표준 점검)로 범위 한정', '본문 내 인라인 자구 제안 (붉은색 자구·푸른색 제안)')
)
Insert-Table $selection $structureHeaders $structureRows

# ===== 2. 공통점 =====
Insert-Heading1 $selection '2. 공통점'
Insert-Para $selection '1. 법리적 결론 일치 — 두 의견서 모두 "본 조례안은 「지방자치법」 제13조에 따른 자치사무로, 법리적으로 특별한 문제 없음"으로 일치'
Insert-Para $selection '2. 상위법 인용 — 둘 다 「인공지능 발전과 신뢰 기반 조성 등에 관한 기본법」을 인용'
Insert-Para $selection '3. 표준 양식 — 자치법규 제명 / 관계 법령 / 검토 의견 / 검토 기간 / 담당자 5행 구조 동일'

# ===== 3. 차이점 (1) =====
Insert-Heading1 $selection '3. 차이점 (1) — 입법조사관에만 있는 것 (LIRA2 보강 필요)'
Insert-Heading2 $selection '3.1 분야 전문지식 활용'
Insert-Bullet $selection '"기밀성은 정보보안의 3원칙(기밀성·무결성·가용성) 요소" → 보안성과 묶어 제6조 통합 제안'
Insert-Bullet $selection '"공정성·투명성·책임성"은 「경기도 인공지능윤리 기반 조성에 관한 조례」 제5조제2항제2호의 기본계획 사항 → 중복 제거 제안'
Insert-Heading2 $selection '3.2 재정부담 검토 (지방자치법 제148조)'
Insert-Bullet $selection '"안 제7조 내지 제8조 → 새로운 재정 부담 → 지방자치단체장 의견 청취 필요"'
Insert-Bullet $selection 'LIRA2의 9개 항목에는 재정부담 자동 점검이 아직 없음 (보조금 점검만 포함)'
Insert-Heading2 $selection '3.3 행정 수단 선택 자유 보호'
Insert-Bullet $selection '"국산 인공지능 모델 및 반도체를 활용하여" → 강행규정으로 행정 수단을 과도 제한 → "적극 활용할 수 있다"로 임의규정화 제안'
Insert-Heading2 $selection '3.4 임의규정 vs 강행규정 구분'
Insert-Bullet $selection '"행정망 보안 강화 등"을 임의적 규정으로 — 집행기관 재량 인정'

# ===== 4. 차이점 (2) =====
Insert-Heading1 $selection '4. 차이점 (2) — LIRA2에만 있는 것 (LIRA2의 강점)'
Insert-Bullet $selection '양성평등·차별 점검 자동 (용어 점검 항목에 포함) — 입법조사관 의견서엔 없음 (별도 절차로 처리되는 듯)'
Insert-Bullet $selection '9개 항목 표준 점검 — 누구나 같은 기준을 일관 적용'
Insert-Bullet $selection '모든 지적에 근거 명시 — 법령 조문·판례·해석례·RAG 청크 출처까지, 검증 가능성 확보'
Insert-Bullet $selection '인용 법령 최신성·무결성 자동 경고 — 폐지·개정·조문이동 탐지. 입법조사관은 "「인공지능…기본법」 등"으로 처리, LIRA2는 관계 법령 풀 인용'
Insert-Bullet $selection '보급성·보안 — 단일 HTML 더블클릭 실행, 키는 메모리 전용(저장 안 함), 설치형 서버 불필요'

# ===== 5. 영역별 평가 =====
Insert-Heading1 $selection '5. 누가 더 우수한가 — 영역별 평가'
$evalHeaders = @('평가 영역', '우수', '이유')
$evalRows = @(
    @('법리적 깊이',       '입법조사관', '분야 전문지식 (보안 3원칙·인공지능윤리 체계)'),
    @('구체성·조문 식별',  '입법조사관', '"제6조제1항제1호의 [기밀성]" 수준'),
    @('행정 수단 보호',    '입법조사관', '강행/임의 구분 + 행정재량 보호 관점'),
    @('재정부담 점검',     '입법조사관', '지방자치법 제148조 명시, LIRA2 미포함'),
    @('체계성·표준화',     'LIRA2',      '9개 항목 일관 적용'),
    @('양성평등·차별 점검','LIRA2',      '용어 점검에서 자동 (입법조사관 의견서엔 없음)'),
    @('근거 검증 가능성',  'LIRA2',      '모든 지적에 법령·판례·해석례·청크 출처 명시'),
    @('인용 법령 무결성',  'LIRA2',      '최신성·폐지·조문이동 자동 경고'),
    @('보급성·접근성',     'LIRA2',      '단일 HTML 더블클릭, 누구나 본인 키로 즉시 사용'),
    @('분량 효율',         '대등',       'v0.11에서 N/A 항목 정리 — 해당 항목만 출력')
)
Insert-Table $selection $evalHeaders $evalRows

# ===== 6. 종합 평가 =====
Insert-Heading1 $selection '6. 종합 평가'
Insert-Para $selection '최종 검토의 법리적 깊이는 현재(v0.11.0)에도 입법조사관 의견이 앞선다. 격차의 핵심은 다음과 같다.'
Insert-Bullet $selection '① 분야 전문지식 — 입법조사관은 보안·AI·재정·행정재량 등 여러 분야 지식을 자유롭게 조합'
Insert-Bullet $selection '② 구체성 — LIRA2는 "안 제2조 정의 정비 검토" 같은 항목 권고, 입법조사관은 "제6조제1항제1호의 [기밀성]은 보안 3원칙이므로 통합" 수준'
Insert-Para $selection ''
Insert-Para $selection '다만 LIRA2는 입법조사관을 대체하려는 도구가 아니라 1차 검토 보조 도구이며, 그 범위에서의 가치는 분명하다.'
Insert-Bullet $selection '1차 점검의 누락 방지 (양성평등·차별·인용 무결성 자동)'
Insert-Bullet $selection '누구나 일관된 표준(9개 항목) 적용'
Insert-Bullet $selection '모든 지적에 근거를 달아 검증 가능'
Insert-Bullet $selection '수정안 생성을 의도적으로 제외해 검토 정확도와 디버깅 용이성에 집중 (직선형 파이프라인)'

# ===== 7. 개선 방향 =====
Insert-Heading1 $selection '7. LIRA2가 최종 검토 수준에 근접하려면 (출품 후 개선 방향)'
Insert-Bullet $selection '① N/A 제거 + 해당 항목만 출력 — v0.11에서 처리 완료'
Insert-Bullet $selection '② 재정부담 자동 점검 추가 — 지방자치법 제148조 자동 검출 ("비용"·"예산"·"지원" 키워드)'
Insert-Bullet $selection '③ 강행/임의 규정 구분 점검 추가'
Insert-Bullet $selection '④ 분야별 chunks 보강 — 보안·AI·교육·환경 등 전문 분야 발췌'
Insert-Bullet $selection '⑤ 분야 전문지식 기반 통합·중복 제거 제안 (입법조사관 수준 구체화)'

# ===== 푸터 =====
Insert-Para $selection ''
Insert-Para $selection ''
Set-DefaultFont $selection
$selection.Font.Size = 9
$selection.Font.Color = 7368816
$selection.Font.Italic = $true
$selection.ParagraphFormat.Alignment = 1
$selection.TypeText('LIRA2 — 조례안 자동검토 어시스턴트 · 경기도의회 법제과 의회 AI 바이브 코딩 워크숍 출품작')
$selection.TypeParagraph()
$selection.TypeText('작성: 이지영 법률자문관 · 2026-06-02')

$doc.SaveAs([ref] $outPath, [ref] 16)
$doc.Close()
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null

Write-Output "생성 완료: $outPath"
Get-Item $outPath | Select-Object Name, Length, LastWriteTime

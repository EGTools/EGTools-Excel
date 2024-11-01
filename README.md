# EGTools-Excel
<br>**EG Tools는 Excel 사용에 도움을 주는 여러 기능과 함수를 제공하는 엑셀추가기능입니다.**
<br>**Excel 2019, 2021, 2024, Microsoft 365에 추가된 신규 함수를 사용할 수 있도록 지원 합니다.**
<br>_Mac용 Excel에서는 VBA코드가 달라 사용할 수 없습니다_

<br>최종버전은 Release페이지에서 다운로드 하세요.
<br>신규 버전도 동일한 이름으로 계속 업데이트 되므로 다운로드 후에 동일한 폴더에 넣으면 됩니다.
<br>https://github.com/EGTools/EGTools-Excel/releases/latest 
<br>
<br>공식배포하기 전에 개선/보완된 Pre-Release가 임시로 올라오는 경우도 있습니다.
<br>기능에 대한 점검이 완료되지 않아 일부 오류가 있을 수 있으므로 사용에 주의를 바랍니다.
<br>https://github.com/EGTools/EGTools-Excel/releases
<br>
<br>문의사항은 Naver Cafe를 이용 바랍니다.
<br>https://cafe.naver.com/egtools
<br>

> [!NOTE]
> v4.5.5부터 ExcelDna-Intellisense Add-in 지원 기능을 삭제하였습니다.

<br>

<p>
<p>
   
   
# 설치방법
<p> Excel 추가기능 설치 방법은 여기를 참고하세요.
<br>https://cafe.naver.com/egtools/14
<br>
<br>
   
# Ribbon Menu 기능
## [보이는 셀만](https://cafe.naver.com/egtools/23)
- 보이는 셀 복사 : 화면에 보이는 셀만 선택해서 복사하는 기능
- 전체 복사 : 화면에 보이는 셀과 보이지 않는 셀 모두 복사하는 기능
   상기 2개 기능은 아래 보이는 셀에 붙여넣기 하기 전에 먼저 수행해야 합니다.
- 모두 붙여넣기 : 셀서식과 값을 모두 보이는 셀에만 붙여넣기
- 값만 붙여넣기 : 값만 보이는 셀에 붙여넣기
- 수식 붙여넣기 : 수식만 보이는 셀에 붙여넣기

## [병합/분할](https://cafe.naver.com/egtools/24)
- 내용 병합 : 셀을 병합하면서 내용까지 모두 같이 병합함
- 열끼리 병합 : 선택된 영역에 대해서 열별로 '내용 병합'을 한 번에 수행함
- 행끼리 병합 : 선택된 영역에 대해서 행별로 '내용 병합'을 한 번에 수행함
<br>상기 2개의 기능은 내용병합을 할 때 텍스트 서식을 유지하는 것과 서식을 제거하는 것이 있음
- 연속값 병합 : 열방향(아래쪽)으로 같은 값이 연속될 때 셀을 자동으로 병합함
- 행 나누기 : 줄바꿈이 있는 셀 내용을 여러 행(줄)으로 나누어 줌
- 열 나누기 : 구분자가 있는 셀 내용을 여러 열(칸)로 나누어 줌
<br>상기 2개의 기능은 텍스트 서식 그대로 진행하는 것과 서식없이 진행하는 것이 있음
- 나누고 채우기 : 병합된 셀을 분리하면서 모두 같은 내용으로 복사해 줌

## [사진/그림](https://cafe.naver.com/egtools/25)
- 선택 삽입 : 선택한 셀에 PC에 저장된 사진/그림을 하나 삽입함
- 폴더 삽입 : 파일명이 셀 내용에 입력된 경우 지정하는 폴더에서 해당 사진/그림을 일괄 삽입함
- 선택 맞춤 : 선택한 사진/그림을 셀에 자동 맞춤 
- 모두 맞춤 : 현재 시트의 모든 사진/그림을 셀에 자동 맞춤
- 전체 저장 : 현재 시트의 모든 사진/그림을 지정한 폴더로 모두 일괄 저장함

## [달력/일정표](https://cafe.naver.com/egtools/27)
- 연간 달력 만들기 : 지정하는 연도의 연간 달력 시트를 삽입 (대한민국의 공휴일 표시)
- 월간 일정표 만들기 : 지정하는 월의 월간 일정표 시트를 삽입(대한민국의 공휴일 및 음력 표시)
- 주간 일정표 만들기 : 지정하는 날짜가 포함된 주간 일정표 시트를 삽입(대한민국의 공휴일, 음력, 주요행사, 시간일정 표시)
- 일일 일정표 만들기 : 지정하는 날짜의 일일 일정표 시트를 삽입(대한민국의 공휴일, 음력, 주요 업무, 시간일정, 업무메모 표시)

## [다단계선택](https://cafe.naver.com/egtools/28)
- 다단계 선택 기준 : 다단계 유효성 검사의 목록(Dropdown List)을 생성
- 다단계 선택 적용 : 선택한 셀에 다단계 유효성 검사를 적용
- 다단계 선택 삭제 : 필요없는 다단계 유효성 검사를 삭제

## [표기능](https://cafe.naver.com/egtools/29)
- 피벗 해제 : 가로세로 Cross Tab이나 Pivot으로 된 표를 해제하여 일반 Data형 표로 만들어 줌
- Cross Tab : 일반 Data형 표를 Cross Tab으로 만들어 집계해 줌
- Table 집계 : 동일 양식의 여러 시트의 자료를 하나로 모아 줌

> [!NOTE]
> EGBarcode.xlam으로 분리되었습니다.
>## 바코드
>- 1D 바코드 :
>- 2D 바코드 :
>- GS1 바코드 :

## 기타기능
- 그림으로 저장 : 선택한 영역을 그림파일로 저장
- 오류 제거 : 현재 시트에서 오류인 셀에 대해서 IFERROR()함수를 자동으로 추가하여 오류가 보이지 않도록 함
- UDF 제거 : 본 EG Tools의 UDF를 사용한 경우 다른 PC로 보낼 때 UDF를 제거하여 값으로 변경할 수 있음
- Style삭제 : 셀스타일이 많을 경우 사용하지 않는 Style이나 Built-in이 아닌 Style을 일괄 삭제함
- [이름 삭제](https://cafe.naver.com/egtools/31) : 보이지 않는 명명된 이름과 유효하지 않은 이름으로 일괄 삭제함
- 빈셀 정리 : 현재 시트에서 길이가 0인 문자열을 제거
- [양쪽 공백 제거](https://cafe.naver.com/egtools/274) : 현재 시트의 모든 셀에서 시작부분과 끝부분의 공백문자를 제거
- 메모 정돈 : 현재 시트에서 모든 메모를 삽입된 셀 바로 옆으로 위치를 재정리
- [메일 머지](https://cafe.naver.com/egtools/32) : 목록과 양식을 이용하여 자동으로 시트나 파일을 만들고, 인쇄 또는 이메일 발송
- [모양 뽑기](https://cafe.naver.com/egtools/33) : 셀바탕색으로 그린 블럭들의 외곽을 따라서 자유도형을 자동 생성

## EGTools관련
- [EGTools 연결 문서 수정](https://cafe.naver.com/egtools/290) : 다른 PC에서 EGTools를 사용한 경로를 수정
- [EGTools 배열 함수 수정](https://cafe.naver.com/egtools/290) : Excel 버전에 따른 EGTools의 배열 및 함수를 수정
- [설명서 보기](https://cafe.naver.com/egtools/8) : 간단한 EG Tools의 설명서를 보여줌
- Version : 현재 버전을 보여주며, 배포된 Update가 있을 경우 링크를 보여줌
- EGTools 사용중지 : EGTools 추가기능을 잠시 사용중지하거나, 사용중단하고 파일을 삭제

<br>
<br>

> [!WARNING]
> UDF를 과다하게 사용할 경우 Excel계산이 많이 느려질 수 있으므로 작업후 값으로 변경하는 것이 좋습니다.

# EXCEL 신규 함수에 대한 호환 UDF
하위버전의 Excel에서 상위 버전에 추가된 함수를 사용할 수 있습니다.<br>
<br>

## Microsoft 365 신규 예정 함수 호환
- [REGEXTEST 함수](https://cafe.naver.com/egtools/161) : 텍스트 일부가 정규 표현식과 일치하는지 확인합니다
- [REGEXEXTRACT 함수](https://cafe.naver.com/egtools/162) : 정규 표현식에 따라 일치하는 하위 문자열을 추출합니다
- [REGEXREPLACE 함수](https://cafe.naver.com/egtools/163) : 정규 표현식을 사용하여 문자열의 일부를 다른 문자열로 바꿉니다
- [TRANSLATE 함수](https://cafe.naver.com/egtools/165) : 문자열을 지정하는 언어로 번역합니다
- [DETECTLANGUAGE 함수](https://cafe.naver.com/egtools/164) : 문자열의 언어를 자동으로 판별합니다
- [TRIMRANGE 함수](https://cafe.naver.com/egtools/231) : 범위/배열에서 빈행과 빈열을 제거합니다

## Microsoft 365 신규 함수 호환
- [GROUPBY 함수](https://cafe.naver.com/egtools/159) : 한 축을 따라 그룹화하고 연결된 값을 집계할 수 있습니다
- [PIVOTBY 함수](https://cafe.naver.com/egtools/160) : 두 축을 따라 그룹화하고 연결된 값을 집계할 수 있습니다
- [PERCENTOF 함수](https://cafe.naver.com/egtools/270) : 해당값을 전체 값으로 나눈 백분율을 계산합니다
 
## Excel 2024 신규 함수 호환
- [TEXTSPLIT 함수](https://cafe.naver.com/egtools/131) : 열과 행 구분 기호를 사용하여 텍스트 문자열을 분할합니다
- [TEXTAFTER 함수](https://cafe.naver.com/egtools/132) : 지정된 문자 또는 문자열 뒤에 나타나는 텍스트를 반환합니다
- [TEXTBEFORE 함수](https://cafe.naver.com/egtools/133) : 지정된 문자 또는 문자열 앞에 나타나는 텍스트를 반환합니다
- [VSTACK 함수](https://cafe.naver.com/egtools/134) : 배열을 세로방향으로 순서대로 추가하여 더 큰 배열을 반환합니다
- [HSTACK 함수](https://cafe.naver.com/egtools/135) : 배열을 가로방향으로 순서대로 추가하여 더 큰 배열을 반환합니다
- [TOCOL 함수](https://cafe.naver.com/egtools/136) : 단일 열의 배열을 반환합니다
- [TOROW 함수](https://cafe.naver.com/egtools/137) : 단일 행의 배열을 반환합니다
- [WRAPCOLS 함수](https://cafe.naver.com/egtools/138) : 지정된 수의 요소 뒤에 있는 열별로 제공된 값 행 또는 열을 래핑하여 새 배열을 구성합니다
- [WRAPROWS 함수](https://cafe.naver.com/egtools/139) : 지정된 수의 요소 뒤에 있는 열별로 제공된 값 행 또는 열을 래핑하여 새 배열을 구성합니다
- [CHOOSECOLS 함수](https://cafe.naver.com/egtools/140) : 배열이나 범위에서 지정된 열 순서대로 재배열한 배열을 반환합니다
- [CHOOSEROWS 함수](https://cafe.naver.com/egtools/141) : 배열이나 범위에서 지정된 행 순서대로 재배열한 배열을 반환합니다
- [TAKE 함수](https://cafe.naver.com/egtools/142) : 배열의 시작 또는 끝에서 지정된 수의 연속 행 또는 열을 반환합니다
- [DROP 함수](https://cafe.naver.com/egtools/143) : 배열의 시작 또는 끝에서 지정된 수의 행 또는 열을 제외합니다
- [EXPAND 함수](https://cafe.naver.com/egtools/144) : 배열을 확장하거나 지정된 행 및 열 차원으로 채웁니다
- [VALUETOTEXT 함수](https://cafe.naver.com/egtools/145)  : 텍스트 값을 변경하지 않고 전달하며 텍스트가 아닌 값을 텍스트로 변환합니다
- [ARRAYTOTEXT 함수](https://cafe.naver.com/egtools/146)  : 배열내의 텍스트 값을 변경하지 않고 전달하며 텍스트가 아닌 값을 텍스트로 변환합니다
- [IMAGE 함수](https://cafe.naver.com/egtools/147) : 인터넷에 올려진 이미지 URL이나 컴퓨터에 저장된 파일명으로 이미지를 삽입합니다

## Excel 2021 신규 함수 호환
- [XMATCH 함수](https://cafe.naver.com/egtools/73) : 배열 또는 셀 범위에서 지정된 항목을 검색한 다음 항목의 상대 위치를 반환합니다
- [XLOOKUP 함수](https://cafe.naver.com/egtools/74) : 테이블 또는 행별 범위에서 항목을 찾습니다
- [XFILTER 함수](https://cafe.naver.com/egtools/75) : 직접 정의한 조건을 바탕으로 일정 범위의 데이터를 필터링합니다
- [XSORT 함수](https://cafe.naver.com/egtools/76) : 범위 또는 배열의 내용을 정렬합니다
- [SORTBY 함수](https://cafe.naver.com/egtools/77) : 대응되는 범위 또는 배열의 값을 기준으로 범위 또는 배열의 내용을 정렬합니다
- [UNIQUE 함수](https://cafe.naver.com/egtools/78) : 목록 또는 범위에서 고유 값의 목록을 반환합니다
- [SEQUENCE 함수](https://cafe.naver.com/egtools/79) : 1, 2, 3, 4와 같이 일련의 연속된 숫자 목록을 생성합니다
- [RANDARRAY 함수](https://cafe.naver.com/egtools/80) : 임의의 숫자 배열을 생성합니다
- [XLET 함수](https://cafe.naver.com/egtools/81) : 계산 결과에 이름을 할당합니다. 중간 계산, 값을 저장하거나 이름을 정의할 수 있습니다

## Excel 2019 신규 함수 호환
- [IFS 함수](https://cafe.naver.com/egtools/38) : 하나 이상의 조건이 충족되는지 확인하고 첫 번째 TRUE 조건에 해당하는 값을 반환합니다
- [MINIFS 함수](https://cafe.naver.com/egtools/39) : 하나 이상의 조건이 모두 충족되는 최소값을 반환합니다
- [MAXIFS 함수](https://cafe.naver.com/egtools/40) : 하나 이상의 조건이 모두 충족되는 최대값을 반환합니다
- [CONCAT 함수](https://cafe.naver.com/egtools/41) : 여러 범위 및/또는 문자열의 텍스트를 결합합니다
- [TEXTJOIN 함수](https://cafe.naver.com/egtools/42) : 여러 범위 및/또는 문자열의 텍스트를 결합하며 구분기호를 포함합니다.
- [SWITCH 함수](https://cafe.naver.com/egtools/43) : 하나의 수식 또는 값을 평가하고 첫 번째 일치하는 값에 해당하는 결과를 반환합니다

## Excel 2013 신규 함수 호환
- FORMULATEXT 함수 : 지정한 셀에 입력된 함수를 보여줍니다
- ENCODEURL 함수 : 값을 브라우저에서 사용할 수 있도록 Encoding합니다
- IFNA 함수 : #N/A 오류일 때 지정한 값으로 변경합니다
- UNICODE 함수 : 첫번째문자의 유니코드 코드값을 반환합니다
- UNICHAR 함수 : 지정한 코드값의 유니코드 문자를 반환합니다

<br>
<br>

# Goolgle 스프레드시트 함수 호환
- [IMPORTRANGE 함수](https://cafe.naver.com/egtools/153) : Google Sheets의 지정하는 범위를 가져옵니다
- [IMPORTHTML 함수](https://cafe.naver.com/egtools/154) : 인터넷 페이지의 표나 목록을 지정하여 자료를 가져옵니다
- [IMPORTDATA 함수](https://cafe.naver.com/egtools/155) : RSS나 ATOM feed 정보를 가져옵니다
- [IMPORTFEED 함수](https://cafe.naver.com/egtools/156) : csv나 tsv 파일의 데이터를 읽어 옵니다
- [GOOGLETRANSLATE 함수](https://cafe.naver.com/egtools/130) : Google의 번역 서비스를 이용한 번역을 제공합니다
- [COUNTUNIQUE 함수](https://cafe.naver.com/egtools/128) : 지정된 값과 범위 목록에서 고유 값의 개수를 셉니다
- [COUNTUNIQUEIFS 함수](https://cafe.naver.com/egtools/129) : 지정된 범위에서 여러 조건에 부합하는 고유 값의 갯수를 셉니다
- [QUERY 함수](https://cafe.naver.com/egtools/127) : 데이터에서 ADODB에 사용하는 언어로 검색을 실행합니다
- [EPOCHTODATE 함수](https://cafe.naver.com/egtools/126) : Unix epoch 타임스탬프를 협정 세계시(UTC) 기준의 날짜 및 시간으로 변환합니다
- [ISBETWEEN 함수](https://cafe.naver.com/egtools/125) : 제공된 값이 다른 두 값 사이에 있는지 확인합니다
- [ISEMAIL 함수](https://cafe.naver.com/egtools/124) : 최상위 도메인을 기준으로 유효한 이메일 주소인지 확인합니다
- [ISURL 함수](https://cafe.naver.com/egtools/123) : 유효한 URL 값인지 확인합니다

<br>
<br>
   
# EGTools 전용 UDF
## 검색 함수
- [MVLOOKUP 함수](https://cafe.naver.com/egtools/107) : Excel의 VLOOKUP함수를 다량으로 실행한 결과를 출력합니다. (mass VLOOKUP)
- [MXLOOKUP 함수](https://cafe.naver.com/egtools/211) : Excel의 XLOOKUP함수를 다량으로 실행한 결과를 출력합니다. (mass XLOOKUP)
- [ILOOKUP 함수](https://cafe.naver.com/egtools/51) : 검색 범위에서 찾는 값 중에서 지정하는 순번에 해당하는 이미지를 가져옵니다.(Image LookUp)
- [NLOOKUP 함수](https://cafe.naver.com/egtools/49) : 검색 범위에서 찾는 값과 일치하는 목록에서 지정하는 순번의 값을 찾습니다
- [MATCHJOIN 함수](https://cafe.naver.com/egtools/48) : 찾는 값이나 조건에 해당하는 내용을 연결 문자를 이용하여 연결
- [COMPARELIST 함수](https://cafe.naver.com/egtools/64) : 두개의 목록에 대해서 개별 값을 비교한 결과 목록을 나열합니다
- [COMPARELISTM 함수](https://cafe.naver.com/egtools/292) : 두개의 목록에 대해서 행별 값을 비교한 결과 목록을 나열합니다
- [SAMPLE 함수](https://cafe.naver.com/egtools/87) : 지정하는 대상 범위에서 무작위 샘플링 추출하여 목록을 생성합니다

## 문자열 함수
- [STREXT 함수](https://cafe.naver.com/egtools/47) : 숫자, 영문, 영숫자, 한글, 일본어, 한자/중국어를 추출하거나 제거합니다
- [MATCHJOIN 함수](https://cafe.naver.com/egtools/48) : 일치하는 내용에 대응하는 결과 값들을 '연결자'를 이용하여 하나의 문자열을 작성합니다
- [TEXTPICK 함수](https://cafe.naver.com/egtools/50) :  문자열을 특정 구분자를 기준으로 분리하여 원하는 순번의 값을 추출합니다
- [TEXTBETWEEN 함수](https://cafe.naver.com/egtools/59) : 지정하는 2개의 문자열 사이에 있는 내용을 추출합니다
- [TEXTJOINIF 함수](https://cafe.naver.com/egtools/86) : 조건에 만족하는 검색범위의 값을 하나의 문자열로 연결합니다
- [CLEANB 함수](https://cafe.naver.com/egtools/105) : 인쇄할 수 없는 문자코드를 제거합니다
- [TRIMENDS 함수](https://cafe.naver.com/egtools/120) : 양쪽 끝의 공백만 제거합니다
   
## 계산 및 집계 함수
- [COUNTER 함수](https://cafe.naver.com/egtools/17) : 범위나 배열 데이터에서 각 요소별 빈도수를 나열합니다
- [EVAL 함수](https://cafe.naver.com/egtools/46) : 주어진 문자열의 Excel에서의 계산 결과를 산출합니다
- [IFVISIBLE 함수](https://cafe.naver.com/egtools/110) : 보이는 셀에 대해서만 각종 통계 함수를 적용합니다
- [AGGREGATEC 함수](https://cafe.naver.com/egtools/83) : 목록 또는 데이터베이스의 숨겨진 셀을 모두 제외할 수 있는 집계를 반환합니다
<br>

> [!NOTE]
> EGqcF.xlam으로 분리되었습니다.
>## QC샘플링 함수 
>- [SAMPLINGSIZE 함수](https://cafe.naver.com/egtools/93) : LOT크기와 AQL, 검사방법에 따라 검사할 시료수를 구합니다
>- [SAMPLINGAC 함수](https://cafe.naver.com/egtools/95) : LOT크기와 AQL, 검사수준에 따라 검사할 합격판정 최대 불량수를 구합니다
>- [SAMPLINGRE 함수](https://cafe.naver.com/egtools/96) : LOT크기와 AQL, 검사수준에 따라 검사할 불합격판정 최소 불량수를 구합니다.
>- [SAMPLINGLABEL 함수](https://cafe.naver.com/egtools/94) : LOT크기와 검사수준에 따른 시료문자를 구합니다
​<br>

> [!NOTE]
> EGBarcode.xlam으로 분리되었습니다.
>## 바코드함수
>- [BARCODE 함수](https://cafe.naver.com/egtools/90) : 1D 및 2D 바코드 이미지를 생성합니다 (11종)
>- [QRCODE 함수](https://cafe.naver.com/egtools/92) : QRCODE 바코드 이미지를 생성합니다
>- [CODE128 함수](https://cafe.naver.com/egtools/91) : CODE128 바코드 이미지를 생성합니다
<br>

## 날짜시간 함수
- [KOREANHOLIDAYS 함수](https://cafe.naver.com/egtools/20) : 대한민국의 공휴일을 나열하는 함수입니다
- [TOLUNAR 함수](https://cafe.naver.com/egtools/60) : 양력날짜를 음력날짜로 변환합니다
- [TOSOLAR 함수](https://cafe.naver.com/egtools/61) : 음력날짜를 양력날짜로 변환합니다
- [DATETIME 함수](https://cafe.naver.com/egtools/67) : 한글, 한자가 포함된 날짜와 시간형 문자열을 날짜와 시간으로 변환합니다
- [MONTHBYWEEK 함수](https://cafe.naver.com/egtools/57) : 지정하는 요일을 기준으로 정한 해당주차의 월을 확인합니다
- [WEEKNUMOFMONTH 함수](https://cafe.naver.com/egtools/58) : 지정하는 요일을 기준으로 정한 해당주차의 월내의 주차수를 구합니다
- [JULIANDAY 함수](https://cafe.naver.com/egtools/102) : 율리우스적일 (Julian Day Number)을 계산합니다
- [JDTODATE 함수](https://cafe.naver.com/egtools/103) : 율리우스적일 (Julian Day Number)을 양력 날짜로 변환합니다
​

## 색상 함수
- [TEXTJOINIFCOLOR 함수](https://cafe.naver.com/egtools/84) : 대상범위의 보이는 색이 기준셀과 같은 색이면 문자열을 구분자를 이용하여 연결합니다
- [DISPLAYCOLOR 함수](https://cafe.naver.com/egtools/70) : 대상셀의 보이는 색으로 바탕색/글자색의 색번호를 반환합니다
- [SUMIFCOLOR 함수](https://cafe.naver.com/egtools/69) : 대상범위의 보이는 색이 기준셀과 같은 바탕색/글자색이면 숫자를 더합니다
- [COUNTIFCOLOR 함수](https://cafe.naver.com/egtools/68) : 대상범위의 보이는 색이 기준셀과 같은 바탕색/글자색이면 숫자를 셉니다
- [RGB 함수](https://cafe.naver.com/egtools/148) : Red, Green, Blue 색상값으로 True Color 색상값을 계산합니다
- [TORGB 함수](https://cafe.naver.com/egtools/149) : True Color 색상값을 Red, Green, Blue 색상값으로 분해합니다
​

## 변환 함수
- [UNPIVOT 함수](https://cafe.naver.com/egtools/303) : 피벗테이블이나 크로스탭을 일반 데이터 표로 변환합니다 
- [JSONPARSE 함수](https://cafe.naver.com/egtools/152) : JSON 문자열의 경로명과 일치하는 값을 검색합니다
- [JSONTOARRAY 함수](https://cafe.naver.com/egtools/151) : JSON 문자열의 경로명 각 단계와 값을 배열로 구성합니다
- [JSONPAIR 함수](https://cafe.naver.com/egtools/150) : JSON 문자열을 경로명과 값의 쌍으로 나열합니다
- [EXRATE 함수](https://cafe.naver.com/egtools/113) : 대한민국 원화의 외환 환율을 조회합니다
- [EXPLODE 함수](https://cafe.naver.com/egtools/108) : 지정하는 열에 대해서 구분자를 기준으로 분해하여 나열합니다
- [TEXTNUMSORT 함수](https://cafe.naver.com/egtools/108) : 문자와 숫자가 섞여 있는 데이터를 정렬할 때, 숫자가 숫자로 정렬하도록 합니다
- [PAPAGOTRANSLATE 함수](https://cafe.naver.com/egtools/104) : 네이버의 Papago API를 이용한 번역을 제공합니다
- [RZ 함수](https://cafe.naver.com/egtools/88) : 0이나 빈셀, 오류를 빈문자열("")로 변환합니다. (Remove Zero)
- IFERRORX 함수 : 
- [HANTONUMBER 함수](https://cafe.naver.com/egtools/62) : 한글이나 한자 및 갖은한자로 입력된 숫자를 아라비아 숫자로 변환합니다
- [US32TODEC 함수](https://cafe.naver.com/egtools/65) : 미국 채권시장의 32분수 표시형식을 십진수로 변환합니다
- [DECTOUS32 함수](https://cafe.naver.com/egtools/66) : 일반 숫자를 미국 채권시장의 32분수 표시형식으로 변환합니다
​

## 대한민국 공개API 함수
- [SEARCHADDRESS 함수](https://cafe.naver.com/egtools/261) : 도로명 주소 검색을 통하여 정보를 조회합니다
- [ZIPCODE 함수](https://cafe.naver.com/egtools/106) : 도로명 주소나 건물명 등의 키워드로 우편번호 및 도로명주소, 지번주소를 검색합니다
- [GEOPOINT 함수](https://cafe.naver.com/egtools/115) : 도로명 주소를 기준으로 해당 주소의 지도 좌표를 확인합니다
- [GEOCONVERT 함수](https://cafe.naver.com/egtools/117) : 지도 좌표를 다른 좌표계로 변환합니다
- [GEODISTANCE 함수](https://cafe.naver.com/egtools/118) : 지도 좌표로 거리를 개략적으로 계산합니다
- [OILPRICE 함수](https://cafe.naver.com/egtools/114) : [OPINET](https://www.opinet.co.kr/user/main/mainView.do)에서 제공하는 API를 이용하여 지역별 유종별 평균유가를 조회합니다
- [GASSTATION 함수](https://cafe.naver.com/egtools/116) : [OPINET](https://www.opinet.co.kr/user/main/mainView.do)에서 제공하는 API를 이용하여 주변 유가를 검색합니다
- [BRNSTATUS 함수](https://cafe.naver.com/egtools/119) : 국세청의 API를 이용하여 사업자등록번호의 현재 상태를 조회합니다


## 기타 함수
- [SHEETSLIST 함수](https://cafe.naver.com/egtools/112) : 현재 Excel 파일의 시트 목록을 작성합니다
- IPINFO 함수 : IP Address 기본정보
- [DIRFOLDER 함수](https://cafe.naver.com/egtools/109) : 지정한 폴더의 파일 목록을 출력합니다
- [IMPORTURL 함수](https://cafe.naver.com/egtools/168) : 인터넷 페이지의 소스를 표시합니다



<br>



# 감사인사
기능에 대한 조언과 테스트를 통해 오류를 잡아 주시는 분들께 항상 감사 드립니다.<br>
<br>



# 사용권한
본 파일은 개인, 회사, 관공서 등 누구나 무료로 사용할 수 있습니다.<br>
본 파일을 사용함으로써 발생하는 모든 책임은 사용자에게 있습니다.<br>
만약 이에 동의하지 않는다면, 사용을 중단하고 파일을 삭제 바랍니다.<br>



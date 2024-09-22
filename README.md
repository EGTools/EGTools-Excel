# EGTools-Excel
<p><b>Excel2010부터 사용할 수 있는 Excel 추가기능(Add-in)입니다.</b><br>
   <b>Mac용 Excel에서는 VBA코드가 달라 사용할 수 없습니다.</b></p>
<br>
<p>최종버전은 Release페이지에서 다운로드 하세요.
<br>신규 버전도 동일한 이름으로 계속 업데이트 되므로 다운로드 후에 동일한 폴더에 넣으면 됩니다.
<br>https://github.com/EGTools/EGTools-Excel/releases/latest 
<p>
<p>공식배포하기 전에 개선/보완된 Pre-Release가 임시로 올라오는 경우도 있습니다.
<br>기능에 대한 점검이 완료되지 않아 일부 오류가 있을 수 있으므로 사용에 주의를 바랍니다.
<br>https://github.com/EGTools/EGTools-Excel/releases
<br>
<p>
문의사항은 Naver Cafe를 이용 바랍니다.
<br>https://cafe.naver.com/egtools
<br>

<p>v4.5.5에서 ExcelDna-Intellisense Add-in 지원 기능을 삭제하였습니다.<br>
<br>

<p>
<p>
   
   
# 설치방법
<p> Excel 추가기능 설치 방법은 여기를 참고하세요.
<br>https://cafe.naver.com/egtools/14
<br>
<br>
<p>
   

   
# 기능설명
<br>
   
## Ribbon Menu 기능
<ul>
   <li> 보이는 셀 복사 : 화면에 보이는 셀만 선택해서 복사하는 기능</li>
   <li> 전체 복사 : 화면에 보이는 셀과 보이지 않는 셀 모두 복사하는 기능</li>
   상기 2개 기능은 아래 보이는 셀에 붙여넣기 하기 전에 먼저 수행해야 합니다.
   <li> 모두 붙여넣기 : 셀서식과 값을 모두 보이는 셀에만 붙여넣기</li>
   <li> 값만 붙여넣기 : 값만 보이는 셀에 붙여넣기</li>
   <li> 수식 붙여넣기 : 수식만 보이는 셀에 붙여넣기</li>
   <br>
   <li> 내용 병합 : 셀을 병합하면서 내용까지 모두 같이 병합함</li>
   <li> 열끼리 병합 : 선택된 영역에 대해서 열별로 '내용 병합'을 한 번에 수행함</li>
   <li> 행끼리 병합 : 선택된 영역에 대해서 행별로 '내용 병합'을 한 번에 수행함</li>
   상기 기능은 내용병합을 할 때 텍스트 서식을 유지하는 것과 서식을 제거하는 것이 있음
   <li> 연속값 병합 : 열방향(아래쪽)으로 같은 값이 연속될 때 셀을 자동으로 병합함</li>
   <li> 행 나누기 : 줄바꿈이 있는 셀 내용을 여러 행(줄)으로 나누어 줌</li>
   <li> 열 나누기 : 구분자가 있는 셀 내용을 여러 열(칸)로 나누어 줌</li>
   상기 2개의 기능은 텍스트 서식 그대로 진행하는 것과 서식없이 진행하는 것이 있음
   <li> 나누고 채우기 : 병합된 셀을 분리하면서 모두 같은 내용으로 복사해 줌</li>  
   <br>
   <li> 선택 삽입 : 선택한 셀에 PC에 저장된 사진/그림을 하나 삽입함</li>
   <li> 폴더 삽입 : 파일명이 셀 내용에 입력된 경우 지정하는 폴더에서 해당 사진/그림을 일괄 삽입함</li>
   <li> 선택 맞춤 : 선택한 사진/그림을 셀에 자동 맞춤 </li>
   <li> 모두 맞춤 : 현재 시트의 모든 사진/그림을 셀에 자동 맞춤</li>
   <li> 전체 저장 : 현재 시트의 모든 사진/그림을 지정한 폴더로 모두 일괄 저장함</li>
   <br>
   <li> 연간 달력 만들기 : 지정하는 연도의 연간 달력 시트를 삽입 (공휴일 표시)</li>
   <li> 월간 일정표 만들기 : 지정하는 월의 월간 일정표 시트를 삽입(공휴일, 음력 표시)</li>
   <li> 주간 일정표 만들기 : 지정하는 날짜가 포함된 주간 일정표 시트를 삽입(공휴일, 음력, 주요행사, 시간일정 표시)</li>
   <li> 일일 일정표 만들기 : 지정하는 날짜의 일일 일정표 시트를 삽입(공휴일, 음력, 주요 업무, 시간일정, 업무메모 표시)</li>
   <br>
   <li> 다단계 선택 기준 : 다단계 유효성 검사의 목록(Dropdown List)을 생성</li>
   <li> 다단계 선택 적용 : 선택한 셀에 다단계 유효성 검사를 적용</li>
   <li> 다단계 선택 삭제 : 필요없는 다단계 유효성 검사를 삭제</li>
   <br>
   <li> 피벗 해제 : 가로세로 Cross Tab이나 Pivot으로 된 표를 해제하여 일반 Data형 표로 만들어 줌</li>
   <li> 오류 제거 : 현재 시트에서 오류인 셀에 대해서 IFERROR()함수를 자동으로 추가하여 오류가 보이지 않도록 함</li>
   <li> UDF 제거 : 본 EG Tools의 UDF를 사용한 경우 다른 PC로 보낼 때 UDF를 제거하여 값으로 변경할 수 있음</li>
   <li> Style제거 : 셀스타일이 많을 경우 사용하지 않는 Style이나 Built-in이 아닌 Style을 일괄 삭제함</li>
   <li> 이름 제거 : 보이지 않는 명명된 이름과 유효하지 않은 이름으로 일괄 삭제함</li>
   <li> 설명서 보기 : 간단한 EG Tools의 설명서를 보여줌</li>
   <li> Version : 현재 버전을 보여주며, 배포된 Update가 있을 경우 링크를 보여줌</li>
</ul>
<br>

# Microsoft Excel 및 365 신규 함수에 대한 호환 UDF
하위버전의 Excel에서 상위 버전에 추가된 함수를 사용할 수 있습니다.<br>
※ 주의사항 : UDF를 과다하게 사용할 경우 Excel계산이 많이 느려질 수 있으므로 작업후 값으로 변경하는 것이 좋습니다.
<br>

 ## Microsoft 365 신규 예정 함수 호환
<ul>
 <li>GROUPBY 함수 : 한 축을 따라 그룹화하고 연결된 값을 집계할 수 있습니다</li>
 <li>PIVOTBY 함수 : 두 축을 따라 그룹화하고 연결된 값을 집계할 수 있습니다</li>
 <li>REGEXTEST 함수 : 텍스트 일부가 정규 표현식과 일치하는지 확인합니다</li>
 <li>REGEXEXTRACT 함수 : 정규 표현식에 따라 일치하는 하위 문자열을 추출합니다</li>
 <li>REGEXREPLACE 함수 : 정규 표현식을 사용하여 문자열의 일부를 다른 문자열로 바꿉니다</li>
 <li>TRANSLATE 함수 : 문자열을 지정하는 언어로 번역합니다</li>
 <li>DETECTLANGUAGE 함수 : 문자열의 언어를 자동으로 판별합니다</li>
 <li>TRIMRANGE 함수 : 범위/배열에서 빈행과 빈열을 제거합니다</li>
</ul>

 ## Microsoft 365 신규 함수 호환
<br>
 
 
 
 ## Excel 2024 신규 함수 호환
<ul>
 <li>TEXTSPLIT 함수 : 열과 행 구분 기호를 사용하여 텍스트 문자열을 분할합니다</li>
 <li>TEXTAFTER 함수 : 지정된 문자 또는 문자열 뒤에 나타나는 텍스트를 반환합니다</li>
 <li>TEXTBEFORE 함수 : 지정된 문자 또는 문자열 앞에 나타나는 텍스트를 반환합니다</li>
 <li>VSTACK 함수 : 배열을 세로방향으로 순서대로 추가하여 더 큰 배열을 반환합니다</li>
 <li>HSTACK 함수 : 배열을 가로방향으로 순서대로 추가하여 더 큰 배열을 반환합니다</li>
 <li>TOCOL 함수 : 단일 열의 배열을 반환합니다</li>
 <li>TOROW 함수 : 단일 행의 배열을 반환합니다</li>
 <li>WRAPCOLS 함수 : 지정된 수의 요소 뒤에 있는 열별로 제공된 값 행 또는 열을 래핑하여 새 배열을 구성합니다</li>
 <li>WRAPROWS 함수 : 지정된 수의 요소 뒤에 있는 열별로 제공된 값 행 또는 열을 래핑하여 새 배열을 구성합니다</li>
 <li>CHOOSECOLS 함수 : 배열이나 범위에서 지정된 열 순서대로 재배열한 배열을 반환합니다</li>
 <li>CHOOSEROWS 함수 : 배열이나 범위에서 지정된 행 순서대로 재배열한 배열을 반환합니다</li>
 <li>TAKE 함수 : 배열의 시작 또는 끝에서 지정된 수의 연속 행 또는 열을 반환합니다</li>
 <li>DROP 함수 : 배열의 시작 또는 끝에서 지정된 수의 행 또는 열을 제외합니다</li>
 <li>EXPAND 함수 : 배열을 확장하거나 지정된 행 및 열 차원으로 채웁니다</li>
 <li>VALUETOTEXT 함수  : 텍스트 값을 변경하지 않고 전달하며 텍스트가 아닌 값을 텍스트로 변환합니다</li>
 <li>ARRAYTOTEXT 함수  : 배열내의 텍스트 값을 변경하지 않고 전달하며 텍스트가 아닌 값을 텍스트로 변환합니다</li>
 <li>IMAGE 함수 : 인터넷에 올려진 이미지 URL이나 컴퓨터에 저장된 파일명으로 이미지를 삽입합니다</li>
</ul>

## Excel 2021 신규 함수 호환
<ul>
 <li>XMATCH 함수 : 배열 또는 셀 범위에서 지정된 항목을 검색한 다음 항목의 상대 위치를 반환합니다</li>
 <li>XLOOKUP 함수 : 테이블 또는 행별 범위에서 항목을 찾습니다</li>
 <li>XFILTER 함수 : 직접 정의한 조건을 바탕으로 일정 범위의 데이터를 필터링합니다</li>
 <li>XSORT 함수 : 범위 또는 배열의 내용을 정렬합니다</li>
 <li>SORTBY 함수 : 대응되는 범위 또는 배열의 값을 기준으로 범위 또는 배열의 내용을 정렬합니다</li>
 <li>UNIQUE 함수 : 목록 또는 범위에서 고유 값의 목록을 반환합니다</li>
 <li>SEQUENCE 함수 : 1, 2, 3, 4와 같이 일련의 연속된 숫자 목록을 생성합니다</li>
 <li>RANDARRAY 함수 : 임의의 숫자 배열을 생성합니다</li>
 <li>XLET 함수 : 계산 결과에 이름을 할당합니다. 중간 계산, 값을 저장하거나 이름을 정의할 수 있습니다</li>
</ul>

## Excel 2019 신규 함수 호환
<ul>
 <li>IFS 함수 : 하나 이상의 조건이 충족되는지 확인하고 첫 번째 TRUE 조건에 해당하는 값을 반환합니다</li>
 <li>MINIFS 함수 : 하나 이상의 조건이 모두 충족되는 최소값을 반환합니다</li>
 <li>MAXIFS 함수 : 하나 이상의 조건이 모두 충족되는 최대값을 반환합니다</li>
 <li>CONCAT 함수 : 여러 범위 및/또는 문자열의 텍스트를 결합합니다</li>
 <li>TEXTJOIN 함수 : 여러 범위 및/또는 문자열의 텍스트를 결합하며 구분기호를 포함합니다.</li>
 <li>SWITCH 함수 : 하나의 수식 또는 값을 평가하고 첫 번째 일치하는 값에 해당하는 결과를 반환합니다</li>
</ul>

## Goolgle 스프레드시트 함수에 대한 호환 UDF
<ul>
 <li>IMPORTRANGE 함수 : Google Sheets의 지정하는 범위를 가져옵니다</li>
 <li>IMPORTHTML 함수 : 인터넷 페이지의 표나 목록을 지정하여 자료를 가져옵니다</li>
 <li>IMPORTDATA 함수 : RSS나 ATOM feed 정보를 가져옵니다</li>
 <li>IMPORTFEED 함수 : csv나 tsv 파일의 데이터를 읽어 옵니다</li>
 <li>GOOGLETRANSLATE 함수 : Google의 번역 서비스를 이용한 번역을 제공합니다</li>
 <li>COUNTUNIQUE 함수 : 지정된 값과 범위 목록에서 고유 값의 개수를 셉니다</li>
 <li>COUNTUNIQUEIFS 함수 : 지정된 범위에서 여러 조건에 부합하는 고유 값의 갯수를 셉니다</li>
 <li>QUERY 함수 : 데이터에서 ADODB에 사용하는 언어로 검색을 실행합니다</li>
 <li>EPOCHTODATE 함수 : Unix epoch 타임스탬프를 협정 세계시(UTC) 기준의 날짜 및 시간으로 변환합니다</li>
 <li>ISBETWEEN 함수 : 제공된 값이 다른 두 값 사이에 있는지 확인합니다</li>
 <li>ISEMAIL 함수 : 최상위 도메인을 기준으로 유효한 이메일 주소인지 확인합니다</li>
 <li>ISURL 함수 : 유효한 URL 값인지 확인합니다</li>
</ul>
<p>
<p>
   
# EG Tools 전용 UDF
## 검색 함수
<ul>
 <li>MVLOOKUP 함수 : Excel의 VLOOKUP함수를 다량으로 실행한 결과를 출력합니다. (mass VLOOKUP)</li>
 <li>ILOOKUP 함수 : 검색 범위에서 찾는 값 중에서 지정하는 순번에 해당하는 이미지를 가져옵니다.(Image LookUp)</li>
 <li>NLOOKUP 함수 : 검색 범위에서 찾는 값과 일치하는 목록에서 지정하는 순번의 값을 찾습니다</li>
 <li>MATCHJOIN 함수 : 찾는 값이나 조건에 해당하는 내용을 연결 문자를 이용하여 연결</li>
 <li>COMPARELIST 함수 : 두개의 목록에 대해서 비교한 결과 목록을 나열합니다</li>
 <li>SAMPLE 함수 : 지정하는 대상 범위에서 무작위 샘플링 추출하여 목록을 생성합니다</li>
</ul>

## 문자열 함수
<ul>
 <li>STREXT 함수 : 숫자, 영문, 영숫자, 한글, 일본어, 한자/중국어를 추출하거나 제거합니다</li>
 <li>MATCHJOIN 함수 : 일치하는 내용에 대응하는 결과 값들을 '연결자'를 이용하여 하나의 문자열을 작성합니다</li>
 <li>TEXTPICK 함수 :  문자열을 특정 구분자를 기준으로 분리하여 원하는 순번의 값을 추출합니다</li>
 <li>TEXTBETWEEN 함수 : 지정하는 2개의 문자열 사이에 있는 내용을 추출합니다</li>
 <li>TEXTJOINIF 함수 : 조건에 만족하는 검색범위의 값을 하나의 문자열로 연결합니다</li>
 <li>CLEANB 함수 : 인쇄할 수 없는 문자코드를 제거합니다</li>
 <li>TRIMENDS 함수 : 양쪽 끝의 공백만 제거합니다.</li>
</ul>
   
## 계산 및 집계 함수
<ul>
 <li>COUNTER 함수 : 범위나 배열 데이터에서 각 요소별 빈도수를 나열합니다</li>
 <li>EVAL 함수 : 주어진 문자열의 Excel에서의 계산 결과를 산출합니다</li>
 <li>IFVISIBLE 함수 : 보이는 셀에 대해서만 각종 통계 함수를 적용합니다</li>
 <li>AGGREGATEC 함수 : 목록 또는 데이터베이스의 숨겨진 셀을 모두 제외할 수 있는 집계를 반환합니다</li>
<br>
<p>아래 함수는 별도 추가기능인 EGqcF.xlam으로 분리되었습니다.</p>
 <li>SAMPLINGSIZE 함수 : LOT크기와 AQL, 검사방법에 따라 검사할 시료수를 구합니다</li>
 <li>SAMPLINGAC 함수 : LOT크기와 AQL, 검사수준에 따라 검사할 합격판정 최대 불량수를 구합니다</li>
 <li>SAMPLINGRE 함수 : LOT크기와 AQL, 검사수준에 따라 검사할 불합격판정 최소 불량수를 구합니다.</li>
 <li>SAMPLINGLABEL 함수 : LOT크기와 검사수준에 따른 시료문자를 구합니다</li>
</ul>​

## 바코드함수
<p>아래 함수는 별도 추가기능인 EGBarcode.xlam으로 분리되었습니다.</p>
<ul>
 <li>BARCODE 함수 : 1D 및 2D 바코드 이미지를 생성합니다 (11종)</li>
 <li>QRCODE 함수 : QRCODE 바코드 이미지를 생성합니다</li>
 <li>CODE128 함수 : CODE128 바코드 이미지를 생성합니다</li>
</ul>

## 날짜시간 함수
<ul>
 <li>KOREANHOLIDAYS 함수 : 대한민국의 공휴일을 나열하는 함수입니다</li>
 <li>TOLUNAR 함수 : 양력날짜를 음력날짜로 변환합니다</li>
 <li>TOSOLAR 함수 : 음력날짜를 양력날짜로 변환합니다</li>
 <li>DATETIME 함수 : 한글, 한자가 포함된 날짜와 시간형 문자열을 날짜와 시간으로 변환합니다</li>
 <li>MONTHBYWEEK 함수 : 지정하는 요일을 기준으로 정한 해당주차의 월을 확인합니다</li>
 <li>WEEKNUMOFMONTH 함수 : 지정하는 요일을 기준으로 정한 해당주차의 월내의 주차수를 구합니다</li>
 <li>JULIANDAY 함수 : 율리우스적일 (Julian Day Number)을 계산합니다</li>
 <li>JDTODATE 함수 : 율리우스적일 (Julian Day Number)을 양력 날짜로 변환합니다</li>
</ul>​

## 색상 함수
<ul>
 <li>TEXTJOINIFCOLOR 함수 : 대상범위의 보이는 색이 기준셀과 같은 색이면 문자열을 구분자를 이용하여 연결합니다</li>
 <li>DISPLAYCOLOR 함수 : 대상셀의 보이는 색으로 바탕색/글자색의 색번호를 반환합니다</li>
 <li>SUMIFCOLOR 함수 : 대상범위의 보이는 색이 기준셀과 같은 바탕색/글자색이면 숫자를 더합니다</li>
 <li>COUNTIFCOLOR 함수 : 대상범위의 보이는 색이 기준셀과 같은 바탕색/글자색이면 숫자를 셉니다</li>
 <li>RGB 함수 : Red, Green, Blue 색상값으로 True Color 색상값을 계산합니다</li>
 <li>TORGB 함수 : True Color 색상값을 Red, Green, Blue 색상값으로 분해합니다</li>
</ul>​

## 변환 함수
<ul>
 <li>JSONPARSE 함수 : JSON 문자열의 경로명과 일치하는 값을 검색합니다</li>
 <li>JSONTOARRAY 함수 : JSON 문자열의 경로명 각 단계와 값을 배열로 구성합니다</li>
 <li>JSONPAIR 함수 : JSON 문자열을 경로명과 값의 쌍으로 나열합니다</li>
 <li>EXRATE 함수 : 외환 환율을 조회합니다
 <li>EXPLODE 함수 : 지정하는 열에 대해서 구분자를 기준으로 분해하여 나열합니다</li>
 <li>TEXTNUMSORT 함수 : 문자와 숫자가 섞여 있는 데이터를 정렬할 때, 숫자가 숫자로 정렬하도록 합니다</li>
 <li>PAPAGOTRANSLATE 함수 : 네이버의 Papago API를 이용한 번역을 제공합니다</li>
 <li>RZ 함수 : 0이나 빈셀, 오류를 빈문자열("")로 변환합니다. (Remove Zero)</li>
 <li>IFERRORX 함수 : </li>
 <li>HANTONUMBER 함수 : 한글이나 한자 및 갖은한자로 입력된 숫자를 아라비아 숫자로 변환합니다</li>
 <li>US32TODEC 함수 : 미국 채권시장의 32분수 표시형식을 십진수로 변환합니다</li>
 <li>DECTOUS32 함수 : 일반 숫자를 미국 채권시장의 32분수 표시형식으로 변환합니다</li>
</ul>​

## 공개API 함수
<ul>
 <li>ZIPCODE 함수 : 도로명 주소나 건물명 등의 키워드로 우편번호 및 도로명주소, 지번주소를 검색합니다</li>
 <li>GEOPOINT 함수 : 도로명 주소를 기준으로 해당 주소의 지도 좌표를 확인합니다</li>
 <li>GEOCONVERT 함수 : 지도 좌표를 다른 좌표계로 변환합니다</li>
 <li>GEODISTANCE 함수 : 지도 좌표로 거리를 개략적으로 계산합니다</li>
 <li>OILPRICE 함수 : OPINET에서 제공하는 API를 이용하여 지역별 유종별 평균유가를 조회합니다</li>
 <li>GASSTATION 함수 : OPINET에서 제공하는 API를 이용하여 주변 유가를 검색합니다</li>
 <li>BRNSTATUS 함수 : 국세청의 API를 이용하여 사업자등록번호의 현재 상태를 조회합니다</li>
</ul>​

## 기타 함수
<ul>
 <li>SHEETSLIST 함수 : 현재 Excel 파일의 시트 목록을 작성합니다</li>
 <li>IPINFO 함수 : IP Address 기본정보</li>
 <li>DIRFOLDER 함수 : 지정한 폴더의 파일 목록을 출력합니다</li>
 <li>IMPORTURL 함수 : 인터넷 페이지의 소스를 표시합니다</li>
</ul>


<br>


# 추가 예정
<ul>
</ul>
<br>



# 감사인사
기능에 대한 조언과 테스트를 통해 오류를 잡아 주시는 분들께 항상 감사 드립니다.<br>
<br>



# 사용권한
본 파일은 개인, 회사, 관공서 등 누구나 무료로 사용할 수 있습니다.<br>
본 파일을 사용함으로써 발생하는 모든 책임은 사용자에게 있습니다.<br>
만약 이에 동의하지 않는다면, 사용을 중단하고 파일을 삭제 바랍니다.<br>



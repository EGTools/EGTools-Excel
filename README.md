# EGTools-Excel
<p><b>Excel2010부터 사용할 수 있는 Excel 추가기능(Add-in)입니다.</b><br>
   <b>Mac용 Excel에서는 VBA코딩이 달라 사용할 수 없습니다.</b></p>
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
문의사항은 토론방을 이용 바랍니다.
<br>https://github.com/EGTools/EGTools-Excel/discussions
<br>

<p>v2.6부터 ExcelDna-Intellisense Add-in을 설치할 경우 셀에 함수 입력시 인라인 도움말을 볼 수 있습니다.<br>
이 Add-in을 설치하지 않아도 EGTools를 사용하는 데에는 전혀 문제가 없습니다.</p>


<ol type="1">
<li> ExcelDna-Intellisense는 여기에서 다운로드 하세요.</li> 

   <p> https://github.com/Excel-DNA/IntelliSense/releases/latest </p>

<li> Excel 버전에 따라 사용하는 파일이 다릅니다. (Windows 버전과 상관없음) </li> 

   Excel이 64비트로 설치된 경우 ExcelDna.IntelliSense64.xll<br>
   Excel이 32비트로 설치된 경우 ExcelDna.IntelliSense.xll 
   
<li> [Excel 추가기능] 에서 다운로드 받은 파일을 찾아 추가합니다. (COM 추가 기능 아님) </li>
</ol>
<br>

<p>
<p>
   
   
# 설치방법
<p> Excel 추가기능 설치 방법은 여기를 참고하세요.
<br>https://github.com/EGTools/EGTools-Excel/wiki/Excel-Add-in-%EC%82%AC%EC%9A%A9-%EC%84%A4%EC%A0%95
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
   상기 기능은 내용병합을 할 때 텍스트 서식을 유지하는 것과 서식을 제거하는 것이 있음
   <li> 연속값 병합 : 열방향(아래쪽)으로 같은 값이 연속될 때 셀을 자동으로 병합함</li>
   <li> 열끼리 병합 : 선택된 영역에 대해서 열별로 '내용 병합'을 한 번에 수행함</li>
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
   <li> 피벗 해제 : 가로세로 Cross Tab이나 Pivot으로 된 표를 해제하여 일반 Data형 표로 만들어 줌</li>
   <li> 오류 제거 : 현재 시트에서 오류인 셀에 대해서 IFERROR()함수를 자동으로 추가하여 오류가 보이지 않도록 함</li>
   <li> UDF 제거 : 본 EG Tools의 UDF를 사용한 경우 다른 PC로 보낼 때 UDF를 제거하여 값으로 변경할 수 있음</li>
   <li> Style제거 : 셀스타일이 많을 경우 사용하지 않는 Style이나 Built-in이 아닌 Style을 일괄 삭제함</li>
   <li> 이름 제거 : 보이지 않는 명명된 이름과 유효하지 않은 이름으로 일괄 삭제함</li>
   <li> 설명서 보기 : 간단한 EG Tools의 설명서를 보여줌</li>
   <li> Version : 현재 버전을 보여주며, 배포된 Update가 있을 경우 링크를 보여줌</li>
</ul>
<br>

## Microsoft Excel 및 365 신규 함수에 대한 호환 UDF
하위버전의 Excel에서 상위 버전에 추가된 함수를 사용할 수 있습니다.<br>
※ 주의사항 : UDF를 과다하게 사용할 경우 Excel계산이 많이 느려질 수 있으므로 작업후 값으로 변경하는 것이 좋습니다.
<br>

 ## Microsoft 365 신규 함수 호환
<ul>
   <li>IMAGE : 셀에 인터넷 URL이나 PC의 전체경로명이 있는 파일명에 해당하는 사진/그림을 삽입함</li>
   <li>TEXTSPLIT : 열과 행 구분 기호를 사용하여 텍스트 문자열을 분할 </li>
   <li>TEXTAFTER : 문자열에서 지정된 부분 문자열 뒤에 발생하는 텍스트 문자열</li>
   <li>TEXTBEFORE : 문자열에서 지정된 부분 문자열 앞에 발생하는 텍스트 문자열 </li>
   <li>VSTACK : 범위의 아래쪽에 있는 첫 번째 빈 셀에 데이터를 추가(세로 쌓기) </li>
   <li>HSTACK : 범위의 오른쪽에 있는 첫 번째 빈 셀에 데이터를 추가(가로 늘리기) </li>
   <li>TOCOL : 원본 배열의 모든 항목을 포함하는 열 벡터를 반환합니다.  원본 배열을 하나의 열에 모두 나열</li>
   <li>TOROW : 원본 배열의 모든 항목을 포함하는 행 벡터를 반환합니다.  원본 배열을 하나의 행에 모두 나열</li>
   <li>WRAPCOLS : 지정한 수로 열을 채운 배열로 나열 (첫번째 열을 지정한 수로 채우고 오른쪽 열로 진행)</li>
   <li>WRAPROWS : 지정한 수로 행을 채운 배열로 나열 (첫번째 행을 지정한 수로 채우고 아래쪽 행으로 진행)</li>
   <li>TAKE : 배열의 시작 또는 끝에서 지정된 수의 연속 행 또는 열을 반환</li>
   <li>DROP : 배열의 시작 또는 끝에서 지정된 수의 연속 행 또는 열을 제외</li>
   <li>CHOOSECOLS : 배열에서 지정된 열(들)을 반환.  열순서를 바꿔서 배열할 수도 있음</li>
   <li>CHOOSEROWS : 배열에서 지정된 행(들)을 반환.  행순서를 바꿔서 배열할 수도 있음</li>
   <li>EXPAND : 지정된 행 및 열 차원으로 확장</li>
   <li>VALUETOTEXT : 지정한 값을 텍스트로 전환, 0은 일반형식을 텍스트로, 1은 수식에 사용할 수 있는 텍스트로 변환</li>
   <li>ARRAYTOTEXT : 범위나 Array를 텍스트로 전환, 0은 일반형식을 텍스트로, 1은 수식에 사용할 수 있는 텍스트로 변환</li>
</ul>

## Excel 2019 신규 함수 호환
<ul>
   <li>XMATCH : 배열 또는 셀 범위에서 지정된 항목을 검색한 다음 항목의 상대 위치를 반환합니다.</li>
   <li>XLOOKUP : 반환 열이 있는 쪽에 관계없이 한 열에서 검색어를 보고 다른 열의 동일한 행에서 결과를 반환함</li>
   <li>xFILTER : 직접 정의한 조건을 바탕으로 일정 범위의 데이터를 필터링(원함수명 FILTER)</li>
   원본함수와 다르게 Ctrl+Shift+Enter로 입력해야 하는 '배열함수'임
   <li>xSORT : 범위 또는 배열의 내용을 정렬  (원함수명 SORT)</li>
   <li>SORTBY : 대응되는 범위 또는 배열의 값을 기준으로 범위 또는 배열의 내용을 정렬. 여러 열 지정 가능</li>
   <li>UNIQUE : 목록 또는 범위에서 고유 값의 목록을 반환하거나, 오직 한 번만 나타나는 목록을 반환</li>
   <li>SEQUENCE : 1, 2, 3, 4와 같이 일련의 연속된 숫자 목록을 생성</li>
   <li>RANDARRAY : 행/열의 수, 최소값/최대값 및 정수/소수값 등을 지정하여 임의의 숫자 배열을 작성</li>
   <li>xLET : 변수 이름과 해당하는 값/수식을 지정하고, 이를 이용한 '사용자 수식'의 결과를 반환 (원함수명 LET)</li>
</ul>

## Excel 2016 신규 함수 호환
<ul>
   <li>IFS : 하나 이상의 조건이 충족될지 여부를 확인하고 첫 번째 TRUE 조건에 해당하는 값을 반환</li>
   <li>MINIFS : 주어진 조건 집합에 맞는 셀에서 최소값</li>
   <li>MAXIFS : 주어진 조건 집합에 맞는 셀에서 최대값</li>
   <li>CONCAT : 여러 범위 및/또는 문자열의 텍스트를 결합(구분자 지정 없음)</li>
   <li>TEXTJOIN : 여러 범위 및/또는 문자열을 구분자를 지정하여 연결</li>
   <li>SWITCH : 찾는값과 그 결과에 따라 반환할 값을 최대 126개까지 지정</li>
</ul>

## Goolgle 스프레드시트 함수에 대한 호환 UDF
<ul>
   <li>Query : 원본데이터에 대해서 Query를 수행</li>
   <li>REGEXEXTRACT : 정규 표현식에 따라 첫 번째로 일치하는 하위 문자열을 추출</li>
   <li>REGEXMATCH : 텍스트 일부가 정규 표현식과 일치하는지 여부를 확인</li>
   <li>REGEXREPLACE : 정규 표현식을 사용하여 텍스트 문자열의 일부를 다른 텍스트 문자열로 대체</li>
   <li>IsBetween : 제공된 값이 다른 두 값들 사이에 있는지 확인</li>
   <li>IsURL : 유효한 URL 값인지 확인</li>
   <li>IsEmail : 국가 또는 지역 코드와 최상위 도메인을 기준으로 유효한 이메일 주소인지 확인</li>
</ul>

## EG Tools 전용 UDF
<ul>
   <li>iLOOKUP : XLOOKUP과 비슷한데, 찾는 값의 지정한 순서에 해당하는 셀의 그림을 복사해 옴 </li>
   <li>TEXTPICK : 대상문자열에 구분자(들)을 마디로 하여 지정한 순번의 문자열을 반환  </li>
   <li>nLOOKUP : XLOOKUP과 비슷한데, 첫번째가 아닌 지정한 순번의 것을 찾음  </li>
   <li>QRCode : 숫자, 영문자, 영숫자, 한글등 유니코드로 된 내용으로 2차원 바코드인 QR Code를 삽입 </li>
   <li>Code128 : 숫자, 영문자, 영숫자 등 ASCII 코드에 해당하는 값을 1차원 Bar Code로 삽입</li>
   <li>MATCHJOIN : TEXTJOIN에 조건을 추가한 것, 조건에 맞는 결과 값만 구분자를 사용하여 하나로 연결 </li>
   <li>STREXT : 옵션에 따라 숫자, 영문, 한글, 한자, 일본어 및 정규식표현을 추출하거나 제거  </li>
   <li>CountInStr : 대상문자열에 찾는문자열이 들어 있는 수를 Count  </li>
   <li>EVAL : Excel 수식을 계산한 결과  </li>
   <li>MonthByWeek : 특정 요일을 기준으로 월을 구분하여 월을 구함 </li>
   <li>WeekNumByWeek : 특정 요일을 기준으로 월내에서의 주차번호를 구함 </li>
   <li>SUMIFBack : 참조범위의 기준셀과 같은 바탕색이면 합산  </li>
   <li>SUMIFFont : 참조범위의 기준셀과 같은 글자색이면 합산  </li>
   <li>CountIFBack : 참조범위의 기준셀과 같은 바탕색인 셀 수  </li>
   <li>CountIFFont : 참조범위의 기준셀과 같은 글자색인 셀 수  </li>
   <li>ToLunar : 양력날짜를 음력날짜로 변환함, 결과는 문자열이며 윤달인 경우 날짜뒤에 "(윤)"이 추가됨  </li>
   <li>ToSolar : 음력날짜를 양력날짜로 변환함, 결과는 날짜형식임 (1900년~2050년)  </li>
   <li>TextBetween : 문자열에서 지정하는 2개의 문자열 사이에 있는 내용을 추출  </li>
   <li>HanToNumber : 한글이나 한자/갖은한자로 된 숫자를 아라비아 숫자로 변환  </li>
   <li>FindFirstData : 참조범위에 데이터가 입력된 첫번째 셀의 순번  </li>
   <li>FindLastData : 참조범위에 데이터가 입력된 마지막 셀의 순번  </li>
   <li>FindIncluded : 찾는내용을 포함하고 있는 셀들의 범위내의 순서를 찾음  </li>
   <li>FindSubstring : 찾는내용의 일부분인 셀들의 범위내의 순서를 찾음  </li>
   <li>NumToXLColumn : 숫자를 Excel 열이름으로 변환  </li>
   <li>NumFromXLColumn : Excel 열 이름을 열번호 숫자로 변환  </li>
   <li>EG10to36 : 10진수를 36진수로 변환  </li>
</ul>
<br>


# 사용권한
본 파일은 개인, 회사, 관공서 등 누구나 무료로 사용할 수 있습니다.<br>
본 파일을 사용함으로써 발생하는 모든 책임은 사용자에게 있습니다.<br>
만약 이에 동의하지 않는다면, 사용을 중단하고 파일을 삭제 바랍니다.<br>



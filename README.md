# EGTools-Excel
<br>**EGTools is an Excel add-in that provides various functions and features to help you use Excel.**
<br>**It supports new functions added to Excel 2019, 2021, 2024, and Microsoft 365.**
<br>
_The VBA code is different, it cannot be used in Excel for Mac._

<br>Please download the latest version from the Release page.
<br>New versions will continue to be updated with the same name, so just put them in the same folder after downloading.
<br>https://github.com/EGTools/EGTools-Excel/releases/latest 
<br>
<br>Before the official distribution, there may be cases where improved/supplemented Pre-Release is temporarily uploaded.
<br>Please use it with caution as there may be some errors because the function inspection has not been completed.
<br>https://github.com/EGTools/EGTools-Excel/releases
<br>
<br>For inquiries, please use the "Discussions"
<br>[Reporting Errors](https://github.com/EGTools/EGTools-Excel/discussions/categories/reporting-errors)
<br>[Q & A](https://github.com/EGTools/EGTools-Excel/discussions/categories/q-a)
<br>

> [!NOTE]
> From v4.5.5, supporting for ExcelDna-Intellisense Add-in has been removed.

<br>

<p>
<p>
   
   
# How to install
<p> For instructions on how to install Excel add-ins, please refer to
<br>https://github.com/EGTools/EGTools-Excel/wiki/Install-Excel-Add%E2%80%90in
<br>
<br>  


> [!WARNING]
> Excessive use of UDF can slow down Excel calculations considerably, so it is recommended to change it to a value after the operation.

# EXCEL New Function Compatible UDFs
You can use functions added in higher versions of Excel in lower versions.<br>
<br>

## [Microsoft 365 New Functions in Preview](https://github.com/EGTools/EGTools-Excel/wiki/Microsoft-365-Preview-New-Functions)
- [REGEXTEST](https://github.com/EGTools/EGTools-Excel/wiki/Microsoft-365-Preview-New-Functions#regextest) : Checks if part of text matches a regular expression.
- [REGEXEXTRACT](https://github.com/EGTools/EGTools-Excel/wiki/Microsoft-365-Preview-New-Functions#regexextract) : Extracts substrings that match a regular expression.
- [REGEXREPLACE](https://github.com/EGTools/EGTools-Excel/wiki/Microsoft-365-Preview-New-Functions#regexreplace) : Replaces part of a string with another string using a regular expression.
- [TRANSLATE](https://github.com/EGTools/EGTools-Excel/wiki/Microsoft-365-Preview-New-Functions#translate) : Translates a string into a specified language.
- [DETECTLANGUAGE](https://github.com/EGTools/EGTools-Excel/wiki/Microsoft-365-Preview-New-Functions#detectlanguage) : Automatically determines the language of a string.
- [TRIMRANGE](https://github.com/EGTools/EGTools-Excel/wiki/Microsoft-365-Preview-New-Functions#trimrange) : Removes blank rows and columns from a range/array.

## [Microsoft 365 New Functions](https://github.com/EGTools/EGTools-Excel/wiki/Microsoft-365-New-Functions)
- [GROUPBY](https://github.com/EGTools/EGTools-Excel/wiki/Microsoft-365-New-Functions#groupby) : Groups along one axis and aggregates associated values.
- [PIVOTBY](https://github.com/EGTools/EGTools-Excel/wiki/Microsoft-365-New-Functions#pivotby) : Groups and aggregates associated values ​​along two axes.
- [xPERCENTOF](https://github.com/EGTools/EGTools-Excel/wiki/Microsoft-365-New-Functions#xpercentof) : Calculates the percentage of a given value divided by the total value.
 
## Excel 2024 New Functions
- [TEXTSPLIT](https://cafe.naver.com/egtools/131) : Splits a text string using column and row delimiters
- [TEXTAFTER](https://cafe.naver.com/egtools/132) : Returns the text that appears after a specified character or string.
- [TEXTBEFORE](https://cafe.naver.com/egtools/133) : Returns the text that appears before a specified character or string.
- [VSTACK](https://cafe.naver.com/egtools/134) : Adds arrays vertically and returns a larger array.
- [HSTACK](https://cafe.naver.com/egtools/135) : Appends arrays horizontally and returns a larger array.
- [TOCOL](https://cafe.naver.com/egtools/136) : Returns an array of a single column.
- [TOROW](https://cafe.naver.com/egtools/137) : Returns an array of a single row.
- [WRAPCOLS](https://cafe.naver.com/egtools/138) : Constructs a new array by wrapping the provided rows or columns of values ​​after a specified number of elements.
- [WRAPROWS](https://cafe.naver.com/egtools/139) : Constructs a new array by wrapping the provided rows or columns of values ​​after a specified number of elements.
- [CHOOSECOLS](https://cafe.naver.com/egtools/140) : Returns an array rearranged in the specified column order from an array or range.
- [CHOOSEROWS](https://cafe.naver.com/egtools/141) : Returns an array with rows rearranged in the specified order in an array or range.
- [TAKE](https://cafe.naver.com/egtools/142) : Returns a specified number of consecutive rows or columns from the beginning or end of an array.
- [DROP](https://cafe.naver.com/egtools/143) : Excludes a specified number of rows or columns from the beginning or end of an array.
- [EXPAND](https://cafe.naver.com/egtools/144) : Expands an array or fills it to specified row and column dimensions.
- [VALUETOTEXT](https://cafe.naver.com/egtools/145)  : Passes text values ​​unchanged and converts non-text values ​​to text.
- [ARRAYTOTEXT](https://cafe.naver.com/egtools/146)  : Passes text values ​​in an array without changing them and converts non-text values ​​to text.
- [IMAGE](https://cafe.naver.com/egtools/147) : Inserts an image using an image URL uploaded on the Internet or a file name stored on your computer.

## Excel 2021 New Functions
- [XMATCH](https://cafe.naver.com/egtools/73) : Searches for a specified item in an array or range of cells and returns the relative position of the item.
- [XLOOKUP](https://cafe.naver.com/egtools/74) :Find items in a table or range by row.
- [XFILTER](https://cafe.naver.com/egtools/75) : Filters a range of data based on conditions you define.
- [XSORT](https://cafe.naver.com/egtools/76) : Sorts the contents of a range or array.
- [SORTBY](https://cafe.naver.com/egtools/77) : Sorts the contents of a range or array based on the values ​​in the corresponding range or array.
- [UNIQUE](https://cafe.naver.com/egtools/78) : Returns a list of unique values ​​from a list or range.
- [SEQUENCE](https://cafe.naver.com/egtools/79) : Generates a list of consecutive numbers, such as 1, 2, 3, 4.
- [RANDARRAY](https://cafe.naver.com/egtools/80) : Generates a random number array.
- [XLET](https://cafe.naver.com/egtools/81) : Assigns a name to the result of a calculation. You can store intermediate calculations, values, or define names.

## Excel 2019 New Functions
- [IFS](https://cafe.naver.com/egtools/38) : Checks if one or more conditions are met and returns the value corresponding to the first TRUE condition.
- [MINIFS](https://cafe.naver.com/egtools/39) : Returns the minimum value for which one or more conditions are met.
- [MAXIFS](https://cafe.naver.com/egtools/40) : Returns the largest value where one or more conditions are met.
- [CONCAT](https://cafe.naver.com/egtools/41) : Combines text from multiple ranges and/or strings.
- [TEXTJOIN](https://cafe.naver.com/egtools/42) : Combines text from multiple ranges and/or strings, including delimiters.
- [SWITCH](https://cafe.naver.com/egtools/43) : Evaluates one formula or value and returns the result corresponding to the first matching value.

## Excel 2013 New Functions
- FORMULATEXT : Shows the function entered in the specified cell.
- ENCODEURL : Encodes a value so that it can be used by the browser.
- IFNA : Changes to a specified value when there is a #N/A error
- UNICODE : Returns the Unicode code value of the first character.
- UNICHAR : Returns the Unicode character with the specified code value.

<br>
<br>

# Google Sheets Compatible Functions
- [IMPORTRANGE](https://cafe.naver.com/egtools/153) : Import a specified range in Google Sheets
- [IMPORTHTML](https://cafe.naver.com/egtools/154) : Imports data by specifying a table or list from an Internet page.
- [IMPORTDATA](https://cafe.naver.com/egtools/155) : Import RSS or ATOM feed information
- [IMPORTFEED](https://cafe.naver.com/egtools/156) : Reads data from a csv or tsv file.
- [GOOGLETRANSLATE](https://cafe.naver.com/egtools/130) : Provides translation using Google's translation service.
- [COUNTUNIQUE](https://cafe.naver.com/egtools/128) : Counts the number of unique values ​​in a list of specified values ​​and ranges.
- [COUNTUNIQUEIFS](https://cafe.naver.com/egtools/129) : Counts the number of unique values ​​that meet multiple conditions in a specified range.
- [QUERY](https://cafe.naver.com/egtools/127) : Executes a search on data in the language used by ADODB.
- [EPOCHTODATE](https://cafe.naver.com/egtools/126) : Converts a Unix epoch timestamp to a date and time in Coordinated Universal Time (UTC).
- [ISBETWEEN](https://cafe.naver.com/egtools/125) : Checks if a given value is between two other values
- [ISEMAIL](https://cafe.naver.com/egtools/124) : Checks if an email address is valid based on the top-level domain
- [ISURL](https://cafe.naver.com/egtools/123) : Check if URL value is valid

<br>
<br>
   
# EGTools UDF
## Search functions
- [MVLOOKUP](https://cafe.naver.com/egtools/107) : Outputs the results of executing Excel's VLOOKUP function in bulk. (mass VLOOKUP)
- [MXLOOKUP](https://cafe.naver.com/egtools/211) : Outputs the results of executing Excel's XLOOKUP function in bulk. (mass XLOOKUP)
- [ILOOKUP](https://cafe.naver.com/egtools/51) : Gets the image corresponding to the specified sequence number among the values ​​found in the search range. (Image LookUp)
- [NLOOKUP](https://cafe.naver.com/egtools/49) : Finds a value with a specified number in a list that matches the value you are looking for in the search range.
- [MATCHJOIN](https://cafe.naver.com/egtools/48) : Connects the content that matches the search value or condition using a connection character.
- [COMPARELIST](https://cafe.naver.com/egtools/64) : Lists the results of comparing individual values ​​in two lists.
- [COMPARELISTM](https://cafe.naver.com/egtools/292) : Lists the results of comparing row-by-row values ​​for two lists.
- [SAMPLE](https://cafe.naver.com/egtools/87) : Generates a list by randomly sampling from a specified target range.

## String functions
- [STREXT](https://cafe.naver.com/egtools/47) : Extract or remove numbers, English letters, alphanumeric characters, Korean letters, Japanese letters, and Chinese letters/Chinese characters.
- [MATCHJOIN](https://cafe.naver.com/egtools/48) : Creates a single string using the result values ​​corresponding to the matching contents using a 'joiner'.
- [TEXTPICK](https://cafe.naver.com/egtools/50) :  Splits a string based on a specific delimiter and extracts the value of the desired sequence.
- [TEXTBETWEEN](https://cafe.naver.com/egtools/59) : Extracts the content between two specified strings.
- [TEXTJOINIF](https://cafe.naver.com/egtools/86) : Concatenates values ​​in the search range that satisfy the conditions into a single string.
- [CLEANB](https://cafe.naver.com/egtools/105) : Removes non-printable character codes.
- [TRIMENDS](https://cafe.naver.com/egtools/120) : Removes only spaces from both ends.
   
## Calculation and Aggregation Functions
- [COUNTER](https://cafe.naver.com/egtools/17) : Lists the frequency of each element in a range or array of data.
- [EVAL](https://cafe.naver.com/egtools/46) : Returns the result of a calculation in Excel for a given string.
- [IFVISIBLE](https://cafe.naver.com/egtools/110) : Applies various statistical functions only to visible cells.
- [AGGREGATEC](https://cafe.naver.com/egtools/83) : Returns an aggregate that excludes all hidden cells in a list or database.
<br>

> [!NOTE]
> It was separated into EGqcF.xlam.
>## QC sampling function
>- [SAMPLINGSIZE](https://cafe.naver.com/egtools/93) : Calculates the number of samples to be tested based on the LOT size, AQL, and testing method.
>- [SAMPLINGAC](https://cafe.naver.com/egtools/95) : Finds the maximum number of defective units to be inspected based on LOT size, AQL, and inspection level.
>- [SAMPLINGRE](https://cafe.naver.com/egtools/96) : Calculates the minimum number of defective units to be inspected based on LOT size, AQL, and inspection level.
>- [SAMPLINGLABEL](https://cafe.naver.com/egtools/94) : Get sample text based on LOT size and inspection level.
​<br>

> [!NOTE]
> Separated into EGBarcode.xlam.
>## Barcode function
>- [BARCODE](https://cafe.naver.com/egtools/90) : Generate 1D and 2D barcode images (11 types)
>- [QRCODE](https://cafe.naver.com/egtools/92) : Generate a QRCODE barcode image.
>- [CODE128](https://cafe.naver.com/egtools/91) : Generates a CODE128 barcode image.
<br>

## DateTime Functions
- [KOREANHOLIDAYS](https://cafe.naver.com/egtools/20) : Lists public holidays in South Korea.
- [TOLUNAR](https://cafe.naver.com/egtools/60) : Converts a solar date to a lunar date.
- [TOSOLAR](https://cafe.naver.com/egtools/61) : Converts a lunar date to a solar date.
- [DATETIME](https://cafe.naver.com/egtools/67) : Converts a date and time string containing Korean and Chinese characters to a date and time.
- [MONTHBYWEEK](https://cafe.naver.com/egtools/57) : Checks the month of the specified week based on the specified day of the week.
- [WEEKNUMOFMONTH](https://cafe.naver.com/egtools/58) : Finds the number of weeks in a month based on the specified day of the week.
- [JULIANDAY](https://cafe.naver.com/egtools/102) : Calculates the Julian Day Number.
- [JDTODATE](https://cafe.naver.com/egtools/103) : Converts a Julian Day Number to a Gregorian date.
​

## Color function
- [TEXTJOINIFCOLOR](https://cafe.naver.com/egtools/84) : Joins strings using a delimiter if the visible color of the target range is the same color as the reference cell.
- [DISPLAYCOLOR](https://cafe.naver.com/egtools/70) : Returns the color number of the background color/text color as the visible color of the target cell.
- [SUMIFCOLOR](https://cafe.naver.com/egtools/69) : Adds numbers if the visible color of the target range is the same background color/font color as the reference cell.
- [COUNTIFCOLOR](https://cafe.naver.com/egtools/68) : Counts the numbers if the visible color of the target range is the same background color/font color as the reference cell.
- [RGB](https://cafe.naver.com/egtools/148) : Calculates a True Color color value using Red, Green, and Blue color values.
- [TORGB](https://cafe.naver.com/egtools/149) : Decomposes a True Color color value into Red, Green, and Blue color values.

## Conversion function
- [UNPIVOT](https://cafe.naver.com/egtools/303) : Converts a pivot table or crosstab to a regular data table.
- [JSONPARSE](https://cafe.naver.com/egtools/152) : Finds values ​​matching a path name in a JSON string.
- [JSONTOARRAY](https://cafe.naver.com/egtools/151) : Converts each step and value of a JSON string's pathname into an array.
- [JSONPAI수](https://cafe.naver.com/egtools/150) : Lists a JSON string as pathname-value pairs.
- [EXRATE](https://cafe.naver.com/egtools/113) : Check the foreign exchange rate for the South Korean Won
- [EXPLODE](https://cafe.naver.com/egtools/108) : Lists the columns you specify by breaking them down into delimiters.
- [TEXTNUMSORT](https://cafe.naver.com/egtools/108) : When sorting data that contains mixed letters and numbers, sorts the numbers as numbers.
- [PAPAGOTRANSLATE](https://cafe.naver.com/egtools/104) : 네이버의 Papago API를 이용한 번역을 제공합니다
- [RZ](https://cafe.naver.com/egtools/88) : 0이나 빈셀, 오류를 빈문자열("")로 변환합니다. (Remove Zero)
- IFERRORX : 
- [HANTONUMBER](https://cafe.naver.com/egtools/62) : Converts numbers entered in Korean, Chinese characters, and various Chinese characters to Arabic numerals.
- [US32TODEC](https://cafe.naver.com/egtools/65) : Converts the 32-fraction representation of the US bond market to decimal.
- [DECTOUS32](https://cafe.naver.com/egtools/66) : Converts a regular number to the 32-digit representation of the U.S. bond market.
​

## Public API functions only for Republic of Korea
- [SEARCHADDRESS](https://cafe.naver.com/egtools/261) : Search for information by road name address.
- [ZIPCODE](https://cafe.naver.com/egtools/106) : Searches for zip codes, road names, and land addresses using keywords such as road name addresses or building names.
- [GEOPOINT](https://cafe.naver.com/egtools/115) : Check the map coordinates of an address based on the road name address.
- [GEOCONVERT](https://cafe.naver.com/egtools/117) : Converts map coordinates to another coordinate system.
- [GEODISTANCE](https://cafe.naver.com/egtools/118) : Roughly calculates distance using map coordinates
- [OILPRICE](https://cafe.naver.com/egtools/114) : [OPINET](https://www.opinet.co.kr/user/main/mainView.do) Search for the average oil price by region and type.
- [GASSTATION](https://cafe.naver.com/egtools/116) : [OPINET](https://www.opinet.co.kr/user/main/mainView.do) Search for nearby oil prices using
- [BRNSTATUS](https://cafe.naver.com/egtools/119) : Check the current status of the business registration number using the National Tax Service API.


## Other functions
- [SHEETSLIST](https://cafe.naver.com/egtools/112) : Creates a list of sheets in the current Excel file.
- IPINFO : IP Address basic information
- [DIRFOLDER](https://cafe.naver.com/egtools/109) : Outputs a list of files in a specified folder.
- [IMPORTURL](https://cafe.naver.com/egtools/168) : Displays the source of an Internet page

<br>
<br>

# Ribbon Menu Features
## [Only visible cells](https://cafe.naver.com/egtools/23)
- Copy Visible Cells: A function to copy only the cells visible on the screen.
- Copy All: Ability to copy both visible and invisible cells on the screen.
- The above two functions must be performed first before pasting into the visible cells below.
- Paste All: Paste only visible cells with both cell format and values.
- Paste values ​​only: Paste into cells where only the values ​​are visible.
- Paste formula: Paste into cells where only the formula is visible

## [Merge/Split](https://cafe.naver.com/egtools/24)
- Merge contents: Merge cells and merge all contents together.
- Merge columns: Performs 'Merge contents' for each column in the selected area at once.
- Merge Rows: Performs 'Merge Contents' for each row in the selected area at once.
- The above two functions include maintaining text formatting and removing formatting when merging content.
- Merge Consecutive Values: Automatically merges cells when the same values ​​are consecutive in the column direction (downward).
- Split Row: Splits cell contents with line breaks into multiple rows (lines).
- Split Column: Divide the cell contents with a separator into multiple columns (columns).
- The above two functions can be used with text format or without format.
- Divide and Fill: Separate merged cells and copy all the same content

## [Photo/Images](https://cafe.naver.com/egtools/25)
- Insert Selection: Insert a photo/picture saved on your PC into the selected cell.
- Insert Folder: When a file name is entered in the cell contents, inserts the corresponding photos/pictures from the specified folder in batches.
- Fit Selection: Automatically fit the selected photo/picture to the cell
- Fit All: Automatically fits all photos/pictures in the current sheet to the cells.
- Save All: Save all photos/pictures in the current sheet to the specified folder.

## [Calendar/Schedule](https://cafe.naver.com/egtools/27)
- Create an annual calendar: Insert an annual calendar sheet for the specified year (showing public holidays in Korea)
- Create a monthly schedule: Insert a monthly schedule sheet for the month you specify (showing Korean public holidays and the lunar calendar)
- Create a weekly schedule: Insert a weekly schedule sheet with the dates you specify (showing Korean public holidays, lunar calendar, major events, and time schedules)
- Create a daily schedule: Insert a daily schedule sheet for the specified date (displaying Korean public holidays, lunar calendar, major tasks, time schedule, and work notes)

## [Multi-Level Selection](https://cafe.naver.com/egtools/28)
- Multi-level selection criteria: Create a dropdown list of multi-step validations
- Apply multi-level selection: Apply multi-step validation to selected cells.
- Remove multi-level selection: Remove unnecessary multi-step validation

## [Table](https://cafe.naver.com/egtools/29)
- UnPivot: De-pivots a table into a normal data type table by de-pivoting it into a Cross Tab or Pivot table.
- Cross Tab: Creates a general data type table as a Cross Tab and aggregates it.
- Table Aggregation: Combine data from multiple sheets of the same format into one.

> [!NOTE]
> Separated into EGBarcode.xlam.
>## Barcode
>- 1D Barcode :
>- 2D Barcode :
>- GS1 Barcode :

## Other features
- Save as Image: Save the selected area as an image file.
- Remove errors: Automatically adds the IFERROR() function to cells that are errors in the current sheet, so that the errors are not visible.
- Remove UDF: If you used UDF of this EG Tools, you can remove UDF and change it to a value when sending it to another PC.
- Delete Style: When there are many cell styles, delete all unused styles or styles that are not built-in.
- [Delete Names](https://cafe.naver.com/egtools/31) : Bulk delete invisible named names and invalid names.
- Clean up empty cells: remove zero-length strings from the current sheet
- [Trim Ends](https://cafe.naver.com/egtools/274) : Removes spaces from the beginning and end of all cells in the current sheet.
- Rearrange notes: Repositions all notes in the current sheet right next to the inserted cell.
- [Mail Merge](https://cafe.naver.com/egtools/32) : Automatically create sheets or files using lists and forms, and print or email them.
- [Outline Shapes](https://cafe.naver.com/egtools/33) :  Automatically generates free-form shapes along the outlines of blocks drawn in cell background color.

## EGTools related
- [Fix EGTools Local Link](https://cafe.naver.com/egtools/290) : Modify the path to use EGTools on another PC
- [Fix EGTools Array Formula Result](https://cafe.naver.com/egtools/290) : Fix EGTools array and function according to Excel version
- [Manual](https://cafe.naver.com/egtools/8) : Shows the documentation for Simple EG Tools
- Version: Shows the current version and shows a link if there is an update released.
- Disable EGTools: Temporarily disable the EGTools add-on, or disable it and delete the files.

<br>
<br>


# Thanks
We are always grateful to those who help us by providing advice and testing on features and by helping us catch errors.<br>
<br>



# Permissions
This file can be used by anyone, including individuals, companies, and government agencies, for free.<br>
All responsibility arising from the use of this file lies with the user.<br>
If you do not agree with this, please stop using it and delete the file.<br>



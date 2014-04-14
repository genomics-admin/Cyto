#include <Excel.au3>
#include <array.au3>

;;-----Statuc variable values

;~ Local $nodeColor[8][4]=[ _
;~         ['Orange','MainNode','255,204,0', '204,0,0'], _
;~         ['Grey','Collaboration','102,102,0', '102,102,0'], _
;~ 		['LBlue','Business Agreement','51,255,255', '51,255,255'], _
;~ 		['DBlue','Divestment/Spinoff','0,51,204', '0,51,204'], _
;~ 		['Green','Aquisition','51,204,0', '51,204,0']]

Local $nodeColor[8][4]=[ _
        ['Orange','MainNode','255,204,0', '204,0,0'], _
        ['Grey','Collaboration','217,211,157', '217,211,157'], _
		['LBlue','Business Agreement','36,215,255', '36,215,255'], _
		['DBlue','Divestment/Spinoff','67,60,224', '67,60,224'], _
		['Green','Aquisition','67,213,89', '67,213,89']]

Local $nodeSizes[3]=['90.0','80.0','50.0']
Local $nodeFontSizes[3]=['25.0','14.0','14.0']

$JsonfileStartStr='' & @CRLF & _
'{' & @CRLF & _
'  "format_version" : "1.0",' & @CRLF & _
'  "generated_by" : "cytoscape-3.1.0",' & @CRLF & _
'  "target_cytoscapejs_version" : "~2.1",' & @CRLF & _
'  "data" : {' & @CRLF & _
'    "selected" : true,' & @CRLF & _
'    "__Annotations" : [ ],' & @CRLF & _
'    "shared_name" : "Union",' & @CRLF & _
'    "SUID" : 322,' & @CRLF & _
'    "name" : "Union"' & @CRLF & _
'  },'

$elementStartStr='' & @CRLF & _
'"elements" : {'

$nodesStr='' & @CRLF & _
'    "nodes" : [ {' & @CRLF

$nodesEndStr='],' & @CRLF

$edgeStartStr='' & @CRLF & _
'    "edges" : [ ' & @CRLF

$edgeEndStr=']' & @CRLF

$elementEndStr = '' & @CRLF & _
'}'

$JsonfileEndStr = '' & @CRLF & _
'}'



;;---Select the source file

$date = @MDAY & "_" & @MON & "_" &  @YEAR & "_" & @HOUR & "_" & @MIN & "_" & @SEC
Local $message = "Please select the excel file with Competitor data."
Local $sFilePath1 = FileOpenDialog($message, @ScriptDir & "\", "Excel (*.xls;*.xlsx)", 1 + 2)


If @error Then
    MsgBox(4096, "", "No File(s) chosen")
Else
    $sFilePath1 = StringReplace($sFilePath1, "|", @CRLF)
    ;MsgBox(4096, "", "You chose " & $sFilePath1)
EndIf

;;---Select the Base company
Local $target = InputBox("Company","Enter company name","Bayer")

if $target = "" Then
	MsgBox(0,"Error","No target mentioned")
	Exit
EndIf

;;---Select the Base year
Local $targetYear = InputBox("Year","Enter the oldest year for analysis","2010")

if $targetYear = "" Then
	MsgBox(0,"Error","No $target Year mentioned")
	Exit
EndIf

;;---Pull the complete data in array for processing
Local $oExcel = _ExcelBookOpen($sFilePath1,0)

If @error = 1 Then
    MsgBox(0, "Error!", "Unable to Create the Excel Object")
    Exit
ElseIf @error = 2 Then
    MsgBox(0, "Error!", "File does not exist - Shame on you!")
    Exit
EndIf


Local $aArray = _ExcelReadSheetToArray($oExcel) ;Using Default Parameters
_ExcelBookClose($oExcel, 0)

_ArrayDisplay($aArray, "Array using Default Parameters")

$rows = $aArray[0][0]
$columns = $aArray[0][1]

ConsoleWrite("total number of lines in excel " & $rows & " X " &$columns & @CRLF) ;

;;-----Find out the column number for the prime columns - Company,First Category, Collaborator, Year

for $i=0 to $columns
	if $aArray[1][$i] == 'Company' Then
		$CompColumn=$i
		ConsoleWrite("$CompColumn " & $CompColumn&@CRLF)
	ElseIf $aArray[1][$i] == 'First Category' Then
		$FCatColumn=$i
		ConsoleWrite("$FCatColumn " & $FCatColumn&@CRLF)
	ElseIf $aArray[1][$i] == 'Collaborator' Then
		$ColabColumn=$i
		ConsoleWrite("$ColabColumn " & $ColabColumn&@CRLF)
	ElseIf $aArray[1][$i] == 'Year' Then
		$YearColumn=$i
		ConsoleWrite("$YearColumn " & $YearColumn&@CRLF)
	ElseIf $aArray[1][$i] == 'Detailed: Nature of Collaboration' Then
		$natureOfColabColumn=$i
		ConsoleWrite("$natureOfColabColumn " & $natureOfColabColumn&@CRLF)
	ElseIf $aArray[1][$i] == 'Primary Grouping: Nature of Collaboration' Then
		$PrimeNatureOfColabColumn=$i
		ConsoleWrite("$PrimeNatureOfColabColumn " & $PrimeNatureOfColabColumn&@CRLF)
	EndIf
Next

for $i=1 to UBound($aArray)-1
	if $aArray[$i][$CompColumn]='Bayer CropScience' Then
		ConsoleWrite($aArray[$i][$ColabColumn]&@CRLF&@CRLF)
	EndIf
Next

;;---opening the cyjs file for writing -

$file = FileOpen(@ScriptDir& "\OutputFiles\"& $target&"_"&$date&".cyjs", 9) ; which is similar to 9 + 8 (append + create dir)

If $file = -1 Then
    MsgBox(0, "Error", "Unable to open file.")
    Exit
EndIf

;;------writing the defaultsyntax that will be witten  at the starting of the file

FileWrite($file,$JsonfileStartStr)
FileWrite($file,$elementStartStr)
FileWrite($file,$nodesStr)

$Master_i=1

;;-----Building and inserting node data - :lvl 1:top most node
		$nodeWriter='' &@CRLF& _
	'	"data" : {' &@CRLF& _
	'        "id" : "'& $Master_i &'",' &@CRLF& _
	'        "cytoscape_alias_list" : [ ],' &@CRLF & _
	'        "node_fillColor" : "'& $nodeColor[0][2] &'",' &@CRLF& _
	'        "SUID" : "'& $Master_i &'",' &@CRLF& _
	'        "node_size" : "'& $nodeSizes[0] &'",' &@CRLF& _
	'        "node_fontSize" : "'& $nodeFontSizes[0] &'",' &@CRLF& _
	'        "node_borderColor" : "'& $nodeColor[0][3] &'",' &@CRLF& _
	'        "selected" : false,' &@CRLF& _
	'        "canonicalName" : "'& $target &'",' &@CRLF& _
	'        "name" : "'& $target &'",' &@CRLF& _
	'        "shared_name" : "'& $target &'"' &@CRLF& _
	'	}' &@CRLF& _
    '}'

	FileWrite($file,$nodeWriter)

;;----------Building the first category data:lvl 2:mid nodes-------------
dim $key[1],$keyDetails1[1][2]
for $i=1 to $rows
	if StringInStr($aArray[$i][$CompColumn],$target) > 0 and $aArray[$i][$FCatColumn]<> '' and int($aArray[$i][$YearColumn]) >= Int($targetYear) Then
		;ConsoleWrite ($aArray[$i][$FCatColumn])
		$key[UBound($key)-1]=$aArray[$i][$FCatColumn]
		ReDim $key[UBound($key)+1]

	EndIf
Next

ReDim $key[UBound($key)-1]
$key=_ArrayUnique($key)
;_ArrayDisplay($key)

for $i=1 to UBound($key)-1

	$keyDetails1[UBound($keyDetails1)-1][0]=$key[$i]
	$keyDetails1[UBound($keyDetails1)-1][1]=$i+1
	ReDim $keyDetails1[UBound($keyDetails1)+1][2]

Next

ReDim $keyDetails1[UBound($keyDetails1)-1][2]
;_ArrayDisplay($keyDetails1)


for $i = 1 to UBound($key)-1
	$Master_i=$Master_i+1
	$nodeWriter=',{' &@CRLF& _
		'	"data" : {' &@CRLF& _
		'        "id" : "'& $Master_i &'",' &@CRLF& _
		'        "cytoscape_alias_list" : [  ],' &@CRLF& _
		'        "SUID" : "'& $Master_i &'",' &@CRLF& _
		'        "node_size" : "'& $nodeSizes[1] &'",' &@CRLF& _
		'        "node_fontSize" : "'& $nodeFontSizes[1] &'",' &@CRLF& _
		'        "selected" : false,' &@CRLF& _
		'        "canonicalName" : "'& $key[$i] &'",' &@CRLF& _
		'        "name" : "'& $key[$i] &'",' &@CRLF& _
		'        "shared_name" : "'& $key[$i] &'"' &@CRLF& _
		'	}' &@CRLF& _
		'}'

	FileWrite($file,$nodeWriter)

Next

;;----------building the collaborator data:lvl 3:end nodes

dim $key[1],$keyDetails2[1][9]
for $i=1 to $rows
	if StringInStr($aArray[$i][$CompColumn],$target) > 0 and $aArray[$i][$ColabColumn]<> '' and int($aArray[$i][$YearColumn]) >= Int($targetYear) Then
		;ConsoleWrite ($aArray[$i][$FCatColumn])

		$key[UBound($key)-1]=$aArray[$i][$ColabColumn]

		$keyDetails2[UBound($keyDetails2)-1][0]=$Master_i+$i
		$keyDetails2[UBound($keyDetails2)-1][1]=$aArray[$i][$ColabColumn]
		$keyDetails2[UBound($keyDetails2)-1][2]=$aArray[$i][$FCatColumn]
		for $j=0 to UBound($keyDetails1)-1
			$Master_i=$Master_i+1
			if $keyDetails1[$j][0]= $aArray[$i][$FCatColumn] Then
				ConsoleWrite($keyDetails1[$j][0] &'|'& $aArray[$i][$FCatColumn]&@CRLF)
				$iIndex = $keyDetails1[$j][1]
				ConsoleWrite($iIndex)
				$keyDetails2[UBound($keyDetails2)-1][3]=$iIndex
			EndIf
		Next

		$keyDetails2[UBound($keyDetails2)-1][4]=$aArray[$i][$YearColumn]
		$keyDetails2[UBound($keyDetails2)-1][5]=$aArray[$i][$natureOfColabColumn]
		$keyDetails2[UBound($keyDetails2)-1][6]=$aArray[$i][$PrimeNatureOfColabColumn]

		if $aArray[$i][$PrimeNatureOfColabColumn]='Collaboration' Then
			$keyDetails2[UBound($keyDetails2)-1][7]=$nodeColor[1][2]
			$keyDetails2[UBound($keyDetails2)-1][8]=$nodeColor[1][3]

		ElseIf $aArray[$i][$PrimeNatureOfColabColumn]='Business Agreement' Then
			$keyDetails2[UBound($keyDetails2)-1][7]=$nodeColor[2][2]
			$keyDetails2[UBound($keyDetails2)-1][8]=$nodeColor[2][3]

		ElseIf $aArray[$i][$PrimeNatureOfColabColumn]='Divestment/Spinoff' Then
			$keyDetails2[UBound($keyDetails2)-1][7]=$nodeColor[3][2]
			$keyDetails2[UBound($keyDetails2)-1][8]=$nodeColor[3][3]

		ElseIf $aArray[$i][$PrimeNatureOfColabColumn]='Acquisition/Equity' Then
			$keyDetails2[UBound($keyDetails2)-1][7]=$nodeColor[4][2]
			$keyDetails2[UBound($keyDetails2)-1][8]=$nodeColor[4][3]

		EndIf

		ReDim $key[UBound($key)+1]
		ReDim $keyDetails2[UBound($keyDetails2)+1][9]
	EndIf
Next

_ArrayDisplay($keyDetails2)


for $i = 0 to UBound($keyDetails2)-2
	$Master_i=$Master_i+1
	$nodeWriter=',{' &@CRLF& _
		'	"data" : {' &@CRLF& _
		'        "id" : "'& $keyDetails2[$i][0] &'",' &@CRLF& _
		'        "cytoscape_alias_list" : [  "'& $keyDetails2[$i][1] &'" ],' &@CRLF& _
		'        "node_fillColor" : "'& $keyDetails2[$i][7] &'",' &@CRLF& _
		'        "Nature_of_Collaboration" : "'& $keyDetails2[$i][5] &'",' &@CRLF& _
		'        "Collaborator" : "'& $keyDetails2[$i][1] &'",' &@CRLF& _
		'        "SUID" : "'& $keyDetails2[$i][0] &'",' &@CRLF& _
		'        "node_borderColor" : "'& $keyDetails2[$i][8] &'",' &@CRLF& _
		'        "selected" : false,' &@CRLF& _
		'        "canonicalName" : "'& $key[$i] &'",' &@CRLF& _
		'        "node_labelPosition" : "W,E,c,0.00,0.00",' &@CRLF& _
		'        "name" : "'& $key[$i] &'",' &@CRLF& _
		'        "shared_name" : "'& $key[$i] &'"' &@CRLF& _
		'	}' &@CRLF& _
		'}'

	FileWrite($file,$nodeWriter)

Next

;;----------Node closing----------
FileWrite($file,$nodesEndStr)
;;-----------Edge Starting-------
FileWrite($file,$edgeStartStr)
;;-----------Edge level1-------

for $i = 0 to UBound($keyDetails1)-1
	$Master_i=$Master_i+1
	$nodeWriter='{' &@CRLF& _
		'	"data" : {' &@CRLF& _
		'        "id" : "'& $Master_i+$i &'",' &@CRLF& _
		'        "source" : "'& $keyDetails1[$i][1] &'",' &@CRLF& _
		'        "target" : "1",' &@CRLF& _
		'        "selected" : false,' &@CRLF& _
		'        "canonicalName" : "'& $target&" (PP) "&$keyDetails1[$i][0] &'",' &@CRLF& _
		'        "SUID" : "'& $Master_i+$i &'",' &@CRLF& _
		'        "name" : "'& $target&" (PP) "&$keyDetails1[$i][0] &'",' &@CRLF& _
		'        "interaction" : "PP",' &@CRLF& _
		'        "shared_interaction" : "PP",' &@CRLF& _
		'        "shared_name" : "'& $target&" (PP) "&$keyDetails1[$i][0] &'"' &@CRLF& _
		'		}' &@CRLF& _
		'	}'

	If $i > 0 Then
		$nodeWriter=','&$nodeWriter
	EndIf
	FileWrite($file,$nodeWriter)

Next

;;-----------Edges level2-------
for $i = 0 to UBound($keyDetails2)-2
	$Master_i=$Master_i+1
	$nodeWriter=',{' &@CRLF& _
		'	"data" : {' &@CRLF& _
		'        "id" : "'& $Master_i+$i &'",' &@CRLF& _
		'        "source" : "'& $keyDetails2[$i][0] &'",' &@CRLF& _
		'        "target" : "'& $keyDetails2[$i][3] &'",' &@CRLF& _
		'        "edge_fontSize" : "'& $nodeFontSizes[1] &'",' &@CRLF& _
		'        "selected" : false,' &@CRLF& _
		'        "canonicalName" : "'& $keyDetails2[$i][1]&" (DefaultEdge) "&$keyDetails2[$i][2] &'",' &@CRLF& _
		'        "SUID" : "'& $Master_i+$i &'",' &@CRLF& _
		'        "name" : "'& $keyDetails2[$i][1]&" (DefaultEdge) "&$keyDetails2[$i][2] &'",' &@CRLF& _
		'        "interaction" : "DefaultEdge",' &@CRLF& _
		'        "shared_interaction" : "DefaultEdge",' &@CRLF& _
		'        "shared_name" : "'& $keyDetails2[$i][1]&" (DefaultEdge) "&$keyDetails2[$i][2] &'",' &@CRLF& _
		'        "edge_label" : "'& $keyDetails2[$i][4] &'"' &@CRLF& _
		'		}' &@CRLF& _
		'	}'


	FileWrite($file,$nodeWriter)

Next

;;------Ending file
FileWrite($file,$edgeEndStr)
FileWrite($file,$elementEndStr)
FileWrite($file,$JsonfileEndStr)

;;----------Closing file
FileClose($file)



;;----Main function end-




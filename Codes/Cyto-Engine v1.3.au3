#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=n
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <Excel.au3>
#include <array.au3>

;;-----Statuc variable values

Local $nodeColor[6][6]=[ _
        ['Orange','MainNode','255,204,0', '255,204,0','#ffcc00','#ffcc00'], _
		['Default','Sub1Node','204,204,0', '204,204,0','#99ff99','#009900'], _
        ['Grey','Collaboration','217,211,157', '217,211,157','#d9d39d','#d9d39d'], _
		['LBlue','Business Agreement','36,216,255', '36,216,255','#24d8ff','#24d8ff'], _
		['DBlue','Divestment/Spinoff','67,60,224', '67,60,224','#433ce0','#433ce0'], _
		['Green','Aquisition','67,213,89', '67,213,89','#43d559','#43d559']]

Local $nodeSizes[3]=['90.0','80.0','25.0']
Local $nodeFontSizes[3]=['25.0','14.0','14.0']

$JsonfileStartStr='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' & @CRLF & _
'<graph label="Union" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#" xmlns:cy="http://www.cytoscape.org" xmlns="http://www.cs.rpi.edu/XGMML"  directed="1">' & @CRLF & _
'  <att name="documentVersion" value="1.1"/>' & @CRLF & _
'  <att name="networkMetadata">' & @CRLF & _
'    <rdf:RDF>' & @CRLF & _
'      <rdf:Description rdf:about="http://www.cytoscape.org/">' & @CRLF & _
'        <dc:type>Protein-Protein Interaction</dc:type>' & @CRLF & _
'        <dc:description>N/A</dc:description>' & @CRLF & _
'        <dc:identifier>N/A</dc:identifier>' & @CRLF & _
'        <dc:date>2014-03-30 01:46:39</dc:date>' & @CRLF & _
'        <dc:title>Union</dc:title>' & @CRLF & _
'        <dc:source>http://www.cytoscape.org/</dc:source>' & @CRLF & _
'        <dc:format>Cytoscape-XGMML</dc:format>' & @CRLF & _
'      </rdf:Description>' & @CRLF & _
'    </rdf:RDF>' & @CRLF & _
'  </att>' & @CRLF & _
'  <att type="string" name="backgroundColor" value="#ffffff"/>' & @CRLF & _
'  <att type="real" name="GRAPH_VIEW_ZOOM" value="0.7822052925967306"/>' & @CRLF & _
'  <att type="real" name="GRAPH_VIEW_CENTER_X" value="341.6669692993164"/>' & @CRLF & _
'  <att type="real" name="GRAPH_VIEW_CENTER_Y" value="390.6387941837311"/>' & @CRLF & _
'  <att type="boolean" name="NODE_SIZE_LOCKED" value="true"/>' & @CRLF & _
'  <att type="string" name="__layoutAlgorithm" value="grid" cy:hidden="true" cy:editable="false"/>'

Dim $nodesStr[18]=['' & @CRLF & _
'  <node label="', _
'" id="-', _
'    <att type="string" name="NODE_TYPE" value="DefaultNode"/>', _
'    <att type="string" name="canonicalName" value="', _
'    <att type="string" name="node.borderColor" value="', _
'    <att type="string" name="node.fillColor" value="', _
'    <att type="string" name="node.fontSize" value="', _
'    <att type="string" name="node.label" value="', _
'    <att type="string" name="node.labelPosition" value="NE,SW,c,0.00,0.00"/>', _
'    <att type="string" name="node.size" value="', _
'    <graphics type="ELLIPSE" x="280.4145812988281" y="621.4956665039062" width="2"  cy:nodeTransparency="0.5882352941176471" cy:nodeLabelFont="SansSerif-0-15" cy:borderLineType="solid" h="', _
'" w="', _
'" fill="', _
'" outline="', _
'" cy:nodeLabel="', _
'    <att type="list" name="cytoscape.alias.list">' & @CRLF & _
'      <att type="string" name="cytoscape.alias.list" value="', _
'    <att type="string" name="Collaborator" value="', _
'    <att type="string" name="Nature of Collaboration" value="']


Dim $edgeStr[9]=['' & @CRLF & _
'  <edge label="', _
'" source="-', _
'" target="-', _
'    <att type="string" name="canonicalName" value="', _
'    <att type="string" name="edge.fontSize" value="', _
'    <att type="string" name="edge.label" value="', _
'    <att type="string" name="interaction"  cy:editable="false" value="', _
'    <graphics width="2" fill="#333333" cy:sourceArrow="0" cy:targetArrow="0" cy:sourceArrowColor="#000000" cy:targetArrowColor="#000000" ' & _
'cy:edgeLabelFont="SanSerif-0-14" cy:edgeLineType="SOLID" cy:curved="STRAIGHT_LINES" cy:edgeLabel="', _
'    <att type="integer" name="Year" value="']

Dim $endStr[7]=['"/>','">','/>','    </att>','  </node>','  </edge>','</graph>']


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

;_ArrayDisplay($aArray, "Array using Default Parameters")

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

;;---opening the .xgmml file for writing -

$file = FileOpen(@ScriptDir& "\OutputFiles\"& $target&"_"&$date&".xgmml", 9) ; which is similar to 9 + 8 (append + create dir)

If $file = -1 Then
    MsgBox(0, "Error", "Unable to open file.")
    Exit
EndIf

;;------writing the defaultsyntax that will be witten  at the starting of the file

FileWrite($file,$JsonfileStartStr)
FileWrite($file,$nodesStr[0])

$Master_i=1

;;-----Building and inserting node data - :lvl 1:top most node
		$nodeWriter=$target& $nodesStr[1] &$Master_i & $endStr[1] &@CRLF& _
					$nodesStr[3] & $target & $endStr[0] &@CRLF& _
					$nodesStr[4] & $nodeColor[0][3] & $endStr[0] &@CRLF& _
					$nodesStr[5] & $nodeColor[0][2] & $endStr[0] &@CRLF& _
					$nodesStr[6] & $nodeFontSizes[0] & $endStr[0] &@CRLF& _
					$nodesStr[9] & $nodeSizes[0] & $endStr[0] &@CRLF& _
					$nodesStr[10] & $nodeSizes[0] & $nodesStr[11] & $nodeSizes[0] & $nodesStr[12] & $nodeColor[0][4] & $nodesStr[13] & $nodeColor[0][5] & $nodesStr[14] & $target & $endStr[0] &@CRLF& $endStr[4]

	FileWrite($file,$nodeWriter)

;;----------Building the first category data:lvl 2:mid nodes-------------

dim $key[1],$keyDetails1[1][2]
for $i=1 to $rows
	if StringInStr($aArray[$i][$CompColumn],$target) > 0 and $aArray[$i][$FCatColumn]<> '' and int($aArray[$i][$YearColumn]) >= Int($targetYear) Then
		;ConsoleWrite ($aArray[$i][$FCatColumn])
		$key[UBound($key)-1]=SpecialCharHandler($aArray[$i][$FCatColumn])
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
	$nodeWriter=$nodesStr[0]& $key[$i] & $nodesStr[1] &$Master_i & $endStr[1] &@CRLF& _
				$nodesStr[3] & $key[$i] & $endStr[0] &@CRLF& _
				$nodesStr[6] & $nodeFontSizes[1] & $endStr[0] &@CRLF& _
				$nodesStr[9] & $nodeSizes[1] & $endStr[0] &@CRLF& _
				$nodesStr[10] & $nodeSizes[1] & $nodesStr[11] & $nodeSizes[1] & $nodesStr[12] & $nodeColor[1][4] & $nodesStr[13] & $nodeColor[1][5] & $nodesStr[14] & $key[$i] & $endStr[0] &@CRLF& $endStr[4]

	FileWrite($file,$nodeWriter)

Next


;;----------building the collaborator data:lvl 3:end nodes

dim $key[1],$keyDetails2[1][11]
for $i=1 to $rows
	if StringInStr($aArray[$i][$CompColumn],$target) > 0 and $aArray[$i][$ColabColumn]<> '' and int($aArray[$i][$YearColumn]) >= Int($targetYear) Then
		;ConsoleWrite ($aArray[$i][$FCatColumn])

		$key[UBound($key)-1]=$aArray[$i][$ColabColumn]

		$keyDetails2[UBound($keyDetails2)-1][0]=$Master_i+1
		$keyDetails2[UBound($keyDetails2)-1][1]=SpecialCharHandler($aArray[$i][$ColabColumn])
		$keyDetails2[UBound($keyDetails2)-1][2]=SpecialCharHandler($aArray[$i][$FCatColumn])
		for $j=0 to UBound($keyDetails1)-1
			$Master_i=$Master_i+1
			if $keyDetails1[$j][0]= SpecialCharHandler($aArray[$i][$FCatColumn]) Then
				ConsoleWrite($keyDetails1[$j][0] &'|'& SpecialCharHandler($aArray[$i][$FCatColumn]&@CRLF))
				$iIndex = $keyDetails1[$j][1]
				ConsoleWrite($iIndex)
				$keyDetails2[UBound($keyDetails2)-1][3]=$iIndex
			EndIf
		Next

		$keyDetails2[UBound($keyDetails2)-1][4]=$aArray[$i][$YearColumn]
		$keyDetails2[UBound($keyDetails2)-1][5]=SpecialCharHandler($aArray[$i][$natureOfColabColumn])
		$keyDetails2[UBound($keyDetails2)-1][6]=SpecialCharHandler($aArray[$i][$PrimeNatureOfColabColumn])

		if $aArray[$i][$PrimeNatureOfColabColumn]='Collaboration' Then
			$keyDetails2[UBound($keyDetails2)-1][7]=$nodeColor[2][2]
			$keyDetails2[UBound($keyDetails2)-1][8]=$nodeColor[2][3]
			$keyDetails2[UBound($keyDetails2)-1][9]=$nodeColor[2][4]
			$keyDetails2[UBound($keyDetails2)-1][10]=$nodeColor[2][5]

		ElseIf $aArray[$i][$PrimeNatureOfColabColumn]='Business Agreement' Then
			$keyDetails2[UBound($keyDetails2)-1][7]=$nodeColor[3][2]
			$keyDetails2[UBound($keyDetails2)-1][8]=$nodeColor[3][3]
			$keyDetails2[UBound($keyDetails2)-1][9]=$nodeColor[3][4]
			$keyDetails2[UBound($keyDetails2)-1][10]=$nodeColor[3][5]

		ElseIf $aArray[$i][$PrimeNatureOfColabColumn]='Divestment/Spinoff' Then
			$keyDetails2[UBound($keyDetails2)-1][7]=$nodeColor[4][2]
			$keyDetails2[UBound($keyDetails2)-1][8]=$nodeColor[4][3]
			$keyDetails2[UBound($keyDetails2)-1][9]=$nodeColor[4][4]
			$keyDetails2[UBound($keyDetails2)-1][10]=$nodeColor[4][5]

		ElseIf $aArray[$i][$PrimeNatureOfColabColumn]='Acquisition/Equity' Then
			$keyDetails2[UBound($keyDetails2)-1][7]=$nodeColor[5][2]
			$keyDetails2[UBound($keyDetails2)-1][8]=$nodeColor[5][3]
			$keyDetails2[UBound($keyDetails2)-1][9]=$nodeColor[5][4]
			$keyDetails2[UBound($keyDetails2)-1][10]=$nodeColor[5][5]

		EndIf

		ReDim $key[UBound($key)+1]
		ReDim $keyDetails2[UBound($keyDetails2)+1][11]
	EndIf
Next

;_ArrayDisplay($keyDetails2)


for $i = 0 to UBound($keyDetails2)-2
	$Master_i=$Master_i+1
	$nodeWriter=$nodesStr[0]& $keyDetails2[$i][1] & $nodesStr[1] &$keyDetails2[$i][0] & $endStr[1] &@CRLF& _
	$nodesStr[2] &@CRLF& _
	$nodesStr[16] & $keyDetails2[$i][1] & $endStr[0] &@CRLF& _
	$nodesStr[17] & $keyDetails2[$i][5] & $endStr[0] &@CRLF& _
	$nodesStr[3] & $keyDetails2[$i][1] & $endStr[0] &@CRLF& _
	$nodesStr[4] & $keyDetails2[$i][8] & $endStr[0] &@CRLF& _
	$nodesStr[5] & $keyDetails2[$i][7] & $endStr[0] &@CRLF& _
	$nodesStr[6] & $nodeFontSizes[2] & $endStr[0] &@CRLF& _
	$nodesStr[7] & $keyDetails2[$i][1] & $endStr[0] &@CRLF& _
	$nodesStr[8] &@CRLF& _
	$nodesStr[9] & $nodeSizes[2] & $endStr[0] &@CRLF& _
	$nodesStr[15] & $keyDetails2[$i][1] & $endStr[0] &@CRLF& $endStr[3] &@CRLF& _
	$nodesStr[10] & $nodeSizes[2] & $nodesStr[11] & $nodeSizes[2] & $nodesStr[12] & $keyDetails2[$i][9] & $nodesStr[13] & $keyDetails2[$i][10] & $nodesStr[14] & $keyDetails2[$i][1] & $endStr[0] &@CRLF& $endStr[4]


	FileWrite($file,$nodeWriter)

Next

;;----------Node End----------
;;-----------Edge Starting-------
;;-----------Edge level 1-------

for $i = 0 to UBound($keyDetails1)-1
	$Master_i=$Master_i+1
	$nodeWriter=$edgeStr[0] & $target & " (PP) " & $keyDetails1[$i][0] & $edgeStr[1] & $keyDetails1[$i][1] & $edgeStr[2] & "1" & $endStr[1] &@CRLF& _
				$edgeStr[3] & $target & " (PP) " & $keyDetails1[$i][0] & $endStr[0] &@CRLF& _
				$edgeStr[6] & "pp" &$endStr[0] &@CRLF& _
				$edgeStr[7] & $endStr[0] &@CRLF& _
				$endStr[5]

	FileWrite($file,$nodeWriter)

Next

;;-----------Edges level2-------

for $i = 0 to UBound($keyDetails2)-2
	$Master_i=$Master_i+1
	$nodeWriter=$edgeStr[0] &  $keyDetails2[$i][1]&" (DefaultEdge) "&$keyDetails2[$i][2] & $edgeStr[1] & $keyDetails2[$i][0] & $edgeStr[2] & $keyDetails2[$i][3] & $endStr[1] &@CRLF& _
				$edgeStr[3] & $keyDetails2[$i][1]&" (DefaultEdge) "&$keyDetails2[$i][2] & $endStr[0] &@CRLF& _
				$edgeStr[4] & $nodeFontSizes[2] & $endStr[0]& @CRLF& _
				$edgeStr[5] & $keyDetails2[$i][4] & $endStr[0]& @CRLF& _
				$edgeStr[6] & "DefaultEdge" &$endStr[0] &@CRLF& _
				$edgeStr[8] & $keyDetails2[$i][4] & $endStr[0]& @CRLF& _
				$edgeStr[7] & $keyDetails2[$i][4]& $endStr[0] &@CRLF& _
				$endStr[5]

	FileWrite($file,$nodeWriter)

Next

;;------Ending edge
FileWrite($file,@CRLF & $endStr[6])
;;----------Closing file
FileClose($file)

;--------Future Development: Node Attribute File building-------------


;-------------------------------------------------


MsgBox(0,"Task : Completed","A new file has been generated." & @CRLF & "Path : " & @ScriptDir& "\OutputFiles\" & @CRLF &"File : "& $target&"_"&$date&".xgmml")
;;----Main function end-

func SpecialCharHandler($str)
	$str = StringReplace($str, "&", "&amp;")
	$str = StringReplace($str, ">", "&gt;")
	$str = StringReplace($str, "<", "&lt;")
	$str = StringReplace($str, "'", "&apos;")
	$str = StringReplace($str, '"', "&quot;")
	$str = StringReplace($str, '®', "")
	$str = StringReplace($str, 'â', "a")
	Return $str
EndFunc

Func URLEncode($urlText)
    $url = ""
    For $i = 1 To StringLen($urlText)
        $acode = Asc(StringMid($urlText, $i, 1))
        Select
            Case ($acode >= 48 And $acode <= 57) Or _
                    ($acode >= 65 And $acode <= 90) Or _
                    ($acode >= 97 And $acode <= 122)
                $url = $url & StringMid($urlText, $i, 1)
            Case $acode = 32
                $url = $url & "+"
			Case $acode = 46
                $url = $url & "."
            Case Else
                $url = $url & "%" & Hex($acode, 2)
        EndSelect
    Next
    Return $url
EndFunc   ;==>URLEncode
#NoTrayIcon
#RequireAdmin
#AutoIt3Wrapper_Run_Debug_Mode=Y

#Include <File.au3>
#Include <String.au3>
#include <Array.au3>

if IsAdmin() = 0 then
	ConsoleWriteError("Exiting，This program requires Local Admistrator rights")
	Exit
 EndIf
$ver="1.0.0"
FileInstall("tcpdump.exe", @TempDir & '\', 1)
Const $wbemFlagReturnImmediately = 0x10
const $wbemFlagForwardOnly = 0x20
$objWMIService = ObjGet("winmgmts:root\cimv2")
Global $oArray[0][13]
Global $headerArray[13][1]
Global $isformatCSV=False
Global $isformatCsvWithHeader=False
Global $isfilterVirtualCard=False
Global $isInteractive=False
$headerArray[0][0]="AdapterName"
$headerArray[1][0]="ProductName"
$headerArray[2][0]="MacAddress"
$headerArray[3][0]="SettingID"
$headerArray[4][0]="Description"
$headerArray[5][0]="SwitchName"
$headerArray[6][0]="SwitchPort"
$headerArray[7][0]="Vlan"
$headerArray[8][0]="SwitchIP"
$headerArray[9][0]="SwitchModel"
$headerArray[10][0]="SwitchDuplex"
$headerArray[11][0]="VTPMgmt"
$headerArray[12][0]="Hostname"

if $cmdline[0] > 0 Then
   for $i=1 to $cmdline[0]
	  $arg=StringLower($cmdline[$i])
	  Switch ($arg)
	  case "-h"
		 print_help()
		 Break
	  case "-?"
		 print_help()
		 Break
	  case "-help"
		 print_help()
		 Break
	  case "--help"
		 print_help()
		 Break
	  case "-csv"
		 $isformatCSV=True
	  case "-csvWithHeader"
		 $isformatCsvWithHeader=True
	  case "-noVirtual"
		 $isfilterVirtualcard=True
	  case "-i"
		 $isInteractive=True
	  case Else
		 ;l($arg)
		 print_help()
		 Break
	  EndSwitch
   Next
Else
    $isformatCSV=False
    $isfilterVirtualCard=False
EndIf
if not $isfilterVirtualCard then
   $colItems = $objWMIService.ExecQuery("select * from win32_networkadapter where netconnectionstatus=2 and ServiceName <> 'VMSMP' ", "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
Else
   $colItems=$objWMIService.ExecQuery("select * from win32_networkadapter where netconnectionstatus=2 and ServiceName <> 'VMSMP'  and  not productname like '%virtualbox%' ")
EndIf
$colItems2 = $objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration")

if not IsObj($colItems2) and not IsObj($colItems) then
	ConsoleWriteError("Run wmi query failed")
	OnExit()
	Exit(1)
Endif


;~ if $isInteractive Then
;~    local $c=0
;~    for $item in $colItems
;~ 	  l(
;~    Next
;~ EndIf

local $i=0,$j

For $objItem In $colItems
   ReDim $oArray[$i+1][13]
	for $j=0 to 12
		$oArray[$i][$j]=""
	Next
	$oArray[$i][0]=$objItem.NetConnectionID
	$oArray[$i][1]=$objItem.ProductName
	$oArray[$i][2]=$objItem.MACAddress
	$oArray[$i][12]=@ComputerName
	For $objItem2 In $colItems2
		If $objItem.Index = $objItem2.Index Then
		$oArray[$i][3]=$objItem2.SettingID
		$oArray[$i][4]=$objItem2.Description
		EndIf
	Next
	$i+=1
Next



GetAllNetworkCDP()

if $isformatCSV  and not $isformatCsvWithHeader Then
   formatcsv(False)
EndIf

if $isformatCsvWithHeader Then
   formatcsv(True)
EndIf


if not $isformatCsvWithHeader and not $isformatcsv Then
   formatlist()
EndIf

OnExit()

func l($str)
   ConsoleWrite($str &@CRLF)
EndFunc

func Print_help()
   l("Version    : " & $ver)
   l("Usage      : " & @ScriptName & "[Option]")
   l("Description: This script is to get cisco CDP info when you are using windows os")
   l("")
   l("Avalable options:")
   l(@TAB & StringFormat("%-25s","-h , /? , -help , --help")& "To get help messages")
   l(@TAB & StringFormat("%-25s","-csv") & "To get csv output,but no header")
   l(@TAB & StringFormat("%-25s","-csvWithHeader") & "To get csv output and with header")
   l(@TAB & StringFormat("%-25s","-noVirtual") & "To filter virutal net adapter like virtualbox")
   OnExit()
   Exit
EndFunc

func GetAllNetworkCDP()
   local $i
   For $i=0 to (ubound($oArray)-1)
	 local $ID= $oArray[$i][3]
	 Local $sOutput = _getDOSOutput(@TempDir & '\tcpdump.exe -i \Device\' & $ID & ' -nn -v -s 1500 -c 1 ether[20:2] == 0x2000 ')
	  if StringLen($sOutput) >0 Then
		 Local $lineArray = StringSplit($sOutput, @CRLF)
		 if not @error Then
			for $line=0 to UBound($lineArray)-1
			   If StringInStr($lineArray[$line], "Device-ID (0x01)") Then
				   $SwitchName = StringSplit($lineArray[$line], "'")
				   $SwitchName = StringUpper($SwitchName[2])
				   $oArray[$i][5] = $SwitchName
			   EndIf
			   If StringInStr($lineArray[$line], "Port-ID (0x03)") Then
				   $SwitchPort = StringSplit($lineArray[$line], "'")
				   $oArray[$i][6]=$SwitchPort[2]
			   EndIf
			   If StringInStr($lineArray[$line], "VLAN ID (0x0a)") Then
				   $VLAN = StringSplit($lineArray[$line], ":")
				   $VLAN = StringStripWS($VLAN[3],8)
				   $oArray[$i][7]= $VLAN
			   EndIf
			   If StringInStr($lineArray[$line], "Address (0x02)") Then
				   $SwitchIP = StringSplit($lineArray[$line], ")")
				   $SwitchIP = StringStripWS($SwitchIP[3],8)
				   $oArray[$i][8]=$SwitchIP
			   EndIf
			   If StringInStr($lineArray[$line], "Platform (0x06)") Then
				   $SwitchModel = StringSplit($lineArray[$line], "'")
				   $SwitchModel = StringTrimLeft (StringUpper($SwitchModel[2]), 6)
				   $oArray[$i][9]=$SwitchModel
			   EndIf
			   If StringInStr($lineArray[$line], "Duplex (0x0b)") Then
				   $Duplex = StringSplit($lineArray[$line], ":")
				   $Duplex = StringLower(StringStripWS($Duplex[3],8))
				   $Duplex = _StringProper($Duplex)
				   $oArray[$i][10]=$Duplex
			   EndIf
			   If StringInStr($lineArray[$line], "VTP Management Domain (0x09)") Then
				   $VTP = StringSplit($lineArray[$line], "'")
				   $oArray[$i][11]=$VTP[2]
				EndIf
			Next
		 Else
			ConsoleWriteError("get cdp info on \Device\" & $ID & " failed" & @CRLF)
		 EndIf
	  EndIf
   next
EndFunc

Func OnExit()
	If ProcessExists("tcpdump.exe") Then ProcessClose("tcpdump.exe")
	;if FileExists(@TempDir & "\CDP.txt") then FileDelete(@TempDir & "\CDP.txt")
	if FileExists(@TempDir & "\tcpdump.exe") then  FileDelete(@TempDir & "\tcpdump.exe")
EndFunc

Func _getDOSOutput($command)
   Local $text = '', $Pid = Run('"' & @ComSpec & '" /c ' & $command, '', @SW_HIDE, 2 + 4)
   if ProcessWaitClose($Pid ,61) then
   Else
	  If ProcessExists("tcpdump.exe") Then
		ProcessClose("tcpdump.exe")
	  Endif
   EndIf

   While 1
	  $text &= StdoutRead($Pid, False, False)
	  If @error Then ExitLoop
	  Sleep(10)
   WEnd
   Return $text
EndFunc


func formatCSV($withHeader=True)
   local $i,$j,$header[11]=[12,0,4,2,5,6,7,8,9,10,11]
   if $withHeader Then
	  local $headerStr=""
	  for $i=0 to (UBound($header)-1)
		 $headerStr=$headerStr & $headerArray[$header[$i]][0] & ","
	  Next
	  $headerStr=StringTrimRight($headerStr, StringLen(","))
	  ConsoleWrite($headerStr & @CRLF)
   EndIf

   for $i=0 to (UBound($oArray)-1)
	  local $line=""
	  for $j=0 to (UBound($header)-1)
		 $line=$line & $oArray[$i][$header[$j]] & ","
	  next
	  $line=StringTrimRight($line, StringLen(","))
	  ConsoleWrite($line & @CRLF)
   Next
EndFunc

func formatlist()
    local $i,$j,$header[11]=[12,0,4,2,5,6,7,8,9,10,11]
    for $i=0 to (UBound($oArray)-1)
	   ConsoleWrite("================================================" & @CRLF)
	   for $j=0 to (UBound($header)-1)
		  ConsoleWrite(StringFormat("%-20s : %s",$headerArray[$header[$j]][0],$oArray[$i][$header[$j]]) & @CRLF)
	   next
	   ConsoleWrite(@CRLF)
	next
EndFunc



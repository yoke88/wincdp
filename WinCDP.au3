;#NoTrayIcon
#RequireAdmin
#AutoIt3Wrapper_Run_Debug_Mode=Y
#Au3Stripper_Parameters=/mo
opt("TrayIconDebug",1)
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=cisco.ico
#AutoIt3Wrapper_Outfile=WinCDP.exe
#AutoIt3Wrapper_Compression=3
#AutoIt3Wrapper_Res_Description=Cisco Discovery Protocol Info Gather
#AutoIt3Wrapper_Res_Fileversion=0.0.1.4
#AutoIt3Wrapper_Res_LegalCopyright=Chris Hall 2010-2012
#AutoIt3Wrapper_Res_requestedExecutionLevel=requireAdministrator
#AutoIt3Wrapper_Res_Field=ProductName|WinCDP
#AutoIt3Wrapper_Res_Field=ProductVersion|1.4
#AutoIt3Wrapper_Res_Field=OriginalFileName|WinCDP.exe
#AutoIt3Wrapper_Run_AU3Check=n
#AutoIt3Wrapper_AU3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

$VER = "1.4"

#include <GuiConstantsEx.au3>
#include <WindowsConstants.au3>
#Include <File.au3>
#Include <String.au3>
#include <GuiButton.au3>
#include <ComboConstants.au3>
#include <FileConstants.au3>
#include <Array.au3>

$WinCDPVer = "WinCDP - v"& $VER  & @YEAR

if IsAdmin() = 0 then
	MsgBox(16,"Exiting","This program requires Local Admistrator rights")
	Exit
	EndIf
FileInstall("tcpdump.exe", @TempDir & '\', 1)
GUISetIcon("cisco.ico")

$log =@TempDir & "\CDP.txt"
$wbemFlagReturnImmediately = 0x10
$wbemFlagForwardOnly = 0x20
$colItems = ""
$strComputer = "localhost"
$Output=""
$Nic_Friend =""
$Hardware=""
$IData=""
SplashTextOn("Please Wait","Enumerating Network Cards via WMI...", 300, 50)
$objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\CIMV2")
$colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapter where NetConnectionID is not null", "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
$colItems2 = $objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration")
Global $oArray[0][13]
;Global $headerArray[13][1]

;$headerArray[0][0]="AdapterName"
;$headerArray[1][0]="ProductName"
;$headerArray[2][0]="MacAddress"
;$headerArray[3][0]="SettingID"
;$headerArray[4][0]="Description"
;$headerArray[5][0]="SwitchName"
;$headerArray[6][0]="SwitchPort"
;$headerArray[7][0]="Vlan"
;$headerArray[8][0]="SwitchIP"
;$headerArray[9][0]="SwitchModel"
;$headerArray[10][0]="SwitchDuplex"
;$headerArray[11][0]="VTPMgmt"
;$headerArray[12][0]="Hostname"

If IsObj($colItems) then
   local $i=0,$j
   For $objItem In $colItems
	  ReDim $oArray[$i+1][13]
	   for $j=0 to 12
		   $oArray[$i][$j]=""
	   Next
	   $oArray[$i][0]="(" & $i & ")" & $objItem.NetConnectionID
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
	;_ArrayDisplay($oArray, "2D display transposed", Default, 1)
Else
   Msgbox(0,"WMI Output","No WMI Objects Found for class: " & "Win32_NetworkAdapterConfiguration" )
Endif


SplashOff()
GUICreate("Cisco Discovery for Windows" & " Runing at " & @ComputerName, 550, 400, (@DesktopWidth - 550) / 2, (@DesktopHeight - 400) / 2, $WS_OVERLAPPEDWINDOW + $WS_VISIBLE + $WS_CLIPSIBLINGS)
GUICtrlCreateGroup("Selection ", 15, 10, 520, 110)
GUICtrlCreateLabel("Adapter:", 30, 35, 100, 20)
$Nic_Friendly = GUICtrlCreateCombo("",145,33,350,20, $CBS_DROPDOWNLIST)
Local $aExtract = _ArrayExtract($oArray,-1,-1,0,0)
GUICtrlSetData($Nic_Friendly ,_ArrayToString($aExtract))
GUICtrlCreateLabel("Network Card:", 30, 62, 100, 20)
$Get = GUICtrlCreateButton("Get CDP Data", 120, 85, 100)
$Save = GUICtrlCreateButton("Save CDP Data", 260, 85, 100)
$Cancel = GUICtrlCreateButton("Cancel", 400, 85, 100)
If RegRead("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System", "EnableLUA") > 0 Then
    GUICtrlSetImage($Get, "imageres.dll", -2, 0)
	 _GUICtrlButton_SetShield($Get)
  EndIf

GUICtrlCreateGroup("Port Info", 15, 130, 530, 170)
GUICtrlCreateLabel("Switch Name:", 30, 150, 100, 20)
GUICtrlCreateLabel("Port:", 30, 180, 100, 20)
GUICtrlCreateLabel("VLAN:", 30, 210, 100, 20)
GUICtrlCreateLabel("Switch IP:", 30, 240, 100, 20)
GUICtrlCreateLabel("Model:", 280, 180, 100, 20)
GUICtrlCreateLabel("Port Duplex:", 280, 210, 100, 20)
GUICtrlCreateLabel("VTP Mgmt Domain:", 280, 240, 100, 20)
GUICtrlCreateLabel("Mac Address:", 30, 270, 100, 20)
GUICtrlCreateGroup("Status ", 15, 310, 530, 65)
GUICtrlCreateLabel($WinCDPVer, 350, 380, 200, 20)

GUISetState()
	While 1
		Switch GUIGetMsg()

		Case $Nic_Friendly
			$Nic_Friend = GUICtrlRead ($Nic_Friendly)
			$index=int(StringSplit(stringmid($Nic_Friend,2),")")[1])
			$hardware=$oArray[$index][1]
			GUICtrlCreateLabel($Hardware, 145, 62, 350, 20)
			ClearResults()
			GUICtrlCreateLabel($oArray[$index][2], 140, 270, 120, 20)
	    Case $Get
			If GUICtrlRead($Nic_Friendly) = "" Then
			   MsgBox(64,"Invalid Selection", "Please select a network card using the dropdown")
			   ContinueLoop
			EndIf
			GetCDP($Nic_Friendly)
		Case $GUI_EVENT_CLOSE
			OnExit()
			ExitLoop
		Case $Cancel
			OnExit()
			ExitLoop
		Case $Save
			SaveData()
		Case Else
				;;;
		EndSwitch
	WEnd
Exit
Func GetCDP($Nic_Friendly)
		$Nic_Friend = GUICtrlRead ($Nic_Friendly)
	    $index=int(StringSplit(stringmid($Nic_Friend,2),")")[1])
		$SaveFile = FileOpen(@TempDir & "\SaveCDP.txt", 2)
		GUICtrlSetState($Get, $GUI_DISABLE)
		GUICtrlSetState($Save, $GUI_DISABLE)
		ClearResults()
		FileWriteLine($SaveFile, $Nic_Friend & " (" & $Hardware & ") is connected to:")
		FileWriteLine($SaveFile, "------------------------------------------------------")
		$ID = $oArray[$index][3]
		$TCPDmpPID = Run(@ComSpec & " /c " & @TempDir & '\tcpdump.exe -i \Device\' & $ID & ' -nn -v -s 1500 -c 1 ether[20:2] == 0x2000 >%temp%\CDP_OUT.txt', "", @SW_HIDE)
		$Secs = 1
		$Status1 = GUICtrlCreateLabel("Running ... May take up to 60 seconds between CDP announcements ...", 40, 327, 400, 20 )
		FileWriteLine($SaveFile, "MacAddress:	" & $oArray[$index][2])
		GUICtrlCreateLabel($oArray[$index][2], 140, 270, 120, 20)
		FileWriteLine($SaveFile, "ComputerName:	" & @ComputerName)
		$iBegin = TimerInit()
		Do
			$msg = GUIGetMsg()
			If $msg = $Cancel Then
				ProcessClose("tcpdump.exe")
				ExitLoop
			EndIf
			If Ceiling(TimerDiff($iBegin)) = ($Secs * 1000) or Ceiling(TimerDiff($iBegin)) > ($Secs * 1000) Then
				GUICtrlCreateLabel(Round($Secs,0) & " Seconds Elapsed", 40, 347, 200, 20 )
				$Secs = $Secs + 1
			EndIf
			$TCPDmpPID = ProcessExists($TCPDmpPID)
		Until $TCPDmpPID = "0" Or TimerDiff($iBegin) > 60000
		GUICtrlDelete($Status1)
		GUICtrlCreateLabel("", 240, 337, 100, 20 )
		GUICtrlCreateLabel("", 210, 317, 200, 20)
	    $file = FileOpen(@TempDir & "\CDP_OUT.txt")
	    $end = _FileCountLines(@TempDir & "\CDP_OUT.txt")
	  If $end > 0 Then
		 $line = 0
		 Do
			 If StringInStr(FileReadLine($file, $line), "Device-ID (0x01)") Then
				 $SwitchName = StringSplit(FileReadLine($file, $line), "'")
				 $SwitchName = StringUpper($SwitchName[2])
				 GUICtrlCreateLabel($SwitchName, 140, 150, 180, 20)
				 $oArray[$index][5] = $SwitchName
				 FileWriteLine($SaveFile, "Switch Name:	" & $SwitchName)
			 EndIf
			 If StringInStr(FileReadLine($file, $line), "Port-ID (0x03)") Then
				 $SwitchPort = StringSplit(FileReadLine($file, $line), "'")
				 GUICtrlCreateLabel($SwitchPort[2], 140, 180, 120, 20)
				 $oArray[$index][6]=$SwitchPort[2]
				 FileWriteLine($SaveFile, "Switch Port:	" & $SwitchPort[2])
			 EndIf
			 If StringInStr(FileReadLine($file, $line), "VLAN ID (0x0a)") Then
				 $VLAN = StringSplit(FileReadLine($file, $line), ":")
				 $VLAN = StringStripWS($VLAN[3],8)
				 GUICtrlCreateLabel($VLAN, 140, 210, 120, 20)
				 $oArray[$index][7]= $VLAN
				 FileWriteLine($SaveFile, "VLAN ID:	" & $VLAN)
			 EndIf
			 If StringInStr(FileReadLine($file, $line), "Address (0x02)") Then
				 $SwitchIP = StringSplit(FileReadLine($file, $line), ")")
				 $SwitchIP = StringStripWS($SwitchIP[3],8)
				 GUICtrlCreateLabel($SwitchIP, 140, 240, 120, 20)
				 $oArray[$index][8]=$SwitchIP
				 FileWriteLine($SaveFile, "Switch IP:	" & $SwitchIP)
			 EndIf
			 If StringInStr(FileReadLine($file, $line), "Platform (0x06)") Then
				 $SwitchModel = StringSplit(FileReadLine($file, $line), "'")
				 $SwitchModel = StringTrimLeft (StringUpper($SwitchModel[2]), 6)
				 GUICtrlCreateLabel($SwitchModel, 390, 180, 120, 20)
				 $oArray[$index][9]=$SwitchModel
				 FileWriteLine($SaveFile, "Switch Model:	" & $SwitchModel)
			 EndIf
			 If StringInStr(FileReadLine($file, $line), "Duplex (0x0b)") Then
				 $Duplex = StringSplit(FileReadLine($file, $line), ":")
				 $Duplex = StringLower(StringStripWS($Duplex[3],8))
				 $Duplex = _StringProper($Duplex)
				 GUICtrlCreateLabel($Duplex, 390, 210, 120, 20)
				 $oArray[$index][10]=$Duplex
				 FileWriteLine($SaveFile, "Switch Duplex:	" & $Duplex)
			 EndIf
			 If StringInStr(FileReadLine($file, $line), "VTP Management Domain (0x09)") Then
				 $VTP = StringSplit(FileReadLine($file, $line), "'")
				 GUICtrlCreateLabel($VTP[2], 390, 240, 120, 20)
				 $oArray[$i][10]=$VTP[2]
				 FileWriteLine($SaveFile, "VTP Mgmt:	" & $VTP[2])
			 EndIf

			 $line = $line + 1
		 Until $line = $end
	  Else
		  If ProcessExists("tcpdump.exe") Then ProcessClose("tcpdump.exe")
		  GUICtrlCreateLabel("NO CDP DATA FOUND ... !", 210, 317, 150, 20)
		  FileClose($SaveFile)
		  FileDelete(@TempDir & "\SaveCDP.txt")
	  EndIf
	FileClose($SaveFile)
	FileClose($file)
	FileDelete(@TempDir & "\CDP_OUT.txt")
	GUICtrlSetState($Get, $GUI_ENABLE)
	GUICtrlSetState($Save, $GUI_ENABLE)
EndFunc

Func ClearResults()
   GUICtrlCreateLabel("", 140, 150, 180, 20)
   GUICtrlCreateLabel("", 140, 180, 120, 20)
   GUICtrlCreateLabel("", 140, 210, 120, 20)
   GUICtrlCreateLabel("", 140, 240, 120, 20)
   GUICtrlCreateLabel("", 390, 180, 120, 20)
   GUICtrlCreateLabel("", 390, 210, 120, 20)
   GUICtrlCreateLabel("", 390, 240, 120, 20)
   GUICtrlCreateLabel("", 140, 270, 120, 20)
EndFunc

 Func SaveData()
	 If FileExists(@TempDir & "\SaveCDP.txt") = 0 Then Return
	 $UserSave = FileSaveDialog("Save CDP Data to","::{20D04FE0-3AEA-1069-A2D8-08002B30309D}","Text Documents (*.txt)", 2)
	 If $UserSave = "" Then Return
	 If StringInStr($UserSave, ".txt") = 0 Then $UserSave = $UserSave & ".txt"
	 FileOpen($UserSave, 1)
	 FileWrite($UserSave, FileRead(@TempDir & "\SaveCDP.txt") & @CRLF)
	 FileClose($UserSave)
 EndFunc

 Func OnExit()
	 If ProcessExists("tcpdump.exe") Then ProcessClose("tcpdump.exe")
	 FileClose($log)
	 FileDelete(@TempDir & "\CDP.txt")
	 FileDelete(@TempDir & "\tcpdump.exe")
	 FileDelete(@TempDir & "\SaveCDP.txt")
  EndFunc

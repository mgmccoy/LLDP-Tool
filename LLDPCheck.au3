#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=network.ico
#AutoIt3Wrapper_Outfile=LLDPCheck.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Description=App for checking LLDP details on link
#AutoIt3Wrapper_Res_Fileversion=2.2.0.0
#AutoIt3Wrapper_Res_LegalCopyright=
#AutoIt3Wrapper_Res_requestedExecutionLevel=requireAdministrator
#AutoIt3Wrapper_Res_Field=ProductName|LLDP Checker
#AutoIt3Wrapper_Res_Field=ProductVersion|1.0
#AutoIt3Wrapper_Res_Field=OriginalFileName|LLDPCheck.exe
#AutoIt3Wrapper_Run_AU3Check=n
#AutoIt3Wrapper_AU3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#NoTrayIcon
#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=network.ico
#AutoIt3Wrapper_Outfile=LLDPCheck.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Description=LLDP Information Display
#AutoIt3Wrapper_Res_Fileversion=1.0
#AutoIt3Wrapper_Res_LegalCopyright=
#AutoIt3Wrapper_Res_requestedExecutionLevel=requireAdministrator
#AutoIt3Wrapper_Res_Field=ProductName|LLDP Check
#AutoIt3Wrapper_Res_Field=ProductVersion|1.0
#AutoIt3Wrapper_Res_Field=OriginalFileName|LLDPCheck.exe
#AutoIt3Wrapper_Run_AU3Check=n
#AutoIt3Wrapper_AU3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;===================================================================================================================================================================
; Based on LDWin - Link Discovery for Windows - Chris Hall 2010-2015.  Modified for ENS by Martin McCoy DCC
;===================================================================================================================================================================
$VER = "1.0"
#include <GuiConstantsEx.au3>
#include <WindowsConstants.au3>
#include <File.au3>
#include <String.au3>
#include <GuiButton.au3>
#include <ComboConstants.au3>


$WinLDPVer = "LLDP Information Display - v" & $VER
If IsAdmin() = 0 Then
	MsgBox(16, "Exiting", "This program requires Local Admistrator rights")
	Exit
EndIf
FileInstall("tcpdump.exe", @TempDir & '\', 1)
GUISetIcon("network.ico")

$LDWinHelp = 99999
$donate = ""
$gotit = ""
$log = FileOpen(@TempDir & "\LinkData.txt", 2)
$wbemFlagReturnImmediately = 0x10
$wbemFlagForwardOnly = 0x20
$colItems = ""
$strComputer = "localhost"
$Output = ""
$Nic_Friend = ""
$Hardware = ""
$IData = ""
SplashTextOn("Please Wait", "Enumerating Network Cards via WMI...", 300, 50)
$objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\CIMV2")
$colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapter", "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
If IsObj($colItems) Then
	For $objItem In $colItems
		FileWriteLine($log, "[" & $objItem.NetConnectionID & "]")
		FileWriteLine($log, "ProductName=" & $objItem.ProductName)
		$value = $objItem.NetConnectionID
		If StringLen($value) > 1 Then $Output = $Output & $value & "|"
		$colItems2 = $objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration", "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
		For $objItem2 In $colItems2
			If $objItem.Index = $objItem2.Index Then
				FileWriteLine($log, "SettingID=" & $objItem2.SettingID)
				FileWriteLine($log, "IPAddress=" & $objItem2.IPAddress(0))
				FileWriteLine($log, "MACAddress=" & $objItem2.MACAddress)
			EndIf
		Next
	Next
Else
	MsgBox(0, "WMI Output", "No WMI Objects Found for class: " & "Win32_NetworkAdapterConfiguration")
EndIf
SplashOff()
$gui = GUICreate("LLDP Information Display", 550, 423, (@DesktopWidth - 550) / 2, (@DesktopHeight - 423) / 2, $WS_OVERLAPPEDWINDOW + $WS_VISIBLE + $WS_CLIPSIBLINGS)
GUICtrlCreateGroup("Selection ", 15, 10, 520, 133)
GUICtrlCreateLabel("Network Interface:", 30, 35, 100, 20)
$Nic_Friendly = GUICtrlCreateCombo("", 145, 33, 350, 20, $CBS_DROPDOWNLIST)
GUICtrlSetData(-1, $Output)
GUICtrlCreateLabel("NIC Details:", 30, 62, 100, 20)
GUICtrlCreateLabel("MAC Address:", 30, 89, 100, 20)
GUICtrlCreateLabel("IPv4 Address:", 280, 89, 100, 20)
$Get = GUICtrlCreateButton("Get LLDP Data", 200, 108, 100)
$Save = GUICtrlCreateButton("Save LLDP Data", 310, 108, 100)
$Cancel = GUICtrlCreateButton("Cancel", 420, 108, 100)

If RegRead("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System", "EnableLUA") > 0 Then
	GUICtrlSetImage($Get, "imageres.dll", -2, 0)
	_GUICtrlButton_SetShield($Get)
EndIf
GUICtrlCreateGroup("LLDP Neighbour Information ", 15, 153, 520, 160)
GUICtrlCreateLabel("Device Name:", 30, 183, 70, 20)
GUICtrlCreateLabel("Port ID:", 30, 213, 70, 20)
GUICtrlCreateLabel("VLAN ID:", 30, 243, 75, 20)
GUICtrlCreateLabel("IPv4 Address:", 30, 273, 90, 20)
GUICtrlCreateLabel("Device Model:", 280, 213, 70, 20)
GUICtrlCreateLabel("Port Duplex:", 280, 243, 70, 20)
GUICtrlCreateLabel("VTP Mgmt Domain:", 280, 273, 95, 20)
GUICtrlCreateGroup("Retrieval Status ", 15, 323, 520, 65)
GUICtrlCreateLabel($WinLDPVer, 20, 398, 450, 20)
GUISetState()
While 1
	$aMsg = GUIGetMsg(1)
	Switch $aMsg[1]
		Case $gui
			Switch $aMsg[0]
				Case $Nic_Friendly
					$Nic_Friend = GUICtrlRead($Nic_Friendly)
					$IData = IniReadSection(@TempDir & "\LinkData.txt", $Nic_Friend)
					$Hardware = $IData[1][1]
					$IPAddr = $IData[3][1]
					$MAC = $IData[4][1]
					GUICtrlCreateLabel($Hardware, 145, 62, 350, 20)
					GUICtrlCreateLabel($IPAddr, 390, 89, 120, 20)
					GUICtrlCreateLabel($MAC, 145, 89, 120, 20)
					ClearResults()
				Case $Get
					If GUICtrlRead($Nic_Friendly) = "" Then
						MsgBox(64, "Invalid Selection", "Please select a network card using the dropdown")
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
			EndSwitch

	EndSwitch
WEnd
Exit

Func GetCDP($Nic_Friendly)
	$SaveFile = FileOpen(@TempDir & "\SaveData.txt", 2)
	GUICtrlSetState($Get, $GUI_DISABLE)
	GUICtrlSetState($Save, $GUI_DISABLE)
	ClearResults()
	FileWriteLine($SaveFile, $Nic_Friend)
	FileWriteLine($SaveFile, "(" & $Hardware & ", " & $MAC & ", " & $IPAddr & ") is connected to:")
	FileWriteLine($SaveFile, "------------------------------------------------------")
	$ID = $IData[2][1]

	;******** DIAG MODE ********
	$TCPDmpPID = Run(@ComSpec & " /c " & @TempDir & '\tcpdump.exe -i \Device\' & $ID & ' -nn -v -s 1500 -c 1 (ether[12:2]==0x88cc or ether[20:2]==0x2000) >%temp%\Data_Out.txt', "", @SW_HIDE)
	;$TCPDmpPID = "0"
	;******** DIAG MODE ********
	$Secs = 1
	$Status1 = GUICtrlCreateLabel("Waiting for LLDP announcement.  May take up to 60 seconds.", 120, 343, 350, 20)
	$iBegin = TimerInit()
	Do
		$msg = GUIGetMsg()
		If $msg = $Cancel Then
			ProcessClose("tcpdump.exe")
			ExitLoop
		EndIf
		If Ceiling(TimerDiff($iBegin)) = ($Secs * 1000) Or Ceiling(TimerDiff($iBegin)) > ($Secs * 1000) Then
			GUICtrlCreateLabel(Round($Secs, 0) & " Seconds elapsed", 240, 363, 100, 20)
			$Secs = $Secs + 1
		EndIf
		$TCPDmpPID = ProcessExists($TCPDmpPID)
	Until $TCPDmpPID = "0" Or TimerDiff($iBegin) > 60000
	GUICtrlDelete($Status1)
	GUICtrlCreateLabel("", 240, 360, 100, 20)
	GUICtrlCreateLabel("", 210, 350, 200, 20)
	$file = FileOpen(@TempDir & "\Data_Out.txt")
	$end = _FileCountLines(@TempDir & "\Data_Out.txt")
	If $end > 0 Then
		$line = 0
		Do
			;===== CDP ==========================================================================
			If StringInStr(FileReadLine($file, $line), "Device-ID (0x01)") Then
				$SwitchName = StringSplit(FileReadLine($file, $line), "'")
				$SwitchName = StringUpper($SwitchName[2])
				GUICtrlCreateLabel($SwitchName, 140, 183, 370, 20)
				FileWriteLine($SaveFile, "Switch Name:	" & $SwitchName)
			EndIf
			If StringInStr(FileReadLine($file, $line), "Port-ID (0x03)") Then
				$SwitchPort = StringSplit(FileReadLine($file, $line), "'")
				GUICtrlCreateLabel($SwitchPort[2], 140, 213, 120, 20)
				FileWriteLine($SaveFile, "Switch Port:	" & $SwitchPort[2])
			EndIf
			If StringInStr(FileReadLine($file, $line), "VLAN ID (0x0a)") Then
				$VLAN = StringSplit(FileReadLine($file, $line), ":")
				$VLAN = StringStripWS($VLAN[3], 8)
				GUICtrlCreateLabel($VLAN, 140, 243, 120, 20)
				FileWriteLine($SaveFile, "VLAN ID:	" & $VLAN)
			EndIf
			If StringInStr(FileReadLine($file, $line), "Address (0x02)") Then
				$SwitchIP = StringSplit(FileReadLine($file, $line), ")")
				$SwitchIP = StringStripWS($SwitchIP[3], 8)
				GUICtrlCreateLabel($SwitchIP, 140, 273, 120, 20)
				FileWriteLine($SaveFile, "Switch IP:	" & $SwitchIP)
			EndIf
			If StringInStr(FileReadLine($file, $line), "Platform (0x06)") Then
				$SwitchModel = StringSplit(FileReadLine($file, $line), "'")
				$SwitchModel = StringUpper($SwitchModel[2])
				If StringInStr($SwitchModel, "CISCO") Then
					$SwitchModel = StringTrimLeft(StringUpper($SwitchModel), 6)
				EndIf
				GUICtrlCreateLabel($SwitchModel, 390, 213, 120, 20)
				FileWriteLine($SaveFile, "Switch Model:	" & $SwitchModel)
			EndIf
			If StringInStr(FileReadLine($file, $line), "Duplex (0x0b)") Then
				$Duplex = StringSplit(FileReadLine($file, $line), ":")
				$Duplex = StringLower(StringStripWS($Duplex[3], 8))
				$Duplex = _StringProper($Duplex)
				GUICtrlCreateLabel($Duplex, 390, 243, 120, 20)
				FileWriteLine($SaveFile, "Switch Duplex:	" & $Duplex)
			EndIf
			If StringInStr(FileReadLine($file, $line), "VTP Management Domain (0x09)") Then
				$VTP = StringSplit(FileReadLine($file, $line), "'")
				GUICtrlCreateLabel($VTP[2], 390, 273, 120, 20)
				FileWriteLine($SaveFile, "VTP Mgmt:	" & $VTP[2])
			EndIf
			;===== LLDP =========================================================================
			If StringInStr(FileReadLine($file, $line), "System Name TLV (5)") Then
				$SwitchName = StringSplit(FileReadLine($file, $line), ":")
				$SwitchName = StringStripWS(StringUpper($SwitchName[2]), 3)
				GUICtrlCreateLabel($SwitchName, 140, 183, 370, 20)
				FileWriteLine($SaveFile, "Switch Name:	" & $SwitchName)
			EndIf
			If StringInStr(FileReadLine($file, $line), "Chassis ID TLV (1)") Then
				$SwitchName = StringSplit(FileReadLine($file, $line), ":")
				If @error Then
					$nextline = $line + 1
					$SwitchName = StringSplit(FileReadLine($file, $nextline), ":")
					$SwitchNameSize = UBound($SwitchName)
					If $SwitchNameSize > 3 Then
						$SWconcat = ""
						For $i = 2 to $SwitchNameSize - 1
							$SWconcat = ($SWconcat & $SwitchName[$i] & ":")
						Next
						$SwitchName = StringTrimRight($SWconcat, 1)
					Else
						$SwitchName = $SwitchName[2]
					EndIf
				Else
					$SwitchName = $SwitchName[2]
				EndIf
				$SwitchName = StringStripWS($SwitchName, 3)
				GUICtrlCreateLabel("", 140, 183, 180, 20)
				GUICtrlCreateLabel($SwitchName, 140, 183, 370, 20)
				FileWriteLine($SaveFile, "Switch Name:	" & $SwitchName)
			EndIf
			If StringInStr(FileReadLine($file, $line), "Port ID TLV (2)") Then
				$SwitchPort = StringSplit(FileReadLine($file, $line), ":")
				If @error Then
					$nextline = $line + 1
					$SwitchPort = StringSplit(FileReadLine($file, $nextline), ":")
					$SwitchPort = $SwitchPort[2]
					GUICtrlCreateLabel("", 140, 213, 120, 20)
					GUICtrlCreateLabel($SwitchPort, 140, 213, 120, 40)

				Else
					$SwitchPort = $SwitchPort[2]
				EndIf
				$SwitchPort = StringStripWS($SwitchPort, 3)
				GUICtrlCreateLabel($SwitchPort, 140, 213, 120, 40)
				FileWriteLine($SaveFile, "Switch Port:	" & $SwitchPort)
			EndIf
			If StringInStr(FileReadLine($file, $line), "Port Description TLV (4)") Then
				$SwitchPort = StringSplit(FileReadLine($file, $line), ":")
				$SwitchPort = StringStripWS($SwitchPort[2], 3)
				FileWriteLine($SaveFile, "Switch Port:	" & $SwitchPort)
			EndIf
			If StringInStr(FileReadLine($file, $line), "port vlan id (PVID)") Then
				$VLAN = StringSplit(FileReadLine($file, $line), ":")
				$VLAN = StringStripWS($VLAN[2], 3)
				GUICtrlCreateLabel($VLAN, 140, 243, 120, 20)
				FileWriteLine($SaveFile, "VLAN ID:	" & $VLAN)
			EndIf
			If StringInStr(FileReadLine($file, $line), "Management Address TLV (8)") Then
				$SwitchIP = StringSplit(FileReadLine($file, $line), ":")
				If @error Then
					$nextline = $line + 1
					$SwitchIP = StringSplit(FileReadLine($file, $nextline), ":")
				Else
					$SwitchIP = $SwitchIP[2]
				EndIf
				$SwitchIP = StringStripWS(StringUpper($SwitchIP[2]), 3)
				GUICtrlCreateLabel($SwitchIP, 140, 273, 120, 20)
				FileWriteLine($SaveFile, "Switch IP:	" & $SwitchIP)
			EndIf
			If StringInStr(FileReadLine($file, $line), "System Description TLV (6)") Then
				$SwitchModel = StringSplit(FileReadLine($file, $line), ":")
				If @error Then
					$nextline = $line + 1
					$SwitchModel = FileReadLine($file, $nextline)
				Else
					$SwitchModel = $SwitchModel[2]
				EndIf
				$SwitchModel = StringStripWS($SwitchModel, 3)
				GUICtrlCreateLabel($SwitchModel, 390, 213, 120, 40)
				FileWriteLine($SaveFile, "Switch Model:	" & $SwitchModel)
			EndIf
			$line = $line + 1
		Until $line = $end
	Else
		If ProcessExists("tcpdump.exe") Then ProcessClose("tcpdump.exe")
		GUICtrlCreateLabel("NO LINK DATA FOUND ... !", 210, 348, 150, 20)
		FileClose($SaveFile)
		FileDelete(@TempDir & "\SaveData.txt")
	EndIf
	FileClose($SaveFile)
	FileClose($file)
	FileDelete(@TempDir & "\Data_Out.txt")
	GUICtrlSetState($Get, $GUI_ENABLE)
	GUICtrlSetState($Save, $GUI_ENABLE)
EndFunc   ;==>GetCDP

Func ClearResults()
	GUICtrlCreateLabel("", 140, 183, 180, 20)
	GUICtrlCreateLabel("", 140, 213, 120, 20)
	GUICtrlCreateLabel("", 140, 243, 120, 20)
	GUICtrlCreateLabel("", 140, 273, 120, 20)
	GUICtrlCreateLabel("", 390, 213, 120, 20)
	GUICtrlCreateLabel("", 390, 243, 120, 20)
	GUICtrlCreateLabel("", 390, 273, 120, 20)
EndFunc   ;==>ClearResults

Func SaveData()
	If FileExists(@TempDir & "\SaveData.txt") = 0 Then Return
	$UserSave = FileSaveDialog("Save Link Data to", "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}", "Text Documents (*.txt)", 2)
	If $UserSave = "" Then Return
	If StringInStr($UserSave, ".txt") = 0 Then $UserSave = $UserSave & ".txt"
	FileOpen($UserSave, 1)
	FileWrite($UserSave, FileRead(@TempDir & "\SaveData.txt") & @CRLF)
	FileClose($UserSave)
EndFunc   ;==>SaveData

Func OnExit()
	If ProcessExists("tcpdump.exe") Then ProcessClose("tcpdump.exe")
	FileClose($log)
	FileDelete(@TempDir & "\LinkData.txt")
	FileDelete(@TempDir & "\tcpdump.exe")
	FileDelete(@TempDir & "\SaveData.txt")
EndFunc   ;==>OnExit



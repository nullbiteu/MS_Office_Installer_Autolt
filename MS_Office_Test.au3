#NoTrayIcon
#RequireAdmin
#include <GUIConstantsEx.au3>
#include <Array.au3>
#include <Crypt.au3>

Global $AudienceId, $LCID, $Version, $Product, $Lang, $ExcludedApps, $ProductsToAdd, $downloading = False, $ScriptDir
Global $dD[5][6]
Global $dDid
Global $hDownload
Global $hTimer, $fDiff
If StringRight(@ScriptDir, 1) = "\" Then
	$ScriptDir = StringTrimRight(@ScriptDir, 1)
Else
	$ScriptDir = @ScriptDir
EndIf

GUICreate("Microsoft Office - yobis_Test", 420, 300)

$guiVersion = GUICtrlCreateGroup("", 10, 10, 220, 275)
GUICtrlCreateLabel("Version:", 20, 44, 60, 20)
$guiProduct = GUICtrlCreateCombo("", 80, 40, 140, 20)
GUICtrlCreateLabel("Kanal:", 20, 74, 60, 20)
$guiChannel = GUICtrlCreateCombo("", 80, 70, 140, 20)
GUICtrlCreateLabel("Sprache:", 20, 104, 60, 20)
$guiLang = GUICtrlCreateCombo("", 80, 100, 140, 20)
GUICtrlCreateLabel("Anzeige:", 20, 134, 60, 20)
$guiDisplay = GUICtrlCreateCombo("", 80, 130, 140, 20)
GUICtrlCreateLabel("Telemetrie:", 20, 164, 60, 20)
$guiTelemetry = GUICtrlCreateCombo("", 80, 160, 140, 20)
$guiInstallOffice = GUICtrlCreateButton("Office Herunterladen / Installieren", 20, 215, 200, 25)
$guiDownloadOffice = GUICtrlCreateButton("Nur Herunterladen", 20, 245, 200, 25)
GUICtrlCreateGroup("", -99, -99, 1, 1)

GUICtrlCreateGroup("Was soll installiert werden?", 240, 10, 130, 300)
$idCBword       = GUICtrlCreateCheckbox("Word", 250, 30, 80, 25)
GUICtrlSetState(-1, $GUI_CHECKED)
$idCBexcel      = GUICtrlCreateCheckbox("Excel", 250, 50, 80, 25)
GUICtrlSetState(-1, $GUI_CHECKED)
$idCBoutlook    = GUICtrlCreateCheckbox("Outlook", 250, 70, 80, 25)
GUICtrlSetState(-1, $GUI_CHECKED)
$idCBpowerpoint = GUICtrlCreateCheckbox("PowerPoint", 250, 90, 80, 25)
GUICtrlSetState(-1, $GUI_CHECKED)
$idCBaccess     = GUICtrlCreateCheckbox("Access", 250, 110, 80, 25)
$idCBpublisher  = GUICtrlCreateCheckbox("Publisher", 250, 130, 80, 25)
$idCBonenote    = GUICtrlCreateCheckbox("OneNote", 250, 150, 80, 25)
$idCBonedrive   = GUICtrlCreateCheckbox("OneDrive", 250, 170, 80, 25)
$idCBlync       = GUICtrlCreateCheckbox("Lync", 250, 190, 80, 25)
$idCBvisio      = GUICtrlCreateCheckbox("Visio", 250, 210, 100, 25)
$idCBproject    = GUICtrlCreateCheckbox("Project", 250, 230, 100, 25)
$idCBseco       = GUICtrlCreateCheckbox("KMS Seco", 250, 250, 100, 25)
GUICtrlCreateGroup("", -99, -99, 1, 1)

GUICtrlSetData($guiProduct, "ProPlus2019Retail|ProPlus2019Volume","ProPlus2019Retail")
GUICtrlSetData($guiChannel, "Insiders::DevMain|Insiders::CC|Production::CC|Production::MEC|Production::LTSC","Production::CC")
GUICtrlSetData($guiLang, "ar-sa|bg-bg|cs-cz|da-dk|de-de|el-gr|en-us|es-es|et-ee|fi-fi|fr-fr|he-il|hi-in|hr-hr|hu-hu|id-id|it-it|ja-jp|kk-kz|ko-kr|lt-lt|lv-lv|ms-my|nb-no|nl-nl|pl-pl|pt-br|pt-pt|ro-ro|ru-ru|sk-sk|sl-si|sr-latn-rs|sv-se|th-th|tr-tr|uk-ua|vi-vn|zh-cn|zh-tw","de-de")
GUICtrlSetData($guiDisplay, "True|False", "True")
GUICtrlSetData($guiTelemetry, "Set Disable|Do not set", "Set Disable")

_KMSsecoVisibility()
SettingChannel()
SettingVersion()
GUICtrlSetData($guiVersion, "Office 2019 " & $Version)
GUISetState(@SW_SHOW)

While 1
	Switch GUIGetMsg()
		Case $GUI_EVENT_CLOSE
			ExitLoop
		Case $guiInstallOffice
			InstallOffice()
		Case $guiDownloadOffice
			DownloadOffice()
		Case $guiProduct
			_KMSsecoVisibility()
		Case $guiChannel
			SettingChannel()
			SettingVersion1()
			GUICtrlSetData($guiVersion, "Office 2019 " & $Version)
	EndSwitch
	If $downloading And TimerDiff($hTimer) > 1000 Then
		If $dD[$dDid][2] Then
			$dDid+=1
		Else
			If $dD[$dDid][3] Then
				If InetGetInfo($hDownload,2) Then
					If InetGetInfo($hDownload,4)<>0 Then
						FileDelete($dD[$dDid][1])
						MsgBox(0, "Error", "Download failed.")
						Exit
					Else
						InetClose($hDownload)
						GUICtrlSetData($guiDownloadOffice, "100 % "&$dDid+1&"/5")
						If $dDid=3 Or $dDid=4 Then
							Sleep(500)
							GUICtrlSetData($guiDownloadOffice, "SHA256 "&$dDid+1&"/5")
							If FileExists(@TempDir&"\abfall_temp") Then DirRemove(@TempDir&"\abfall_temp", 1)
							DirCreate(@TempDir&"\abfall_temp")
							If FileExists($dD[$dDid][1]) Then
								$dHash = _Crypt_HashFile($dD[$dDid][1], $CALG_SHA_256)
								RunWait('EXPAND "'&$dD[$dDid-3][1]&'" /f:'&$dD[$dDid][5]&' "'&@TempDir&'\abfall_temp"', @ScriptDir, @SW_HIDE)
								$hFileOpen = FileOpen(@TempDir&"\abfall_temp\"&$dD[$dDid][5], 32)
								$sFileRead = FileRead($hFileOpen)
								If $dHash <> "0x"&$sFileRead Then
									MsgBox(0, "Error1", $dHash & @CRLF & "0x"&$sFileRead)
									MsgBox(0, "Error", "Download failed.")
									Exit
								EndIf
							Else
								MsgBox(0, "Error", "Download failed.")
								Exit
							EndIf
							DirRemove(@TempDir&"\abfall_temp", 1)
						EndIf
						$dDid+=1
					EndIf
				Else
					GUICtrlSetData($guiDownloadOffice, Floor((InetGetInfo($hDownload,0)/$dD[$dDid][4])*100)&" % "&$dDid+1&"/5")
				EndIf
			Else
				$dD[$dDid][4] = InetGetSize($dD[$dDid][0],1)
				$hDownload = InetGet($dD[$dDid][0], $dD[$dDid][1],1,1)
				$dD[$dDid][3]=True
				GUICtrlSetData($guiDownloadOffice, "0 % "&$dDid+1&"/5")
			EndIf
		EndIf
		If $dDid = 5 Then
			$downloading = False
			GUICtrlSetData($guiDownloadOffice, "Only Download")
			GUICtrlSetState($guiInstallOffice, $GUI_ENABLE)
			GUICtrlSetState($guiDownloadOffice, $GUI_ENABLE)
			$hTimer = 0
		EndIf
		$hTimer = TimerInit()
	Endif
WEnd

Func SettingChannel()
	$AudienceData = GUICtrlRead($guiChannel)
	If $AudienceData = "Insiders::DevMain"   Then $AudienceId = "5440FD1F-7ECB-4221-8110-145EFAA6372F"
	If $AudienceData = "Insiders::CC"        Then $AudienceId = "64256AFE-F5D9-4F86-8936-8840A6A4F5BE"
	If $AudienceData = "Production::CC"      Then $AudienceId = "492350F6-3A01-4F97-B9C0-C7C6DDF67D60"
	If $AudienceData = "Production::MEC"     Then $AudienceId = "55336B82-A18D-4DD6-B5F6-9E5095C314A6"
	If $AudienceData = "Production::LTSC"    Then $AudienceId = "F2E724C1-748F-4B47-8FB8-8E0D210E9208"
EndFunc

Func SettingLang()
	$Lang = GUICtrlRead($guiLang)
	If $Lang = "ar-sa" Then $LCID = "1025"
	If $Lang = "bg-bg" Then $LCID = "1026"
	If $Lang = "cs-cz" Then $LCID = "1029"
	If $Lang = "da-dk" Then $LCID = "1030"
	If $Lang = "de-de" Then $LCID = "1031"
	If $Lang = "el-gr" Then $LCID = "1032"
	If $Lang = "en-us" Then $LCID = "1033"
	If $Lang = "es-es" Then $LCID = "3082"
	If $Lang = "et-ee" Then $LCID = "1061"
	If $Lang = "fi-fi" Then $LCID = "1035"
	If $Lang = "fr-fr" Then $LCID = "1036"
	If $Lang = "he-il" Then $LCID = "1037"
	If $Lang = "hi-in" Then $LCID = "1081"
	If $Lang = "hr-hr" Then $LCID = "1050"
	If $Lang = "hu-hu" Then $LCID = "1038"
	If $Lang = "id-id" Then $LCID = "1057"
	If $Lang = "it-it" Then $LCID = "1040"
	If $Lang = "ja-jp" Then $LCID = "1041"
	If $Lang = "kk-kz" Then $LCID = "1087"
	If $Lang = "ko-kr" Then $LCID = "1042"
	If $Lang = "lt-lt" Then $LCID = "1063"
	If $Lang = "lv-lv" Then $LCID = "1062"
	If $Lang = "ms-my" Then $LCID = "1086"
	If $Lang = "nb-no" Then $LCID = "1044"
	If $Lang = "nl-nl" Then $LCID = "1043"
	If $Lang = "pl-pl" Then $LCID = "1045"
	If $Lang = "pt-br" Then $LCID = "1046"
	If $Lang = "pt-pt" Then $LCID = "2070"
	If $Lang = "ro-ro" Then $LCID = "1048"
	If $Lang = "ru-ru" Then $LCID = "1049"
	If $Lang = "sk-sk" Then $LCID = "1051"
	If $Lang = "sl-si" Then $LCID = "1060"
	If $Lang = "sr-latn-rs" Then $LCID = "9242"
	If $Lang = "sv-se" Then $LCID = "1053"
	If $Lang = "th-th" Then $LCID = "1054"
	If $Lang = "tr-tr" Then $LCID = "1055"
	If $Lang = "uk-ua" Then $LCID = "1058"
	If $Lang = "vi-vn" Then $LCID = "1066"
	If $Lang = "zh-cn" Then $LCID = "2052"
	If $Lang = "zh-tw" Then $LCID = "1028"
EndFunc

Func InstallOffice()
	GUICtrlSetState($guiInstallOffice, $GUI_DISABLE)
	GUICtrlSetState($guiDownloadOffice, $GUI_DISABLE)
	sleep(500)
	If _IsChecked($idCBseco) Then _KMSsecoInstall()
	If FileExists(@TempDir&"\abfall_temp") Then DirRemove(@TempDir&"\abfall_temp", 1)
	DirCreate(@TempDir&"\abfall_temp")
	SettingLang()
	SettingProduct()
	SettingExcludedApps()
	If Not FileExists(@CommonFilesDir & "\microsoft shared\ClickToRun\OfficeClickToRun.exe") Then
		InetGet("https://officecdn.microsoft.com/pr/"&$AudienceId&"/Office/Data/"&$Version&"/i640.cab", @TempDir & "\abfall_temp\i640.cab", 1)
		DirCreate(@CommonFilesDir & "\microsoft shared\ClickToRun")
		RunWait('EXPAND "'&@TempDir&'\abfall_temp\i640.cab" /f:* "'&@CommonFilesDir&'\microsoft shared\ClickToRun"', @ScriptDir, @SW_HIDE)
	EndIf
	RunWait (@CommonFilesDir & "\microsoft shared\ClickToRun\OfficeClickToRun.exe" & _
		" cdnbaseurl.16=http://officecdn.microsoft.com/pr/"&$AudienceId & _
		" baseurl.16=http://officecdn.microsoft.com/pr/"&$AudienceId & _
		" version.16="&$Version & _
		" platform=x64" & _
		" culture="&$Lang & _
		" displaylevel=" & GUICtrlRead($guiDisplay) & _
		" productstoadd="&$ProductsToAdd & _
		" deliverymechanism="&$AudienceId & _
		$ExcludedApps)
	DirRemove(@TempDir&"\abfall_temp", 1)
	$Telemetry = GUICtrlRead($guiTelemetry)
	If $Telemetry = "Set Disable" Then _DisableTelemetry()
	GUICtrlSetState($guiInstallOffice, $GUI_ENABLE)
	GUICtrlSetState($guiDownloadOffice, $GUI_ENABLE)
EndFunc

Func DownloadOffice()
	GUICtrlSetState($guiInstallOffice, $GUI_DISABLE)
	GUICtrlSetState($guiDownloadOffice, $GUI_DISABLE)
	sleep(500)
	$downloading = True
	SettingLang()
	SettingProduct()
	If Not FileExists($ScriptDir&"\Office\Data\"&$Version) Then DirCreate($ScriptDir&"\Office\Data\"&$Version)
	$dD[0][1]=$ScriptDir&"\Office\Data\"&$Version&"\s64"&$LCID&".cab"
	If Not FileExists($dD[0][1]) Then
		$dD[0][0]="http://officecdn.microsoft.com/pr/"&$AudienceId&"/Office/Data/"&$Version&"/s64"&$LCID&".cab"
		$dD[0][2]=False
	Else
		$dD[0][2]=True
	EndIf
	$dD[1][1]=$ScriptDir&"\Office\Data\"&$Version&"\s640.cab"
	If Not FileExists($dD[1][1]) Then
		$dD[1][0]="http://officecdn.microsoft.com/pr/"&$AudienceId&"/Office/Data/"&$Version&"/s640.cab"
		$dD[1][2]=False
	Else
		$dD[1][2]=True
	EndIf
	$dD[2][1]=$ScriptDir&"\Office\Data\"&$Version&"\i640.cab"
	If Not FileExists($dD[2][1]) Then
		$dD[2][0]="http://officecdn.microsoft.com/pr/"&$AudienceId&"/Office/Data/"&$Version&"/i640.cab"
		$dD[2][2]=False
	Else
		$dD[2][2]=True
	EndIf
	$dD[3][1]=$ScriptDir & "\Office\Data\"&$Version&"\stream.x64."&$Lang&".dat"
	If Not FileExists($dD[3][1]) Then
		$dD[3][0]="http://officecdn.microsoft.com/pr/"&$AudienceId&"/Office/Data/"&$Version&"/stream.x64."&$Lang&".dat"
		$dD[3][5]="stream.x64."&$Lang&".hash"
		$dD[3][2]=False
	Else
		$dD[3][2]=True
	EndIf
	$dD[4][1]=$ScriptDir&"\Office\Data\"&$Version&"\stream.x64.x-none.dat"
	If Not FileExists($dD[4][1]) Then
		$dD[4][0]="http://officecdn.microsoft.com/pr/"&$AudienceId&"/Office/Data/"&$Version&"/stream.x64.x-none.dat"
		$dD[4][5]="stream.x64.x-none.hash"
		$dD[4][2]=False
	Else
		$dD[4][2]=True
	EndIf
	$hTimer = TimerInit()
	$dDid=0
EndFunc

Func SettingProduct()
	$Product = GUICtrlRead($guiProduct)
	$ProductsToAdd = $Product&".16_"&$Lang&"_x-none"
	If $Product = "ProPlus2019Volume" And _IsChecked($idCBvisio) Then $ProductsToAdd &= "|VisioPro2019Volume.16_"&$Lang&"_x-none"
	If $Product = "ProPlus2019Retail" And _IsChecked($idCBvisio) Then $ProductsToAdd &= "|VisioPro2019Retail.16_"&$Lang&"_x-none"
	If $Product = "ProPlus2019Volume" And _IsChecked($idCBproject) Then $ProductsToAdd &= "|ProjectPro2019Volume.16_"&$Lang&"_x-none"
	If $Product = "ProPlus2019Retail" And _IsChecked($idCBproject) Then $ProductsToAdd &= "|ProjectPro2019Retail.16_"&$Lang&"_x-none"
EndFunc

Func SettingVersion()
	Local $arr[5] = ["5440FD1F-7ECB-4221-8110-145EFAA6372F", "64256AFE-F5D9-4F86-8936-8840A6A4F5BE", "492350F6-3A01-4F97-B9C0-C7C6DDF67D60", "55336B82-A18D-4DD6-B5F6-9E5095C314A6", "F2E724C1-748F-4B47-8FB8-8E0D210E9208"]
	For $i = 0 To 4
		$txt = BinaryToString(InetRead("https://mrodevicemgr.officeapps.live.com/mrodevicemgrsvc/api/v2/C2RReleaseData?audienceFFN=" & $arr[$i], 1))
		$aArrayAvailableBuild = StringRegExp($txt, '(?i)"AvailableBuild": "(.*?)"', 2)
		Assign (StringReplace($arr[$i], "-", ""), $aArrayAvailableBuild[1],2)
	Next
	$Version = Eval(StringReplace($AudienceId, "-", ""))
EndFunc

Func SettingVersion1()
	$Version = Eval(StringReplace($AudienceId, "-", ""))
EndFunc

Func SettingExcludedApps()
	$ExcludedApps = ""
	Local $aArrayExcludedApps[0]
	If Not _IsChecked($idCBword)       Then _ArrayAdd($aArrayExcludedApps, "word")
	If Not _IsChecked($idCBexcel)      Then _ArrayAdd($aArrayExcludedApps, "excel")
	If Not _IsChecked($idCBoutlook)    Then _ArrayAdd($aArrayExcludedApps, "outlook")
	If Not _IsChecked($idCBpowerpoint) Then _ArrayAdd($aArrayExcludedApps, "powerpoint")
	If Not _IsChecked($idCBaccess)     Then _ArrayAdd($aArrayExcludedApps, "access")
	If Not _IsChecked($idCBpublisher)  Then _ArrayAdd($aArrayExcludedApps, "publisher")
	If Not _IsChecked($idCBonenote)    Then _ArrayAdd($aArrayExcludedApps, "onenote")
	If Not _IsChecked($idCBonedrive)   Then _ArrayAdd($aArrayExcludedApps, "onedrive")
	                                        _ArrayAdd($aArrayExcludedApps, "groove")
	If Not _IsChecked($idCBlync)       Then _ArrayAdd($aArrayExcludedApps, "lync")
	If UBound($aArrayExcludedApps) Then $ExcludedApps = " " & $Product & ".excludedapps.16=" & _ArrayToString($aArrayExcludedApps, ",")
	If $Product = "ProPlus2019Volume" And _IsChecked($idCBvisio) And Not _IsChecked($idCBonedrive) Then $ExcludedApps &= " VisioPro2019Volume.excludedapps.16=onedrive,groove"
	If $Product = "ProPlus2019Retail" And _IsChecked($idCBvisio) And Not _IsChecked($idCBonedrive) Then $ExcludedApps &= " VisioPro2019Retail.excludedapps.16=onedrive,groove"
	If $Product = "ProPlus2019Volume" And _IsChecked($idCBproject) And Not _IsChecked($idCBonedrive) Then $ExcludedApps &= " ProjectPro2019Volume.excludedapps.16=onedrive,groove"
	If $Product = "ProPlus2019Retail" And _IsChecked($idCBproject) And Not _IsChecked($idCBonedrive) Then $ExcludedApps &= " ProjectPro2019Retail.excludedapps.16=onedrive,groove"
EndFunc

Func _IsChecked($idControlID)
    Return BitAND(GUICtrlRead($idControlID), $GUI_CHECKED) = $GUI_CHECKED
EndFunc

Func _KMSsecoVisibility()
	If Not FileExists (@SystemDir & "\SppExtComObjHook.dll") And (FileExists($ScriptDir & "\SppExtComObjHook.dll") Or (FileExists($ScriptDir & "\x64.dll"))) And GUICtrlRead($guiProduct) = "ProPlus2019Volume" Then
		GUICtrlSetState($idCBseco, $GUI_SHOW + $GUI_CHECKED)
	Else
		GUICtrlSetState($idCBseco, $GUI_HIDE + $GUI_UNCHECKED)
	EndIf
EndFunc

Func _KMSsecoInstall()
	If FileExists($ScriptDir & "\SppExtComObjHook.dll") Then
		FileCopy ($ScriptDir & "\SppExtComObjHook.dll", @SystemDir, 1)
	ElseIf FileExists($ScriptDir & "\x64.dll") Then
		FileCopy ($ScriptDir & "\x64.dll", @SystemDir & "\SppExtComObjHook.dll", 1)
	EndIf
	RegDelete("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\55c92734-d682-4d71-983e-d6ec3f16059f")
	RegDelete("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663")
	RegWrite("HKLM\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\Software Protection Platform", "NoGenTicket", "REG_DWORD", 1)
	RegWrite("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\SppExtComObj.exe", "VerifierDlls", "REG_SZ", "SppExtComObjHook.dll")
	RegWrite("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\SppExtComObj.exe", "GlobalFlag", "REG_DWORD", 256)
	$oWMIService = ObjGet("winmgmts:\\.\root\cimv2")
	If IsObj($oWMIService) Then
		$oCollection = $oWMIService.ExecQuery("SELECT Version FROM SoftwareLicensingService")
		If IsObj($oCollection) Then
			For $oItem In $oCollection
				$oItem.SetKeyManagementServiceMachine("172.16.16.16")
				$oItem.SetKeyManagementServicePort("1688")
			Next
		EndIf
	EndIf
EndFunc

Func _DisableTelemetry()
	RegWrite("HKLM\Software\Policies\Microsoft\Office\16.0\osm", "Enablelogging", "REG_DWORD", 0)
	RegWrite("HKLM\Software\Policies\Microsoft\Office\16.0\osm", "EnableUpload", "REG_DWORD", 0)
	RegWrite("HKLM\Software\Microsoft\Office\Common\ClientTelemetry", "DisableTelemetry", "REG_DWORD", 1)
	RegWrite("HKCU\Software\Policies\Microsoft\Office\16.0\osm", "Enablelogging", "REG_DWORD", 0)
	RegWrite("HKCU\Software\Policies\Microsoft\Office\16.0\osm", "EnableUpload", "REG_DWORD", 0)
	RegWrite("HKCU\Software\Microsoft\Office\Common\ClientTelemetry", "DisableTelemetry", "REG_DWORD", 1)
EndFunc

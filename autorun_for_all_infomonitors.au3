#pragma compile(ProductVersion, 0.1)
#pragma compile(UPX, true)
#pragma compile(CompanyName, 'ООО Клиника ЛМС')
#pragma compile(FileDescription, Скрипт для запуска ПО на инфомониторах и настройке обоев)
#pragma compile(LegalCopyright, Грашкин Павел Павлович - Нижний Новгород - 31-555 - nn-admin@nnkk.budzdorov.su)
#pragma compile(ProductName, autorun_for_all_infomonitors)

#include <AD.au3>
#include <File.au3>
#include <FileConstants.au3>


#Region ==========================    Check for temp folder and create log    ==========================
Local $error = False
Local $oMyError = ObjEvent("AutoIt.Error","HandleComError")
Local $messageToSend = ""
Local $current_pc_name = @ComputerName ;"NNKK-MON-1-101";
Local $tempFolder = StringSplit(@SystemDir, "\")[1] & "\Temp\"
Local $errStr = "===ERROR=== "
ConsoleWrite("current_pc_name: " & $current_pc_name & @CRLF)

If Not FileExists($tempFolder) Then
   If Not DirCreate($tempFolder) Then
	  $error = True
	  ConsoleWrite($errStr & "Cannot create folder " & $tempFolder & @CRLF)
   EndIf
EndIf

Local $logFilePath = $tempFolder & "autorun_for_all_monitors.log"
Local $logFile = FileOpen($logFilePath, $FO_OVERWRITE)
ToLog($current_pc_name)
ToLog(@CRLF & "---Check for temp folder and create log---")

If $logFile = -1 Then
   $error = True
   ToLog($errStr & "Cannot create log file at " & $tempFolder)
EndIf
#EndRegion

#Region ==========================    Variables    ==========================
Local $iniFile = @ScriptDir & "\autorun_for_all_infomonitors_settings.ini"
Local $generalSection = "general"
Local $mailSection = "mail"

Local $server_backup = "172.16.6.6"
Local $login_backup = "autorun_for_all_monitors@nnkk.budzdorov.su"
Local $password_backup = "giodxalf"
Local $to_backup = "nn-admin@nnkk.budzdorov.su"
Local $send_email_backup = "1"
#EndRegion

#Region ==========================    Reading the main settings    ==========================
ToLog(@CRLF & "---Reading the main settings---")

Local $server = IniRead($iniFile, $mailSection, "server", $server_backup)
Local $login = IniRead($iniFile, $mailSection, "login", $login_backup)
Local $password = IniRead($iniFile, $mailSection, "password", $password_backup)
Local $to = IniRead($iniFile, $mailSection, "to", $to_backup)
Local $send_email = IniRead($iniFile, $mailSection, "send_email", $send_email_backup)

If $send_email = "" Then
   $send_email = "1"
EndIf

If Not FileExists($iniFile) Then
   ToLog($errStr & "Cannot find settings file: " & $iniFile)
   SendEmail()
EndIf

ToLog("server: " & $server)
ToLog("login: " & $login)
ToLog("password: " & $password)
ToLog("to: " & $to)
ToLog("send_mail: " & $send_email)

Local $infoscreen_restart_path = IniRead($iniFile, $generalSection, "infoscreen_restart_path", "")
Local $wallpaper_path = IniRead($iniFile, $generalSection, "wallpaper_path", "")
Local $default_infoscreen_path = IniRead($iniFile, $generalSection, "default_infoscreen_path", "")
Local $infoscreen_standard = IniRead($iniFile, $generalSection, "infoscreen_standard", "")
Local $infoscreen_timetable = IniRead($iniFile, $generalSection, "infoscreen_timetable", "")
#EndRegion

#Region ==========================    Check for the settings error    ==========================
ToLog(@CRLF & "---Check for the settings error---")
If $infoscreen_restart_path = "" Then
   $error = True
   ToLog($errStr & "Cannot find key: infoscreen_restart_path")
EndIf

If $wallpaper_path = "" Then
   $error = True
   ToLog($errStr & "Cannot find key: wallpaper_path")
EndIf

If $default_infoscreen_path = "" Then
   $error = True
   ToLog($errStr & "Cannot find key: default_infoscreen_path")
EndIf

If $infoscreen_standard = "" Then
   $error = True
   ToLog($errStr & "Cannot find key: infoscreen_standard")
EndIf

If $infoscreen_timetable = "" Then
   $error = True
   ToLog($errStr & "Cannot find key: infoscreen_timetable")
EndIf

If $error Then
   SendEmail()
EndIf

ToLog("infoscreen_path: " & $infoscreen_restart_path)
ToLog("wallpaper_path: " & $wallpaper_path)
ToLog("default_infoscreen_path: " & $default_infoscreen_path)
ToLog("infoscreen_standard: " & $infoscreen_standard)
ToLog("infoscreen_timetable: " & $infoscreen_timetable)
#EndRegion

#Region ==========================    Set autohide taskbar    ==========================
HideTaskbar()
#EndRegion

#Region ==========================    Set wallpaper    ==========================
SetWallpaper($wallpaper_path)
#EndRegion

#Region ==========================    Reading the current clinic settings     ==========================
ToLog(@CRLF & "---Reading the current clinic settings---")
Local $clinicName = ""
If StringInStr($current_pc_name, "-") Then
   $clinicName = StringSplit($current_pc_name, "-")
   $clinicName = $clinicName[1]
EndIf
ToLog("clinicName: " & $clinicName)

Local $clinicSettings = IniReadSection($iniFile, $clinicName)
If Not IsArray($clinicSettings) Then
   ToLog($errStr & "Cannot read the clinic section: " & $clinicName)
   SendEmail()
EndIf

_AD_Open()
$current_ou_name = _AD_GetObjectOU($current_pc_name & "$")
ToLog("current_ou_name: " & $current_ou_name)
_AD_Close()

If $current_ou_name = "" Or Not StringInStr($current_ou_name, "OU=") Then
   ToLog($errStr & "Cannot read the computer OU")
   SendEmail()
EndIf

$current_ou_name = StringSplit($current_ou_name, ",")
$current_ou_name = StringReplace($current_ou_name[1], "OU=", "")
ToLog("current_ou_name: " & $current_ou_name)
#EndRegion

#Region ==========================    Parse the current clinic settings    ==========================
ToLog(@CRLF & "---Parse the current clinic settings---")
Local $popupimage = IniRead($iniFile, $clinicName, "popupimage", "")
Local $popupimage_ou_to_run = IniRead($iniFile, $clinicName, "popupimage_ou_to_run", "")
Local $presentation = IniRead($iniFile, $clinicName, "presentation", "")
Local $presentation_ou_to_run = IniRead($iniFile, $clinicName, "presentation_ou_to_run", "")
Local $infoscreen = IniRead($iniFile, $clinicName, "infoscreen", "")
Local $infoscreen_ou_to_run = IniRead($iniFile, $clinicName, "infoscreen_ou_to_run", "")
Local $infoscreen_timetable_ou_to_run = IniRead($iniFile, $clinicName, "infoscreen_timetable_ou_to_run", "")

ToLog("popupimage: " & $popupimage)
ToLog("popupimage_ou_to_run: " & $popupimage_ou_to_run)
ToLog("presentation: " & $presentation)
ToLog("presentation_ou_to_run: " & $presentation_ou_to_run)
ToLog("infoscreen_ou_to_run: " & $infoscreen_ou_to_run)
ToLog("infoscreen_timetable_ou_to_run: " & $infoscreen_timetable_ou_to_run)
#EndRegion

#Region ==========================    Checking for run PopUpImage/presentation/Infoscreen     ==========================
ToLog(@CRLF & "---Checking for run PopUpImage/presentation/Infoscreen---")
If $popupimage Then
   If $popupimage_ou_to_run Then
	  TryToRun($popupimage_ou_to_run, $popupimage, "")
   EndIf
EndIf

If $presentation Then
   If $presentation_ou_to_run Then
	  TryToRun($presentation_ou_to_run, $presentation, "")
   EndIf
EndIf

If $infoscreen Then
   If $infoscreen = "default" Then
	  $infoscreen = StringReplace($default_infoscreen_path, "*", $clinicName)
   EndIf

   If $infoscreen_ou_to_run Then
	  TryToRun($infoscreen_ou_to_run, $infoscreen_restart_path & $infoscreen_standard, $infoscreen)
   EndIf

   If $infoscreen_timetable_ou_to_run Then
	  TryToRun($infoscreen_timetable_ou_to_run, $infoscreen_restart_path & $infoscreen_timetable, $infoscreen)
   EndIf
EndIf

If $error Then
   SendEmail()
EndIf
#EndRegion

#Region ==========================    Functions     ==========================
Func HideTaskbar()
   ToLog(@CRLF & "---HideTaskbar---")
   Const $HKCU = 0x80000001
   Local $objReg = ObjGet("winmgmts:{impersonationLevel=impersonate}root\default:StdRegProv")
   Local $objWMI = ObjGet("winmgmts:{impersonationLevel=impersonate}root\cimv2")
   Local $arrVal[1]

   $objReg.GetBinaryValue($HKCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\StuckRects2", "Settings", $arrVal)
   $arrVal[8] = BitOR(BitAND($arrVal[8], 0x07), 0x01)
   $objReg.SetBinaryValue($HKCU, "Software\Microsoft\Windows\CurrentVersion\Explorer\StuckRects2", "Settings", $arrVal)

   RunWait('"' & @ComSpec & '"' & " /c " & 'taskkill /f /im ' & "explorer.exe", "", @SW_HIDE)
   RunWait('"' & @ComSpec & '"' & " /c " & 'start ' & @WindowsDir & "\explorer.exe", "", @SW_HIDE)
   If @error Then
	  $error = True
	  ToLog($errStr & "Cannot run explorer.exe")
   EndIf
EndFunc

Func SetWallpaper($pathToFind)
   ToLog(@CRLF & "---Set wallpaper---")
   Local $width = @DesktopWidth
   Local $height = @DesktopHeight
   ToLog("desktop width: " & $width)
   ToLog("desktop height: " & $height)

   If Not FileExists($wallpaper_path) Then
	  $error = True
	  ToLog($errStr & "Cannot find the directory with wallpapers: " & $wallpaper_path)
	  Return
   EndIf

   Local $toSet = $width & "_" & $height & ".jpg"

   If Not FileExists($wallpaper_path & $toSet) Then
	  $error = True
	  ToLog($errStr & "Cannot find the needed wallpaper: " & $toSet & " at: " & $wallpaper_path)
	  $toSet = ""

	  Local $wallpapers = _FileListToArray($wallpaper_path, "*_*.jpg")
	  ;_ArrayDisplay($wallpapers)

	  If Not IsArray($wallpapers) Then
		 $error = True
		 ToLog($errStr & "Cannot find any wallpaper at: " & $wallpaper_path)
		 Return
	  Else
		 $toSet = $wallpapers[$wallpapers[0]]
	  EndIf
   EndIf

   ToLog("Wallpaper to use: " & $toSet)

   If Not FileCopy($wallpaper_path & $toSet, $tempFolder, BitOR($FC_OVERWRITE,  $FC_CREATEPATH)) Then
	  ToLog($errStr & "Cannot copy file to " & $tempFolder)
	  Return
   EndIf

   RegWrite('HKCU\Control Panel\Desktop', 'TileWallpaper', 'reg_sz', '0')
   RegWrite('HKCU\Control Panel\Desktop', 'WallpaperStyle', 'reg_sz', '2')
   RegWrite('HKCU\Control Panel\Desktop', 'Wallpaper', 'reg_sz', $tempFolder & $toSet)
   DllCall("User32.dll", "int", "SystemParametersInfo", "int", 20, "int", 0,"str", $tempFolder & $toSet, "int", 0)
EndFunc

Func ToLog($message)
   $message &= @CRLF
   $messageToSend &= $message
   ConsoleWrite($message)
   _FileWriteLog($logFile, $message)
EndFunc

Func TryToRun($ou, $prog, $attr)
   $ou = StringSplit($ou, ";")
   For $tmp In $ou
	  If $current_ou_name = $tmp Then
		 ToLog("Attempt to run: " & $prog & " with attribute: " & $attr)
		 If FileExists($prog) Then
			If Not ShellExecute('"' & $prog & '"', '"' & $attr & '"') Then
			   $error = True
			   ToLog($errStr & "Cannot run the : " & $prog)
			EndIf
		 Else
			$error = True
			ToLog($errStr & "Cannot find the : " & $prog)
		 EndIf
	  EndIf
   Next
EndFunc

Func SendEmail()
   If Not $send_email Then
	  FileClose($logFile)
	  Exit
   EndIf

   ToLog(@CRLF & "---Sending email---")
   If _INetSmtpMailCom($server, "Autorun for all infomonitors", $login, $to, _
		 $current_pc_name & ": error(s) occurred", _
		 $messageToSend, "", "", "", $login, $password) <> 0 Then

	  _INetSmtpMailCom($server_backup, "Autorun for all infomonitors", $login_backup, $to_backup, _
		 $current_pc_name & ": error(s) occurred", _
		 $messageToSend, "", "", "", $login_backup, $password_backup)
   EndIf

   FileClose($logFile)
   Exit
EndFunc

Func _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, _
   $s_Subject = "", $as_Body = "", $s_AttachFiles = "", $s_CcAddress = "", _
   $s_BccAddress = "", $s_Username = "", $s_Password = "",$IPPort=25, $ssl=0)

   Local $objEmail = ObjCreate("CDO.Message")
   Local $i_Error = 0
   Local $i_Error_desciption = ""

   $objEmail.From = '"' & $s_FromName & '" <' & $s_FromAddress & '>'
   $objEmail.To = $s_ToAddress

   If $s_CcAddress <> "" Then $objEmail.Cc = $s_CcAddress
   If $s_BccAddress <> "" Then $objEmail.Bcc = $s_BccAddress

   $objEmail.Subject = $s_Subject

   If StringInStr($as_Body,"<") and StringInStr($as_Body,">") Then
	  $objEmail.HTMLBody = $as_Body
   Else
	  $objEmail.Textbody = $as_Body & @CRLF
   EndIf

   If $s_AttachFiles <> "" Then
	  Local $S_Files2Attach = StringSplit($s_AttachFiles, ";")
	  For $x = 1 To $S_Files2Attach[0] - 1
		 $S_Files2Attach[$x] = _PathFull ($S_Files2Attach[$x])
		 If FileExists($S_Files2Attach[$x]) Then
			$objEmail.AddAttachment ($S_Files2Attach[$x])
		 Else
			$i_Error_desciption = $i_Error_desciption & @lf & 'File not found to attach: ' & $S_Files2Attach[$x]
			SetError(1)
			return 0
		 EndIf
	  Next
   EndIf

   $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
   $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = $s_SmtpServer
   $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = $IPPort

   If $s_Username <> "" Then
	  $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
	  $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = $s_Username
	  $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = $s_Password
   EndIf

   If $Ssl Then
	  $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
   EndIf

   $objEmail.Configuration.Fields.Update
   $objEmail.Send

   if @error then
	  SetError(2)
   EndIf

   Return @error
EndFunc

Func HandleComError()
   ToLog($errStr & "ComError occured: " & $oMyError.source & " " & $oMyError.description)
Endfunc
#EndRegion
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=dbf.ico
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Icon_Add=dbf.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#Region ;**** 参数创建于 ACNWrapper_GUI ****
#EndRegion ;**** 参数创建于 ACNWrapper_GUI ****
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
func setup()
fileinstall("dbfcnv.exe",@scriptdir&"\dbfcnv.exe",0)
EndFunc
opt("TrayIconHide",1)
func del()
filedelete(@scriptdir&"\dbfcnv.exe")
EndFunc

Func a()
RegWrite ('HKEY_CLASSES_ROOT\.dbf',"","REG_SZ","dbffile")
RegWrite ('HKEY_CLASSES_ROOT\dbffile',"","REG_SZ","FoxBase 数据库文件")
RegWrite ('HKEY_CLASSES_ROOT\dbffile\DefaultIcon',"","REG_SZ",@ScriptFullPath&",4")
RegWrite ('HKEY_CLASSES_ROOT\dbffile\shell',"","REG_SZ","")
EndFunc
func b()
RegWrite ('HKEY_CLASSES_ROOT\dbffile\shell\Convert',"","REG_SZ","转换为 EXCEL 文档(&E)")
RegWrite ('HKEY_CLASSES_ROOT\dbffile\shell\Convert\command',"","REG_SZ",@ScriptFullPath&" %1")
EndFunc
func c()
RegWrite ('HKEY_CLASSES_ROOT\dbffile\shell\Convertall',"","REG_SZ","批量转换为 EXCEL 文档(&L)")
RegWrite ('HKEY_CLASSES_ROOT\dbffile\shell\Convertall\command',"","REG_SZ",@ScriptFullPath&" /B %1")
EndFunc
func d()
	RegWrite ('HKEY_CLASSES_ROOT\dbffile\shell\Convertb',"","REG_SZ","自定义转换(&N)...")
RegWrite ('HKEY_CLASSES_ROOT\dbffile\shell\Convertb\command',"","REG_SZ",@ScriptDir&"\dbfcnv.exe")
EndFunc
Func aa()
RegDelete ('HKEY_CLASSES_ROOT\.dbf')
RegDelete ('HKEY_CLASSES_ROOT\dbffile')
Regdelete ('HKEY_CURRENT_USER\Software\dbfconverter')
EndFunc
func bb()
Regdelete('HKEY_CLASSES_ROOT\dbffile\shell\Convert')
EndFunc
func cc()
Regdelete ('HKEY_CLASSES_ROOT\dbffile\shell\Convertall')
EndFunc
func dd()
regdelete('HKEY_CLASSES_ROOT\dbffile\shell\Convertb')
EndFunc
if $CmdLine[0]=0 then 
	#Region ### START Koda GUI section ### Form=
$Form1 = GUICreate("关联", 249, 105, 401, 310, BitOR($WS_SYSMENU,$WS_CAPTION,$WS_POPUP,$WS_POPUPWINDOW,$WS_BORDER,$WS_CLIPSIBLINGS))
$Checkbox1 = GUICtrlCreateCheckbox("关联 DBF 文件(&X)", 8, 8, 145, 17)
if regread('HKEY_CLASSES_ROOT\.dbf',"")="dbffile" then Guictrlsetstate($Checkbox1,$GUI_CHECKED)
$Button1 = GUICtrlCreateButton("确定(&O)", 176, 3, 65, 24, 0)
$Group1 = GUICtrlCreateGroup(" 选择右键菜单项 ", 8, 26, 233, 73)

$Checkbox2 = GUICtrlCreateCheckbox("转换为 Excel 文档(&E)", 24, 42, 185, 17)
if regread('HKEY_CLASSES_ROOT\dbffile\shell\Convert',"")="转换为 EXCEL 文档(&E)" then Guictrlsetstate($Checkbox2,$GUI_CHECKED)
$Checkbox3 = GUICtrlCreateCheckbox("批量转换为 Excel 文档(&L)", 24, 58, 209, 17)
if regread('HKEY_CLASSES_ROOT\dbffile\shell\Convertall',"")="批量转换为 EXCEL 文档(&L)" then Guictrlsetstate($Checkbox3,$GUI_CHECKED)
$Checkbox4 = GUICtrlCreateCheckbox("自定义转换(&N)...", 24, 74, 137, 17)
if regread('HKEY_CLASSES_ROOT\dbffile\shell\Convertb',"")="自定义转换(&N)..." then Guictrlsetstate($Checkbox4,$GUI_CHECKED)

if guictrlread($Checkbox1)=4 Then
GUICtrlSetState($Group1,$GUI_DISABLE)
GUICtrlSetState($Checkbox2,$GUI_DISABLE)
GUICtrlSetState($Checkbox3,$GUI_DISABLE)
GUICtrlSetState($Checkbox4,$GUI_DISABLE)
EndIf
if guictrlread($Checkbox1)=1 Then
GUICtrlSetState($Group1,$GUI_ENABLE)
GUICtrlSetState($Checkbox2,$GUI_ENABLE)
GUICtrlSetState($Checkbox3,$GUI_ENABLE)
GUICtrlSetState($Checkbox4,$GUI_ENABLE)
EndIf

GUICtrlCreateGroup("", -99, -99, 1, 1)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
			Case $GUI_EVENT_CLOSE
;~ 			del()
			Exit
		Case $Button1
			if guictrlread($Checkbox1)=1 Then 
				a()
				setup()
			EndIf

			if guictrlread($checkbox2)=1 then b()
			if guictrlread($checkbox3)=1 then c()
			if guictrlread($checkbox4)=1 then d()
			if guictrlread($Checkbox1)=4 Then
				aa()
				del()
			endif
			if guictrlread($checkbox2)=4 then bb()
			if guictrlread($checkbox3)=4 then cc()
			if guictrlread($checkbox4)=4 then dd()
;~ 			del()
			Exit
		Case $Checkbox1
			if GUICtrlRead($Checkbox1)=1 Then 
				GUICtrlSetState($Group1,$GUI_ENABLE)
GUICtrlSetState($Checkbox2,$GUI_ENABLE)
GUICtrlSetState($Checkbox3,$GUI_ENABLE)
GUICtrlSetState($Checkbox4,$GUI_ENABLE)
Guictrlsetstate($Checkbox4,$GUI_CHECKED)
Guictrlsetstate($Checkbox3,$GUI_CHECKED)
Guictrlsetstate($Checkbox2,$GUI_CHECKED)
				EndIf
			if GUICtrlRead($Checkbox1)=4 Then
				GUICtrlSetState($Group1,$GUI_DISABLE)
GUICtrlSetState($Checkbox2,$GUI_DISABLE)
GUICtrlSetState($Checkbox3,$GUI_DISABLE)
GUICtrlSetState($Checkbox4,$GUI_DISABLE)
Guictrlsetstate($Checkbox4,$GUI_unCHECKED)
Guictrlsetstate($Checkbox3,$GUI_unCHECKED)
Guictrlsetstate($Checkbox2,$GUI_unCHECKED)
				EndIf

	EndSwitch
WEnd


endif
;~ msgbox(1,$CmdLine[$CmdLine[0]],$CmdLine[$CmdLine[0]])
if $CmdLine[0]=1 Then
$fullpath=FileGetLongName($CmdLine[$CmdLine[0]])

$local=StringLeft($fullpath,stringlen($fullpath)-4)
;~ msgbox(1,$CmdLine[$CmdLine[0]],$local)
;~ msgbox(1,$fullpath,$local&".xls")
run(@scriptdir&"\dbfcnv.exe "&'"'&$fullpath&'"'&" "&'"'&$local&".xls"&'"')
;~ msgbox(1,$fullpath,@scriptdir&"\dbfcnv.exe "&'"'&$fullpath&'"'&" "&'"'&$local&".xls"&'"')
;~ del()
EndIf

if $CmdLine[0]=2 Then
	$fullpath=FileGetLongName($CmdLine[$CmdLine[0]])
;~ 	$fullpath=$CmdLine[$CmdLine[0]]
	$LNo=stringinstr($fullpath,"\",0,-1)
	$fullpath=stringmid($fullpath,1,$LNo)
	$fullpath1=$fullpath&"*.dbf"
;~ 	msgbox(1,$fullpath1,@scriptdir&"\dbfcnv.exe "&'"'&$fullpath1&'"'&" "&'"'&$fullpath&'"'&" /xls")
;~ 	msgbox(1,$fullpath1,@scriptdir&"\dbfcnv.exe "&$fullpath1&" "&$fullpath&" /XLS")
	
	run(@scriptdir&"\dbfcnv.exe "&'"'&$fullpath1&'"'&" "&'"'&$fullpath&'"'&" /TOXLS")	
;~ 	run(@scriptdir&"\dbfcnv.exe "&$fullpath1&" "&$fullpath&" /TOXLS")
;~ del()
EndIf


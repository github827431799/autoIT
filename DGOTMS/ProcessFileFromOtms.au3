#include<IE.au3>
#include <Date.au3>
#include <MsgBoxConstants.au3>

;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;Open IE and login Logicaldoc and check out file
Sleep(10000)
;Local $oIE = _IECreate("http://192.168.99.92:17891/logicaldoc/login.jsp");DE
Local $oIE = _IECreate("http://192.168.20.41:15789/logicaldoc/login.jsp");TEST
Local $doc = _IEDocGetObj($oIE)
Local $title=$doc.title
WinSetState($title, "", @SW_MAXIMIZE)
WinWaitActive($title)
Sleep(20000)
MouseClick("left", 790, 214, 1);Get active field for user
Sleep(4000)
Send("otms.bi");User
Sleep(4000)
Send("{TAB}")
Sleep(4000)
Send("1qaz@WSX");Password
Sleep(4000)
MouseClick("left", 925, 317, 1);Select language
Sleep(4000)
MouseClick("left", 796, 497, 1);Select Chinese
Sleep(4000)
MouseClick("left", 882, 363, 1);Login
Sleep(20000)
MouseClick("left", 64, 274, 2);Expand BI
Sleep(4000)
MouseClick("left", 108, 301, 2);Expand external files
Sleep(4000)
MouseClick("left", 144, 323, 2);Expand 01
Sleep(4000)
MouseClick("left", 157, 348, 1);Open upload path
Sleep(4000)
MouseClick("right", 443, 304, 1);Right-click file
Sleep(4000)
MouseClick("left", 526, 584, 1);Check out
Sleep(4000)
If FileExists("E:\autoIT\DGOTMS\Email\德高daily.xlsx") Then
   FileDelete("E:\autoIT\DGOTMS\Email\德高daily.xlsx");Delete file
EndIf
Sleep(4000)
MouseClick("left", 1153, 837, 1);Press save list
Sleep(4000)
MouseClick("left", 1226, 813, 1);Press save as
Sleep(4000)
Send("E:\autoIT\DGOTMS\Email\德高daily.xlsx");Input path
Sleep(4000)
Send("{ENTER}");Press ENTER
Sleep(4000)


;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;Open IE and login 21cn and download file
Sleep(4000)
Local $oIE = _IECreate("https://mail.21cn.net")
Local $doc = _IEDocGetObj($oIE)
Local $title=$doc.title
WinSetState($title, "", @SW_MAXIMIZE)
WinWaitActive($title)
Sleep(20000)
MouseClick("left", 1168, 362, 1);Clear email input field
Sleep(4000)
MouseClick("left", 1277, 362, 1);Clear email input field
Sleep(4000)
Send("bi.otms@davco.cn");Email user
Sleep(4000)
Send("{TAB}")
Sleep(4000)
Send("1qaz@otms");Email password
Sleep(4000)
MouseClick("left", 1150, 550, 1);login
Sleep(20000)
MouseClick("left", 55, 242, 1);Press inbox
Sleep(4000)
MouseClick("left", 670, 253, 1);Press lastest email title
Sleep(4000)
If FileExists("E:\autoIT\DGOTMS\Email\德高daily.xlsx") Then
   MouseMove(279, 575);Move to download attachment
   Sleep(4000)
   MouseClick("left", 279, 540, 1);Press download
   Sleep(4000)
   MouseClick("left", 1153, 838, 1);Press save list
   Sleep(4000)
   MouseClick("left", 1226, 813, 1);Press save as
   Sleep(4000)
   Send("E:\autoIT\DGOTMS\Email\德高daily.xlsx");Input path
   Sleep(4000)
   Send("{ENTER}");Press ENTER
   Sleep(4000)
   Send("{TAB}");Press ENTER
   Sleep(4000)
   Send("{ENTER}");Press ENTER
EndIf
Sleep(4000)
MouseClick("left", 1582, 9, 1);Close IE

;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;Check in file on Logicalfile
Sleep(4000)
MouseClick("right", 447, 304, 1);Right-click file
Sleep(4000)
MouseClick("left", 530, 617, 1);Check in
Sleep(4000)
MouseClick("left", 831, 477, 1);Open path
Sleep(4000)
Send("E:\autoIT\DGOTMS\Email\德高daily.xlsx");Input path
Sleep(4000)
Send("{ENTER}");Press ENTER
Sleep(4000)
MouseClick("left", 743, 441, 1);Move to comment
Sleep(4000)
Local $now = _Now()
Send($now)
Sleep(4000)
MouseClick("left", 653, 556, 1);Upload
Sleep(4000)
MouseClick("left", 460, 588, 1);Check in
Sleep(4000)
MouseClick("left", 1582, 9, 1);Close IE
Sleep(4000)
If FileExists("E:\autoIT\DGOTMS\Email\德高daily.xlsx") Then
   FileDelete("E:\autoIT\DGOTMS\Email\德高daily.xlsx");Delete file
EndIf








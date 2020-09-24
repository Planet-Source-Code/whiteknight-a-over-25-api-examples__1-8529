VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAPIstuff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Yes I Used Text1 Text2 Command1 ect. Only because This is an example I usually dont _
do that and I don't sugest that you do it either.  I Call them What they Do (If a button exits the _
app then I'd Call it bttn_exit). This Is just an example of how to do Some things with API.  If you _
Like My example then please vote for me.  Why is it all in a class well I plan on making a dll with
'these functions and More too.  If you have any comments of sugestions please e-mail them _
me at witenite87@excite.com.  Feel Free to use this in any way you want.  You can Modify it any way _
you like as long as the comments remain and you give credit where credit is do.  These are not that hard _
of API Function I myself am just learning API.  Please Visit my site http://camalot.virtualave.net.
'If you have any thing you would like me to add to this just email it to me and If I add it your name And Info _
Will Go under each Sub/Function you add.

Private Declare Function DoFileDownload Lib "shdocvw.dll" (ByVal lpszFile As String) As Long
Private Declare Function DrawCaption Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long, pcRect As RECT, ByVal un As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, ByVal lpDrawTextParams As Any) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Sub FatalAppExit Lib "kernel32" Alias "FatalAppExitA" (ByVal uAction As Long, ByVal lpMessageText As String)
Private Declare Function fCreateShellLink Lib "VB5STKIT.DLL" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArgs As String) As Long
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Private Declare Function FlashWindowEx Lib "user32" (pfwi As FLASHWINFO) As Boolean
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet" (ByVal hInet As Long) As Integer
Private Declare Function InternetReadFile Lib "wininet" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwflags As Long, ByVal dwContext As Long) As Long
Private Declare Function IsCharAlpha Lib "user32" Alias "IsCharAlphaA" (ByVal cChar As Byte) As Long
Private Declare Function IsCharAlphaNumeric Lib "user32" Alias "IsCharAlphaNumericA" (ByVal cChar As Byte) As Long
Private Declare Function IsCharLower Lib "user32" Alias "IsCharLowerA" (ByVal cChar As Byte) As Long
Private Declare Function IsCharUpper Lib "user32" Alias "IsCharUpperA" (ByVal cChar As Byte) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long


Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function WriteProfileSection Lib "kernel32" Alias "WriteProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String) As Long
Private Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long

'Internet Const
 Const INTERNET_OPEN_TYPE_DIRECT = 1
 Const INTERNET_OPEN_TYPE_PROXY = 3
 Const INTERNET_FLAG_RELOAD = &H80000000
'Exit Window Const
 Const EWX_LOGOFF = 0
 Const EWX_SHUTDOWN = 1
 Const EWX_REBOOT = 2
 Const EWX_FORCE = 4
'Flash Window EX Const
 Const FLASHW_STOP = 0 'Stop flashing. The system restores the window to its original state.
 Const FLASHW_CAPTION = &H1 'Flash the window caption.
 Const FLASHW_TRAY = &H2 'Flash the taskbar button.
 Const FLASHW_ALL = (FLASHW_CAPTION Or FLASHW_TRAY) 'Flash both the window caption and taskbar button. This is equivalent to setting the FLASHW_CAPTION Or FLASHW_TRAY flags.
 Const FLASHW_TIMER = &H4 'Flash continuously, until the FLASHW_STOP flag is set.
 Const FLASHW_TIMERNOFG = &HC 'Flash continuously until the window comes to the foreground.
'Flash Window Const
Const INVERT = 1
'Key Spy const
 Const DT_CENTER = &H1
 Const DT_WORDBREAK = &H10
'Draw Edges
 Const DC_ACTIVE = &H1
 Const DC_ICON = &H4
 Const DC_TEXT = &H8
 Const BDR_SUNKENOUTER = &H2
 Const BDR_RAISEDINNER = &H4
 Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
 Const BF_BOTTOM = &H8
 Const BF_LEFT = &H1
 Const BF_RIGHT = &H4
 Const BF_TOP = &H2
 Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
 Const DFC_BUTTON = 4
 Const DFC_POPUPMENU = 5            'Only Win98/2000 !!
 Const DFCS_BUTTON3STATE = &H10
 Const DC_GRADIENT = &H20          'Only Win98/2000 !!
'Registry Const
Const HKEY_CLASSES_ROOT = &H80000000
'Tray Const
Private Const WM_USER = &H400
Private Const WM_LBUTTONUP = &H202
Private Const WM_MBUTTONUP = &H208
Private Const WM_RBUTTONUP = &H205
Private Const WM_LBUTTONDBLCLICK = &H203
Private Const WM_MOUSEMOVE = &H200

'Show Window Functions
Private Const SW_MAXIMIZE = 3
Private Const SW_MINIMIZE = 6
Private Const SW_HIDE = 0
Private Const SW_SHOW = 5
Private Const SW_NORMAL = 1


Private Const TRAY_CALLBACK = (WM_USER + 1001&)
Private Const GWL_WNDPROC = (-4)
Private Const GWL_USERDATA = (-21)
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIM_ADD = &H0
Private Const NIF_MESSAGE = &H1
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
'Flash Window Types
Private Type FLASHWINFO
    cbSize As Long
    hwnd As Long
    dwflags As Long
    uCount As Long
    dwTimeout As Long
End Type
'RECT
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'NOTIFYICONDATA
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Cnt As Long
Private sSave As String
Private sOld As String
Private strUrl As String
Private Ret As String
Dim Tel As Long
Dim MSG As VbMsgBoxResult
Dim encKEY() As Long
Public Sub DownLoadFile(sUrl As String, sFile As String)
' I set it up so the url is different then the file so all you have to do is set up the url once then just change the File Name
Dim DL As Long
On Error GoTo errHandler
If sUrl$ = "" Then sUrl$ = strUrl$
If strUrl$ = "" Then strUrl$ = sUrl$
If Left(strUrl$, 4) <> "http" Then strUrl$ = "http://" & strUrl$ 'See if  there's a http:// at the begining if not add it
If Right$(strUrl$, 1) <> "/" Then strUrl$ = strUrl$ & "/" 'see if there is an "/" at the end of the url if theres not add it
If Left$(sFile$, 1) = "/" Then sFile = Mid$(sFile$, 2) 'since we added an "/" we have to make sure there isnt one at the front of the file"
 DL& = DoFileDownload(StrConv(strUrl$ & sFile$, vbUnicode))
 Exit Sub
errHandler:
Debug.Print "Error Source:", Err.Source
Debug.Print "Error Description:", Err.Description
Debug.Print "Error Number:", Err.Number
MsgBox "An error has occured attempting to start download to " & sUrl & sFile$ & ".", vbApplicationModal + vbCritical + vbDefaultButton1, "Error"
End Sub
  
Public Sub RestartWindows(Optional msgPrompt As Boolean = False, Optional ShowErrors As Boolean = False)
Dim Ret As Long

On Error GoTo errHandler
If msgPrompt = True Then
MSG = MsgBox("This program is going to reboot your computer. Press OK to continue or Cancel to stop.", vbCritical + vbOKCancel + 256, App.Title)
    If MSG = vbCancel Then End
End If
    'reboot the computer
    Ret& = ExitWindowsEx(EWX_FORCE Or EWX_REBOOT, 0)
    Exit Sub
errHandler:
Debug.Print "Error Source:", Err.Source
Debug.Print "Error Description:", Err.Description
Debug.Print "Error Number:", Err.Number
If ShowErrors = True Then
Call MsgBox("An error has occured attempting to shut down your computer.", vbApplicationModal + vbCritical + vbDefaultButton1, "Error")
End If
End Sub

Public Sub LogOffWindows(Optional msgPrompt As Boolean = False, Optional ShowErrors As Boolean = False)
Dim Ret As Long
On Error GoTo errHandler
If msgPrompt = True Then
MSG = MsgBox("This program is going to Log Off  your computer. Press OK to continue or Cancel to stop.", vbCritical + vbOKCancel + 256, App.Title)
    If MSG = vbCancel Then End
End If
    'LogOff  the computer
    Ret& = ExitWindowsEx(EWX_LOGOFF, 0)
    Exit Sub
errHandler:
Debug.Print "Error Source:", Err.Source
Debug.Print "Error Description:", Err.Description
Debug.Print "Error Number:", Err.Number
If ShowErrors = True Then
Call MsgBox("An error has occured attempting to shut down your computer.", vbApplicationModal + vbCritical + vbDefaultButton1, "Error")
End If
End Sub

Public Sub ShutDownWindows(Optional msgPrompt As Boolean = False, Optional ShowErrors As Boolean = False)
Dim Ret As Long
On Error GoTo errHandler
If msgPrompt = True Then
MSG = MsgBox("This program is going to shut down your computer. Press OK to continue or Cancel to stop.", vbCritical + vbOKCancel + 256, App.Title)
    If MSG = vbCancel Then End
End If
    'Shut Down  the computer
    Ret& = ExitWindowsEx(EWX_SHUTDOWN, 0)
Exit Sub
errHandler:
Debug.Print "Error Source:", Err.Source
Debug.Print "Error Description:", Err.Description
Debug.Print "Error Number:", Err.Number
If ShowErrors = True Then
Call MsgBox("An error has occured attempting to shut down your computer.", vbApplicationModal + vbCritical + vbDefaultButton1, "Error")
End If
End Sub

Public Sub FatalErrorExit(ErrorMSG As String)
On Error GoTo errHandler
'0 = white fatal error
'1 =
FatalAppExit 1, ErrorMSG$

Exit Sub
errHandler:
Debug.Print "Error Source:", Err.Source
Debug.Print "Error Description:", Err.Description
Debug.Print "Error Number:", Err.Number
End Sub

Public Sub DesktopShortcut(sLinkName As String, sFilePath As String, Optional ShowErrors As Boolean = False)
'Adds A link to your desktop _
sLinkName Is the Link "Caption" _
sFilePath$ is the Path of the App _
ShowErrors = True will Show A msgbox When and error happens
Dim lngresult  As Long
On Error GoTo errHandler
lngresult = fCreateShellLink("..\..\Desktop", sLinkName$, sFilePath$, "")
Exit Sub
errHandler:
Debug.Print "Error Source:", Err.Source
Debug.Print "Error Description:", Err.Description
Debug.Print "Error Number:", Err.Number
If ShowErrors = True Then
Call MsgBox("An error has occured attempting to create shortcut for " & sFilePath$ & "named " & sLinkName$ & ".", vbApplicationModal + vbCritical + vbDefaultButton1, "Error")
End If
End Sub

Public Sub FlashWinEXON(lhwnd As Long, iTimeout As Integer, iCount As Integer)
On Error GoTo errorHandler
Dim FlashInfo As FLASHWINFO
    'the size of the structure.
    FlashInfo.cbSize = Len(FlashInfo)
    'the flash status
    FlashInfo.dwflags = FLASHW_ALL Or FLASHW_TIMER
    'the rate, in milliseconds, at which the window will be flashed. If dwTimeout is zero, the function uses the default cursor blink rate.
    FlashInfo.dwTimeout = iTimeout%
    'Handle to the window to be flashed. The window can be either opened or minimized.
    FlashInfo.hwnd = lhwnd&
    'the number of times to flash the window.
    FlashInfo.uCount = iCount%
    FlashWindowEx FlashInfo
    Exit Sub
errorHandler:
Debug.Print "Error Source:", Err.Source
Debug.Print "Error Description:", Err.Description
Debug.Print "Error Number:", Err.Number
End Sub
Public Sub FlashWin(lhwnd As Long, flashOn As Boolean)
Dim a As Long
On Error Resume Next
If flashOn = True Then
        DoEvents
        a& = FlashWindow(lhwnd, 1)
End If
 If flashOn = False Then
    DoEvents
    a& = FlashWindow(lhwnd, 0)
End If
End Sub
Public Sub FlashWinEXOFF(lhwnd As Long)
'This one i could not test because I dont have this Function on my Computer
On Error GoTo errorHandler
Dim FlashInfo As FLASHWINFO
    'size of the structure.
    FlashInfo.cbSize = Len(FlashInfo)
    'Stop The Flash
    FlashInfo.dwflags = FLASHW_STOP
    FlashInfo.dwTimeout = 0
    'Handle to the window to Stop
    FlashInfo.hwnd = lhwnd&
    FlashInfo.uCount = 0
    FlashWindowEx FlashInfo
    Exit Sub
errorHandler:
Debug.Print "Error Source:", Err.Source
Debug.Print "Error Description:", Err.Description
Debug.Print "Error Number:", Err.Number
End Sub


Function GetPressedKey() As String
    For Cnt = 32 To 128
        'Get the keystate of a specified key
        If GetAsyncKeyState(Cnt) <> 0 Then
            GetPressedKey = Chr$(Cnt)
            Exit For
        End If
    Next Cnt
End Function
Private Sub TimerProc(ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
    ' I Used For A key Logger
    Ret = GetPressedKey
    If Ret <> sOld Then
        sOld = Ret
        sSave = sSave + sOld
    End If
End Sub

Public Function FreeDiscSpace(sDriveLetter As String) As String
    Dim Sectors As Long, Bytes As Long, FreeC As Long, TotalC As Long, Total As Long, Freeb As Long, sTemp As String
    'Retrieve information about the sDriveLetter
    GetDiskFreeSpace sDriveLetter$ & ":\", Sectors, Bytes, FreeC, TotalC
    sTemp$ = "Sectors per Cluster:" & Str$(Sectors) & vbCrLf
    sTemp$ = sTemp$ & "Bytes per sector:" & Str$(Bytes) & vbCrLf
    sTemp$ = sTemp$ & "Number Of Free Clusters:" & Str$(FreeC) & vbCrLf
    sTemp$ = sTemp$ & "Total Number Of Clusters:" & Str$(TotalC) & vbCrLf
    Total = TotalC& * Sectors& * Bytes&
    sTemp$ = sTemp$ & "Total number of bytes in path:" & Str$(Total) & vbCrLf
    Freeb = FreeC& * Sectors& * Bytes&
    sTemp$ = sTemp$ & "Free bytes:" & Str$(Freeb) & vbCrLf
    FreeDiscSpace$ = sTemp$
End Function

Public Function Get_DriveType(sDriveLetter As String) As String
Select Case GetDriveType(sDriveLetter$ & ":\")
        Case 2
            Get_DriveType$ = "Removable"
        Case 3
            Get_DriveType$ = "Drive Fixed"
        Case Is = 4
            Get_DriveType$ = "Remote"
        Case Is = 5
            Get_DriveType$ = "CD-Rom"
        Case Is = 6
            Get_DriveType$ = "Ram disk"
        Case Else
            Get_DriveType$ = "Unrecognized"
    End Select
End Function
Public Function CharType(KeyAscii As Integer) As String
Dim strsave As String
    'Get the character information
    If IsCharAlpha(KeyAscii) Then strsave$ = "Alpha"
    If IsCharAlphaNumeric(KeyAscii) Then strsave$ = " AlphaNumeric"
    If IsCharLower(KeyAscii) Then strsave$ = strsave$ & " Lower Case"
    If IsCharUpper(KeyAscii) Then strsave$ = strsave$ & " Upper Case"
    'save the infomation
    CharType$ = strsave$
End Function

Public Function GetURLSource(sUrl As String) As String
Dim lOpen As Long, lFile As Long, sBuffer As String, Ret As Long
    'Create a buffer for the file we're going to download
    sBuffer$ = Space(1000)
    'Create an internet connection
    lOpen& = InternetOpen(App.Title, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    'Open the url
    lFile& = InternetOpenUrl(lOpen, sUrl$, vbNullString, ByVal 0&, INTERNET_FLAG_RELOAD, ByVal 0&)
    'Read the first 1000 bytes of the file
    InternetReadFile lFile&, sBuffer$, 1000, Ret&
    'clean up
    InternetCloseHandle lFile&
    InternetCloseHandle lOpen&
    'Show our file
    GetURLSource$ = sBuffer$
End Function

Public Sub DrawFormEdges(Frm As Form, iWidth As Integer, Optional Outter As Boolean = True)
Dim r As RECT
Dim i As Integer
    'Clear the form
    Frm.Cls
    'API uses pixels
    Frm.ScaleMode = vbPixels
    'Depending on the thickness Draw The Form Edge
    For i% = 0 To iWidth% - 1
    SetRect r, 0 + i%, 0 + i%, Frm.ScaleWidth - i%, Frm.ScaleHeight - i%
    If Outter = True Then DrawFrameControl Frm.hdc, r, DFC_BUTTON, DFCS_BUTTON3STATE
    If Outter = False Then DrawEdge Frm.hdc, r, EDGE_ETCHED, BF_RECT
    Next i%
End Sub

'Public Sub DefaultApp(ByVal sAppName As String, ByVal sEXE As String, ByVal sSubFileTypeKey As String)
'Dim lRegKey As Long
'Call RegCreateKey(HKEY_CLASSES_ROOT, sSubFileTypeKey, lRegKey)
'Call RegSetValueEx(lRegKey, "", 0&, 1, ByVal sAppName, Len(sAppName))
'Call RegCloseKey(lRegKey)
'Call RegCreateKey(HKEY_CLASSES_ROOT, sAppName & "\Shell\Open\Command", lRegKey) ' adds info into the shell open command
'Call RegSetValueEx(lRegKey, "", 0&, 1, ByVal sEXE, Len(sEXE))
'Call RegCloseKey(lRegKey)
'End Sub
Public Sub StartMenuShortcut(strshortcutdir As String, strshortcutname As String, strshortcutpath As String)
'I did not write this I think it is from izekial32.bas
    Dim strwinshortcutdir As String, strwinshortcutname As String, strwinshortcutexepath As String, lngretval As Long
     strwinshortcutdir$ = strshortcutdir$
     strwinshortcutname$ = strshortcutname$
     strwinshortcutexepath$ = strshortcutpath$
     lngretval& = fCreateShellLink("", strwinshortcutname$, strwinshortcutexepath$, "")
    Name "c:\windows\start menu\programs\" & strwinshortcutname$ & ".lnk" As strwinshortcutdir$ & "\" & strwinshortcutname$ & ".lnk"
End Sub
Public Sub LaunchWebBrowser(lhwnd As Long, Site As String)
'Launch Web Browser To a site
On Error Resume Next
Dim IE As Long
IE& = ShellExecute(lhwnd, "Open", Site, vbNullString, vbNullString, SW_MAXIMIZE)
End Sub
Sub MaxWindow(hwnd)
'Maximize a Window
Dim Max
Max = ShowWindow(hwnd, SW_MAXIMIZE)
End Sub

Sub MinWindow(hwnd)
'Minimizes a Window
Dim min
min = ShowWindow(hwnd, SW_MINIMIZE)
End Sub
Sub NormWindow(hwnd)
'Minimizes a Window
Dim nor
nor = ShowWindow(hwnd, SW_NORMAL)
End Sub
Sub OpenEXE(Path As String)
'Do I really Need To explain This?
On Error GoTo oops
Dim x As Integer
Dim NoFreeze As Integer
x% = Shell(Path, 1): NoFreeze% = DoEvents(): Exit Sub
Exit Sub
oops:
    Debug.Print "Error description", Err.Description
    Debug.Print "Error source:", Err.Source
    Debug.Print "Error Number:", Err.Number
End Sub
Function GetPercent(Complete As Integer, Total As Integer, TotalOutput As Integer) As Integer
'Do I really Need To explain This?
    On Error GoTo oops
    GetPercent% = Int(Complete% / Total% * TotalOutput%)
    Exit Function
oops:
    Debug.Print "Error description", Err.Description
    Debug.Print "Error source:", Err.Source
    Debug.Print "Error Number:", Err.Number
End Function
Public Function Encrypt(ByVal Plain As String, sEncKey As String) As String
'This Encrypts A string by converting it to its ASCII number but the difference is it uses a Key String it converts the keystring _
to ASCII and adds it to the first ASCII Value the key is needed to decrypt the text.  I do plain on changing this some what but For _
Now its ok.  I've only seen it cause an error when the wrong Key was entered while decrypting.
    On Error GoTo oops
    Dim i As Integer
    Dim Letter As String
    Dim KeyNum As String
    Dim encstr As String
    Dim Math As Long
    If sEncKey = "" Then sEncKey = "WhiteKnight"
    ReDim encKEY(1 To Len(sEncKey))
    For i = 1 To Len(sEncKey$)
    KeyNum = Mid$(sEncKey$, i, 1)
        encKEY(i) = Asc(KeyNum)
        If i = 1 Then Math = encKEY(i): GoTo nextone
        If i >= 2 And encKEY(i) >= Math And encKEY(i) <= encKEY(i - 1) Then Math = Math - encKEY(i)
        If i >= 2 And encKEY(i) <= Math And encKEY(i) <= encKEY(i - 1) Then Math = Math - encKEY(i)
        If i >= 2 And encKEY(i) >= Math And encKEY(i) >= encKEY(i - 1) Then Math = Math + encKEY(i)
        If i >= 2 And encKEY(i) <= Math And encKEY(i) >= encKEY(i - 1) Then Math = Math + encKEY(i)
nextone:
    Next i
    For i = 1 To Len(Plain)
        Letter = Mid$(Plain, i, 1)
        encstr = encstr & Asc(Letter) + Math & " "
    Next i
Encrypt = encstr
oops:
    Debug.Print "Error description", Err.Description
    Debug.Print "Error source:", Err.Source
    Debug.Print "Error Number:", Err.Number
End Function

Public Function Decrypt(ByVal Encrypted As String, sEncKey As String) As String
'This Encrypts A string by converting it to its ASCII number but the difference is it uses a Key String it converts the keystring _
to ASCII and adds it to the first ASCII Value the key is needed to decrypt the text.  I do plain on changing this some what but For _
Now its ok.  I've only seen it cause an error when the wrong Key was entered while decrypting.
On Error GoTo oops
Dim i As Integer
Dim Letter As String
    Dim KeyNum As String
    Dim EncNum As String
    Dim encbuffer As Long
    Dim strDecrypted As String
    Dim Math As Long
    If sEncKey = "" Then sEncKey = "WhiteKnight"
       ReDim encKEY(1 To Len(sEncKey))
    For i = 1 To Len(sEncKey$)
    KeyNum = Mid$(sEncKey$, i, 1)
        encKEY(i) = Asc(KeyNum)
        If i = 1 Then Math = encKEY(i): GoTo nextone
        If i >= 2 And encKEY(i) >= Math And encKEY(i) <= encKEY(i - 1) Then Math = Math - encKEY(i)
        If i >= 2 And encKEY(i) <= Math And encKEY(i) <= encKEY(i - 1) Then Math = Math - encKEY(i)
        If i >= 2 And encKEY(i) >= Math And encKEY(i) >= encKEY(i - 1) Then Math = Math + encKEY(i)
        If i >= 2 And encKEY(i) <= Math And encKEY(i) >= encKEY(i - 1) Then Math = Math + encKEY(i)
nextone:

    Next i
    For i = 1 To Len(Encrypted)
    
        Letter = Mid$(Encrypted, i, 1)
        EncNum = EncNum & Letter
        If Letter = " " Then
        encbuffer = CLng(Mid(EncNum, 1, Len(EncNum) - 1))
        strDecrypted$ = strDecrypted & Chr(encbuffer - Math)
        'MsgBox strDecrypted
        EncNum = ""
        End If
        
    Next i
    Decrypt = strDecrypted
    Exit Function
oops:
    MsgBox "You Have Entered The WRONG Encrypt Key or Have the Worng Encrypted Message. " & vbCrLf & "If you think you received this Message in error please contact WhiteKnight @ witenite87@excite.com", vbApplicationModal + vbCritical + vbCritical + vbMsgBoxSetForeground, "Decryption Error"
End Function

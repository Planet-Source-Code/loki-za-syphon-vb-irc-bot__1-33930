Attribute VB_Name = "Environment"
Option Explicit
'********************************************************
'* Name: Environment Module
'* Type: VB Module
'* Author: Darryn Britton
'* Date: 15/04/2002
'********************************************************
' Global Constants, Types, Enums and Variables as well as Global
' Methods needed by all Forms and Objects. Defines all API
'********************************************************
'API defines
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
'err consts
Public Const ERR_PROPERTYGET = 4000
Public Const ERR_PROPERTYLET = 4001
Public Const ERR_METHOD = 4002
Public Const ERR_SOCKET = 4003
'env consts
Public Const FILE_CONFIG = "\media\settings.conf"
Public Const FILE_HASH = "\media\commands.hash"
Public Const FILE_DEFINE = "\media\defines.hash"
Public Const PATH_LOGDATA = "\media\logs\"
Public Const DEF_INLINEVAR = "%v"
'API consts
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
'types
Private Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type
'enums
Public Enum ENUM_LOGTYPE
    socketEvent = 0
    botEvent = 1
    appEvent = 2
    ircEvent = 3
End Enum
Public Enum ENUM_DIRECTION
    directionIn = 0
    directionOut = 1
    directionBoth = 2
    directionNull = 3
End Enum
Public Enum ENUM_BOTSTATE
    botIdle = 0
    botConnecting = 1
    botWorking = 2
    botClosing = 3
    botRegistered = 4
End Enum
Public Enum ENUM_IRCCOMMAND
    commandNone = 0
End Enum
'variables
Private ShellIcon As NOTIFYICONDATA
'Replace defined escape characters with the VB equivelant (C Style)
Public Function ReplaceEscapeChars(ByVal DataString As String) As String
On Error GoTo err_method_ReplaceEscapeChars
Dim strReturn As String
    'tab
    strReturn = Replace(DataString, "\t", vbTab)
    'newline
    strReturn = Replace(strReturn, "\n", vbCrLf)
    'carriage return
    strReturn = Replace(strReturn, "\cr", Chr(13))
    'line feed
    strReturn = Replace(strReturn, "\lf", Chr(10))
    'bold
    strReturn = Replace(strReturn, "\b", Chr(2))
    'sepchar
    strReturn = Replace(strReturn, "\s", Chr(1))
    ReplaceEscapeChars = strReturn
Exit Function
err_method_ReplaceEscapeChars:
    Call KillWithError(ERR_METHOD, "Environment::Method::ReplaceEscapeChars", Err.Description)
End Function
Public Function ParseBotCommand(ByVal BotCommand As String)

End Function
'VB Equivelant of C's "printf" where a string is filled with defined variable replacement
'characters and an array of values passed..
Public Function VBPrintF(ByVal Data As String, ParamArray Argv() As Variant) As String
On Error GoTo err_method_VBPrintF
Dim arrTextParts() As String
Dim lngPartCounter As Long
Dim lngPositionBuffer As Long
Dim strReturn As String
Dim blnConversionComplete As Boolean
    If (UBound(Argv) = -1) Then Exit Function
    ReDim arrTextParts(0)
    ReDim arrValueParts(0)
    ReDim arrValueParts(UBound(Argv))
    lngPositionBuffer = 1
    Data = ReplaceEscapeChars(Data)
    blnConversionComplete = False
    Do While Not blnConversionComplete
        If (InStr(lngPositionBuffer, Data, "%v") = 0) Then
            blnConversionComplete = True
            If (Len(Data) > lngPositionBuffer) Then
                'get the last of the string
                arrTextParts(UBound(arrTextParts)) = arrTextParts(UBound(arrTextParts)) & Mid(Data, lngPositionBuffer)
            End If
            Exit Do
        End If
        ReDim Preserve arrTextParts(UBound(arrTextParts) + 1)
        arrTextParts(UBound(arrTextParts)) = Mid(Data, lngPositionBuffer, InStr(lngPositionBuffer, Data, "%v") - lngPositionBuffer)
        arrTextParts(UBound(arrTextParts)) = arrTextParts(UBound(arrTextParts)) & Argv(UBound(arrTextParts) - 1)
        lngPositionBuffer = InStr(lngPositionBuffer, Data, "%v") + 2
    Loop
    strReturn = ""
    VBPrintF = Join(arrTextParts, "")
Exit Function
err_method_VBPrintF:
    Call KillWithError(ERR_METHOD, "Environment::Method::VBPrintF", Err.Description)
End Function
'Add a specific form to the systray
Public Sub AddTrayIcon(ByVal FormToTray As Form, ByVal ToolTip As String)
On Error GoTo err_method_AddTrayIcon
    With ShellIcon
        .cbSize = Len(ShellIcon)
        .hIcon = FormToTray.Icon
        .hwnd = FormToTray.hwnd
        .szTip = ToolTip & vbNullChar
        .uID = vbNull
        .uCallbackMessage = WM_MOUSEMOVE
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    End With
    Call Shell_NotifyIcon(NIM_ADD, ShellIcon)
Exit Sub
err_method_AddTrayIcon:
    Call KillWithError(ERR_METHOD, "Environment::Method::AddTrayIcon", Err.Description)
End Sub
'Edit a "form" instance in the systray
Public Sub ChangeTrayIcon(ByVal FormToEdit As Form, ByVal NewToolTip As String)
On Error GoTo err_method_ChangeTrayIcon
    With ShellIcon
        .hIcon = FormToEdit.Icon
        .hwnd = FormToEdit.hwnd
        .szTip = NewToolTip & vbNullChar
    End With
    Call Shell_NotifyIcon(NIM_MODIFY, ShellIcon)
Exit Sub
err_method_ChangeTrayIcon:
    Call KillWithError(ERR_METHOD, "Environment::Method::ChangeTrayIcon", Err.Description)
End Sub
'Remove a "form" from the systray
Public Sub RemoveTrayIcon(ByVal FormToEdit As Form)
On Error GoTo err_method_RemoveTrayIcon
    Call Shell_NotifyIcon(NIM_DELETE, ShellIcon)
Exit Sub
err_method_RemoveTrayIcon:
    Call KillWithError(ERR_METHOD, "Environment::Method::RemoveTrayIcon", Err.Description)
End Sub
Public Function CSecondsToString(ByVal SecondValue As Long)
On Error GoTo err_method_CSecondsToString
    CSecondsToString = Format(Fix(SecondValue / 3600), "#0") & "hr(s), " & Format(Fix((SecondValue Mod 3600) / 60), "#0") & "min(s), " & Format(SecondValue Mod 60, "00") & "sec(s)"
Exit Function
err_method_CSecondsToString:
    Call KillWithError(ERR_METHOD, "Environment::Method::CSecondsToString", Err.Description)
End Function
Public Sub KillWithError(ByVal ErrCode As Long, ByVal ErrSource As String, ByVal ErrDesc As String)
    Call MsgBox("Error (" & ErrCode & ")" & vbCrLf & ErrSource & vbCrLf & vbCrLf & ErrDesc, vbOKOnly Or vbCritical, "Error: " & ErrCode)
    End
End Sub

Attribute VB_Name = "IO"
Option Explicit
'********************************************************
'* Name: IO Module
'* Type: VB Module
'* Author: Darryn Britton
'* Date: 15/04/2002
'********************************************************
' Functions for interaction with the filesystem
'********************************************************
Private lngFileHandle As Long
Private arrFileLine As Variant
Private strName As String
Private strValue As String
'Get a value from a defined config file based on the name.
'Style = "Name::Value"
Public Function GetConfigString(ByVal Search As String) As String
On Error GoTo err_method_GetConfigString
Dim strConfigFile As String
    strConfigFile = App.Path & "\" & FILE_CONFIG
    lngFileHandle = FreeFile()
    If Dir(strConfigFile) <> "" Then
        Open strConfigFile For Input As lngFileHandle
        Do While Not EOF(lngFileHandle)
            Input #lngFileHandle, arrFileLine
            'skip comments
            If Left(arrFileLine, 2) <> "##" And arrFileLine <> "" Then
                arrFileLine = Split(arrFileLine, "::")
                If (UBound(arrFileLine) = 1) Then
                    If (arrFileLine(0) = Search) Then
                        GetConfigString = arrFileLine(1)
                        Exit Do
                    End If
                End If
            End If
        Loop
        Close lngFileHandle
    End If
Exit Function
err_method_GetConfigString:
    Call Err.Raise(ERR_METHOD, "IO::Method::GetConfigString", Err.Description)
End Function
'Change a value from a defined config file based on the name.
'Style = "Name::Value"
Public Function EditConfigString(ByVal Search As String, ByVal NewValue As String) As String
On Error GoTo err_method_EditConfigString
Dim strConfigFile As String
Dim strFileContents As String
    strConfigFile = App.Path & "\" & FILE_CONFIG
    lngFileHandle = FreeFile()
    If Dir(strConfigFile) <> "" Then
        Open strConfigFile For Input As lngFileHandle
        strFileContents = ""
        Do While Not EOF(lngFileHandle)
            Input #lngFileHandle, arrFileLine
            'skip comments
            If Left(arrFileLine, 2) <> "##" Then
                If (arrFileLine <> "") Then
                    arrFileLine = Split(arrFileLine, "::")
                    If (UBound(arrFileLine) = 1) Then
                        If (arrFileLine(0) = Search) Then
                            strFileContents = strFileContents & arrFileLine(0) & "::" & NewValue & vbCrLf
                        Else
                            strFileContents = strFileContents & arrFileLine(0) & "::" & arrFileLine(1) & vbCrLf
                        End If
                    Else
                        strFileContents = strFileContents & arrFileLine(0) & vbCrLf
                    End If
                End If
            Else
                strFileContents = strFileContents & arrFileLine & vbCrLf
            End If
        Loop
        Close lngFileHandle
    End If
    If (strFileContents <> "") Then
        strConfigFile = App.Path & "\" & FILE_CONFIG
        lngFileHandle = FreeFile()
        Open strConfigFile For Output As lngFileHandle
        Print #lngFileHandle, strFileContents
        Close lngFileHandle
    End If
Exit Function
err_method_EditConfigString:
    Call Err.Raise(ERR_METHOD, "IO::Method::EditConfigString", Err.Description)
End Function
'Put a value into a defined config file.
'Style = "Name::Value"
Public Function PutConfigString(ByVal Name As String, ByVal Value As String)
On Error GoTo err_method_PutConfigString
Dim strConfigFile As String
    strConfigFile = App.Path & "\" & FILE_CONFIG
    lngFileHandle = FreeFile()
    Open strConfigFile For Append As lngFileHandle
    Print #lngFileHandle, Name & "::" & Value & vbCrLf
    Close lngFileHandle
Exit Function
err_method_PutConfigString:
    Call Err.Raise(ERR_METHOD, "IO::Method::PutConfigString", Err.Description)
End Function
'Get a command from a defined hash file based on the name.
'Style = "Name::Command"
Public Function GetBotCommand(ByVal BotCommand As String) As String
On Error GoTo err_method_GetBotCommand
Dim strCommandHashFile As String
    strCommandHashFile = App.Path & "\" & FILE_HASH
    lngFileHandle = FreeFile()
    If Dir(strCommandHashFile) <> "" Then
        Open strCommandHashFile For Input As lngFileHandle
        Do While Not EOF(lngFileHandle)
            Input #lngFileHandle, arrFileLine
            'skip comments
            If Left(arrFileLine, 2) <> "##" And arrFileLine <> "" Then
                arrFileLine = Split(arrFileLine, "::")
                If (UBound(arrFileLine) = 1) Then
                    If (arrFileLine(0) = BotCommand) Then
                        GetBotCommand = arrFileLine(1)
                        Exit Do
                    End If
                End If
            End If
        Loop
        Close lngFileHandle
    End If
Exit Function
err_method_GetBotCommand:
    Call Err.Raise(ERR_METHOD, "IO::Method::GetBotCommand", Err.Description)
End Function
'Get a replacement definition from a command.
'Style = "Command::Replacement"
Public Function ReplaceBotCommandDefines(ByVal BotCommandString As String) As String
On Error GoTo err_method_ReplaceBotCommandDefines
Dim strCommandHashFile As String
Dim strReplacedDefines As String
    strCommandHashFile = App.Path & "\" & FILE_DEFINE
    lngFileHandle = FreeFile()
    If Dir(strCommandHashFile) <> "" Then
        Open strCommandHashFile For Input As lngFileHandle
        strReplacedDefines = BotCommandString
        Do While Not EOF(lngFileHandle)
            Input #lngFileHandle, arrFileLine
            'skip comments
            If Left(arrFileLine, 2) <> "##" And arrFileLine <> "" Then
                arrFileLine = Split(arrFileLine, "::")
                If (UBound(arrFileLine) = 1) Then
                    strReplacedDefines = Replace(strReplacedDefines, arrFileLine(0), arrFileLine(1))
                End If
            End If
        Loop
        Close lngFileHandle
    End If
    ReplaceBotCommandDefines = strReplacedDefines
Exit Function
err_method_ReplaceBotCommandDefines:
    Call Err.Raise(ERR_METHOD, "IO::Method::ReplaceBotCommandDefines", Err.Description)
End Function
'Add an entry to the log file. The log file is generated based on the date.
Public Sub AddLogEntry(ByVal LogData As String, ByVal Direction As ENUM_DIRECTION, ByVal EventType As ENUM_LOGTYPE)
On Error GoTo err_method_AddLogEntry
Dim strLogFile As String
    strLogFile = App.Path & PATH_LOGDATA & Replace(Date, "/", "_") & ".log"
    LogData = ReplaceEscapeChars(LogData)
    LogData = Replace(LogData, Chr(10), "\cr")
    LogData = Replace(LogData, Chr(13), "\lf")
    lngFileHandle = FreeFile()
    Open strLogFile For Append As #lngFileHandle
    Print #lngFileHandle, "[" & Time() & "] " & EventType & " " & Direction & " " & LogData
    Close lngFileHandle
Exit Sub
err_method_AddLogEntry:
    Call Err.Raise(ERR_METHOD, "IO::Method::AddLogEntry", Err.Description)
End Sub


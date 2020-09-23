VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'********************************************************
'* Name: Bot Object
'* Type: VB Class Module
'* Author: Darryn Britton
'* Date: 15/04/2002
'********************************************************
' Instance of a bot. Once object has been created, it can connect
' to an IRC server, receive and parse messages, send data and connect
' to a (MS)SQL Server and log data
'********************************************************
Dim m_Nick As String
Dim m_Username As String
Dim m_DateStarted As Date
Dim m_State As ENUM_BOTSTATE
Dim m_Address As String
Dim m_Port As Integer
Dim m_Version As String
Dim m_Owner As String
Dim m_Password As String
Public Property Let Password(ByVal p_Password As String)
On Error GoTo err_let_Password
    m_Password = p_Password
Exit Property
err_let_Password:
    Call Err.Raise(ERR_PROPERTYLET, "Bot::PropLet::Password", Err.Description)
End Property
Public Property Get Password() As String
On Error GoTo err_get_Password
    Password = m_Password
Exit Property
err_get_Password:
    Call Err.Raise(ERR_PROPERTYGET, "Bot::PropGet::Password", Err.Description)
End Property
Public Property Let Owner(ByVal p_Owner As String)
On Error GoTo err_let_Owner
    m_Owner = p_Owner
Exit Property
err_let_Owner:
    Call Err.Raise(ERR_PROPERTYLET, "Bot::PropLet::Owner", Err.Description)
End Property
Public Property Get Owner() As String
On Error GoTo err_get_Owner
    Owner = m_Owner
Exit Property
err_get_Owner:
    Call Err.Raise(ERR_PROPERTYGET, "Bot::PropGet::Owner", Err.Description)
End Property
Public Property Let Version(ByVal p_Version As String)
On Error GoTo err_let_Version
    m_Version = p_Version
Exit Property
err_let_Version:
    Call Err.Raise(ERR_PROPERTYLET, "Bot::PropLet::Version", Err.Description)
End Property
Public Property Get Version() As String
On Error GoTo err_get_Version
    Version = m_Version
Exit Property
err_get_Version:
    Call Err.Raise(ERR_PROPERTYGET, "Bot::PropGet::Version", Err.Description)
End Property
Public Property Let Nick(ByVal p_Nick As String)
On Error GoTo err_let_Nick
    m_Nick = p_Nick
Exit Property
err_let_Nick:
    Call Err.Raise(ERR_PROPERTYLET, "Bot::PropLet::Nick", Err.Description)
End Property
Public Property Get Nick() As String
On Error GoTo err_get_Nick
    Nick = m_Nick
Exit Property
err_get_Nick:
    Call Err.Raise(ERR_PROPERTYGET, "Bot::PropGet::Nick", Err.Description)
End Property
Public Property Let Username(ByVal p_Username As String)
On Error GoTo err_let_Username
    m_Username = p_Username
Exit Property
err_let_Username:
    Call Err.Raise(ERR_PROPERTYLET, "Bot::PropLet::Username", Err.Description)
End Property
Public Property Get Username() As String
On Error GoTo err_get_Username
    Username = m_Username
Exit Property
err_get_Username:
    Call Err.Raise(ERR_PROPERTYGET, "Bot::PropGet::Username", Err.Description)
End Property
Public Property Let DateStarted(ByVal p_DateStarted As Date)
On Error GoTo err_let_DateStarted
    m_DateStarted = p_DateStarted
Exit Property
err_let_DateStarted:
    Call Err.Raise(ERR_PROPERTYLET, "Bot::PropLet::DateStarted", Err.Description)
End Property
Public Property Get DateStarted() As Date
On Error GoTo err_get_DateStarted
    DateStarted = m_DateStarted
Exit Property
err_get_DateStarted:
    Call Err.Raise(ERR_PROPERTYGET, "Bot::PropGet::DateStarted", Err.Description)
End Property
Public Property Let State(ByVal p_State As ENUM_BOTSTATE)
On Error GoTo err_let_State
    m_State = p_State
Exit Property
err_let_State:
    Call Err.Raise(ERR_PROPERTYLET, "Bot::PropLet::State", Err.Description)
End Property
Public Property Get State() As ENUM_BOTSTATE
On Error GoTo err_get_State
    State = m_State
Exit Property
err_get_State:
    Call Err.Raise(ERR_PROPERTYGET, "Bot::PropGet::State", Err.Description)
End Property
Public Property Let Address(ByVal p_Address As String)
On Error GoTo err_let_Address
    m_Address = p_Address
Exit Property
err_let_Address:
    Call Err.Raise(ERR_PROPERTYLET, "Bot::PropLet::Address", Err.Description)
End Property
Public Property Get Address() As String
On Error GoTo err_get_Address
    Address = m_Address
Exit Property
err_get_Address:
    Call Err.Raise(ERR_PROPERTYGET, "Bot::PropGet::Address", Err.Description)
End Property
Public Property Let Port(ByVal p_Port As Integer)
On Error GoTo err_let_Port
    m_Port = p_Port
Exit Property
err_let_Port:
    Call Err.Raise(ERR_PROPERTYLET, "Bot::PropLet::Port", Err.Description)
End Property
Public Property Get Port() As Integer
On Error GoTo err_get_Port
    Port = m_Port
Exit Property
err_get_Port:
    Call Err.Raise(ERR_PROPERTYGET, "Bot::PropGet::Port", Err.Description)
End Property
'Load the bot's details from the config file
Public Sub LoadConfig()
    Me.Address = GetConfigString("server")
    Me.Nick = GetConfigString("nick")
    Me.Port = GetConfigString("port")
    Me.State = botIdle
    Me.Username = GetConfigString("username")
    Me.Owner = GetConfigString("owner")
    Me.Password = GetConfigString("password")
    Me.Version = GetConfigString("version") & " v" & App.Major & "." & App.Minor
    Call AddLogEntry("Loaded configuration file", directionIn, botEvent)
End Sub
'Connect the bot to it's designated server
Public Sub Connect()
    'if the server address is nothing or the port isn't set, log the error and leave
    If ((Me.Address = "") Or (Me.Port <= 0)) Then
        Call AddLogEntry("Connection failed, address details missing", directionNull, botEvent)
    End If
    'start winsock and connect the bot
    Me.DateStarted = Now()
    Call AddLogEntry(VBPrintF("Connecting to %v:%v", Me.Address, Me.Port), directionNull, botEvent)
    frmMain.sckClient.RemoteHost = Me.Address
    frmMain.sckClient.RemotePort = Me.Port
    frmMain.sckClient.Connect
End Sub
'Parse an IRC Message
Public Sub ParseData(ByVal DataString As String)
Dim Msg As New Message
Dim strSourceUser As String
Dim strCommand As String
Dim strMessage As String
Dim strErrorMessage As String
    'if it isn't a proper message, check for errors, otherwise discard it
    If (Right(DataString, 2) <> vbCrLf) Then Exit Sub
    If (Left(DataString, 1) = ":") Then
        strSourceUser = Mid(DataString, 2, InStr(1, DataString, " ") - 2)
        'it's a normal chatter and not the server
        'strSourceUser = "furry-animal!furryone@192.26.56.103"
        Msg.Received = Now()
        If (InStr(1, strSourceUser, "!") > 0 And InStr(1, strSourceUser, "@") > 0) Then
            Msg.Source.Nick = Mid(strSourceUser, 1, InStr(1, strSourceUser, "!") - 1)
            Msg.Source.Username = Mid(strSourceUser, InStr(1, strSourceUser, "!") + 1, (InStr(1, strSourceUser, "@") - Len(Msg.Source.Nick)) - 2)
            Msg.Source.Host = Mid(strSourceUser, InStr(1, strSourceUser, "@") + 1)
        Else
            'save the nick as the servers name
            Msg.Source.Nick = strSourceUser
        End If
        If InStr(2, DataString, ":") Then
            Msg.Command = Mid(DataString, InStr(1, DataString, " ") + 1, InStr(2, DataString, ":") - InStr(1, DataString, " ") - 2)
            If InStr(Msg.Command, " ") Then
                Msg.Destination = Mid(Msg.Command, InStr(1, Msg.Command, " ") + 1)
                Msg.Command = Mid(Msg.Command, 1, InStr(1, Msg.Command, " ") - 1)
            End If
        End If
        Msg.Message = Replace(Mid(DataString, InStr(2, DataString, ":") + 1), vbCrLf, "")
        'parse the command
        Call Me.ParseCommand(Msg)
    'not a "real" message, might be an error message though
    ElseIf (Left(DataString, 5) = "ERROR") Then
        strErrorMessage = Mid(DataString, InStr(1, DataString, "(") + 1, InStr(1, DataString, ")") - (InStr(1, DataString, "(") + 1))
        Call AddLogEntry(VBPrintF("ERROR(IRC) encountered: %v", strErrorMessage), directionIn, ircEvent)
        frmMain.sckClient.Close
    ElseIf (Left(DataString, 4) = "PING") Then
        Call Me.Send(VBPrintF("PONG %v", Mid(DataString, 6)))
    End If
    If (IsObject(Msg)) Then Set Msg = Nothing
End Sub
'Parse a command from the IRC message
Public Sub ParseCommand(ByVal Msg As Message)
Dim strAction As String
Dim arrParams As Variant
Dim lngParamCounter As Long
    If ((Msg.Command = "") Or (Msg.Message = "")) Then Exit Sub
    Select Case UCase(Msg.Command)
        'connected
        Case "001"
            Me.State = botWorking
        'server is requesting us to send them our details
        Case "NOTICE"
            Select Case UCase(Msg.Destination)
                Case "AUTH"
                    'if we're being asked to AUTH, send bot details!
                    If (Me.State <> botRegistered) Then
                        Call Me.Send(VBPrintF("NICK %v\n", Me.Nick))
                        Call Me.Send(VBPrintF("USER %v %v %v :%v\n", Me.Username, "NULL", "NULL", Me.Version))
                        Me.State = botRegistered
                    End If
                Case UCase(Me.Nick)
                    If (UCase(Msg.Source.Nick) = "NICKSERV") Then
                        'if nickserv is asking us to IDENTIFY with our password, send it!
                        If InStr(1, Msg.Message, "IDENTIFY") > 1 Then
                            Call Me.Send(VBPrintF("PRIVMSG %v :IDENTIFY %v\n", Msg.Source.Nick, Me.Password))
                        End If
                    End If
            End Select
        'private message
        Case "PRIVMSG"
            If (UCase(Msg.Destination) = UCase(Me.Nick)) Or Left(Msg.Destination, 1) = "#" Then
                'someone trying to CTCP VERSION the bot
                If (Msg.Message = "VERSION") Then
                    Call Me.Send(VBPrintF("NOTICE %v :VERSION %v\n", Msg.Source.Nick, Me.Version))
                ElseIf (Msg.Message = "REHASH") Then
                    Call Me.Send(VBPrintF("NOTICE %v :%v\n", Msg.Source.Nick, "Rehashed.."))
                'someone trying to CTCP PING the bot
                ElseIf (Msg.Message Like "PING*") Then
                    Call Me.Send(VBPrintF("NOTICE %v :PING %v\n", Msg.Source.Nick, "NULL Reply"))
                'someones trying to PRIVMSG the bot with general text,
                'is it a command or a message?
                ElseIf (Left(Msg.Message, 1) = ".") Then
                'command, parse it..
                    If InStr(1, Msg.Message, " ") Then
                        Msg.Action = Mid(Msg.Message, 2, InStr(2, Msg.Message, " ") - 2)
                    Else
                        Msg.Action = Mid(Msg.Message, 2)
                    End If
                    Msg.Message = Mid(Msg.Message, 2)
                    Msg.Message = Mid(Msg.Message, Len(Msg.Action) + 2)
                    arrParams = Split(Msg.Message, " ")
                    For lngParamCounter = LBound(arrParams) To UBound(arrParams)
                        Call Msg.Params.Add(arrParams(lngParamCounter), "k" & lngParamCounter)
                    Next lngParamCounter
                    Call ParseAction(Msg)
                Else
                    If (UCase(Msg.Destination) = UCase(Me.Nick)) Then
                        Call Me.Send(VBPrintF("PRIVMSG %v :<%v> %v\n", Me.Owner, Msg.Source.Nick, Msg.Message))
                    End If
                    'general message .. ignore it.
                    If (Left(Msg.Message, Len(Me.Nick) + 1) = Me.Nick & ":") Then
                        If InStr(Msg.Message, "test") > 0 Then
                            Msg.Message = "test"
                            'Call Me.Send(VBPrintF("PRIVMSG %v :%v\n", Msg.Destination, Msg.Message))
                        End If
                    End If
                End If
            End If
        Case "NICK"
            If UCase(Msg.Source.Nick) = UCase(Me.Nick) Then
                Me.Nick = Msg.Message
            End If
    End Select
End Sub
'send data via the socket to the server
Public Sub Send(ByVal DataString As String)
    If (DataString = "") Then Exit Sub
    If frmMain.sckClient.State = sckConnected Then frmMain.sckClient.SendData (DataString)
    Call AddLogEntry(VBPrintF("Data out (%vb): %v", Len(DataString), DataString), directionOut, socketEvent)
End Sub
'parse a defined action from a user via the commands hash file
Public Sub ParseAction(ByVal Msg As Message)
Dim strCommandString As String
Dim lngPartCounter As Long
    Select Case LCase(Msg.Action)
        Case "nick"
            If (Msg.Params.Count > 0) Then
                Me.Nick = Msg.Message
                Call Me.Send(VBPrintF("NICK %v\n", Msg.Params(1)))
            Else
                Call Me.Send(VBPrintF("NOTICE %v :%v\n", Msg.Source.Nick, "Error. Syntax: .nick [newnick]"))
            End If
        Case "uptime"
            Call Me.Send(VBPrintF("PRIVMSG %v :\bUp since\b: %v: %v\n", Msg.Source.Nick, Me.DateStarted, CSecondsToString(DateDiff("s", Me.DateStarted, Now()))))
        Case "join"
            If (Msg.Params.Count > 0) Then
                Call Me.Send(VBPrintF("JOIN %v\n", Msg.Params(1)))
            Else
                Call Me.Send(VBPrintF("NOTICE %v :%v\n", Msg.Source.Nick, "Error. Syntax: .join #[channel]"))
            End If
        Case "part"
            If (Msg.Params.Count > 0) Then
                Call Me.Send(VBPrintF("PART %v\n", Msg.Params(1)))
            Else
                Call Me.Send(VBPrintF("NOTICE %v :%v\n", Msg.Source.Nick, "Error. Syntax: .part #[channel]"))
            End If
        Case "quit"
            If (Msg.Params.Count > 0) Then
                Call Me.Send(VBPrintF("QUIT :%v\n", Msg.Message))
            Else
                Call Me.Send(VBPrintF("NOTICE %v :%v\n", Msg.Source.Nick, "Error. Syntax: .quit [quitmessage]"))
            End If
        Case "topic"
            If (Msg.Params.Count > 2) Then
                For lngPartCounter = 2 To Msg.Params.Count
                    strCommandString = strCommandString & Msg.Params(lngPartCounter) & " "
                Next lngPartCounter
                strCommandString = Left(strCommandString, Len(strCommandString) - 1)
                Call Me.Send(VBPrintF("TOPIC %v :%v\n", Msg.Params(1), strCommandString))
            Else
                Call Me.Send(VBPrintF("NOTICE %v :%v\n", Msg.Source.Nick, "Error. Syntax: .topic #[channel] [newtopic]"))
            End If
        Case "say"
            If (Msg.Params.Count > 1) Then
                For lngPartCounter = 2 To Msg.Params.Count
                    strCommandString = strCommandString & Msg.Params(lngPartCounter) & " "
                Next lngPartCounter
                strCommandString = Left(strCommandString, Len(strCommandString) - 1)
                Call Me.Send(VBPrintF("PRIVMSG %v :%v\n", Msg.Params(1), strCommandString))
            Else
                Call Me.Send(VBPrintF("NOTICE %v :%v\n", Msg.Source.Nick, "Error. Syntax: .say [#channel]|[nick] [message]"))
            End If
        Case "notice"
            If (Msg.Params.Count > 1) Then
                For lngPartCounter = 2 To Msg.Params.Count
                    strCommandString = strCommandString & Msg.Params(lngPartCounter) & " "
                Next lngPartCounter
                strCommandString = Left(strCommandString, Len(strCommandString) - 1)
                Call Me.Send(VBPrintF("NOTICE %v :%v\n", Msg.Params(1), strCommandString))
            Else
                Call Me.Send(VBPrintF("NOTICE %v :%v\n", Msg.Source.Nick, "Error. Syntax: .notice [#channel]|[nick] [message]"))
            End If
        Case "action"
            If (Msg.Params.Count > 1) Then
                For lngPartCounter = 2 To Msg.Params.Count
                    strCommandString = strCommandString & Msg.Params(lngPartCounter) & " "
                Next lngPartCounter
                strCommandString = Left(strCommandString, Len(strCommandString) - 1)
                Call Me.Send(VBPrintF("PRIVMSG %v :ACTION %v\n", Msg.Params(1), strCommandString))
            Else
                Call Me.Send(VBPrintF("NOTICE %v :%v\n", Msg.Source.Nick, "Error. Syntax: .action [#channel]|[nick] [action]"))
            End If
        Case "kick"
            If (Msg.Params.Count > 2) Then
                For lngPartCounter = 3 To Msg.Params.Count
                    strCommandString = strCommandString & Msg.Params(lngPartCounter) & " "
                Next lngPartCounter
                strCommandString = Left(strCommandString, Len(strCommandString) - 1)
                Call Me.Send(VBPrintF("KICK %v %v :%v\n", Msg.Params(1), Msg.Params(2), strCommandString))
            Else
                Call Me.Send(VBPrintF("NOTICE %v :%v\n", Msg.Source.Nick, "Error. Syntax: .kick [#channel] [nick] [message]"))
            End If
        Case "mode"
            If (Msg.Params.Count > 2) Then
                For lngPartCounter = 3 To Msg.Params.Count
                    strCommandString = strCommandString & Msg.Params(lngPartCounter) & " "
                Next lngPartCounter
                strCommandString = Left(strCommandString, Len(strCommandString) - 1)
                Call Me.Send(VBPrintF("MODE %v %v %v\n", Msg.Params(1), Msg.Params(2), strCommandString))
            Else
                Call Me.Send(VBPrintF("NOTICE %v :%v\n", Msg.Source.Nick, "Error. Syntax: .mode [#channel] [mode prefix] [mode]"))
            End If
    End Select
End Sub
'parse a defined action from a user (within a channel)
Public Sub ParseChannelAction(ByVal CurrentMessage As Message, ByVal Action As String, MessageString As String)
    Select Case LCase(Action)
        
    End Select
End Sub

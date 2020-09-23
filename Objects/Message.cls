VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'********************************************************
'* Name: Message Object
'* Type: VB Class Module
'* Author: Darryn Britton
'* Date: 15/04/2002
'********************************************************
' Instance of an IRC Message. 1 Message object is created for each
' IRC message (datastream) the system receives
'********************************************************
Dim m_Source As New User
Dim m_Destination As String
Dim m_Message As String
Dim m_Command As String
Dim m_Params As New Collection
Dim m_Received As Date
Dim m_Action As String
Public Property Let Action(ByVal p_Action As String)
On Error GoTo err_let_Action
    m_Action = p_Action
Exit Property
err_let_Action:
    Call Err.Raise(ERR_PROPERTYLET, "Bot::PropLet::Action", Err.Description)
End Property
Public Property Get Action() As String
On Error GoTo err_get_Action
    Action = m_Action
Exit Property
err_get_Action:
    Call Err.Raise(ERR_PROPERTYGET, "Bot::PropGet::Action", Err.Description)
End Property
Public Property Let Source(ByVal p_Source As User)
On Error GoTo err_let_Source
    Set m_Source = p_Source
Exit Property
err_let_Source:
    Call Err.Raise(ERR_PROPERTYLET, "Bot::PropLet::Source", Err.Description)
End Property
Public Property Get Source() As User
On Error GoTo err_get_Source
    Set Source = m_Source
Exit Property
err_get_Source:
    Call Err.Raise(ERR_PROPERTYGET, "Bot::PropGet::Source", Err.Description)
End Property
Public Property Let Destination(ByVal p_Destination As String)
On Error GoTo err_let_Destination
    m_Destination = p_Destination
Exit Property
err_let_Destination:
    Call Err.Raise(ERR_PROPERTYLET, "Bot::PropLet::Destination", Err.Description)
End Property
Public Property Get Destination() As String
On Error GoTo err_get_Destination
    Destination = m_Destination
Exit Property
err_get_Destination:
    Call Err.Raise(ERR_PROPERTYGET, "Bot::PropGet::Destination", Err.Description)
End Property
Public Property Let Message(ByVal p_Message As String)
On Error GoTo err_let_Message
    m_Message = p_Message
Exit Property
err_let_Message:
    Call Err.Raise(ERR_PROPERTYLET, "Bot::PropLet::Message", Err.Description)
End Property
Public Property Get Message() As String
On Error GoTo err_get_Message
    Message = m_Message
Exit Property
err_get_Message:
    Call Err.Raise(ERR_PROPERTYGET, "Bot::PropGet::Message", Err.Description)
End Property
Public Property Let Command(ByVal p_Command As String)
On Error GoTo err_let_Command
    m_Command = p_Command
Exit Property
err_let_Command:
    Call Err.Raise(ERR_PROPERTYLET, "Bot::PropLet::Command", Err.Description)
End Property
Public Property Get Command() As String
On Error GoTo err_get_Command
    Command = m_Command
Exit Property
err_get_Command:
    Call Err.Raise(ERR_PROPERTYGET, "Bot::PropGet::Command", Err.Description)
End Property
Public Property Let Params(ByVal p_Params As Collection)
On Error GoTo err_let_Params
    Set m_Params = p_Params
Exit Property
err_let_Params:
    Call Err.Raise(ERR_PROPERTYLET, "Bot::PropLet::Params", Err.Description)
End Property
Public Property Get Params() As Collection
On Error GoTo err_get_Params
    Set Params = m_Params
Exit Property
err_get_Params:
    Call Err.Raise(ERR_PROPERTYGET, "Bot::PropGet::Params", Err.Description)
End Property
Public Property Let Received(ByVal p_Received As Date)
On Error GoTo err_let_Received
    m_Received = p_Received
Exit Property
err_let_Received:
    Call Err.Raise(ERR_PROPERTYLET, "Bot::PropLet::Received", Err.Description)
End Property
Public Property Get Received() As Date
On Error GoTo err_get_Received
    Received = m_Received
Exit Property
err_get_Received:
    Call Err.Raise(ERR_PROPERTYGET, "Bot::PropGet::Received", Err.Description)
End Property

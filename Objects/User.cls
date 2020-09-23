VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'********************************************************
'* Name: User Object
'* Type: VB Class Module
'* Author: Darryn Britton
'* Date: 15/04/2002
'********************************************************
' Instance of an IRC User (chatter). Contains details and addresses
' about a user for sending/receiving messages and logging stats.
'********************************************************
Dim m_Nick As String
Dim m_Username As String
Dim m_Host As String
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
Public Property Let Host(ByVal p_Host As String)
On Error GoTo err_let_Host
    m_Host = p_Host
Exit Property
err_let_Host:
    Call Err.Raise(ERR_PROPERTYLET, "Bot::PropLet::Host", Err.Description)
End Property
Public Property Get Host() As String
On Error GoTo err_get_Host
    Host = m_Host
Exit Property
err_get_Host:
    Call Err.Raise(ERR_PROPERTYGET, "Bot::PropGet::Host", Err.Description)
End Property


VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   525
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   1560
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   525
   ScaleWidth      =   1560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock sckClient 
      Left            =   120
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuSys 
      Caption         =   "mnuSys"
      Begin VB.Menu mnuSysClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Client As New Bot
Private Sub Form_Load()
    Call AddLogEntry(VBPrintF("Application started with params: ""%v""", Command), directionIn, appEvent)
    Call AddTrayIcon(Me, "Syphon [Idle]")
    Me.Hide
    Client.LoadConfig
    Client.Connect
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngMessage As Long
    lngMessage = X / Screen.TwipsPerPixelX
    Select Case lngMessage
        Case WM_RBUTTONUP
            Call Me.PopupMenu(Me.mnuSys)
    End Select
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call AddLogEntry(VBPrintF("Application terminated with killCode: %v", UnloadMode), directionIn, appEvent)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If (Me.sckClient.State = sckOpen Or Me.sckClient.State = sckConnected) Then Me.sckClient.Close
    Call RemoveTrayIcon(Me)
End Sub
Private Sub mnuSysClose_Click()
    Call Unload(Me)
End Sub
Private Sub sckClient_Connect()
    Call AddLogEntry(VBPrintF("Connected. ID: %v", Me.sckClient.SocketHandle), directionNull, socketEvent)
End Sub
Private Sub sckClient_DataArrival(ByVal bytesTotal As Long)
On Error GoTo err_method_sckClient_DataArrival
Dim varData As Variant
Dim lngDataPartCounter As Long
    Call Me.sckClient.GetData(varData, vbString)
    If (varData <> "") Then
        varData = Split(varData, vbCrLf)
        If (UBound(varData) >= 0) Then
            For lngDataPartCounter = LBound(varData) To UBound(varData)
                If (varData(lngDataPartCounter) <> "") Then
                    Call AddLogEntry(VBPrintF("Data in:\t%v", varData(lngDataPartCounter)), directionIn, socketEvent)
                    Call Client.ParseData(varData(lngDataPartCounter) & vbCrLf)
                End If
            Next lngDataPartCounter
        End If
    End If
Exit Sub
err_method_sckClient_DataArrival:
    Call KillWithError(ERR_METHOD, "Form::Method::DataArrival", Err.Description)
End Sub
Private Sub sckClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call AddLogEntry(VBPrintF("Error ""%v"": %v", Number, Description), directionNull, socketEvent)
End Sub

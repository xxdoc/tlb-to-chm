VERSION 5.00
Begin VB.Form frmProcWnd 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   765
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   5130
   Icon            =   "frmProcWnd.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   420
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "frmProcWnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mblnCancel As Boolean
Private mblnVisible As Boolean
Private mintMode As Integer
Private mblnEnableMenu As Boolean

Public Sub ShowProcess(ByVal msg As String)
   ShowMe
   Label1.Caption = msg
   
   If mintMode < 1 Then
      Label2.Visible = False
      Me.Height = 900
      mintMode = 1
   End If
   
   DoEvents
End Sub

Public Sub ShowProcess2(ByVal msg As String)
   ShowMe
   Label2.Caption = msg
   
   If mintMode < 2 Then
      Label2.Visible = True
      mintMode = 2
      Me.Height = 1200
   End If
   
   DoEvents
End Sub

Private Sub ShowMe()
   If Not mblnVisible Then
      Me.Show vbModeless, frmMDI
      mblnVisible = True
   End If
   
   While frmMDI.mnuPauseConvertionProcess.Checked And Not mblnCancel
      Sleep 200
      DoEvents
   Wend
End Sub

Public Sub Prepare(ByVal blnEnableMenu As Boolean)
   mblnEnableMenu = blnEnableMenu
   mintMode = 0
   mblnCancel = False
   If blnEnableMenu Then
      frmMDI.mnuPauseConvertionProcess.Checked = False
      frmMDI.mnuPauseConvertionProcess.Enabled = True
   End If
End Sub

Public Sub Finish()
   Me.Hide
   mblnVisible = False
   frmMDI.mnuPauseConvertionProcess.Checked = False
   frmMDI.mnuPauseConvertionProcess.Enabled = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyEscape Then
      mblnCancel = True
      mblnVisible = False
      Me.Hide
   End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If mblnVisible And UnloadMode = vbFormControlMenu Then
      Cancel = True
      mblnCancel = True
      mblnVisible = False
      Me.Hide
   End If
End Sub

Private Sub Form_Initialize()
   LogInitialize Me
End Sub

Private Sub Form_Terminate()
   LogTerminate Me
End Sub


VERSION 5.00
Object = "{03B209C4-7ADD-4264-A128-4BFAD583CACD}#1.0#0"; "PuppyResizer.ocx"
Begin VB.MDIForm frmMDI 
   BackColor       =   &H8000000C&
   Caption         =   "tlb-to-chm"
   ClientHeight    =   6915
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11850
   Icon            =   "frmMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin PuppyResizerLib.PuppyResizer PuppyResizer1 
      Left            =   180
      Top             =   300
      _ExtentX        =   529
      _ExtentY        =   529
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   2775
      Left            =   0
      ScaleHeight     =   2715
      ScaleWidth      =   11790
      TabIndex        =   0
      Top             =   4140
      Width           =   11850
      Begin VB.CommandButton cmdOpenDir 
         Caption         =   "@"
         Height          =   315
         Left            =   9600
         TabIndex        =   7
         Tag             =   "L"
         ToolTipText     =   "Open output directory with explorer."
         Top             =   60
         Width           =   315
      End
      Begin VB.ComboBox comboOutputDir 
         Height          =   300
         Left            =   1860
         TabIndex        =   6
         Tag             =   "W"
         ToolTipText     =   "Input output directory's path."
         Top             =   60
         Width           =   7095
      End
      Begin VB.CommandButton cmdSelectDir 
         Caption         =   "..."
         Height          =   315
         Left            =   9060
         TabIndex        =   4
         Tag             =   "L"
         ToolTipText     =   "Select output directory."
         Top             =   60
         Width           =   435
      End
      Begin VB.CommandButton cmdGenerateCHM 
         Caption         =   "&Generate CHM"
         Default         =   -1  'True
         Height          =   315
         Left            =   10020
         TabIndex        =   2
         Tag             =   "L"
         Top             =   60
         Width           =   1635
      End
      Begin VB.ListBox lstMsg 
         Height          =   2100
         IntegralHeight  =   0   'False
         Left            =   180
         TabIndex        =   3
         Tag             =   "HW"
         Top             =   480
         Width           =   11475
      End
      Begin VB.Label Label2 
         Caption         =   "&Output directory:"
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Note: all of the chm files should be in the same folder."
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   180
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   5040
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuTreatHelpStringAsHtml 
         Caption         =   "&Treat help string as html"
      End
      Begin VB.Menu mnuClearSelectedControlHistory 
         Caption         =   "&Clear selected control's history"
      End
      Begin VB.Menu mnuSpliter1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenFinishedChmFiles 
         Caption         =   "&Open finished chm files"
      End
      Begin VB.Menu mnuPauseConvertionProcess 
         Caption         =   "&Pause convertion process"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuModes 
      Caption         =   "&Modes"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About ..."
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mdblDelta As Double
Private WithEvents mobjLogMgr As LogManagement
Attribute mobjLogMgr.VB_VarHelpID = -1

Public Sub FormActivate(ByVal frm As Form)
   If frm Is frmListMode Or frm Is frmNameMode Then
      Picture1.Visible = True
   ElseIf frm Is frmMergeMode Then
      Picture1.Visible = False
   Else
      Debug.Assert False
   End If
   cmdGenerateCHM.Enabled = True
   lstMsg.Enabled = True
End Sub

Private Sub SetChildrenEnabled(ByVal blnEnabled As Boolean)
   Dim frm As Form
   For Each frm In Forms
      If Not TypeOf frm Is MDIForm Then
         If frm.MDIChild Then
            frm.Enabled = blnEnabled
         End If
      End If
   Next
End Sub

Private Sub SetPictureBoxControlsEnabled(ByVal blnEnabled As Boolean)
   On Error Resume Next
   Dim ctrl As Control
   For Each ctrl In Me.Controls
      If TypeOf ctrl Is TextBox Then
         ctrl.Enabled = blnEnabled
      ElseIf TypeOf ctrl Is ListBox Then
         ctrl.Enabled = blnEnabled
      ElseIf TypeOf ctrl Is Label Then
         ctrl.Enabled = blnEnabled
      ElseIf TypeOf ctrl Is CommandButton Then
         ctrl.Enabled = blnEnabled
      End If
   Next
End Sub

Private Sub cmdGenerateCHM_Click()
   On Error GoTo hErr
   
   gblnError = False
   
   If comboOutputDir.Visible Then
      If Not fso.FolderExists(comboOutputDir.Text) Then
         MsgBox "Output directory """ & comboOutputDir.Text & """ does not exists!", vbCritical
         Exit Sub
      End If
      
      AddTextToComboList comboOutputDir
   End If
   
   SetChildrenEnabled False
   SetPictureBoxControlsEnabled False
   lstMsg.Enabled = True
   
   gobjTaskMgr.Prepare
   
   If ActiveForm Is frmListMode Then
      gobjTaskMgr.Begin frmListMode.GetSelectedPaths, comboOutputDir.Text
   ElseIf ActiveForm Is frmNameMode Then
      gobjTaskMgr.Begin frmNameMode.GetSelectedPaths, comboOutputDir.Text
'   ElseIf ActiveForm Is frmMergeMode Then
'      gobjTaskMgr.Begin2 frmMergeMode.GetSelectedPaths
   End If
   
   SetPictureBoxControlsEnabled True
   SetChildrenEnabled True
   
   If gblnError Then
      MsgBox "The task completed with some error!", vbExclamation
   Else
      MsgBox "The task completed successfully!", vbInformation
   End If
   
   Exit Sub
hErr:
   gobjLogMgr.LogErrorNR "@cmdGenerateCHM_Click()"
   gblnError = True
   Resume Next
End Sub

Private Sub cmdOpenDir_Click()
   Shell "explorer /e, " & comboOutputDir.Text, vbNormalFocus
End Sub

Private Sub cmdSelectDir_Click()
   comboOutputDir.Text = ShowFolderSelection(Me.hwnd, "Select output folder:", comboOutputDir.Text)
End Sub

Private Sub lstMsg_DblClick()
   MsgBox lstMsg.List(lstMsg.ListIndex), vbInformation
End Sub

Private Sub MDIForm_Load()
   Set mobjLogMgr = gobjLogMgr
   
'   comboOutputDir.Text = gobjConfig.OutDir
   LoadComboSetting comboOutputDir, "c:\"
   Caption = "tlb-to-chm " & App.major & "." & App.minor & " Rev " & App.Revision
   
   mnuOpenFinishedChmFiles.Checked = gobjConfig.OpenFinishedChmFiles
   mnuTreatHelpStringAsHtml.Checked = gobjConfig.TreatHelpStringAsHtml
   
   mdblDelta = Picture1.Height / Height
   
'   frmMergeMode.Show
   frmMergeMode.Show
   frmListMode.Show
   frmNameMode.Show
End Sub

Private Sub MDIForm_Resize()
   If mdblDelta < Picture1.Height / Height - 0.01 Or mdblDelta > Picture1.Height / Height + 0.01 Then
      Picture1.Height = Height * mdblDelta
   End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
   SaveComboSetting comboOutputDir
'   gobjConfig.OutDir = txtOutputDir.Text
End Sub

Private Sub mnuAbout_Click()
   frmAbout.Show vbModal, Me
End Sub

Private Sub mnuClearOutputDirectoryHistory_Click()
End Sub

Private Sub mnuClearSelectedControlHistory_Click()
   If TypeOf Screen.ActiveControl Is ComboBox Then
      Select Case MsgBox("Are you sure to delete history?", vbOKCancel + vbQuestion)
      Case vbOK
         ClearComboList Screen.ActiveControl
      End Select
   Else
      MsgBox "You must select a combobox!", vbExclamation
   End If
End Sub

Private Sub mnuExit_Click()
   Unload Me
End Sub

Private Sub mnuTreatHelpStringAsHtml_Click()
   With mnuTreatHelpStringAsHtml
      .Checked = Not .Checked
      gobjConfig.TreatHelpStringAsHtml = .Checked
   End With
End Sub

Private Sub mnuOpenFinishedChmFiles_Click()
   With mnuOpenFinishedChmFiles
      .Checked = Not .Checked
      gobjConfig.OpenFinishedChmFiles = .Checked
   End With
End Sub

Private Sub mnuPauseConvertionProcess_Click()
   With mnuPauseConvertionProcess
      .Checked = Not .Checked
   End With
End Sub

Private Sub mobjLogMgr_LogMsg(ByVal msg As String)
   gobjLogMgr.AddMsgToList lstMsg, msg
End Sub

Private Sub mobjLogMgr_Reset()
   lstMsg.Clear
End Sub

Private Sub txtOutputDir_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error GoTo hErr
   If Data.GetFormat(vbCFText) Then
      comboOutputDir.Text = Data.GetData(vbCFText)
   ElseIf Data.GetFormat(vbCFFiles) Then
      comboOutputDir.Text = Data.Files(1)
   End If
   Exit Sub
hErr:
   gobjLogMgr.LogErrorNR
   Resume Next
End Sub


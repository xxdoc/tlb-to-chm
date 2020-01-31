VERSION 5.00
Object = "{03B209C4-7ADD-4264-A128-4BFAD583CACD}#1.0#0"; "PuppyResizer.ocx"
Begin VB.Form frmListMode 
   Caption         =   "Registry List Mode"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7275
   Icon            =   "frmListMode.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3045
   ScaleWidth      =   7275
   Begin PuppyResizerLib.PuppyResizer PuppyResizer1 
      Left            =   240
      Top             =   1920
      _ExtentX        =   529
      _ExtentY        =   529
      Mode            =   0
      FormInfo        =   "7515,3630"
      ItemCount       =   9
      Item1           =   "cmdReload,CommandButton,-1,1080,2580,1095,315,,,,,,0,0"
      Item2           =   "cmdClearLog,CommandButton,-1,5760,2580,1395,315,,,,,,0,0"
      Item3           =   "cmdInvertSelection,CommandButton,-1,3780,2580,1875,315,,,,,,0,0"
      Item4           =   "cmdSelectAll,CommandButton,-1,2280,2580,1395,315,,,,,,0,0"
      Item5           =   "txtFilter,TextBox,-1,900,120,4215,270,,,,,,0,0"
      Item6           =   "lstTlbs,ListBox,-1,120,480,7035,1860,,,,,,0,0"
      Item7           =   "Label4,Label,-1,120,2640,450,180,,,,,,0,0"
      Item8           =   "Label3,Label,-1,5220,180,1875,195,,,,,,0,0"
      Item9           =   "Label1,Label,-1,120,180,675,195,,,,,,0,0"
   End
   Begin VB.CommandButton cmdReload 
      Caption         =   "&Reload"
      Height          =   315
      Left            =   1080
      TabIndex        =   5
      Tag             =   "TL"
      Top             =   2580
      Width           =   1095
   End
   Begin VB.CommandButton cmdClearLog 
      Caption         =   "&Clear Log"
      Height          =   315
      Left            =   5760
      TabIndex        =   8
      Tag             =   "TL"
      Top             =   2580
      Width           =   1395
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   6600
      Top             =   1920
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   6120
      Top             =   1920
   End
   Begin VB.CommandButton cmdInvertSelection 
      Caption         =   "&Invert Selection"
      Height          =   315
      Left            =   3780
      TabIndex        =   7
      Tag             =   "TL"
      Top             =   2580
      Width           =   1875
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "&Select All"
      Height          =   315
      Left            =   2280
      TabIndex        =   6
      Tag             =   "TL"
      Top             =   2580
      Width           =   1395
   End
   Begin VB.TextBox txtFilter 
      Height          =   270
      Left            =   900
      TabIndex        =   1
      Tag             =   "W"
      Top             =   120
      Width           =   4215
   End
   Begin VB.ListBox lstTlbs 
      Height          =   1860
      IntegralHeight  =   0   'False
      Left            =   120
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   3
      Tag             =   "WH"
      Top             =   480
      Width           =   7035
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "[0/0]"
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Tag             =   "T"
      Top             =   2640
      Width           =   450
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "[Regular Expression]"
      BeginProperty Font 
         Name            =   "ËÎÌו"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5220
      TabIndex        =   2
      Tag             =   "L"
      Top             =   180
      Width           =   1875
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "&Filter:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   675
   End
End
Attribute VB_Name = "frmListMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private tlbmgr As New TlbManagement
Private mblnInitialized As Boolean

Private Sub cmbFilterMode_Click()
   UpdateStatus
End Sub

Private Sub cmdApplyFilter_Click()
   UpdateStatus
End Sub

Private Sub cmdInvertSelection_Click()
   Dim temp As Long
   temp = lstTlbs.ListIndex
   
   Dim i As Long
   For i = 0 To lstTlbs.ListCount - 1
      lstTlbs.Selected(i) = Not lstTlbs.Selected(i)
   Next
   
   lstTlbs.ListIndex = temp
End Sub

Private Sub cmdSelectAll_Click()
   Dim temp As Long
   temp = lstTlbs.ListIndex
   
   Dim i As Long
   For i = 0 To lstTlbs.ListCount - 1
      lstTlbs.Selected(i) = True
   Next

   lstTlbs.ListIndex = temp
End Sub

Public Function GetSelectedPaths() As Collection
   Set GetSelectedPaths = New Collection
   
   Dim i As Long
   For i = 0 To lstTlbs.ListCount - 1
      If lstTlbs.Selected(i) Then
         GetSelectedPaths.Add tlbmgr.Path(lstTlbs.ItemData(i))
      End If
   Next
End Function

Private Sub cmdClearLog_Click()
   gobjLogMgr.Reset
End Sub

Private Sub cmdReload_Click()
   Timer1.Enabled = True
End Sub

Private Sub UpdateStatus(Optional ByVal UpdateList As Boolean = True)
   On Error GoTo hErr
   If UpdateList Then
      tlbmgr.FillList lstTlbs, txtFilter
   End If
   Label4.Caption = "[" & lstTlbs.SelCount & "/" & lstTlbs.ListCount & "]"
   Exit Sub
hErr:
   gobjLogMgr.LogErrorNR
   Resume Next
End Sub

Private Sub Form_Activate()
   If Not gblnInitializing And Not mblnInitialized And frmMDI.ActiveForm Is Me Then
      Timer1.Enabled = True
      mblnInitialized = True
   End If
   
   KeepMaximized Me
   frmMDI.FormActivate Me
End Sub

Private Sub Form_Initialize()
   LogInitialize Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If UnloadMode = vbFormControlMenu Then
      Cancel = True
      WindowState = vbMinimized
   End If
End Sub

Private Sub Form_Resize()
   KeepMaximized Me
End Sub

Private Sub Form_Terminate()
   LogTerminate Me
End Sub

Private Sub lstTlbs_ItemCheck(Item As Integer)
   Timer2.Enabled = True
End Sub

Private Sub Timer1_Timer()
   On Error Resume Next
   Timer1.Enabled = False
   tlbmgr.LoadTypeLibraryPaths
   UpdateStatus
End Sub

Private Sub Timer2_Timer()
   Timer2.Enabled = False
   UpdateStatus False
End Sub

Private Sub txtFilter_Change()
   UpdateStatus
End Sub


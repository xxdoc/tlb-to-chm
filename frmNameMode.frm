VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{03B209C4-7ADD-4264-A128-4BFAD583CACD}#1.0#0"; "puppyresizer.ocx"
Begin VB.Form frmNameMode 
   Caption         =   "ClassName & FilePath Mode"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   5925
   Icon            =   "frmNameMode.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2595
   ScaleWidth      =   5925
   Begin PuppyResizerLib.PuppyResizer PuppyResizer1 
      Left            =   900
      Top             =   1920
      _ExtentX        =   529
      _ExtentY        =   529
      Mode            =   0
      FormInfo        =   "6045,3000"
      ItemCount       =   6
      Item1           =   "cmdRemove,CommandButton,-1,4620,540,1155,315,,,,,,0,0"
      Item2           =   "optFilePath,OptionButton,-1,120,600,2955,240,,,,,,0,0"
      Item3           =   "optClassName,OptionButton,-1,120,180,1275,195,,,,,,0,0"
      Item4           =   "txtClassName,TextBox,-1,1500,120,4275,315,,,,,,0,0"
      Item5           =   "cmdAdd,CommandButton,-1,3360,540,1155,315,,,,,,0,0"
      Item6           =   "lstPaths,ListBox,-1,120,960,5655,1500,,,,,,0,0"
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   315
      Left            =   4620
      TabIndex        =   4
      Tag             =   "L"
      Top             =   540
      Width           =   1155
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   300
      Top             =   1860
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      MaxFileSize     =   32767
   End
   Begin VB.OptionButton optFilePath 
      Caption         =   "&Select File(s):"
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2955
   End
   Begin VB.OptionButton optClassName 
      Caption         =   "&ClassName:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Value           =   -1  'True
      Width           =   1275
   End
   Begin VB.TextBox txtClassName 
      Height          =   315
      Left            =   1500
      TabIndex        =   1
      Tag             =   "W"
      Text            =   "Scripting.Dictionary"
      Top             =   120
      Width           =   4275
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add ..."
      Default         =   -1  'True
      Height          =   315
      Left            =   3360
      TabIndex        =   3
      Tag             =   "L"
      Top             =   540
      Width           =   1155
   End
   Begin VB.ListBox lstPaths 
      Height          =   1500
      IntegralHeight  =   0   'False
      Left            =   120
      MultiSelect     =   2  'Extended
      OLEDropMode     =   1  'Manual
      TabIndex        =   5
      Tag             =   "WH"
      ToolTipText     =   "You can Drag and Drop files from Windows Explorer."
      Top             =   960
      Width           =   5655
   End
End
Attribute VB_Name = "frmNameMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function GetSelectedPaths() As Collection
   Set GetSelectedPaths = New Collection
   
   If optClassName.value Then
      Dim obj As Object
      Set obj = CreateObject(txtClassName)
      
      GetSelectedPaths.Add ClassInfoFromObject(obj).parent.ContainingFile
   ElseIf optFilePath.value Then
      Dim i As Long
      For i = 0 To lstPaths.ListCount - 1
         GetSelectedPaths.Add lstPaths.List(i)
      Next
   Else
      Debug.Assert False
   End If
End Function

Private Sub cmdAdd_Click()
   On Error GoTo hErr
   With CommonDialog1
      .Flags = cdlOFNAllowMultiselect Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNExplorer
      .Filter = "All supported files|*.dll;*.tlb;*.olb;*.exe;*.ocx|*.*|*.*"
      .FileName = ""
      .ShowOpen
      
      Dim arr() As String
      arr = Split(.FileName, Chr(0))
      
      If Array_GetCount(arr) = 1 Then
         lstPaths.AddItem arr(0)
      Else
         Dim strDir As String
         strDir = arr(0)
         
         Dim i As Long
         For i = 1 To UBound(arr)
            lstPaths.AddItem fso.BuildPath(arr(0), arr(i))
         Next
      End If
   End With
   UpdateStatus
   Exit Sub
hErr:
   gobjLogMgr.LogErrorNR
   Resume Next
End Sub

Private Sub cmdRemove_Click()
   Dim i As Long
   For i = lstPaths.ListCount - 1 To 0 Step -1
      If lstPaths.Selected(i) Then
         lstPaths.RemoveItem i
      End If
   Next
   UpdateStatus
End Sub

Private Sub Form_Activate()
   KeepMaximized Me
   frmMDI.FormActivate Me
End Sub

Private Sub Form_Initialize()
   LogInitialize Me
End Sub

Private Sub Form_Load()
   UpdateStatus
End Sub

Private Sub Form_Resize()
   KeepMaximized Me
End Sub

Private Sub Form_Terminate()
   LogTerminate Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If UnloadMode = vbFormControlMenu Then
      Cancel = True
      WindowState = vbMinimized
   End If
End Sub

Private Sub UpdateStatus()
   If optClassName.value Then
      lstPaths.BackColor = &H8000000F
      lstPaths.Enabled = False
      txtClassName.BackColor = &H80000005
      txtClassName.Enabled = True
      cmdAdd.Enabled = False
      cmdRemove.Enabled = False
   ElseIf optFilePath.value Then
      lstPaths.BackColor = &H80000005
      lstPaths.Enabled = True
      txtClassName.BackColor = &H8000000F
      txtClassName.Enabled = False
      cmdAdd.Enabled = True
      cmdRemove.Enabled = True
   Else
      Debug.Assert False
   End If
   
   optFilePath.Caption = "&Select File(s): " & lstPaths.ListCount
End Sub

Private Sub optClassName_Click()
   UpdateStatus
End Sub

Private Sub optFilePath_Click()
   UpdateStatus
End Sub

Private Sub lstPaths_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error GoTo hErr
   If Data.GetFormat(vbCFText) Then
      Dim arr() As String
      arr = Split(Data.GetData(vbCFText), vbCrLf)
      
      Dim i As Long
      For i = LBound(arr) To UBound(arr)
         lstPaths.AddItem arr(i)
      Next
   ElseIf Data.GetFormat(vbCFFiles) Then
      Dim varPath As Variant
      For Each varPath In Data.Files
         lstPaths.AddItem varPath
      Next
   End If
   UpdateStatus
   Exit Sub
hErr:
   gobjLogMgr.LogErrorNR
   Resume Next
End Sub


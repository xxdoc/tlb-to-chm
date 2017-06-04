VERSION 5.00
Object = "{03B209C4-7ADD-4264-A128-4BFAD583CACD}#1.0#0"; "PuppyResizer.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMergeMode 
   Caption         =   "Merge Chm Mode"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9075
   Icon            =   "frmMergeMode.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4350
   ScaleWidth      =   9075
   Begin PuppyResizerLib.PuppyResizer PuppyResizer1 
      Left            =   1200
      Top             =   2880
      _ExtentX        =   529
      _ExtentY        =   529
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox comboWindowTitle 
      Height          =   300
      Left            =   6120
      TabIndex        =   9
      Tag             =   "T"
      Top             =   3900
      Width           =   2835
   End
   Begin VB.ComboBox comboOutputFileName 
      Height          =   300
      Left            =   1860
      TabIndex        =   8
      Tag             =   "T"
      Top             =   3900
      Width           =   2835
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add ..."
      Default         =   -1  'True
      Height          =   315
      Left            =   3960
      TabIndex        =   2
      Tag             =   "L"
      Top             =   180
      Width           =   1155
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   315
      Left            =   5220
      TabIndex        =   6
      Tag             =   "L"
      Top             =   180
      Width           =   1155
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "&Up"
      Height          =   315
      Left            =   7740
      TabIndex        =   3
      Tag             =   "L"
      Top             =   180
      Width           =   1155
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "&Down"
      Height          =   315
      Left            =   6480
      TabIndex        =   7
      Tag             =   "L"
      Top             =   180
      Width           =   1155
   End
   Begin VB.ListBox lstFiles 
      Height          =   2985
      ItemData        =   "frmMergeMode.frx":000C
      Left            =   180
      List            =   "frmMergeMode.frx":000E
      MultiSelect     =   2  'Extended
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Tag             =   "WH"
      Top             =   600
      Width           =   8715
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Output File &Name:"
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Tag             =   "T"
      Top             =   3960
      Width           =   1530
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "&Window Title:"
      Height          =   180
      Left            =   4860
      TabIndex        =   5
      Tag             =   "T"
      Top             =   3960
      Width           =   1170
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      Caption         =   "File Count: 0"
      Height          =   180
      Left            =   180
      TabIndex        =   1
      Top             =   240
      Width           =   1170
   End
End
Attribute VB_Name = "frmMergeMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private fso As New FileSystemObject
Private WithEvents ce As CompileEngine
Attribute ce.VB_VarHelpID = -1

Private Sub ce_Log(ByVal msg As String)
    LogMsg msg
End Sub

Private Sub ce_Proc(ByVal msg As String)
    DoEvents
End Sub

Private Sub cmdAdd_Click()
    With CommonDialog1
        .FileName = ""
        .ShowOpen
        
        If .FileName <> "" Then
            Dim arr() As String
            arr = Split(.FileName, Chr(0))
            
            If UBound(arr) = 0 Then
                lstFiles.AddItem arr(0)
            Else
                Dim i As Integer
                For i = 1 To UBound(arr)
                    lstFiles.AddItem fso.BuildPath(arr(0), arr(i))
                Next
            End If
        End If
    End With
    UpdateFileCount
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdRemove_Click()
    Dim i As Integer
    For i = lstFiles.ListCount - 1 To 0 Step -1
        If lstFiles.Selected(i) Then
            lstFiles.RemoveItem i
        End If
    Next
    UpdateFileCount
End Sub

Public Function GetSelectedPaths() As Collection
    Set GetSelectedPaths = New Collection
    
    With GetSelectedPaths
        .Add comboOutputFileName.Text
        .Add comboWindowTitle.Text
        
        Dim i As Long
        For i = 0 To lstFiles.ListCount - 1
            .Add lstFiles.List(i)
        Next
    End With
    
    AddTextToComboList comboOutputFileName
    AddTextToComboList comboWindowTitle
End Function

Private Sub cmdUp_Click()
    If lstFiles.ListCount = 0 Then
        Exit Sub
    End If
    
    If lstFiles.Selected(0) Then
        Exit Sub
    End If

    Dim i As Integer
    For i = 1 To lstFiles.ListCount - 1
        If lstFiles.Selected(i) Then
            Dim temp As String
            Dim temp2 As Boolean
            temp = lstFiles.List(i - 1)
            temp2 = lstFiles.Selected(i - 1)
            lstFiles.List(i - 1) = lstFiles.List(i)
            lstFiles.Selected(i - 1) = lstFiles.Selected(i)
            lstFiles.List(i) = temp
            lstFiles.Selected(i) = temp2
        End If
    Next
End Sub

Private Sub cmdDown_Click()
    If lstFiles.ListCount = 0 Then
        Exit Sub
    End If
    
    If lstFiles.Selected(lstFiles.ListCount - 1) Then
        Exit Sub
    End If

    Dim i As Integer
    For i = lstFiles.ListCount - 2 To 0 Step -1
        If lstFiles.Selected(i) Then
            Dim temp As String
            Dim temp2 As Boolean
            temp = lstFiles.List(i + 1)
            temp2 = lstFiles.Selected(i + 1)
            lstFiles.List(i + 1) = lstFiles.List(i)
            lstFiles.Selected(i + 1) = lstFiles.Selected(i)
            lstFiles.List(i) = temp
            lstFiles.Selected(i) = temp2
        End If
    Next
End Sub

Public Sub WriteFile(ByVal filepath As String, ByVal content As String)
    Dim fp As Integer
    fp = FreeFile()
    
    Open filepath For Output As fp
    Print #fp, content
    Close fp
End Sub

Public Function ReadFile(ByVal filepath As String) As String
    Dim fp As Integer
    fp = FreeFile()
    
    Open filepath For Input As fp
    
    While Not EOF(fp)
        Dim line As String
        Line Input #fp, line
        
        ReadFile = ReadFile & line & vbCrLf
    Wend
    
    Close fp
End Function

Private Sub LogMsg(ByVal msg As String)
    gobjLogMgr.LogMsg msg
End Sub

Private Sub Form_Activate()
    KeepMaximized Me
    frmMDI.FormActivate Me
End Sub

Private Sub Form_Load()
    Set ce = New CompileEngine
    UpdateFileCount
    
    LoadComboSetting comboOutputFileName, "help.chm"
    LoadComboSetting comboWindowTitle, "Help"
'    Me.Caption = "MergeChm " & App.major & "." & App.minor & " Rev " & App.Revision
End Sub

Private Sub UpdateFileCount()
    lblCount.Caption = "File Count: " & lstFiles.ListCount
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

Private Sub Form_Unload(Cancel As Integer)
    SaveComboSetting comboOutputFileName
    SaveComboSetting comboWindowTitle
End Sub

Private Sub lstFiles_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyDelete
        cmdRemove.Value = True
    End Select
End Sub

Private Sub lstFiles_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo hErr
    If Data.GetFormat(vbCFText) Then
        Dim arr() As String
        arr = Split(Data.GetData(vbCFText), vbCrLf)
        
        Dim i As Long
        For i = LBound(arr) To UBound(arr)
            lstFiles.AddItem arr(i)
        Next
    ElseIf Data.GetFormat(vbCFFiles) Then
        Dim varPath As Variant
        For Each varPath In Data.Files
            lstFiles.AddItem varPath
        Next
    End If
    Exit Sub
hErr:
    gobjLogMgr.LogErrorNR
    Resume Next
End Sub

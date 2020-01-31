VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{03B209C4-7ADD-4264-A128-4BFAD583CACD}#1.0#0"; "PuppyResizer.ocx"
Begin VB.Form frmMergeMode 
   Caption         =   "Merge Chm Mode"
   ClientHeight    =   7740
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7740
   ScaleWidth      =   9420
   Begin PuppyResizerLib.PuppyResizer PuppyResizer1 
      Left            =   8520
      Top             =   3120
      _ExtentX        =   529
      _ExtentY        =   529
      Mode            =   0
      FormInfo        =   "9660,8325"
      ItemCount       =   15
      Item1           =   "txtLog,TextBox,-1,180,6060,9075,1455,,,,,,0,0"
      Item2           =   "Frame2,Frame,-1,180,4560,9075,795,,,,,,0,0"
      Item3           =   "txtWindowTitle,TextBox,-1,6120,300,2415,315,,,,,,0,0"
      Item4           =   "txtOutputFileName,TextBox,-1,1980,300,2415,315,,,,,,0,0"
      Item5           =   "Label3,Label,-1,4860,360,1170,180,,,,,,0,0"
      Item6           =   "Label2,Label,-1,360,360,1530,180,,,,,,0,0"
      Item7           =   "cmdStartMerge,CommandButton,-1,7560,5520,1635,375,,,,,,0,0"
      Item8           =   "Frame1,Frame,-1,180,180,9075,4275,,,,,,0,0"
      Item9           =   "cmdDown,CommandButton,-1,7920,3720,915,375,,,,,,0,0"
      Item10          =   "cmdUp,CommandButton,-1,6900,3720,915,375,,,,,,0,0"
      Item11          =   "cmdRemove,CommandButton,-1,5880,3720,915,375,,,,,,0,0"
      Item12          =   "cmdAdd,CommandButton,-1,4860,3720,915,375,,,,,,0,0"
      Item13          =   "lstFiles,ListBox,-1,240,300,8595,3120,,,,,,0,0"
      Item14          =   "lblCount,Label,-1,300,3840,1170,180,,,,,,0,0"
      Item15          =   "Label1,Label,-1,180,5640,5040,180,,,,,,0,0"
   End
   Begin VB.TextBox txtLog 
      BackColor       =   &H8000000F&
      Height          =   1455
      Left            =   180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Tag             =   "TW"
      Top             =   6060
      Width           =   9075
   End
   Begin VB.Frame Frame2 
      Caption         =   "&Options"
      Height          =   795
      Left            =   180
      TabIndex        =   7
      Tag             =   "TW"
      Top             =   4560
      Width           =   9075
      Begin VB.TextBox txtWindowTitle 
         Height          =   315
         Left            =   6120
         TabIndex        =   11
         Text            =   "Help"
         Top             =   300
         Width           =   2415
      End
      Begin VB.TextBox txtOutputFileName 
         Height          =   315
         Left            =   1980
         TabIndex        =   9
         Text            =   "1.chm"
         Top             =   300
         Width           =   2415
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Window Title:"
         Height          =   180
         Left            =   4860
         TabIndex        =   10
         Top             =   360
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Output File &Name:"
         Height          =   180
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   1530
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7920
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.chm|*.chm|*.*|*.*"
      Flags           =   524800
      MaxFileSize     =   32767
   End
   Begin VB.CommandButton cmdStartMerge 
      Caption         =   "&Start merge"
      Height          =   375
      Left            =   7560
      TabIndex        =   13
      Tag             =   "TL"
      Top             =   5520
      Width           =   1635
   End
   Begin VB.Frame Frame1 
      Caption         =   "&File List"
      Height          =   4275
      Left            =   180
      TabIndex        =   0
      Tag             =   "WH"
      Top             =   180
      Width           =   9075
      Begin VB.CommandButton cmdDown 
         Caption         =   "&Down"
         Height          =   375
         Left            =   7920
         TabIndex        =   6
         Tag             =   "TL"
         Top             =   3720
         Width           =   915
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "&Up"
         Height          =   375
         Left            =   6900
         TabIndex        =   5
         Tag             =   "TL"
         Top             =   3720
         Width           =   915
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   375
         Left            =   5880
         TabIndex        =   4
         Tag             =   "TL"
         Top             =   3720
         Width           =   915
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   4860
         TabIndex        =   3
         Tag             =   "TL"
         Top             =   3720
         Width           =   915
      End
      Begin VB.ListBox lstFiles 
         Height          =   3120
         ItemData        =   "frmMergeMode.frx":0000
         Left            =   240
         List            =   "frmMergeMode.frx":0002
         MultiSelect     =   2  'Extended
         TabIndex        =   1
         Tag             =   "WH"
         Top             =   300
         Width           =   8595
      End
      Begin VB.Label lblCount 
         AutoSize        =   -1  'True
         Caption         =   "File Count: 0"
         Height          =   180
         Left            =   300
         TabIndex        =   2
         Tag             =   "T"
         Top             =   3840
         Width           =   1170
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Note: all of the chm files should be in the same folder."
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   180
      TabIndex        =   12
      Tag             =   "T"
      Top             =   5640
      Width           =   5040
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

Public Function ReadHHCName(ByVal strPath As String)
   Dim fr As Integer
   fr = FreeFile()
   
   Open strPath For Binary As fr
   Seek #fr, &HE6
   
   While Not EOF(fr)
      Dim c As Byte
      Get #fr, , c
      
      If Chr(c) = "." Then
         Get #fr, , c
         If LCase(Chr(c)) = "h" Then
            Get #fr, , c
            If LCase(Chr(c)) = "h" Then
               Get #fr, , c
               If LCase(Chr(c)) = "c" Then
                  Dim be As Byte
                  Get #fr, , be
                  
                  If be = &H1 Then
                     Seek #fr, Seek(fr) - 1
                     
                     ' 0 <-> 1
                     Dim pos As Long
                     pos = Seek(fr)
                     
                     Seek #fr, Seek(fr) - 5
                     Get #fr, , c
                     
                     While Chr(c) <> "/" And Seek(fr) - 1 >= &HE6
                        Seek #fr, Seek(fr) - 2
                        Get #fr, , c
                     Wend
                     
                     Seek #fr, Seek(fr) - 2
                     Get #fr, , be
                     
                     If be <= &H20 Then
                        Seek #fr, Seek(fr) + 1
                        
                        Dim length As Long
                        length = pos - Seek(fr) - 1
                        
                        Dim hhc() As Byte
                        ReDim hhc(length)
                        
                        Get #fr, , hhc
                        ReadHHCName = Replace(StrConv(hhc, vbUnicode), Chr(1), "")
                        
                        Close fr
                        Exit Function
                     Else
                        Seek #fr, pos
                     End If
                  Else
                     Seek #fr, Seek(fr) - 4
                  End If
               Else
                  Seek #fr, Seek(fr) - 3
               End If
            Else
               Seek #fr, Seek(fr) - 2
            End If
         Else
            Seek #fr, Seek(fr) - 1
         End If
      End If
   Wend
   
   Close fr
End Function

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

Private Sub cmdRemove_Click()
   Dim i As Integer
   For i = lstFiles.ListCount - 1 To 0 Step -1
      If lstFiles.Selected(i) Then
         lstFiles.RemoveItem i
      End If
   Next
   UpdateFileCount
End Sub

Private Sub cmdStartMerge_Click()
   txtLog.Text = ""

   Dim outputdir As String
   outputdir = fso.GetParentFolderName(lstFiles.List(0))
   
   Dim chmpath As String
   chmpath = fso.BuildPath(outputdir, txtOutputFileName)
   
   If fso.FileExists(chmpath) Then
      Select Case MsgBox("Target chm file """ & chmpath & """ already exists, do you want to overwrite it?", vbOKCancel)
      Case vbCancel
         Exit Sub
      End Select
   End If

   Dim tempdir As String
   tempdir = fso.BuildPath(App.Path, "temp")
   
   If Not fso.FolderExists(tempdir) Then
      fso.CreateFolder tempdir
   End If
   
   fso.CopyFile fso.BuildPath(App.Path, "default.htm"), fso.BuildPath(tempdir, "default.htm"), True
   fso.CopyFile fso.BuildPath(App.Path, "merge.hhk"), fso.BuildPath(tempdir, "merge.hhk"), True
   
   Dim content As String
   content = ReadFile(fso.BuildPath(App.Path, "merge.hhp"))
   
   content = Replace(content, "<Title>", txtWindowTitle)
   content = Replace(content, "<Title bar text>", txtWindowTitle)
   content = Replace(content, "<Compiled file>", chmpath)
   
   content = content & "[MERGE FILES]" & vbCrLf
   
   Dim i As Integer
   For i = 0 To lstFiles.ListCount - 1
      content = content & fso.GetFileName(lstFiles.List(i)) & vbCrLf
   Next
   
   WriteFile fso.BuildPath(tempdir, "merge.hhp"), content
   
   content = _
         "<HTML>" & vbCrLf & _
         "<BODY>" & vbCrLf
         
   For i = 0 To lstFiles.ListCount - 1
      Dim hhc As String
      hhc = ReadHHCName(lstFiles.List(i))
      
      If hhc <> "" Then
         content = content & _
               "<OBJECT type=""text/sitemap"">" & vbCrLf & _
               "<param name=""Merge"" value=""" & _
               fso.GetFileName(lstFiles.List(i)) & "::\" & _
               hhc & """></OBJECT>" & vbCrLf
      Else
         LogMsg "Warning: file """ & lstFiles.List(i) & """ does not have hhc." & vbCrLf
      End If
      
      DoEvents
   Next
         
   content = content & _
         "</BODY>" & vbCrLf & _
         "</HTML>" & vbCrLf

   WriteFile fso.BuildPath(tempdir, "merge.hhc"), content
   
   LogMsg vbCrLf
   
   If ce.CompileHHP(fso.BuildPath(tempdir, "merge.hhp")) Then
      Select Case MsgBox("Completed! Click OK to open the file.", vbInformation + vbOKCancel)
      Case vbOK
         ShellExecute 0&, vbNullString, chmpath, vbNullString, vbNullString, vbNormalFocus
      End Select
   Else
      MsgBox "Failed!", vbCritical
   End If
End Sub

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
   txtLog.SelStart = Len(txtLog.Text)
   txtLog.SelText = msg
End Sub

Private Sub Form_Activate()
   KeepMaximized Me
   frmMDI.FormActivate Me
End Sub

Private Sub Form_Load()
   Set ce = New CompileEngine
   UpdateFileCount
End Sub

Private Sub UpdateFileCount()
   lblCount.Caption = "File Count: " & lstFiles.ListCount
End Sub

Private Sub Form_Resize()
   KeepMaximized Me
End Sub

Private Sub lstFiles_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case vbKeyDelete
      cmdRemove.value = True
   End Select
End Sub

Attribute VB_Name = "modCommon"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type BrowseInfo
   hWndOwner As Long
   pIDLRoot As Long
   pszDisplayName As Long
   lpszTitle As Long
   ulFlags As Long
   lpfnCallback As Long
   lParam As Long
   iImage As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_EDITBOX = &H10
Private Const BIF_NEWFOLDERBUTTON = &H40
Private Const MAX_PATH = 260

Public Const errUserCancel As Long = vbObjectError + 512

Public fso As New FileSystemObject
Public gobjTaskMgr As TaskManagement
Public gobjConfig As Config
Public gobjLogMgr As LogManagement
Public gblnInitializing As Boolean
Public gblnError As Boolean

Public Function Array_GetCount(ByVal arr As Variant) As Long
   Array_GetCount = UBound(arr) - LBound(arr) + 1
End Function

' 打开 Windows 的选择目录对话框
' hwnd 为窗口句柄(通常设为 Me.hwnd), Prompt 为指示字符串
Public Function ShowFolderSelection(ByVal hwnd As Long, ByVal strPrompt As String, ByVal strDefault As String) As String
   Dim iNull As Integer
   Dim lpIDList As Long
   Dim lResult As Long
   Dim sPath As String
   Dim udtBI As BrowseInfo
   With udtBI
      .hWndOwner = hwnd
      .lpszTitle = lstrcat(strPrompt, "")
      .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_NEWFOLDERBUTTON Or BIF_EDITBOX
   End With
   lpIDList = SHBrowseForFolder(udtBI)
   If lpIDList Then
      sPath = String$(MAX_PATH, 0)
      lResult = SHGetPathFromIDList(lpIDList, sPath)
      CoTaskMemFree lpIDList
      iNull = InStr(sPath, vbNullChar)
      If iNull Then sPath = Left$(sPath, iNull - 1)
   End If
   ShowFolderSelection = IIf(sPath <> "", sPath, strDefault)
End Function

Public Sub Main()
   Set gobjTaskMgr = New TaskManagement
   Set gobjConfig = New Config
   Set gobjLogMgr = New LogManagement
   
   gblnInitializing = True
   frmMDI.Show
   gblnInitializing = False
End Sub

Public Sub WriteFile(ByVal filepath As String, ByVal content As String, ByVal appendmode As Boolean)
   Dim fp As Integer
   fp = FreeFile()
   
   If appendmode Then
      Open filepath For Append As fp
   Else
      Open filepath For Output As fp
   End If
   
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

Public Sub LogInitialize(ByVal obj As Object)
   Debug.Print TypeName(obj) & " Initialize."
End Sub

Public Sub LogTerminate(ByVal obj As Object)
   Debug.Print TypeName(obj) & " Terminate."
End Sub

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

' 参数可以是数组或者任何支持枚举的集合类.
Public Function ConvertToPathString(ByVal varPaths As Variant, ByVal blnFirstIsDir As Boolean) As String
   Dim strBuffer As String
   strBuffer = ""
   
   If IsArray(varPaths) Or IsObject(varPaths) Then
      Dim blnFirst As Boolean
      blnFirst = True
      
      Dim varPath As Variant
      For Each varPath In varPaths
         If blnFirst Then
            If blnFirstIsDir Then
               strBuffer = strBuffer & """" & varPath & """ "
            Else
               strBuffer = strBuffer & """" & fso.GetParentFolderName(varPath) & """ "
            End If
         End If
         
         If Not (blnFirst And blnFirstIsDir) Then
            strBuffer = strBuffer & """" & fso.GetFileName(varPath) & """ "
         End If
         
         blnFirst = False
      Next
   Else
      Debug.Assert False
   End If
   
   ConvertToPathString = Trim(strBuffer)
End Function

Public Function ConvertFromPathString(ByVal strPaths As String) As Collection
   Set ConvertFromPathString = New Collection
   
   Dim arr() As String
   arr = Split(strPaths, """")
   
   Dim col As Collection
   Set col = New Collection
   
   Dim i As Long
   For i = LBound(arr) To UBound(arr)
      Dim strTemp As String
      strTemp = Trim(arr(i))
      
      If Len(strTemp) > 0 Then
         col.Add strTemp
      End If
   Next
   
   If col.count = 1 Then
      ConvertFromPathString.Add col(1)
   Else
      For i = 2 To col.count
         ConvertFromPathString.Add fso.BuildPath(col(1), col(i))
      Next
   End If
End Function

Public Sub KeepMaximized(ByVal frm As Form)
   If frm.WindowState = vbNormal And frmMDI.ActiveForm Is frm Then
      frm.WindowState = vbMaximized
   End If
End Sub

Public Sub ClearComboList(ByVal combo As ComboBox)
   Dim sTemp As String
   sTemp = combo.Text
   
   combo.Clear
   
   combo.Text = sTemp
End Sub

Public Function AddTextToComboList(ByVal combo As ComboBox) As Boolean
   If combo.Text = "" Then
      Exit Function
   End If

   Dim bFound As Boolean
   bFound = False
   
   Dim i As Long
   For i = 1 To combo.ListCount
      If LCase(combo.List(i - 1)) = LCase(combo.Text) Then
         bFound = True
         Exit For
      End If
   Next
   
   If Not bFound Then
      combo.AddItem combo.Text, 0
      AddTextToComboList = True
   Else
      AddTextToComboList = False
   End If
End Function

Public Sub SaveComboSetting(ByVal combo As ComboBox, Optional ByVal bTextOnly As Boolean = False)
   With combo
      SaveSetting App.EXEName, "main", .name & ".Text", .Text
      If Not bTextOnly Then
         SaveSetting App.EXEName, "main", .name & ".ListCount", .ListCount
         
         Dim i As Long
         For i = 1 To .ListCount
            SaveSetting App.EXEName, "main", .name & ".ListItem" & i, .List(i - 1)
         Next
      End If
   End With
End Sub

Public Sub LoadComboSetting( _
   ByVal combo As ComboBox, _
   ByVal default As String, _
   Optional ByVal bTextOnly As Boolean = False)
   
   With combo
      If Not bTextOnly Then
         Dim nCount As Long
         nCount = GetSetting(App.EXEName, "main", .name & ".ListCount", "0")
         
         .Clear
         
         Dim i As Long
         For i = 1 To nCount
            .AddItem GetSetting(App.EXEName, "main", .name & ".ListItem" & i)
         Next
      End If
   
      .Text = GetSetting(App.EXEName, "main", .name & ".Text", default)
   End With
End Sub

Public Function IsFormLoaded(ByVal frm As Form) As Boolean
   IsFormLoaded = False
   
   Dim f As Form
   For Each f In Forms
      If f Is frm Then
         IsFormLoaded = True
         Exit Function
      End If
   Next
End Function

Public Sub SaveCheckBoxSetting(ByRef chk As CheckBox)
   SaveSetting App.EXEName, "main", chk.name, chk.value
End Sub

Public Sub LoadCheckBoxSetting(ByRef chk As CheckBox, ByVal default As CheckBoxConstants)
   chk.value = GetSetting(App.EXEName, "main", chk.name, default)
End Sub

Public Sub SaveListSetting(ByRef lst As ListBox)
   With lst
      SaveSetting App.EXEName, "main", .name & ".ListCount", .ListCount
      
      Dim i As Long
      For i = 1 To .ListCount
         SaveSetting App.EXEName, "main", .name & ".ListItem" & i, .List(i - 1)
      Next
      If lst.Style = vbListBoxCheckbox Then
         For i = 1 To .ListCount
            SaveSetting App.EXEName, "main", .name & ".Selected" & i, .Selected(i - 1)
         Next
      End If
   End With
End Sub

Public Sub LoadListSetting(ByRef lst As ListBox)
   With lst
      Dim nCount As Long
      nCount = GetSetting(App.EXEName, "main", .name & ".ListCount", "0")
      
      .Clear
      
      Dim i As Long
      For i = 1 To nCount
         .AddItem GetSetting(App.EXEName, "main", .name & ".ListItem" & i)
         .Selected(.NewIndex) = GetSetting(App.EXEName, "main", .name & ".Selected" & i)
      Next
      If lst.Style = vbListBoxCheckbox Then
         For i = 1 To nCount
            .Selected(i) = GetSetting(App.EXEName, "main", .name & ".Selected" & i)
         Next
      End If
   End With
End Sub

Public Function AddStrToList(ByRef lst As ListBox, ByVal newstr As String) As Boolean
   AddStrToList = False
   Dim i As Long
   For i = 1 To lst.ListCount
      If LCase(lst.List(i - 1)) = LCase(newstr) Then
         Exit Function
      End If
   Next
   
   lst.AddItem newstr
   AddStrToList = True
End Function



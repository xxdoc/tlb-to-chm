VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TaskManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public mcolPaths As Collection
Public strCompiledFile As String
Private lngCurrentIndex As Long
Private WithEvents ce As CompileEngine
Attribute ce.VB_VarHelpID = -1
Private cvt As New Tli2XmlConverter
Private mstrCurrentName As String

Public Property Get Cancel() As Boolean
   Cancel = frmProcWnd.mblnCancel
End Property

Private Property Get Status() As String
   Status = "Processing (" & lngCurrentIndex & " of " & mcolPaths.count & ") ... [" & mstrCurrentName & "]"
End Property

Private Sub UpdateStatus()
   frmProcWnd.ShowProcess Status
   gobjLogMgr.LogMsg Status
   DoEvents
End Sub

Private Sub ce_Log(ByVal msg As String)
   DoEvents
   gobjLogMgr.LogMsg msg
   If Cancel Then
      ce.Cancel = True
   End If
End Sub

Private Sub ce_Proc(ByVal msg As String)
   frmProcWnd.ShowProcess2 "Compiling ... " & msg
   DoEvents
   If Cancel Then
      ce.Cancel = True
   End If
End Sub

Public Sub Begin(ByVal colPaths As Collection, ByVal strOutDir As String)
   On Error GoTo hErr
   
   Set mcolPaths = colPaths
   
   lngCurrentIndex = 1
   
   frmProcWnd.Prepare True
   
   Dim varPath As Variant
   For Each varPath In colPaths
      Dim tlb As TypeLibInfo
      Set tlb = TypeLibInfoFromFile(varPath)
      
      mstrCurrentName = tlb.name
      UpdateStatus
      
      Dim xml As DOMDocument
      Set xml = cvt.StartConvert(tlb)
      ConvertXml2Hhp xml, strOutDir
      
      DoEvents
      
      If ce.CompileHHP(fso.BuildPath(TempDir, "output.hhp")) Then
         If gobjConfig.OpenFinishedChmFiles Then
            gobjLogMgr.LogMsg "Opening: """ & strCompiledFile & """ ..."
            ShellExecute 0&, vbNullString, strCompiledFile, vbNullString, vbNullString, vbMinimizedNoFocus
         End If
      End If
      
      lngCurrentIndex = lngCurrentIndex + 1
   Next
   
   frmProcWnd.Finish
   Exit Sub
hErr:
   gobjLogMgr.LogErrorR
   gblnError = True
   Resume Next
End Sub

Public Sub Prepare()
   If Not fso.FolderExists(TempDir) Then
      fso.CreateFolder TempDir
   End If
End Sub

Private Property Get TempDir() As String
   TempDir = fso.BuildPath(App.Path, "temp")
End Property

Public Sub Begin2(ByVal colPaths As Collection)
   On Error GoTo hErr
   
   Dim strOutputFile As String
   strOutputFile = colPaths(1)
   colPaths.Remove 1
   
   Dim strTitle As String
   strTitle = colPaths(1)
   colPaths.Remove 1
   
   Dim outputdir As String
   outputdir = fso.GetParentFolderName(colPaths(1))
   
   Dim chmpath As String
   chmpath = fso.BuildPath(outputdir, strOutputFile)
   
   If fso.FileExists(chmpath) Then
      Select Case MsgBox("Target chm file """ & chmpath & """ already exists, do you want to overwrite it?", vbOKCancel)
      Case vbCancel
         Exit Sub
      End Select
   End If

   frmProcWnd.ShowProcess "Merging """ & strOutputFile & """ ..."
   
   fso.CopyFile fso.BuildPath(App.Path, "default.htm"), fso.BuildPath(TempDir, "default.htm"), True
   fso.CopyFile fso.BuildPath(App.Path, "merge.hhk"), fso.BuildPath(TempDir, "merge.hhk"), True
   
   Dim content As String
   content = ReadFile(fso.BuildPath(App.Path, "merge.hhp"))
   
   content = Replace(content, "<Title>", strTitle)
   content = Replace(content, "<Title bar text>", strTitle)
   content = Replace(content, "<Compiled file>", chmpath)
   
   content = content & "[MERGE FILES]" & vbCrLf
   
   Dim i As Integer
   For i = 1 To colPaths.count
      content = content & fso.GetFileName(colPaths(i)) & vbCrLf
   Next
   
   WriteFile fso.BuildPath(TempDir, "merge.hhp"), content, False
   
   content = _
         "<HTML>" & vbCrLf & _
         "<BODY>" & vbCrLf
         
   For i = 1 To colPaths.count
      Dim hhc As String
      hhc = ReadHHCName(colPaths(i))
      
      If hhc <> "" Then
         content = content & _
               "<OBJECT type=""text/sitemap"">" & vbCrLf & _
               "<param name=""Merge"" value=""" & _
               fso.GetFileName(colPaths(i)) & "::\" & _
               hhc & """></OBJECT>" & vbCrLf
      Else
         gobjLogMgr.LogMsg "Warning: file """ & colPaths(i) & """ does not have hhc." & vbCrLf
      End If
      
      DoEvents
   Next
         
   content = content & _
         "</BODY>" & vbCrLf & _
         "</HTML>" & vbCrLf

   WriteFile fso.BuildPath(TempDir, "merge.hhc"), content, False
   
   gobjLogMgr.LogMsg vbCrLf
   
   If ce.CompileHHP(fso.BuildPath(TempDir, "merge.hhp")) Then
      If gobjConfig.OpenFinishedChmFiles Then
         gobjLogMgr.LogMsg "Opening: """ & chmpath & """ ..."
         ShellExecute 0&, vbNullString, chmpath, vbNullString, vbNullString, vbNormalFocus
      End If
   End If
   
   frmProcWnd.Finish
   Exit Sub
hErr:
   gobjLogMgr.LogErrorR
   Resume Next
End Sub

Private Sub ConvertXml2Hhp(ByVal xml As DOMDocument, ByVal OutDir As String)
   On Error GoTo hErr
   
   Dim strErrInfo As String
   
   Dim configdir As String
   configdir = fso.BuildPath(App.Path, "config")
   
   Dim xsl As New DOMDocument
   xsl.Load fso.BuildPath(configdir, IIf(gobjConfig.TreatHelpStringAsHtml, "htm2.xsl", "htm1.xsl"))
   
   Dim htmldir As String
   htmldir = fso.BuildPath(TempDir, "html")
   
   If fso.FolderExists(htmldir) Then
      fso.DeleteFolder htmldir, True
   End If
   fso.CreateFolder htmldir
   
   ' * Ignore Unions/UnionInfo here, because vb can not use union.
   Dim nodes As IXMLDOMNodeList
   Set nodes = xml.selectNodes("//TypeLibInfo|//CoclassInfo|//InterfaceInfo" & _
      "|//ConstantInfo|//DeclarationInfo|//RecordInfo|//MemberInfo" & _
      "|//Coclasses|//Interfaces|//Constants|//Declarations|//Records")
   
   Dim i As Long
   i = 1
   
'   Dim colPaths As Collection
'   Set colPaths = New Collection
   
   Dim node As IXMLDOMElement
   For Each node In nodes
      Dim result As String
      result = node.transformNode(xsl)
      
      If result <> "" Then
         Dim filepath As String
         filepath = fso.BuildPath(htmldir, node.getAttribute("htmlname") & ".htm")
         
         strErrInfo = filepath
         WriteFile filepath, result, False
'         colPaths.Add filepath
      End If
      
      frmProcWnd.ShowProcess2 (node.nodeName & " (" & i & "/" & nodes.length & _
         "," & 100 * i \ nodes.length & "%) ... " & _
         node.getAttribute("name"))
      CheckCancel
      i = i + 1
   Next
   
   strErrInfo = "hhc.xsl"
   xsl.Load fso.BuildPath(configdir, "hhc.xsl")
   result = xml.transformNode(xsl)
   
   ' This is a bug of compiler, without charactor "CrLf" there would be an error.
   Dim pos As Long
   pos = InStr(1, result, "<param name=""Merge""", vbTextCompare)
   
   If pos > 0 Then
      result = Mid(result, 1, pos - 1) & vbCrLf & Mid(result, pos)
   End If
   
   WriteFile fso.BuildPath(TempDir, "output.hhc"), result, False
   CheckCancel
   
   strErrInfo = "hhk.xsl"
   xsl.Load fso.BuildPath(configdir, "hhk.xsl")
   result = xml.transformNode(xsl)
   WriteFile fso.BuildPath(TempDir, "output.hhk"), result, False
   CheckCancel
   
   strErrInfo = "output.hhp"
   Dim content As String
   content = ReadFile(fso.BuildPath(App.Path, "output.hhp"))
   
   content = Replace(content, "<Default topic>", "html/" & _
      GetAttr(xml.documentElement.firstChild, "htmlname") & ".htm")
      
   Dim libname As String
   libname = GetAttr(xml.documentElement.firstChild, "name") & _
      "(" & GetAttr(xml.documentElement.firstChild, "majorversion") & "." & _
      GetAttr(xml.documentElement.firstChild, "minorversion") & ")"
      
   strCompiledFile = fso.BuildPath(OutDir, libname & ".chm")
   content = Replace(content, "<Compiled file>", strCompiledFile)
   content = Replace(content, "<Title>", libname)
   content = Replace(content, "<Title bar text>", libname)
   
   Dim helpfile As String
   helpfile = GetAttr(xml.documentElement.firstChild, "helpfile")
   
'   content = content & vbCrLf & "[FILES]" & vbCrLf
'
'   Dim varPath As Variant
'   For Each varPath In colPaths
'      content = content & varPath & vbCrLf
'   Next
   
   If helpfile <> "" Then
      Dim ref As String
      ref = GetAttr(xml.documentElement.firstChild, "helpref")
      
      fso.CopyFile helpfile, fso.BuildPath(OutDir, ref), True
      content = content & vbCrLf & "[MERGE FILES]" & vbCrLf & ref
   End If
   
   
   WriteFile fso.BuildPath(TempDir, "output.hhp"), content, False
   Exit Sub
hErr:
   gobjLogMgr.LogErrorR "@ConvertXml2Hhp(" & strErrInfo & ")"
   gblnError = True
   Resume Next
End Sub

Private Function GetAttr(ByVal src As IXMLDOMElement, ByVal name As String) As String
   GetAttr = src.getAttribute(name)
End Function

Public Sub CheckCancel()
   DoEvents
   If Not ce Is Nothing Then
      ce.Cancel = True
   End If
   If Cancel Then
      Err.Raise errUserCancel
   End If
End Sub

Private Sub Class_Initialize()
   Set ce = New PUPPYCHMCOMPILER2Lib.CompileEngine
   LogInitialize Me
End Sub

Private Sub Class_Terminate()
   LogTerminate Me
End Sub


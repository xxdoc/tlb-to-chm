VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TlbManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private colTlbPaths As New Collection
Private reg As New Registry
Private mblnLoaded As Boolean

Public Property Get Item(ByVal Index As Long) As TypeLibInfo
   Set Item = TypeLibInfoFromFile(colTlbPaths.Item(Index))
End Property

Public Property Get Path(ByVal Index As Long) As String
   Path = colTlbPaths.Item(Index)
End Property

Private Function ExpandEnvironmentVariable(ByVal Path As String) As String
   Static reg As RegExp
   If reg Is Nothing Then
      Set reg = New RegExp
      reg.MultiLine = False
      reg.IgnoreCase = True
      reg.Global = True
      reg.Pattern = "%(\w+)%"
   End If
   
   Dim m As Match
   For Each m In reg.Execute(Path)
      Dim name As String
      name = m.SubMatches(0)
      
      Dim value As String
      value = Environ(name)
      
      If value <> "" Then
         Path = Replace(Path, "%" & name & "%", value)
      End If
   Next
   
   ExpandEnvironmentVariable = Path
End Function

Private Function Add(ByVal strPath As String) As Boolean
   On Error GoTo hErr
   
   Dim tlb As TypeLibInfo
   Set tlb = TypeLibInfoFromFile(strPath)
   
   colTlbPaths.Add strPath
   Add = True
   Exit Function
hErr:
   gobjLogMgr.LogErrorR strPath
   Add = False
End Function

Public Sub LoadTypeLibraryPaths()
   On Error GoTo hErr
   mblnLoaded = True
   
   Dim reg As New Registry
   frmProcWnd.Prepare False
   
   Dim i As Long
   For i = 1 To colTlbPaths.count
      colTlbPaths.Remove 1
   Next
   
   Dim regpath As String
   regpath = "Software\Classes\TypeLib"
   
   Dim guid As Variant
   For Each guid In reg.GetSubKeys(iHKEY_LOCAL_MACHINE, regpath)
      Dim regpath2 As String
      regpath2 = regpath & "\" & guid
      
      Dim ver As Variant
      For Each ver In reg.GetSubKeys(iHKEY_LOCAL_MACHINE, regpath2)
         Dim regpath3 As String
         regpath3 = regpath2 & "\" & ver
         
         Dim arr() As String
         arr = Split(ver, ".")
         
         If IsNumeric(ver) And UBound(arr) - LBound(arr) + 1 = 2 Then
            Dim major As Long
            major = arr(0)
            
            Dim minor As Long
            minor = arr(1)
            
            Dim lcid As Variant
            For Each lcid In reg.GetSubKeys(iHKEY_LOCAL_MACHINE, regpath3)
               If IsNumeric(lcid) Then
                  If CLng(lcid) = lcid Then
                     Dim regpath4 As String
                     regpath4 = regpath3 & "\" & lcid
                     
                     ' maybe less when no sys
                     Dim sys As Variant
                     For Each sys In reg.GetSubKeys(iHKEY_LOCAL_MACHINE, regpath4)
                        Dim regpath5 As String
                        regpath5 = regpath4 & "\" & sys
                        
                        Dim filepath As String
                        reg.GetValue iHKEY_LOCAL_MACHINE, regpath5, "", filepath
                        
                        filepath = ExpandEnvironmentVariable(filepath)
                        
                        If Add(filepath) Then
                           frmProcWnd.ShowProcess "Loading from registry ... " & fso.GetFileName(filepath)
                        End If
                        
                        CheckCancel
                     Next
                     
                     CheckCancel
                  End If
               End If
            Next
            
         End If
      Next
      CheckCancel
   Next
   frmProcWnd.Finish
   Exit Sub
hErr:
   gobjLogMgr.LogErrorR "@LoadTypeLibraryPaths(" & filepath & ")"
   Resume Next
End Sub

Private Sub CheckCancel()
   DoEvents
   If frmProcWnd.mblnCancel Then
      Err.Raise errUserCancel
   End If
End Sub

Private Function RegExpTest(ByVal src As String, ByVal ptn As String) As Boolean
   On Error GoTo hErr
   Static reg As RegExp
   If reg Is Nothing Then
      Set reg = New RegExp
      With reg
         .Global = True
         .MultiLine = False
         .IgnoreCase = True
         .Pattern = ptn
      End With
   End If
   reg.Pattern = ptn
   RegExpTest = reg.Test(src)
   Exit Function
hErr:
   RegExpTest = False
End Function

Public Sub FillList(ByRef lb As ListBox, Optional ByVal ptn As String)
   On Error GoTo hErr
   With lb
      .Clear
      
      Dim i As Long
      i = 1
      
      Dim varPath As Variant
      For Each varPath In colTlbPaths
         Dim tlb As TypeLibInfo
         Set tlb = TypeLibInfoFromFile(varPath)
         
         Dim strItem As String
         strItem = tlb.name & " (" & tlb.MajorVersion & "." & tlb.MinorVersion & ")"
         
         If tlb.helpfile <> "" Then
            strItem = strItem & " *"
         End If
         
         If tlb.HelpString <> "" Then
            strItem = strItem & " - " & tlb.HelpString
         End If
         
         If RegExpTest(strItem, ptn) Then
            .AddItem strItem
            .ItemData(.NewIndex) = i
         End If
         
         i = i + 1
      Next
   End With
   Exit Sub
hErr:
   gobjLogMgr.LogErrorR "@FillList()"
   Resume Next
End Sub

Private Sub Class_Initialize()
   LogInitialize Me
End Sub

Private Sub Class_Terminate()
   LogTerminate Me
End Sub


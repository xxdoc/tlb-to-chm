VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LogManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event LogMsg(ByVal msg As String)
Public Event Reset()

Public Sub Reset()
   RaiseEvent Reset
End Sub

Public Sub AddMsgToList(ByVal lst As ListBox, ByVal msg As String)
   With lst
      .AddItem Replace(msg, vbCrLf, " ")
      If .ListCount > 500 Then
         .RemoveItem 0
      End If
      .ListIndex = .ListCount - 1
   End With
End Sub

Private Property Get LogFilePath() As String
   LogFilePath = fso.BuildPath(App.Path, App.EXEName & ".log")
End Property

Public Sub LogMsg(Optional ByVal msg As Variant)
   On Error Resume Next
   msg = Now & " >> " & msg
   WriteFile LogFilePath, msg, True
   RaiseEvent LogMsg(msg)
   Debug.Print msg
End Sub

Private Function IsCancelError(ByVal num As Long) As Boolean
   Select Case num
   Case errUserCancel
      IsCancelError = True
   Case Else
      IsCancelError = False
   End Select
End Function

Public Sub LogErrorNR(Optional ByVal msg As Variant)
   If IsCancelError(Err.Number) Then
      Exit Sub
   End If
     
   Dim strError As String
   strError = Err.Number & ": " & Err.Description

   If Not IsMissing(msg) Then
      strError = strError & " => " & msg
   End If
   
   LogMsg strError
End Sub

Public Sub LogErrorR(Optional ByVal msg As Variant)
   Dim arr(0 To 4) As Variant
   arr(0) = Err.Number
   arr(1) = Err.Source
   arr(2) = Err.Description
   arr(3) = Err.helpfile
   arr(4) = Err.HelpContext

   If IsCancelError(Err.Number) Then
      Err.Raise arr(0), arr(1), arr(2), arr(3), arr(4)
   Else
      LogErrorNR msg
   End If
End Sub

Private Sub Class_Initialize()
   LogInitialize Me
End Sub

Private Sub Class_Terminate()
   LogTerminate Me
End Sub


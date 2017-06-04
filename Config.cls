VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Public OutDir As String
Public TreatHelpStringAsHtml As Boolean
Public OpenFinishedChmFiles As Boolean

Private Sub Class_Initialize()
'   OutDir = GetSetting(App.EXEName, "Config", "OutDir", "c:\")
   TreatHelpStringAsHtml = GetSetting(App.EXEName, "Config", "TreatHelpStringAsHtml", False)
   OpenFinishedChmFiles = GetSetting(App.EXEName, "Config", "OpenFinishedChmFiles", True)
   
   LogInitialize Me
End Sub

Private Sub Class_Terminate()
'   SaveSetting App.EXEName, "Config", "OutDir", OutDir
   SaveSetting App.EXEName, "Config", "TreatHelpStringAsHtml", TreatHelpStringAsHtml
   SaveSetting App.EXEName, "Config", "OpenFinishedChmFiles", OpenFinishedChmFiles
   
   LogTerminate Me
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tli2XmlConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public AutoImportExternal As Boolean

Private xml As New DOMDocument
Private mlngCount As Long

Private Static Function GenHtmlName(ByVal name As String) As String
   mlngCount = mlngCount + 1
   GenHtmlName = name & "_t2c_" & mlngCount
End Function

Public Sub Reset()
   With xml
      .loadXML ""
      .appendChild .createProcessingInstruction("xml", "version='1.0'")
      Set .documentElement = .createElement("TypeLibs")
      mlngCount = 0
   End With
End Sub

Private Function AddXmlElement(ByRef parent As IXMLDOMElement, ByVal name As String) As IXMLDOMElement
   Set AddXmlElement = parent.appendChild(xml.createElement(name))
   frmProcWnd.ShowProcess2 "Generating xml element ... " & name
   gobjTaskMgr.CheckCancel
End Function

Private Sub SetXmlAttribute(ByRef elem As IXMLDOMElement, ByVal name As String, ByVal value As String)
   elem.setAttribute name, value
End Sub

Private Sub SetAttributeStrings(ByRef elem As IXMLDOMElement, ByVal obj As Object)
   Dim arr() As String
   Dim count As Long
   count = obj.AttributeStrings(arr)
   
   Dim str As String
   Dim i As Long
   For i = 1 To count
      str = str & arr(i)
      If i < count Then
         str = str & ","
      End If
   Next
   
   elem.setAttribute "attributestrings", str
   gobjTaskMgr.CheckCancel
End Sub

Private Sub ConvertCustomData(ByRef elem As IXMLDOMElement, ByRef cd As CustomData)
   On Error GoTo hErr
   Dim node As IXMLDOMElement
   Set node = AddXmlElement(elem, "CustomData")

   With cd
      SetXmlAttribute node, "guid", .guid
      SetXmlAttribute node, "value", .value
   End With
   Exit Sub
hErr:
   gobjLogMgr.LogErrorR "@ConvertCustomData()"
   Resume Next
End Sub

Private Sub ConvertCustomDataCollection(ByRef elem As IXMLDOMElement, ByRef col As CustomDataCollection)
   On Error GoTo hErr
   Dim node As IXMLDOMElement
   Set node = AddXmlElement(elem, "CustomDataCollection")
   
   Dim cd As CustomData
   For Each cd In col
      ConvertCustomData node, cd
   Next
   Exit Sub
hErr:
   gobjLogMgr.LogErrorR "@ConvertCustomDataCollection()"
   Resume Next
End Sub

Public Function StartConvert(ByVal tlinfo As TypeLibInfo) As DOMDocument
   Reset
   ConvertTypeLibInfo tlinfo
   xml.save App.Path & "\" & App.EXEName & ".xml"
   Set StartConvert = xml
End Function

Private Function GetHelpFilePath(ByRef tlinfo As TypeLibInfo) As String
   GetHelpFilePath = ""
   With tlinfo
      If .helpfile = "" Then
         Exit Function
      ElseIf fso.FileExists(.helpfile) Then
         GetHelpFilePath = .helpfile
      Else
         Dim Path As String
         Path = fso.BuildPath(fso.GetParentFolderName(.ContainingFile), .helpfile)
         
         If fso.FileExists(Path) Then
            GetHelpFilePath = Path
         Else
            Dim arr() As String
            arr = Split(Environ("path"), ";")
            
            Dim i As Long
            For i = LBound(arr) To UBound(arr)
               Path = fso.BuildPath(arr(i), .helpfile)
               
               If fso.FileExists(Path) Then
                  GetHelpFilePath = Path
                  Exit For
               End If
               
               gobjTaskMgr.CheckCancel
            Next
         End If
      End If
   End With
End Function

Private Function GetHhcName(ByRef tlinfo As TypeLibInfo) As String
   Dim Path As String
   Path = GetHelpFilePath(tlinfo)
   
   If fso.FileExists(Path) Then
      GetHhcName = ReadHHCName(Path)
   End If
End Function

Private Function GetShortHelpString(ByVal strHelpString As String) As String
   Dim lngPos As Long
   lngPos = InStr(1, strHelpString, ".", vbTextCompare)
   
   Dim strResult As String
   If lngPos > 0 Then
      strResult = Left(strHelpString, lngPos)
   Else
      strResult = strHelpString
   End If
   
   Static reg As RegExp
   If reg Is Nothing Then
      Set reg = New RegExp
      reg.Global = True
      reg.IgnoreCase = True
      reg.MultiLine = True
      reg.Pattern = "<.+?>"
   End If
   GetShortHelpString = reg.Replace(strResult, "")
End Function

Private Sub ConvertTypeLibInfo(ByRef tlinfo As TypeLibInfo)
   On Error GoTo hErr
   Dim node As IXMLDOMElement
   Set node = AddXmlElement(xml.documentElement, "TypeLibInfo")
   
   With tlinfo
      SetXmlAttribute node, "name", .name
      SetAttributeStrings node, .Me
      SetXmlAttribute node, "containingfile", .ContainingFile
      SetXmlAttribute node, "guid", .guid
      SetXmlAttribute node, "helpstring", .HelpString
      SetXmlAttribute node, "shorthelpstring", GetShortHelpString(.HelpString)
      SetXmlAttribute node, "lcid", .lcid
      SetXmlAttribute node, "majorversion", .MajorVersion
      SetXmlAttribute node, "minorversion", .MinorVersion
      SetXmlAttribute node, "syskind", .SysKind
      SetXmlAttribute node, "htmlname", GenHtmlName(.name)
      SetXmlAttribute node, "helpfile", GetHelpFilePath(tlinfo)
      SetXmlAttribute node, "helpref", _
         .name & "(" & .MajorVersion & "." & .MinorVersion & ")_ref.chm"
      SetXmlAttribute node, "hhcname", GetHhcName(tlinfo)
      
      ConvertCustomDataCollection node, .CustomDataCollection
      ConvertCoclasses node, .CoClasses
      ConvertInterfaces node, .Interfaces
      ConvertDeclarations node, .Declarations
      ConvertConstants node, .Constants
      ConvertRecords node, .Records
      ConvertUnions node, .Unions
   End With
   Exit Sub
hErr:
   gobjLogMgr.LogErrorR "@ConvertTypeLibInfo()"
   Resume Next
End Sub

Private Sub ConvertCoclasses(ByRef elem As IXMLDOMElement, ByRef col As CoClasses)
   On Error GoTo hErr
   Dim node As IXMLDOMElement
   Set node = AddXmlElement(elem, "Coclasses")
   
   SetXmlAttribute node, "htmlname", GenHtmlName("Coclasses")

   Dim cci As CoClassInfo
   For Each cci In col
      ConvertCoclassInfo node, cci
   Next
   Exit Sub
hErr:
   gobjLogMgr.LogErrorR "@ConvertCoclasses()"
   Resume Next
End Sub

Private Sub ConvertCoclassInfo(ByRef elem As IXMLDOMElement, ByRef cci As CoClassInfo)
   On Error GoTo hErr
   Dim node As IXMLDOMElement
   Set node = AddXmlElement(elem, "CoclassInfo")
   
   With cci
      SetXmlAttribute node, "name", .name
      SetXmlAttribute node, "guid", .guid
      SetXmlAttribute node, "typekindstring", .TypeKindString
      SetAttributeStrings node, .Me
      SetXmlAttribute node, "helpstring", .HelpString
      SetXmlAttribute node, "shorthelpstring", GetShortHelpString(.HelpString)
      SetXmlAttribute node, "majorversion", .MajorVersion
      SetXmlAttribute node, "minorversion", .MinorVersion
      SetXmlAttribute node, "htmlname", GenHtmlName(.name)
      If Not .DefaultEventInterface Is Nothing Then
         SetXmlAttribute node, "defaulteventinterfaceguid", .DefaultEventInterface.guid
      End If
      If Not .DefaultInterface Is Nothing Then
         SetXmlAttribute node, "defaultinterfaceguid", .DefaultInterface.guid
      End If
      
      ConvertCustomDataCollection node, .CustomDataCollection
   End With
   Exit Sub
hErr:
   gobjLogMgr.LogErrorR "@ConvertCoclassInfo()"
   Resume Next
End Sub

Private Sub ConvertInterfaces( _
   ByRef elem As IXMLDOMElement, _
   ByRef col As Interfaces, _
   Optional ByVal elemname As String = "Interfaces")
   On Error GoTo hErr
   
   Dim node As IXMLDOMElement
   Set node = AddXmlElement(elem, elemname)

   SetXmlAttribute node, "htmlname", GenHtmlName("Interfaces")

   Dim ifi As InterfaceInfo
   For Each ifi In col
      ConvertInterfaceInfo node, ifi
   Next
   Exit Sub
hErr:
   gobjLogMgr.LogErrorR "@ConvertInterfaces()"
   Resume Next
End Sub

Private Sub ConvertInterfaceInfo( _
   ByRef elem As IXMLDOMElement, _
   ByRef ifi As InterfaceInfo, _
   Optional ByVal elemname As String = "InterfaceInfo")
   On Error GoTo hErr
   
   If ifi Is Nothing Then
      Exit Sub
   End If
   
   Dim node As IXMLDOMElement
   Set node = AddXmlElement(elem, elemname)
   
   With ifi
      SetXmlAttribute node, "name", .name
      SetAttributeStrings node, .Me
      SetXmlAttribute node, "guid", .guid
      SetXmlAttribute node, "typekindstring", .TypeKindString
      SetXmlAttribute node, "helpstring", .HelpString
      SetXmlAttribute node, "shorthelpstring", GetShortHelpString(.HelpString)
      SetXmlAttribute node, "majorversion", .MajorVersion
      SetXmlAttribute node, "minorversion", .MinorVersion
      SetXmlAttribute node, "htmlname", GenHtmlName(.name)
   
      ConvertCustomDataCollection node, .CustomDataCollection
      ConvertInterfaces node, .ImpliedInterfaces, "ImpliedInterfaces"
      ConvertMembers node, .Members, .TypeKind
   End With
   Exit Sub
hErr:
   gobjLogMgr.LogErrorR "@ConvertInterfaceInfo()"
   Resume Next
End Sub

Private Sub ConvertMembers(ByRef elem As IXMLDOMElement, ByRef col As Members, ByVal tk As TypeKinds)
   On Error GoTo hErr
   Dim node As IXMLDOMElement
   Set node = AddXmlElement(elem, "Members")

   SetXmlAttribute node, "htmlname", GenHtmlName("Members")

   Dim mi As MemberInfo
   For Each mi In col
      ConvertMember node, mi, tk
   Next
   Exit Sub
hErr:
   gobjLogMgr.LogErrorR "@ConvertMembers()"
   Resume Next
End Sub

Private Function GetVbMemberType(ByRef mi As MemberInfo, ByVal tk As TypeKinds) As String
   With mi
      Select Case tk
      Case TKIND_ENUM, TKIND_RECORD, TKIND_UNION
         GetVbMemberType = "Const"
      Case Else
         Select Case mi.InvokeKind
         Case INVOKE_FUNC, INVOKE_EVENTFUNC
            Select Case .ReturnType.VarType
            Case VT_VOID, VT_HRESULT
               GetVbMemberType = "Sub"
            Case Else
               GetVbMemberType = "Function"
            End Select
         Case Else
            GetVbMemberType = "Property"
         End Select
      End Select
   End With
End Function

Private Function GetInvokeKindString(ByRef ik As InvokeKinds) As String
   Select Case ik
   Case INVOKE_CONST
      GetInvokeKindString = "Const"
   Case INVOKE_EVENTFUNC
      GetInvokeKindString = "Event"
   Case INVOKE_FUNC
      GetInvokeKindString = "Function"
   Case INVOKE_PROPERTYGET
      GetInvokeKindString = "Get"
   Case INVOKE_PROPERTYPUT
      GetInvokeKindString = "Let"
   Case INVOKE_PROPERTYPUTREF
      GetInvokeKindString = "Set"
   Case INVOKE_UNKNOWN
      GetInvokeKindString = "Unknown"
   End Select
End Function

Private Sub ConvertMember(ByRef elem As IXMLDOMElement, ByRef mi As MemberInfo, ByVal tk As TypeKinds)
   On Error GoTo hErr
   Dim node As IXMLDOMElement
   Set node = AddXmlElement(elem, "MemberInfo")
   
   With mi
      SetXmlAttribute node, "name", .name
      SetAttributeStrings node, .Me
      SetXmlAttribute node, "desckind", .DescKind
      SetXmlAttribute node, "callconv", .CallConv
      SetXmlAttribute node, "helpstring", .HelpString
      SetXmlAttribute node, "shorthelpstring", GetShortHelpString(.HelpString)
      SetXmlAttribute node, "invokekind", .InvokeKind
      SetXmlAttribute node, "invokekindstring", GetInvokeKindString(.InvokeKind)
      SetXmlAttribute node, "htmlname", GenHtmlName(.name)
      
      If .DescKind = DESCKIND_VARDESC And (tk = TKIND_ENUM Or tk = TKIND_MODULE Or tk = TKIND_ALIAS) Then
         SetXmlAttribute node, "value", .value
      End If
      
      SetXmlAttribute node, "vbtype", GetVbMemberType(.Me, tk)
      
      ConvertCustomDataCollection node, .CustomDataCollection
      ConvertParameters node, .Parameters
      ConvertVarTypeInfo node, .ReturnType, "ReturnType"
   End With
   Exit Sub
hErr:
   gobjLogMgr.LogErrorR "ConvertMember(" & node.getAttribute("name") & ")"
   Resume Next
End Sub

Private Sub ConvertParameters(ByRef elem As IXMLDOMElement, ByRef col As Parameters)
   On Error GoTo hErr
   Dim node As IXMLDOMElement
   Set node = AddXmlElement(elem, "Parameters")

   SetXmlAttribute node, "htmlname", GenHtmlName("Parameters")

   With col
      SetXmlAttribute node, "defaultcount", .DefaultCount
      SetXmlAttribute node, "optionalcount", .OptionalCount
   End With

   Dim pi As ParameterInfo
   For Each pi In col
      ConvertParameterInfo node, pi
   Next
   Exit Sub
hErr:
   gobjLogMgr.LogErrorR "@ConvertParameters()"
   Resume Next
End Sub

Private Sub ConvertParameterInfo( _
   ByRef elem As IXMLDOMElement, _
   ByRef pi As ParameterInfo)
   On Error GoTo hErr
   
   Dim node As IXMLDOMElement
   Set node = AddXmlElement(elem, "ParameterInfo")
   
   With pi
      SetXmlAttribute node, "name", .name
      SetXmlAttribute node, "default", .default
      Dim DefVal As Variant
      If .default Then
         'Static util As New Utilities
         'DefVal = util.ConvertVar(.DefaultValue)
         If VarType(.DefaultValue) = vbObject Then
            Set DefVal = .DefaultValue
         Else
            DefVal = .DefaultValue
         End If
      Else
         DefVal = Empty
      End If
      ' Here maybe a unknown value error.
      'If Not (IsNull(DefVal) Or IsEmpty(DefVal)) Then
      '   SetXmlAttribute node, "defaultvalue", DefVal
      'End If
      SetXmlAttribute node, "flags", .Flags
      SetXmlAttribute node, "optional", .Optional
      SetXmlAttribute node, "htmlname", GenHtmlName(.name)
      
      ConvertCustomDataCollection node, .CustomDataCollection
      ConvertVarTypeInfo node, .VarTypeInfo, "VarTypeInfo", DefVal
   End With
   Exit Sub
hErr:
   gobjLogMgr.LogErrorR "@ConvertParameterInfo()"
   Resume Next
End Sub

' 获取VB的类型名。
Private Function GetVbVarTypeName(ByRef vti As VarTypeInfo, ByVal elemname As String) As String
   With vti
      If Not vti.TypeInfo Is Nothing Then 'VbArrayTypeOnly(.Me) Or (elemname = "ReturnType" And vti.VarType = 0) Then
         Dim ti As TypeInfo
         Set ti = GetVbTypeInfo(.Me)
         
         If ti Is Nothing Then 'Error
            GetVbVarTypeName = "?"
         ElseIf .IsExternalType Then ' 外部库。。。
            GetVbVarTypeName = .TypeLibInfoExternal.name & "." & ti.name
         Else
            GetVbVarTypeName = ti.name
         End If
      Else
         Select Case elemname
         Case "VarTypeInfo"
            GetVbVarTypeName = TypeName(.TypedVariant)
         Case "ReturnType"
            Select Case vti.VarType
            Case VT_VARIANT, VT_VOID, VT_HRESULT
            Case Else
               GetVbVarTypeName = TypeName(.TypedVariant)
            End Select
         Case Else
            Debug.Assert False
         End Select
         
         If IsVbArrayType(.Me) Then
            GetVbVarTypeName = Replace(GetVbVarTypeName, "()", "")
         End If
      End If
   End With
End Function

' 判断是否为值传递。
Private Function IsVbByVal(ByRef vti As VarTypeInfo) As Boolean
   With vti
      If VbArrayTypeOnly(.Me) Then
         Select Case GetVbTypeKind(.Me)
         Case TKIND_INTERFACE, TKIND_COCLASS, TKIND_DISPATCH
            IsVbByVal = (.PointerLevel = 1)
         Case TKIND_RECORD
            'Records not passed ByVal in VB
            IsVbByVal = False
         Case Else
            IsVbByVal = (.PointerLevel = 0)
         End Select
      Else
         IsVbByVal = (.PointerLevel = 0)
      End If
   End With
End Function

Private Function GetVbTypeKind(ByVal vti As VarTypeInfo) As TypeKinds
   Dim ti As TypeInfo
   Set ti = GetVbTypeInfo(vti)
   
   If Not ti Is Nothing Then
      GetVbTypeKind = ti.TypeKind
   Else
      GetVbTypeKind = TKIND_MAX
   End If
End Function

Private Function GetVbTypeGuid(ByRef vti As VarTypeInfo) As String
   Dim ti As TypeInfo
   Set ti = GetVbTypeInfo(vti)
   
   If Not ti Is Nothing Then
      Select Case ti.TypeKind
      Case TKIND_INTERFACE, TKIND_COCLASS, TKIND_DISPATCH
         GetVbTypeGuid = ti.guid
      Case Else
         GetVbTypeGuid = ""
      End Select
   Else
      GetVbTypeGuid = ""
   End If
End Function

Private Function GetVbDefVal(ByVal vti As VarTypeInfo, ByVal DefVal As Variant) As String
   Dim ti As TypeInfo
   Set ti = GetVbTypeInfo(vti)
   
   If ti Is Nothing Then
      Select Case VarType(DefVal)
      Case vbString
         If Len(DefVal) > 0 Then
            GetVbDefVal = """" & DefVal & """"
         End If
      Case vbBoolean 'Always show for Boolean
         GetVbDefVal = DefVal
      Case vbDate
         If DefVal Then
            GetVbDefVal = "#" & DefVal & "#"
         End If
      Case Else 'Numeric Values
         If IsNumeric(DefVal) Then
            GetVbDefVal = DefVal
         End If
      End Select
   Else
      'See if we have an enum and track the matching member
      'If the type is an object, then there will never be a
      'default value other than Nothing
      If ti.TypeKind = TKIND_ENUM Then
         Dim mi As MemberInfo
         For Each mi In ti.Members
            If CLng(mi.value) = CLng(DefVal) Then
               GetVbDefVal = mi.name
               Exit For
            End If
         Next
      End If
   End If
End Function

Private Function VbArrayTypeOnly(ByVal vti As VarTypeInfo) As Boolean
   VbArrayTypeOnly = ((vti.VarType And Not (VT_ARRAY Or VT_VECTOR)) = 0)
End Function

Private Function IsVbArrayType(ByVal vti As VarTypeInfo) As Boolean
   IsVbArrayType = ((vti.VarType And (VT_ARRAY Or VT_VECTOR)) <> 0)
End Function

Private Function GetVbTypeInfo(ByRef vti As VarTypeInfo) As TypeInfo
   Set GetVbTypeInfo = GetResolvedType(vti.TypeInfo)
'   If Not GetVbTypeInfo Is Nothing Then
'      If Left(GetVbTypeInfo.name, 6) = "__MIDL" Then
'         gLogMgr.LogMessage "Warning: GetVbTypeInfo " & GetVbTypeInfo.name
''         Debug.Assert False
'      End If
'   End If
End Function

' 转换VarTypeInfo为XML。
Private Sub ConvertVarTypeInfo( _
   ByRef elem As IXMLDOMElement, _
   ByRef vti As VarTypeInfo, _
   ByVal elemname As String, _
   Optional ByVal DefVal As Variant)
   On Error GoTo hErr

   Dim node As IXMLDOMElement
   Set node = AddXmlElement(elem, elemname)
   
   With vti
      SetXmlAttribute node, "isexternaltype", .IsExternalType
      SetXmlAttribute node, "pointerlevel", .PointerLevel
      SetXmlAttribute node, "typeinfonumber", .TypeInfoNumber
      SetXmlAttribute node, "vartype", .VarType
      
      SetXmlAttribute node, "vbarray", IsVbArrayType(.Me)
      SetXmlAttribute node, "vbbyval", IsVbByVal(.Me)
      SetXmlAttribute node, "vbtypename", GetVbVarTypeName(.Me, elemname)
      SetXmlAttribute node, "vbtypeguid", GetVbTypeGuid(.Me)
      SetXmlAttribute node, "vbdefval", GetVbDefVal(vti, DefVal)
      
      If IsVbArrayType(.Me) Then
         SetXmlAttribute node, "elementpointerlevel", .ElementPointerLevel
         
         If (.VarType And VT_VECTOR) <> 0 Then
            ConvertArrayBounds node, vti
         End If
      End If
      
      If .IsExternalType And AutoImportExternal Then
         ConvertTypeLibInfo .TypeLibInfoExternal
      End If
   End With
   Exit Sub
hErr:
   gobjLogMgr.LogErrorR "@ConvertVarTypeInfo()"
   Resume Next
End Sub

Private Sub ConvertArrayBounds(ByRef elem As IXMLDOMElement, ByRef vti As VarTypeInfo)
   On Error GoTo hErr
   Dim node As IXMLDOMElement
   Set node = AddXmlElement(elem, "ArrayBounds")

   Dim arr() As Long
   Dim i As Long
   For i = 1 To vti.ArrayBounds(arr)
      ConvertArrayBound node, arr, i
   Next
   Exit Sub
hErr:
   gobjLogMgr.LogErrorR "@ConvertArrayBounds()"
   Resume Next
End Sub

Private Sub ConvertArrayBound(ByRef elem As IXMLDOMElement, ByRef arr() As Long, ByVal i As Long)
   On Error GoTo hErr
   Dim node As IXMLDOMElement
   Set node = AddXmlElement(elem, "ArrayBound")
   
   SetXmlAttribute node, "lbound", arr(i, 1)
   SetXmlAttribute node, "ubound", arr(i, 2)
   Exit Sub
hErr:
   gobjLogMgr.LogErrorR "@ConvertArrayBound()"
   Resume Next
End Sub

Private Function GetResolvedType(ByVal ti As TypeInfo) As TypeInfo
   If ti Is Nothing Then
      Set GetResolvedType = Nothing
   ElseIf ti.TypeInfoNumber <= 0 Then
      Set GetResolvedType = ti
   ElseIf Left(ti.name, 1) <> "_" Then
      Set GetResolvedType = ti
   Else
      ' 改为使用IndexedItem
      Set GetResolvedType = ti.parent.TypeInfos.IndexedItem(ti.TypeInfoNumber)
   End If
End Function

Private Sub ConvertTypeInfo( _
   ByRef elem As IXMLDOMElement, _
   ByVal ti As TypeInfo, _
   ByVal elemname As String)
   On Error GoTo hErr
   
   Dim node As IXMLDOMElement
   Set node = AddXmlElement(elem, elemname)
   
   ' Do not resolve the type information here, only the vartype needs to be resolved.
   With ti 'GetResolvedType(ti)
      SetXmlAttribute node, "name", .name
      SetXmlAttribute node, "guid", .guid
      SetAttributeStrings node, .Me
      SetXmlAttribute node, "typekindstring", .TypeKindString
      SetXmlAttribute node, "helpstring", .HelpString
      SetXmlAttribute node, "shorthelpstring", GetShortHelpString(.HelpString)
      SetXmlAttribute node, "majorversion", .MajorVersion
      SetXmlAttribute node, "minorversion", .MinorVersion
      SetXmlAttribute node, "htmlname", GenHtmlName(.name)
      
      ConvertCustomDataCollection node, .CustomDataCollection
      ConvertMembers node, .Members, .TypeKind
   End With
   Exit Sub
hErr:
   gobjLogMgr.LogErrorR "@ConvertTypeInfo()"
   Resume Next
End Sub

Private Sub ConvertConstants(ByRef elem As IXMLDOMElement, ByRef col As tli.Constants)
   On Error GoTo hErr
   Dim node As IXMLDOMElement
   Set node = AddXmlElement(elem, "Constants")

   SetXmlAttribute node, "htmlname", GenHtmlName("Constants")
   
   Dim ci As ConstantInfo
   For Each ci In col
      ConvertConstantInfo node, ci
   Next
   Exit Sub
hErr:
   gobjLogMgr.LogErrorR "@ConvertConstants()"
   Resume Next
End Sub

Private Sub ConvertConstantInfo(ByRef elem As IXMLDOMElement, ByRef ci As ConstantInfo)
   ConvertTypeInfo elem, ci, "ConstantInfo"
End Sub

Private Sub ConvertDeclarations(ByRef elem As IXMLDOMElement, ByRef col As Declarations)
   On Error GoTo hErr
   Dim node As IXMLDOMElement
   Set node = AddXmlElement(elem, "Declarations")

   SetXmlAttribute node, "htmlname", GenHtmlName("Declarations")

   Dim di As DeclarationInfo
   For Each di In col
      ConvertDeclarationInfo node, di
   Next
   Exit Sub
hErr:
   gobjLogMgr.LogErrorR "@ConvertDeclarations()"
   Resume Next
End Sub

Private Sub ConvertDeclarationInfo(ByRef elem As IXMLDOMElement, ByRef di As DeclarationInfo)
   ConvertTypeInfo elem, di, "DeclarationInfo"
End Sub

Private Sub ConvertRecords(ByRef elem As IXMLDOMElement, ByRef col As Records)
   On Error GoTo hErr
   Dim node As IXMLDOMElement
   Set node = AddXmlElement(elem, "Records")

   SetXmlAttribute node, "htmlname", GenHtmlName("Records")

   Dim ri As RecordInfo
   For Each ri In col
      ConvertRecordInfo node, ri
   Next
   Exit Sub
hErr:
   gobjLogMgr.LogErrorR "@ConvertRecords()"
   Resume Next
End Sub

Private Sub ConvertRecordInfo(ByRef elem As IXMLDOMElement, ByRef ri As RecordInfo)
   ConvertTypeInfo elem, ri, "RecordInfo"
End Sub

Private Sub ConvertUnions(ByRef elem As IXMLDOMElement, ByRef col As Unions)
   On Error GoTo hErr
   Dim node As IXMLDOMElement
   Set node = AddXmlElement(elem, "Unions")

   SetXmlAttribute node, "htmlname", GenHtmlName("Unions")

   Dim ri As UnionInfo
   For Each ri In col
      ConvertUnionInfo node, ri
   Next
   Exit Sub
hErr:
   gobjLogMgr.LogErrorR "@ConvertUnions()"
   Resume Next
End Sub

Private Sub ConvertUnionInfo(ByRef elem As IXMLDOMElement, ByRef ri As UnionInfo)
   ConvertTypeInfo elem, ri, "UnionInfo"
End Sub

Private Sub Class_Initialize()
   LogInitialize Me
End Sub

Private Sub Class_Terminate()
   LogTerminate Me
End Sub


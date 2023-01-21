VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodeModuleSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>%AppFolder%/source/CodeModuleSupport.cls</file>
'</codelib>
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Private Const ErrNumber91_ObjectVariableOrWithBlockVariableNotSet As Long = 91

Private Const BAC_EarlyBindingConstName As String = "BAC_EarlyBinding"
Private Const ConditionalCompilationArgumentsOptionName As String = "Conditional Compilation Arguments"
Private Const BetterAccessChartsReferenceName As String = "BetterAccessCharts"

#Const EarlyBinding = 0
#Const DEBUGMODE = 0

#If EarlyBinding = 1 Then
#Else

   Public Enum vbext_ComponentType
      vbext_ct_StdModule = 1
      vbext_ct_ClassModule = 2
      vbext_ct_MSForm = 3
      vbext_ct_ActiveXDesigner = 11
      vbext_ct_Document = 100
   End Enum
   
#End If

#If EarlyBinding = 1 Then
Private Property Get CurrentVbProject() As VBIDE.VBProject
#Else
Private Property Get CurrentVbProject() As Object
#End If

#If EarlyBinding = 1 Then
   Dim Proj As VBProject
   Dim Temp As VBProject
#Else
   Dim Proj As Object
   Dim Temp As Object
#End If
   Dim strCurrentDbName As String

   Set Proj = VBE.ActiveVBProject
   'Prüfen, ob das richtige VbProject gewählt wurde (muss das von CurrentDb sein)
   strCurrentDbName = UncPath(CurrentDb.Name)
   If Proj.FileName <> strCurrentDbName Then
      Set Proj = Nothing
      For Each Temp In VBE.VBProjects
         If Temp.FileName = strCurrentDbName Then
            Set Proj = Temp
            Exit For
         End If
      Next
   End If
   
   Set CurrentVbProject = Proj

End Property

Public Property Get EarlyBinding() As Boolean

   Dim ConditionalCompilationArgumentsString As String
   Dim ConditionalCompilationArguments() As String
   Dim TestString As String
   Dim i As Long
   
   ConditionalCompilationArgumentsString = Trim(Nz(Application.GetOption(ConditionalCompilationArgumentsOptionName), vbNullString))
   
   If Len(ConditionalCompilationArgumentsString) = 0 Then
      EarlyBinding = False
      Exit Property
   End If
   
   ConditionalCompilationArguments = GetConditionalCompilationArgumentsArray(ConditionalCompilationArgumentsString)
   For i = LBound(ConditionalCompilationArguments) To UBound(ConditionalCompilationArguments)
      TestString = Replace(ConditionalCompilationArguments(i), " ", vbNullString)
      If TestString Like BAC_EarlyBindingConstName & "=*" Then
         EarlyBinding = (TestString <> BAC_EarlyBindingConstName & "=0")
         Exit For
      End If
   Next
   
End Property

Public Property Let EarlyBinding(ByVal UseEarlyBinding As Boolean)

   Dim ConditionalCompilationArgumentsString As String
   Dim ConditionalCompilationArguments() As String
   Dim TestString As String
   Dim ArgExists As Boolean
   Dim i As Long
   
   ConditionalCompilationArgumentsString = Trim(Nz(Application.GetOption("Conditional Compilation Arguments"), vbNullString))
   
   If Len(ConditionalCompilationArgumentsString) = 0 Then
      Application.SetOption ConditionalCompilationArgumentsOptionName, BAC_EarlyBindingConstName & " = " & Abs(UseEarlyBinding)
      Exit Property
   End If
   
   ConditionalCompilationArguments = GetConditionalCompilationArgumentsArray(ConditionalCompilationArgumentsString)
   For i = LBound(ConditionalCompilationArguments) To UBound(ConditionalCompilationArguments)
      If Replace(ConditionalCompilationArguments(i), " ", vbNullString) Like BAC_EarlyBindingConstName & "=*" Then
         ConditionalCompilationArguments(i) = BAC_EarlyBindingConstName & " = " & Abs(UseEarlyBinding)
         ArgExists = True
         Exit For
      End If
   Next
   
   If ArgExists Then
      ConditionalCompilationArgumentsString = Join(ConditionalCompilationArguments, ":")
   Else
      ConditionalCompilationArgumentsString = ConditionalCompilationArgumentsString & " : " & BAC_EarlyBindingConstName & " = " & Abs(UseEarlyBinding)
   End If

   Application.SetOption ConditionalCompilationArgumentsOptionName, ConditionalCompilationArgumentsString
   
   If UseEarlyBinding Then
      AddReference
   Else
      RemoveReference
   End If

End Property

Private Function GetConditionalCompilationArgumentsArray(ByVal FullString As String) As String()
   GetConditionalCompilationArgumentsArray = Split(FullString, ":")
End Function

Private Sub AddReference()
   If Not ReferenceExits() Then
    Application.References.AddFromFile CodeProject.FullName
   End If
End Sub

Private Function GetAddInLocation() As String
   Dim strLocation As String
   strLocation = CodeProject.Path
   GetAddInLocation = strLocation
End Function

Private Function ReferenceExits() As Boolean
   Dim ref As Reference
   Dim BacRef As Reference
   For Each ref In Application.References
      If ref.Name = BetterAccessChartsReferenceName Then
         ReferenceExits = True
         Exit Function
      End If
   Next
   ReferenceExits = False
End Function

Private Sub RemoveReference()
   Dim ref As Reference
   Dim BacRef As Reference
   For Each ref In Application.References
      If ref.Name = BetterAccessChartsReferenceName Then
         Set BacRef = ref
      End If
   Next
   If Not (BacRef Is Nothing) Then
      Application.References.Remove BacRef
   End If
End Sub


Public Sub CreateModulesInApp()
   CreateModuleInApp BAC_LoaderModuleName, GetBetterAccessChartsLoaderModuleString
   CreateModuleInApp BAC_FactoryModuleName, GetBetterAccessChartsFactoryModuleString
End Sub

Public Sub CreateModuleInApp(ByVal ModuleName As String, ByVal CodeString As String)

On Error GoTo Fehler

#If EarlyBinding = 1 Then
    Dim cm As CodeModule
    Dim vbc As VBComponent
    Dim vbp As VBProject
#Else
    Dim cm As Object
    Dim vbc As Object
    Dim vbp As Object
#End If
    
    If CodeModuleExists(BAC_FactoryModuleName) Then
        Set vbc = CurrentVbProject.VBComponents(ModuleName)
    Else 'Modul anlegen und benennen
        Set vbc = CurrentVbProject.VBComponents.Add(vbext_ct_StdModule)
        vbc.Name = BAC_FactoryModuleName
    End If

   'Code einfügen
    Set cm = vbc.CodeModule
    cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString CodeString
    
    'Leerzeilen am Ende säubern
    Do While Len(Trim$(cm.Lines(cm.CountOfLines, 1))) = 0 And cm.CountOfLines > 1
        cm.DeleteLines cm.CountOfLines
    Loop
    
    Call SaveModule
    
Ende_Cleanup:
    On Error Resume Next
    Set cm = Nothing
    Set vbc = Nothing
    Set vbp = Nothing
    Exit Sub

Fehler:
    If Err.Number = ErrNumber91_ObjectVariableOrWithBlockVariableNotSet Then
        RaiseError ErrNumber91_ObjectVariableOrWithBlockVariableNotSet, "TransferCodeModule.CreateModuleInApp", Err.Source, "Keine Datenbank geöffnet!"
    Else
        RaiseError Err.Number, "TransferCodeModule.CreateModuleInApp", Err.Source, Err.Description
    End If
    Resume Ende_Cleanup
    
End Sub

Private Sub SaveModule()
    DoCmd.Save acModule, BAC_FactoryModuleName
    DoCmd.Close acModule, BAC_FactoryModuleName
End Sub

Public Sub DeleteModuleFromApp()

On Error GoTo Fehler

#If EarlyBinding = 1 Then
    Dim vbc As VBComponent
#Else
    Dim vbc As Object
#End If
    
    Set vbc = CurrentVbProject.VBComponents(BAC_FactoryModuleName)
    CurrentVbProject.VBComponents.Remove vbc
    
Ende_Cleanup:
    On Error Resume Next
    Set vbc = Nothing
    Exit Sub

Fehler:
    If Err.Number = ErrNumber91_ObjectVariableOrWithBlockVariableNotSet Then
        RaiseError ErrNumber91_ObjectVariableOrWithBlockVariableNotSet, "TransferCodeModule.DeleteModuleFromApp", Err.Source, "Keine Datenbank geöffnet!"
    Else
        RaiseError Err.Number, "TransferCodeModule.DeleteModuleFromApp", Err.Source, Err.Description
    End If
    Resume Ende_Cleanup

End Sub

Public Function CodeModuleExists( _
                ByVal CodeModulName As String) As Boolean

#If EarlyBinding = 1 Then
   Dim VbcCol As VBIDE.VBComponents
   Dim vbc As VBIDE.VBComponent
#Else
   Dim VbcCol As Object
   Dim vbc As Object
#End If
   
   Set VbcCol = CurrentVbProject.VBComponents
   For Each vbc In VbcCol
      If vbc.Name = CodeModulName Then
         CodeModuleExists = True
         Exit For
      End If
   Next
   
End Function

Private Sub RaiseError( _
                    ByVal ErrNumber As Long, _
                    ByVal ErrSource As String, _
                    ByVal ErrSubSource As String, _
                    ByVal ErrDescription As String)

    If ErrSubSource <> "sqlDebugPrint" Then
        ErrSource = ErrSource & "->" & ErrSubSource
    End If
    Err.Raise ErrNumber, ErrSource, ErrDescription

End Sub
                    
Private Function GetBetterAccessChartsFactoryModuleString() As String

    Dim str As String
    
    str = "Option Compare Database" & vbNewLine & _
         "Option Explicit" & vbNewLine & _
         "" & vbNewLine & _
         "Private m_Bac As Object" & vbNewLine & _
         "" & vbNewLine & _
         "#If BAC_EarlyBinding Then" & vbNewLine & _
         "#Else" & vbNewLine
         
   str = str & vbNewLine & _
"Public Enum chChartType" & vbNewLine & _
"    Line = 1" & vbNewLine & _
"    Bar = 2" & vbNewLine & _
"    HorizontalBar = 3" & vbNewLine & _
"    Radar = 4" & vbNewLine & _
"    Doughnut = 5" & vbNewLine & _
"    Pie = 6" & vbNewLine & _
"    PolarArea = 7" & vbNewLine & _
"    Bubble = 8" & vbNewLine & _
"    Scatter = 9" & vbNewLine & _
"End Enum" & vbNewLine & _
"" & vbNewLine & _
"Public Enum chDataSourceType" & vbNewLine & _
"    dstDemo" & vbNewLine & _
"    dstTableName" & vbNewLine & _
"    dstQueryName" & vbNewLine & _
"    dstSQLStament" & vbNewLine & _
"    dstRecordset" & vbNewLine & _
"    dstData" & vbNewLine & _
"    dstEmpty" & vbNewLine & _
"End Enum" & vbNewLine

   str = str & vbNewLine & _
"Public Enum chPosition" & vbNewLine & _
"    posTop = 1" & vbNewLine & _
"    posLeft = 2" & vbNewLine & _
"    posBottom = 3" & vbNewLine & _
"    posRight = 4" & vbNewLine & _
"End Enum" & vbNewLine

   str = str & vbNewLine & _
"Public Enum chScriptSource" & vbNewLine & _
"    CDN = 1" & vbNewLine & _
"    LocalFile = 2" & vbNewLine & _
"End Enum" & vbNewLine

   str = str & vbNewLine & _
"Public Enum chAlign" & vbNewLine & _
"    alStart = 1" & vbNewLine & _
"    alCenter = 2" & vbNewLine & _
"    alEnd = 3" & vbNewLine & _
"End Enum" & vbNewLine

   str = str & vbNewLine & _
"Public Enum chDataLabelAnchor" & vbNewLine & _
"    anStart = 1" & vbNewLine & _
"    anCenter = 2" & vbNewLine & _
"    anEnd = 3" & vbNewLine & _
"End Enum" & vbNewLine

   str = str & vbNewLine & _
"Public Enum chDisplayIn" & vbNewLine & _
"    chWebBrowserControl = 1" & vbNewLine & _
"    chWebBrowserActiveX = 2" & vbNewLine & _
"    chImageControl = 3" & vbNewLine & _
"    chSystemBrowser = 4" & vbNewLine & _
"End Enum" & vbNewLine

   str = str & vbNewLine & _
"Public Enum chEasing" & vbNewLine & _
"    linear = 0" & vbNewLine & _
"    easeInQuad = 1" & vbNewLine & _
"    easeOutQuad = 2" & vbNewLine & _
"    easeInOutQuad = 3" & vbNewLine & _
"    easeInCubic = 4" & vbNewLine & _
"    easeOutCubic = 5" & vbNewLine & _
"    easeInOutCubic = 6" & vbNewLine & _
"    easeInQuart = 7" & vbNewLine & _
"    easeOutQuart = 8" & vbNewLine & _
"    easeInOutQuart = 9" & vbNewLine & _
"    easeInQuint = 10" & vbNewLine & _
"    easeOutQuint = 11" & vbNewLine & _
"    easeInOutQuint = 12" & vbNewLine & _
"    easeInSine = 13" & vbNewLine & _
"    easeOutSine = 14" & vbNewLine & _
"    easeInOutSine = 15" & vbNewLine & _
"    easeInExpo = 16" & vbNewLine

   str = str & _
"    easeOutExpo = 17" & vbNewLine & _
"    easeInOutExpo = 18" & vbNewLine & _
"    easeInCirc = 19" & vbNewLine & _
"    easeOutCirc = 20" & vbNewLine & _
"    easeInOutCirc = 21" & vbNewLine & _
"    easeInElastic = 22" & vbNewLine & _
"    easeOutElastic = 23" & vbNewLine & _
"    easeInOutElastic = 24" & vbNewLine & _
"    easeInBack = 25" & vbNewLine & _
"    easeOutBack = 26" & vbNewLine & _
"    easeInOutBack = 27" & vbNewLine & _
"    easeInBounce = 28" & vbNewLine & _
"    easeOutBounce = 29" & vbNewLine & _
"    easeInOutBounce = 30" & vbNewLine & _
"End Enum" & vbNewLine
   
   str = str & vbNewLine & _
"#End If"

   str = str & vbNewLine & _
"' Factory" & vbNewLine & _
"#If BAC_EarlyBinding Then" & vbNewLine & _
"Public Function BAC() As BAC__Factory" & vbNewLine & _
"#Else" & vbNewLine & _
"Public Function BAC() As Object" & vbNewLine & _
"#End If" & vbNewLine & _
"    If m_Bac Is Nothing Then" & vbNewLine & _
"        Set m_Bac = BetterAccessChartsLoader.GetBetterAccessChartsFactory" & vbNewLine & _
"    End If" & vbNewLine & _
"    Set BAC = m_Bac" & vbNewLine & _
"End Function" & vbNewLine
       
    GetBetterAccessChartsFactoryModuleString = str

End Function

Private Function GetBetterAccessChartsLoaderModuleString() As String

   Dim str As String
   
str = _
"Option Compare Database" & vbNewLine & _
"Option Explicit" & vbNewLine & _
"" & vbNewLine & _
"' Set BAC_EarlyBinding in project property for compiler arguments" & vbNewLine & _
"' BAC_EarlyBinding = 1 ... CheckBetterAccessChartsReference add reference to accda (BAC created by accda-Reference)" & vbNewLine & _
"' BAC_EarlyBinding = 0 ... CheckBetterAccessChartsReference removes reference (BAC created by add-in)" & vbNewLine & _
"" & vbNewLine & _
"Private Const BetterAccessChartsFileName As String = ""BetterAccessCharts""" & vbNewLine & _
"Private Const BetterAccessChartsReferenceName As String = ""BetterAccessCharts""" & vbNewLine & _
"Private Const BetterAccessChartsFactory As String = ""BAC""" & vbNewLine

   str = str & vbNewLine & _
"Public Sub CheckBetterAccessChartsReference()" & vbNewLine & _
"   CheckReference" & vbNewLine & _
"End Sub" & vbNewLine

   str = str & vbNewLine & _
"Public Function GetBetterAccessChartsFactory() As Object" & vbNewLine & _
"   CheckReference" & vbNewLine & _
"#If BAC_EarlyBinding Then" & vbNewLine & _
"    Set GetBetterAccessChartsFactory = BetterAccessCharts.BetterAccessCharts" & vbNewLine & _
"#Else" & vbNewLine & _
"    Set GetBetterAccessChartsFactory = Application.Run(GetAddInLocation & BetterAccessChartsFileName & ""."" & BetterAccessChartsFactory)" & vbNewLine & _
"#End If" & vbNewLine & _
"End Function" & vbNewLine

   str = str & vbNewLine & _
"Private Function GetAddInLocation() As String" & vbNewLine & _
"   GetAddInLocation = GetAppDataLocation & ""\Microsoft\AddIns\""" & vbNewLine & _
"   'GetAddInLocation = CodeProject.Path & ""\""" & vbNewLine & _
"   ' ... welcher Speicherort ist besser?" & vbNewLine & _
"End Function" & vbNewLine

   str = str & vbNewLine & _
"Private Function GetAppDataLocation()" & vbNewLine & _
"   With CreateObject(""WScript.Shell"")" & vbNewLine & _
"      GetAppDataLocation = .ExpandEnvironmentStrings(""%APPDATA%"") & """"" & vbNewLine & _
"   End With" & vbNewLine & _
"End Function" & vbNewLine

   str = str & vbNewLine & _
"Private Sub CheckReference()" & vbNewLine & _
"   Static m_ReferenceChecked As Boolean" & vbNewLine & _
"   Static m_UseEarlyBindingState As Boolean" & vbNewLine & _
"#If BAC_EarlyBinding Then" & vbNewLine & _
"   If m_UseEarlyBindingState = False Then" & vbNewLine & _
"      m_ReferenceChecked = False" & vbNewLine & _
"      m_UseEarlyBindingState = True" & vbNewLine & _
"   End If" & vbNewLine & _
"#Else" & vbNewLine & _
"   If m_UseEarlyBindingState = True Then" & vbNewLine & _
"      m_ReferenceChecked = False" & vbNewLine & _
"      m_UseEarlyBindingState = False" & vbNewLine & _
"   End If" & vbNewLine & _
"#End If" & vbNewLine & _
"   If m_ReferenceChecked Then" & vbNewLine & _
"      Exit Sub" & vbNewLine & _
"   End If" & vbNewLine & _
"#If BAC_EarlyBinding Then" & vbNewLine & _
"    AddReference" & vbNewLine & _
"#Else" & vbNewLine & _
"    RemoveReference" & vbNewLine & _
"#End If" & vbNewLine & _
"End Sub" & vbNewLine

   str = str & vbNewLine & _
"Private Sub AddReference()" & vbNewLine & _
"   RemoveReference" & vbNewLine & _
"   Application.References.AddFromFile GetAddInLocation & BetterAccessChartsFileName & "".accda""" & vbNewLine & _
"End Sub" & vbNewLine

   str = str & vbNewLine & _
"Private Sub RemoveReference()" & vbNewLine & _
"   Dim ref As Reference" & vbNewLine & _
"   Dim BacRef As Reference" & vbNewLine & _
"   For Each ref In Application.References" & vbNewLine & _
"      If ref.Name = BetterAccessChartsReferenceName Then" & vbNewLine & _
"         Set BacRef = ref" & vbNewLine & _
"      End If" & vbNewLine & _
"   Next" & vbNewLine & _
"   If Not (BacRef Is Nothing) Then" & vbNewLine & _
"      Application.References.Remove BacRef" & vbNewLine & _
"   End If" & vbNewLine & _
"End Sub" & vbNewLine

   GetBetterAccessChartsLoaderModuleString = str

End Function

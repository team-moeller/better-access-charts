Attribute VB_Name = "BetterAccessChartsLoader"
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>%AppFolder%/source/BetterAccessChartsLoader.bas</file>
'</codelib>
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

' Set BAC_EarlyBinding in project property for compiler arguments
' BAC_EarlyBinding = 1 ... CheckBetterAccessChartsReference add reference to accda (BAC created by accda reference)
' BAC_EarlyBinding = 0 ... CheckBetterAccessChartsReference removes reference (BAC created by add-in)

Private Const BetterAccessChartsFileName As String = "BetterAccessCharts"
Private Const BetterAccessChartsReferenceName As String = "BetterAccessCharts"
Private Const BetterAccessChartsFactory As String = "BAC"
Private Const BetterAccessChartsAddInTools As String = "BACx"

Public Sub CheckBetterAccessChartsReference()
   CheckReference
End Sub

Public Function GetBetterAccessChartsFactory() As Object
   CheckReference
#If BAC_EarlyBinding Then
    Set GetBetterAccessChartsFactory = BetterAccessCharts.BetterAccessCharts
#Else
    Set GetBetterAccessChartsFactory = Application.Run(GetAddInLocation & BetterAccessChartsFileName & "." & BetterAccessChartsFactory)
#End If
End Function

Public Function GetBetterAccessChartsAddInTools() As Object
   CheckReference
#If BAC_EarlyBinding Then
    Set GetBetterAccessChartsAddInTools = BetterAccessCharts.BetterAccessCharts
#Else
    Set GetBetterAccessChartsAddInTools = Application.Run(GetAddInLocation & BetterAccessChartsFileName & "." & BetterAccessChartsAddInTools)
#End If
End Function

Private Function GetAddInLocation() As String
   Dim strLocation As String
   strLocation = GetAppDataLocation & "\Microsoft\AddIns\"
   'GetAddInLocation = CodeProject.Path & "\"
   ' ... welcher Speicherort ist besser?
   If Len(VBA.Dir(strLocation & BetterAccessChartsFileName & ".accda")) = 0 Then
      strLocation = CodeProject.Path & "\"
      If Len(VBA.Dir(strLocation & BetterAccessChartsFileName & ".accda")) = 0 Then
         Err.Raise vbObjectError, "BetterAccessChartsLoader.GetAddInLocation", "Add-In file is missing"
      End If
   End If
   GetAddInLocation = strLocation
End Function

Private Function GetAppDataLocation()
   With CreateObject("WScript.Shell")
      GetAppDataLocation = .ExpandEnvironmentStrings("%APPDATA%") & ""
   End With
End Function

Private Sub CheckReference()
   Static m_ReferenceChecked As Boolean
   Static m_UseEarlyBindingState As Boolean
#If BAC_EarlyBinding Then
   If m_UseEarlyBindingState = False Then
      m_ReferenceChecked = False
      m_UseEarlyBindingState = True
   End If
#Else
   If m_UseEarlyBindingState = True Then
      m_ReferenceChecked = False
      m_UseEarlyBindingState = False
   End If
#End If
   If m_ReferenceChecked Then
      Exit Sub
   End If
#If BAC_EarlyBinding Then
    AddReference
#Else
    RemoveReference
#End If
End Sub

Private Sub AddReference()
   If Not ReferenceExits() Then
    Application.References.AddFromFile GetAddInLocation & BetterAccessChartsFileName & ".accda"
   End If
End Sub

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

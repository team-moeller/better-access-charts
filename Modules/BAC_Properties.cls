VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_Properties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Private m_Properties As Collection

Private Sub Class_Initialize()
   Set m_Properties = New Collection
End Sub

Public Function NewProperty(ByVal Name As String) As BAC_Property
   With New BAC_Property
      .Name = Name
      Append .Self
      Set NewProperty = .Self
   End With
End Function

Public Sub Append(ByVal Property As BAC_Property)
   m_Properties.Add Property, Property.Name
End Sub

Public Property Get Item(ByVal Index As Variant) As BAC_Property
   Set Item = m_Properties.Item(Index)
End Property

Public Property Get Count() As Long
   Count = m_Properties.Count
End Property

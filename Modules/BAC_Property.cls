VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_Property"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Private m_Name As String
Private m_Value As Variant
Private m_Properties As BAC_Properties

Public Property Get Self() As BAC_Property
   Set Self = Me
End Property

Public Property Get Name() As String
   Name = m_Name
End Property

Friend Property Let Name(ByVal NewValue As String)
   m_Name = NewValue
End Property

Public Property Get Value() As Variant
   Value = m_Value
End Property

Public Property Let Value(ByVal NewValue As Variant)
   m_Value = NewValue
End Property

Public Property Get Properties() As BAC_Properties
   If m_Properties Is Nothing Then
      Set m_Properties = New BAC_Properties
   End If
   Set Properties = m_Properties
End Property

Public Property Set Properties(ByVal NewRef As BAC_Properties)
   Set m_Properties = NewRef
End Property

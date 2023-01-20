VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_Font"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###########################################################################################
'# Copyright (c) 2020 - 2022 Thomas Moeller, supported by K.D.Gundermann                   #
'# MIT License  => https://github.com/team-moeller/better-access-charts/blob/main/LICENSE  #
'# Version 2.30.02  published: 30.12.2022                                                  #
'###########################################################################################

Option Compare Database
Option Explicit


'### Members
Private m_Properties As BAC_Properties

' ### Properties
Public Property Get Properties() As BAC_Properties
   Set Properties = m_Properties
End Property

Public Property Let Color(ByVal This_Color As String)
    m_Properties.Item("Color").Value = This_Color
End Property
Public Property Get Color() As String
    Color = Nz(m_Properties.Item("Color").Value, vbNullString)
End Property

Public Property Let Family(ByVal This_Family As String)
    m_Properties.Item("Family").Value = This_Family
End Property
Public Property Get Family() As String
    Family = Nz(m_Properties.Item("Family").Value, vbNullString)
End Property

Public Property Let Size(ByVal This_Size As Integer)
    m_Properties.Item("Size").Value = This_Size
End Property
Public Property Get Size() As Integer
    Size = Nz(m_Properties.Item("Size").Value, 0)
End Property

Public Property Let IsItalic(ByVal This_IsItalic As Boolean)
    m_Properties.Item("IsItalic").Value = This_IsItalic
End Property
Public Property Get IsItalic() As Boolean
    IsItalic = Nz(m_Properties.Item("IsItalic").Value, False)
End Property

Public Property Let IsBold(ByVal This_IsBold As Boolean)
    m_Properties.Item("IsBold").Value = This_IsBold
End Property
Public Property Get IsBold() As Boolean
    IsBold = Nz(m_Properties.Item("IsBold").Value, False)
End Property


' ### Functions

Friend Property Get SectionIsNeeded(Optional Parent As String) As Boolean

    Select Case Parent
        Case "DataLabel"
            SectionIsNeeded = (Family <> vbNullString) Or (Size > 0) Or (IsBold = True) Or (IsItalic = True)
        Case "Legend"
            SectionIsNeeded = (Color <> vbNullString) Or (Family <> vbNullString) Or (Size > 0) Or (IsBold = True) Or (IsItalic = True)
        Case Else
            'Do nothing
    End Select
    
End Property

Private Sub Class_Initialize()
   InitProperties
End Sub

Private Sub InitProperties()
   Set m_Properties = New BAC_Properties
   m_Properties.NewProperty "Color"
   m_Properties.NewProperty "Family"
   m_Properties.NewProperty "Size"
   m_Properties.NewProperty "IsItalic"
   m_Properties.NewProperty "IsBold"
End Sub

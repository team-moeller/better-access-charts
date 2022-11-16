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
'# Version 2.25.03  published: 16.11.2022                                                  #
'###########################################################################################

Option Compare Database
Option Explicit


'### Members

Private m_Color As String
Private m_Family As String
Private m_Size As Integer
Private m_IsItalic As Boolean
Private m_IsBold As Boolean


' ### Properties

Public Property Let Color(ByVal This_Color As String)
    m_Color = This_Color
End Property
Public Property Get Color() As String
    Color = m_Color
End Property

Public Property Let Family(ByVal This_Family As String)
    m_Family = This_Family
End Property
Public Property Get Family() As String
    Family = m_Family
End Property

Public Property Let Size(ByVal This_Size As Integer)
    m_Size = This_Size
End Property
Public Property Get Size() As Integer
    Size = m_Size
End Property

Public Property Let IsItalic(ByVal This_IsItalic As Boolean)
    m_IsItalic = This_IsItalic
End Property
Public Property Get IsItalic() As Boolean
    IsItalic = m_IsItalic
End Property

Public Property Let IsBold(ByVal This_IsBold As Boolean)
    m_IsBold = This_IsBold
End Property
Public Property Get IsBold() As Boolean
    IsBold = m_IsBold
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

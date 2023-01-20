VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_ColorSchema"
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


' ### Members

Private m_Properties As BAC_Properties

Private Sub Class_Initialize()
   InitProperties
End Sub

Private Sub InitProperties()
   Set m_Properties = New BAC_Properties
   m_Properties.NewProperty "Name"
End Sub

' ### Properties
Public Property Get Properties() As BAC_Properties
   Set Properties = m_Properties
End Property

Public Property Let Name(ByVal NewValue As String)
    m_Properties.Item("Name").Value = NewValue
End Property
Public Property Get Name() As String
    Name = Nz(m_Properties.Item("Name").Value, vbNullString)
End Property


' ### Generator Function

Friend Function getColorSchema() As String

    With New BAC_StringBuilder
        .AddLine "                colorschemes: {"
        .AddLine "                    scheme: '" & Me.Name & "'"
        .AddLine "                    },"
        
        getColorSchema = .ToString
    End With
    
End Function

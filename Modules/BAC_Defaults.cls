VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_Defaults"
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

Private m_Font As BAC_Font
Private m_Properties As BAC_Properties

' ### Properties
Public Property Get Properties() As BAC_Properties
   Set Properties = m_Properties
End Property

Public Property Get Font() As BAC_Font
    Set Font = m_Font
End Property


'### Instancing

Private Sub Class_Initialize()
    Set m_Font = New BAC_Font
    InitProperties
End Sub

Private Sub InitProperties()
   Set m_Properties = New BAC_Properties
   Set m_Properties.NewProperty("Font").Properties = m_Font.Properties
End Sub

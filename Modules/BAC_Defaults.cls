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
'# Copyright (c) 2020 - 2022 Thomas Möller, supported by K.D.Gundermann                    #
'# MIT License  => https://github.com/team-moeller/better-access-charts/blob/main/LICENSE  #
'# Version 2.04.04  published: 15.01.2022                                                  #
'###########################################################################################

Option Compare Database
Option Explicit


Private m_Font As BAC_Font

Private Sub Class_Initialize()
    Set m_Font = New BAC_Font
End Sub

Public Property Get Font() As BAC_Font
    Set Font = m_Font
End Property

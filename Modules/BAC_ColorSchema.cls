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
'# Copyright (c) 2020, 2021 Thomas Möller, supported by K.D.Gundermann                     #
'# MIT License  => https://github.com/team-moeller/better-access-charts/blob/main/LICENSE  #
'# Version 1.31.12  published: 08.01.2022                                                  #
'###########################################################################################

Option Compare Database
Option Explicit


Private m_ColorSchema As String

' ### Properties

Public Property Let ColorSchema(ByVal This_ColorSchema As String)
    m_ColorSchema = This_ColorSchema
End Property
Public Property Get ColorSchema() As String
    ColorSchema = m_ColorSchema
End Property

' ### Generator Function

Friend Function getColorSchema() As String
    With New BAC_StringBuilder
        .AddLine "                colorschemes: {"
        .AddLine "                    scheme: '" & ColorSchema & "'"
        .AddLine "                    },"
        
        getColorSchema = .ToString
    End With
End Function


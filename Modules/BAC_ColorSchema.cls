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
'# Version 2.25.03  published: 16.11.2022                                                  #
'###########################################################################################

Option Compare Database
Option Explicit


' ### Members

Private m_Name As String


' ### Properties

Public Property Let Name(ByVal newValue As String)
    m_Name = newValue
End Property
Public Property Get Name() As String
    Name = m_Name
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


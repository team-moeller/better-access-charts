VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_Tooltip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###########################################################################################
'# Copyright (c) 2020 - 2024 Thomas Moeller, supported by K.D.Gundermann                   #
'# MIT License  => https://github.com/team-moeller/better-access-charts/blob/main/LICENSE  #
'# Version 4.21.09  published: 07.04.2024                                                  #
'###########################################################################################

Option Compare Database
Option Explicit


'### Members

Private m_Enabled As Boolean


'### Properties

Public Property Let Enabled(ByVal newValue As Boolean)
    m_Enabled = newValue
End Property
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property


'### Instancing

Private Sub Class_Initialize()

    m_Enabled = True
    
End Sub


'### Generator Function

Friend Function getTootip() As String
    
    With New BAC_StringBuilder
        If m_Enabled = False Then
            .AddLine "                tooltip: {"
            .AddLine "                    enabled: false"
            .AddLine "                },"
        Else
            .AddLine "                tooltip: {"
            .AddLine "                    enabled: true,"
            .CutTrailingComma
            .AddLine "                },"
        End If
        getTootip = .ToString
    End With
    
End Function

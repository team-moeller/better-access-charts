VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_Legend_Title"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###########################################################################################
'# Copyright (c) 2020 - 2024 Thomas Moeller, supported by K.D.Gundermann                   #
'# MIT License  => https://github.com/team-moeller/better-access-charts/blob/main/LICENSE  #
'# Version 4.26.05  published: 26.09.2024                                                  #
'###########################################################################################

Option Compare Database
Option Explicit


'### Members

Private m_Show As Boolean
Private m_Text As String


' ### Properties

Public Property Let Show(ByVal newValue As Boolean)
    m_Show = newValue
End Property
Public Property Get Show() As Boolean
    Show = m_Show
End Property

Public Property Let Text(ByVal This_Text As String)
    m_Text = This_Text
End Property
Public Property Get Text() As String
    Text = m_Text
End Property


'### Instancing

Private Sub Class_Initialize()

    'Nothing so far
    
End Sub


' ### Generator Function

Friend Function getTitle() As String

    With New BAC_StringBuilder
        If Me.Show = False Then
            'Nothing
        Else
            .AddLine "                    title: {"
            .AddLine "                        display: true,"
            .AddLine "                        text: '" & Me.Text & "',"
            .CutTrailingComma
            .AddLine "                    },"
        End If
        getTitle = .ToString
    End With
    
End Function





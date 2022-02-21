VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_Legend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###########################################################################################
'# Copyright (c) 2020 - 2022 Thomas Moeller, supported by K.D.Gundermann                   #
'# MIT License  => https://github.com/team-moeller/better-access-charts/blob/main/LICENSE  #
'# Version 2.18.06  published: 21.02.2022                                                  #
'###########################################################################################

Option Compare Database
Option Explicit


'### Members

Private m_Font As BAC_Font
Private m_Show As Boolean
Private m_Position As chPosition


' ### Properties

Public Property Get Font() As BAC_Font
    Set Font = m_Font
End Property

Public Property Let Show(ByVal This_Show As Boolean)
    m_Show = This_Show
End Property
Public Property Get Show() As Boolean
    Show = m_Show
End Property

Public Property Let Position(ByVal This_Position As chPosition)
    m_Position = This_Position
End Property
Public Property Get Position() As chPosition
    Position = m_Position
End Property


'### Instancing

Private Sub Class_Initialize()
    Set m_Font = New BAC_Font
End Sub


' ### Generator Function

Friend Function getLegend() As String

    With New BAC_StringBuilder
        If Me.Show = False Then
            .AddLine "            legend: {"
            .AddLine "                display: false"
            .AddLine "            },"
        Else
            .AddLine "            legend: {"
            .AddLine "                display: true,"
            .AddLine "                position: " & BAC.Helper.getPositionString(Me.Position) & ","
            If Me.Font.SectionIsNeeded("Legend") = True Then
                .AddLine "                labels: {"
                If Me.Font.Family <> vbNullString Then
                    .AddLine "                    fontFamily: '" & Me.Font.Family & "',"
                End If
                If Me.Font.Size > 0 Then
                    .AddLine "                    fontSize: " & Me.Font.Size & ","
                End If
                If Me.Font.Color <> vbNullString Then
                    .AddLine "                    fontColor: '" & Me.Font.Color & "',"
                End If
                If Me.Font.IsBold = True And Me.Font.IsItalic = True Then
                    .AddLine "                    fontStyle: 'bold italic'"
                ElseIf Me.Font.IsBold = True Then
                    .AddLine "                    fontStyle: 'bold'"
                ElseIf Me.Font.IsItalic = True Then
                    .AddLine "                    fontStyle: 'italic'"
                Else
                    'Do nothing
                End If
                .CutTrailingComma
                .AddLine "                }"
            End If
            .AddLine "            },"
        End If
        getLegend = .ToString
    End With
    
End Function


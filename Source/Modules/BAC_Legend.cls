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
'# Copyright (c) 2020, 2021 Thomas Möller, supported by K.D.Gundermann                     #
'# MIT License  => https://github.com/team-moeller/better-access-charts/blob/main/LICENSE  #
'# Version 1.31.12  published: 29.10.2021                                                  #
'###########################################################################################

Option Compare Database
Option Explicit

Private m_Show As Boolean
Private m_Position As chPosition


' ### Properties

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


' ### Generator Function

Friend Function getLegend() As String

    'Declarations
    Dim strText As String

    If Me.Show = False Then
        strText = strText & "            legend: {" & vbCrLf
        strText = strText & "                display: false" & vbCrLf
        strText = strText & "            }," & vbCrLf
    Else
        strText = strText & "            legend: {" & vbCrLf
        strText = strText & "                display: true," & vbCrLf
        strText = strText & "                position: " & BAC.Helper.getPositionString(Me.Position) & vbCrLf
        strText = strText & "            }," & vbCrLf
    End If

    getLegend = strText

End Function


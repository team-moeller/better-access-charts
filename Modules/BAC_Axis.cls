VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_Axis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###########################################################################################
'# Copyright (c) 2020, 2021 Thomas Möller, supported by K.D.Gundermann                     #
'# MIT License  => https://github.com/team-moeller/better-access-charts/blob/main/LICENSE  #
'# Version xxxxx  published: 04.11.2021                                                  #
'###########################################################################################

Option Compare Database
Option Explicit

Public Enum chCartesianAxis
    apXAxis = 1
    apYAxis = 2
End Enum

Private m_Axis As chCartesianAxis
Private m_Show As Boolean
Private m_LabelText As String

Private m_BeginAtZero As Boolean
Private m_TicksMin As String
Private m_TicksMax As String

Private m_TicksNeeded As Boolean


' ### Initializer

Friend Sub Init(ByVal Axis As chCartesianAxis)
    m_Axis = Axis
End Sub


' ### Properties

Public Property Let Show(ByVal This_ShowAxis As Boolean)
    m_Show = This_ShowAxis
End Property
Public Property Get Show() As Boolean
    Show = m_Show
End Property

Public Property Let LabelText(ByVal This_LabelText As String)
    m_LabelText = This_LabelText
End Property
Public Property Get LabelText() As String
    LabelText = m_LabelText
End Property

Public Property Let BeginAtZero(ByVal This_BeginAtZero As Boolean)
    m_BeginAtZero = This_BeginAtZero
    If This_BeginAtZero = True Then
        m_TicksNeeded = True
    End If
End Property
Public Property Get BeginAtZero() As Boolean
    BeginAtZero = m_BeginAtZero
End Property

Public Property Let TicksMin(ByVal This_TicksMin As String)
    m_TicksMin = This_TicksMin
    If Len(Nz(This_TicksMin, vbNullString)) > 0 Then
        m_TicksNeeded = True
    End If
End Property
Public Property Get TicksMin() As String
    TicksMin = m_TicksMin
End Property

Public Property Let TicksMax(ByVal This_TicksMax As String)
    m_TicksMax = This_TicksMax
    If Len(Nz(This_TicksMax, vbNullString)) > 0 Then
        m_TicksNeeded = True
    End If
End Property
Public Property Get TicksMax() As String
    TicksMax = m_TicksMax
End Property


' ### Generator Function
Friend Function getAxis(ByVal StackedChart As Boolean)

    Dim strText As String

    strText = strText & "                ï¿½Axes: [{" & vbCrLf
    strText = strText & "                    id: 'first-ï¿½-Axis'," & vbCrLf
    
    If StackedChart = True Then
        strText = strText & "                    stacked: true," & vbCrLf
    End If
    
    If Me.Show = False Then
        strText = strText & "                    display: false" & vbCrLf
    Else
        strText = strText & "                    display: true," & vbCrLf
        If Len(Me.LabelText) > 0 Then
            strText = strText & "                    scaleLabel: {" & vbCrLf
            strText = strText & "                       display: true," & vbCrLf
            strText = strText & "                       labelString: '" & Me.LabelText & "'" & vbCrLf
            strText = strText & "                    }," & vbCrLf
        End If
        If m_TicksNeeded = True Then
            strText = strText & "                    ticks: {" & vbCrLf
        End If
        If Me.BeginAtZero = True Then
            strText = strText & "                       beginAtZero: true," & vbCrLf
        End If
        If Len(Me.TicksMin) > 0 Then
            If IsNumeric(Me.TicksMin) = True Then
                strText = strText & "                       min: " & Str$(Me.TicksMin) & " ," & vbCrLf
            Else
                strText = strText & "                       min: '" & Me.TicksMin & "' ," & vbCrLf
            End If
        End If
        If Len(Me.TicksMax) > 0 Then
            If IsNumeric(Me.TicksMax) = True Then
                strText = strText & "                       max: " & Str$(Me.TicksMax) & " ," & vbCrLf
            Else
                strText = strText & "                       max: '" & Me.TicksMax & "' ," & vbCrLf
            End If
        End If
        strText = BAC.Helper.CutTrailingComma(strText)
        If m_TicksNeeded = True Then
            strText = strText & "                    }," & vbCrLf
        End If
    End If
    strText = BAC.Helper.CutTrailingComma(strText)
    strText = strText & "                    }]," & vbCrLf
    
    If m_Axis = apXAxis Then
        strText = Replace(strText, "ï¿½", "x")
    ElseIf m_Axis = apYAxis Then
        strText = Replace(strText, "ï¿½", "y")
    End If
    getAxis = strText
    
End Function


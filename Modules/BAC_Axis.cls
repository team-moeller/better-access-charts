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
'# Copyright (c) 2020 - 2022 Thomas Moeller, supported by K.D.Gundermann                   #
'# MIT License  => https://github.com/team-moeller/better-access-charts/blob/main/LICENSE  #
'# Version 2.25.03  published: 16.11.2022                                                  #
'###########################################################################################

Option Compare Database
Option Explicit


'### Enums

Public Enum chCartesianAxis
    apXAxis = 1
    apYAxis = 2
End Enum


'### Members

Private m_Axis As chCartesianAxis
Private m_Show As Boolean
Private m_LabelText As String
Private m_BeginAtZero As Boolean
Private m_TicksMin As String
Private m_TicksMax As String

Private m_TicksNeeded As Boolean


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

Public Property Let beginAtZero(ByVal This_BeginAtZero As Boolean)
    m_BeginAtZero = This_BeginAtZero
    If This_BeginAtZero = True Then
        m_TicksNeeded = True
    End If
End Property
Public Property Get beginAtZero() As Boolean
    beginAtZero = m_BeginAtZero
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


' ### Initializer

Friend Sub Init(ByVal Axis As chCartesianAxis)
    m_Axis = Axis
End Sub


' ### Generator Functions

Friend Function getCartesianAxis(ByVal StackedChart As Boolean) As String

    With New BAC_StringBuilder
        .AddLine "                #Axes: [{"
        .AddLine "                    id: 'first-#-Axis',"
    
        If StackedChart = True Then
            .AddLine "                    stacked: true,"
        End If
    
        If Me.Show = False Then
            .AddLine "                    display: false"
        Else
            .AddLine "                    display: true,"
            If Len(Me.LabelText) > 0 Then
                .AddLine "                    scaleLabel: {"
                .AddLine "                       display: true,"
                .AddLine "                       labelString: '" & Me.LabelText & "'"
                .AddLine "                    },"
            End If
            If m_TicksNeeded = True Then
                .AddLine "                    ticks: {"
            End If
            If Me.beginAtZero = True Then
                .AddLine "                       beginAtZero: true,"
            End If
            If Len(Me.TicksMin) > 0 Then
                If IsNumeric(Me.TicksMin) = True Then
                    .AddLine "                       min: " & Str$(Me.TicksMin) & " ,"
                Else
                    .AddLine "                       min: '" & Me.TicksMin & "' ,"
                End If
            End If
            If Len(Me.TicksMax) > 0 Then
                If IsNumeric(Me.TicksMax) = True Then
                    .AddLine "                       max: " & Str$(Me.TicksMax) & " ,"
                Else
                    .AddLine "                       max: '" & Me.TicksMax & "' ,"
                End If
            End If
            .CutTrailingComma
            If m_TicksNeeded = True Then
                .AddLine "                    },"
            End If
        End If
        .CutTrailingComma
        .AddLine "                    }],"
    
        If m_Axis = apXAxis Then
            getCartesianAxis = Replace(.ToString, "#", "x")
        ElseIf m_Axis = apYAxis Then
            getCartesianAxis = Replace(.ToString, "#", "y")
        End If
    
    End With
    
End Function

Friend Function getRadialAxis() As String

    With New BAC_StringBuilder
        If Me.Show = False Then
            .AddLine "                display: false"
        Else
            .AddLine "                display: true,"
            If m_TicksNeeded = True Then
                .AddLine "                ticks: {"
            End If
            If Me.beginAtZero = True Then
                .AddLine "                    beginAtZero: true,"
            End If
            If Len(Me.TicksMin) > 0 Then
                If IsNumeric(Me.TicksMin) = True Then
                    .AddLine "                    min: " & Str$(Me.TicksMin) & " ,"
                End If
            End If
            If Len(Me.TicksMax) > 0 Then
                If IsNumeric(Me.TicksMax) = True Then
                    .AddLine "                    max: " & Str$(Me.TicksMax) & " ,"
                End If
            End If
            .CutTrailingComma
            If m_TicksNeeded = True Then
                .AddLine "                }"
            End If
        End If
        getRadialAxis = .ToString
    End With

End Function

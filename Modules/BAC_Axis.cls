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
'# Version 2.30.02  published: 30.12.2022                                                  #
'###########################################################################################

Option Compare Database
Option Explicit


'### Enums

Public Enum chCartesianAxis
    apXAxis = 1
    apYAxis = 2
End Enum


'### Members

Private m_Properties As BAC_Properties

Private m_Axis As chCartesianAxis


' ### Properties

Public Property Get Properties() As BAC_Properties
   Set Properties = m_Properties
End Property

Public Property Let Show(ByVal This_ShowAxis As Boolean)
    m_Properties.Item("Show").Value = This_ShowAxis
End Property
Public Property Get Show() As Boolean
    Show = Nz(m_Properties.Item("Show").Value, False)
End Property

Public Property Let LabelText(ByVal This_LabelText As String)
    m_Properties.Item("LabelText").Value = This_LabelText
End Property
Public Property Get LabelText() As String
    LabelText = Nz(m_Properties.Item("LabelText").Value, vbNullString)
End Property

Public Property Let BeginAtZero(ByVal This_BeginAtZero As Boolean)
    m_Properties.Item("BeginAtZero").Value = This_BeginAtZero
End Property
Public Property Get BeginAtZero() As Boolean
    BeginAtZero = Nz(m_Properties.Item("BeginAtZero").Value, False)
End Property

Public Property Let TicksMin(ByVal This_TicksMin As String)
    m_Properties.Item("TicksMin").Value = This_TicksMin
End Property
Public Property Get TicksMin() As String
    TicksMin = Nz(m_Properties.Item("TicksMin").Value, vbNullString)
End Property

Public Property Let TicksMax(ByVal This_TicksMax As String)
    m_Properties.Item("TicksMax").Value = This_TicksMax
End Property
Public Property Get TicksMax() As String
    TicksMax = Nz(m_Properties.Item("TicksMax").Value, vbNullString)
End Property

Private Property Get TicksNeeded() As Boolean
   If BeginAtZero Then
      TicksNeeded = True
   ElseIf Len(TicksMin) > 0 Then
      TicksNeeded = True
   ElseIf Len(TicksMax) > 0 Then
      TicksNeeded = True
   Else
      TicksNeeded = False
   End If
End Property


' ### Initializer
Private Sub Class_Initialize()
   InitProperties
End Sub

Private Sub InitProperties()

   Set m_Properties = New BAC_Properties
   
   m_Properties.NewProperty "Show"
   m_Properties.NewProperty "LabelText"
   m_Properties.NewProperty "BeginAtZero"
   m_Properties.NewProperty "TicksMin"
   m_Properties.NewProperty "TicksMax"
   
End Sub

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
            If TicksNeeded = True Then
                .AddLine "                    ticks: {"
            End If
            If Me.BeginAtZero = True Then
                .AddLine "                       beginAtZero: true,"
            End If
            If Len(Me.TicksMin) > 0 Then
                If IsNumeric(Me.TicksMin) = True Then
                    .AddLine "                       min: " & str$(Me.TicksMin) & " ,"
                Else
                    .AddLine "                       min: '" & Me.TicksMin & "' ,"
                End If
            End If
            If Len(Me.TicksMax) > 0 Then
                If IsNumeric(Me.TicksMax) = True Then
                    .AddLine "                       max: " & str$(Me.TicksMax) & " ,"
                Else
                    .AddLine "                       max: '" & Me.TicksMax & "' ,"
                End If
            End If
            .CutTrailingComma
            If TicksNeeded = True Then
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
            If TicksNeeded = True Then
                .AddLine "                ticks: {"
            End If
            If Me.BeginAtZero = True Then
                .AddLine "                    beginAtZero: true,"
            End If
            If Len(Me.TicksMin) > 0 Then
                If IsNumeric(Me.TicksMin) = True Then
                    .AddLine "                    min: " & str$(Me.TicksMin) & " ,"
                End If
            End If
            If Len(Me.TicksMax) > 0 Then
                If IsNumeric(Me.TicksMax) = True Then
                    .AddLine "                    max: " & str$(Me.TicksMax) & " ,"
                End If
            End If
            .CutTrailingComma
            If TicksNeeded = True Then
                .AddLine "                }"
            End If
        End If
        getRadialAxis = .ToString
    End With

End Function

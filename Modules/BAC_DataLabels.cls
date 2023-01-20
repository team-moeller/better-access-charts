VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_DataLabels"
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

'### Members
Private m_Properties As BAC_Properties
Private m_Font As BAC_Font

' ### Properties
Public Property Get Properties() As BAC_Properties
   Set Properties = m_Properties
End Property

Public Property Let Show(ByVal This_Show As Boolean)
    m_Properties.Item("Show").Value = This_Show
End Property
Public Property Get Show() As Boolean
    Show = Nz(m_Properties.Item("Show").Value, True)
End Property

Public Property Let Anchor(ByVal This_Anchor As chDataLabelAnchor)
    m_Properties.Item("Anchor").Value = This_Anchor
End Property
Public Property Get Anchor() As chDataLabelAnchor
    Anchor = Nz(m_Properties.Item("Anchor").Value, 0)
End Property

Public Property Get Font() As BAC_Font
    Set Font = m_Font
End Property

Private Property Get SectionIsNeeded() As Boolean
   SectionIsNeeded = (Anchor > 0)
End Property

' ### Instanzing

Private Sub Class_Initialize()

    Set m_Font = New BAC_Font
    InitProperties
    
End Sub

Private Sub InitProperties()

   Set m_Properties = New BAC_Properties
   
   Set m_Properties.NewProperty("Font").Properties = m_Font.Properties
   m_Properties.NewProperty "Show"
   m_Properties.NewProperty "Anchor"
   
End Sub


' ### Generator Function

Friend Function getDataLabel() As String

    If Not Me.Show Then Exit Function
    If Not SectionIsNeeded And Not Me.Font.Color <> vbNullString And Not Me.Font.SectionIsNeeded("DataLabel") = True Then
      Exit Function
    End If
    
    With New BAC_StringBuilder
        .AddLine "                datalabels: {"
        If Me.Anchor > 0 Then
            .AddLine "                    anchor: '" & getAnchorString(Me.Anchor) & "',"
        End If
        If Me.Font.Color <> vbNullString Then
            .AddLine "                    color: '" & Me.Font.Color & "',"
        End If
        If Me.Font.SectionIsNeeded("DataLabel") Then
            .AddLine "                    font: {"
            If Me.Font.Family <> vbNullString Then
                .AddLine "                        family: '" & Me.Font.Family & "',"
            End If
            If Me.Font.Size > 0 Then
                .AddLine "                        size: " & Me.Font.Size & ","
            End If
            If Me.Font.IsBold = True Then
                .AddLine "                        weight: 'bold',"
            End If
            If Me.Font.IsItalic = True Then
                .AddLine "                        style: 'italic',"
            End If
            .CutTrailingComma
            .AddLine "                        },"
        End If
        .CutTrailingComma
        .AddLine "                    },"
        
        getDataLabel = .ToString
    End With
    
End Function

Private Function getAnchorString(ByVal ThisAnchor As chDataLabelAnchor) As String

    Select Case ThisAnchor
        Case chDataLabelAnchor.anStart
            getAnchorString = "start"
        Case chDataLabelAnchor.anCenter
            getAnchorString = "center"
        Case chDataLabelAnchor.anEnd
            getAnchorString = "end"
        Case Else
            getAnchorString = "center"
    End Select
    
End Function

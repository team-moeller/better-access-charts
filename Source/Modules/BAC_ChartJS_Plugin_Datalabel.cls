VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_ChartJS_Plugin_Datalabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Source As chScriptSource
Private m_CDN As String
Private m_PathFilename As String
Private m_Show As Boolean
Private m_Anchor As chDataLabelAnchor
Private m_FontColor As String
Private m_FontFamily As String
Private m_FontSize As Integer
Private m_FontIsItalic As Boolean
Private m_FontIsBold As Boolean

Private m_Needed As Boolean
Private m_FontNeeded As Boolean

Public Property Let Source(ByVal This_Source As chScriptSource)
    m_Source = This_Source
End Property
Public Property Get Source() As chScriptSource
    Source = m_Source
End Property

Public Property Let CDN(ByVal This_CDN As String)
    m_CDN = This_CDN
End Property
Public Property Get CDN() As String
    CDN = m_CDN
End Property

Public Property Let PathFilename(ByVal This_PathFilename As String)
    m_PathFilename = This_PathFilename
End Property
Public Property Get PathFilename() As String
    PathFilename = m_PathFilename
End Property

Public Property Let Show(ByVal This_Show As Boolean)
    m_Show = This_Show
End Property
Public Property Get Show() As Boolean
    Show = m_Show
End Property

Public Property Let Anchor(ByVal This_Anchor As chDataLabelAnchor)
    m_Anchor = This_Anchor
    If This_Anchor > 0 Then
        m_Needed = True
    End If
End Property
Public Property Get Anchor() As chDataLabelAnchor
    Anchor = m_Anchor
End Property

Public Property Let FontColor(ByVal This_FontColor As String)
    m_FontColor = This_FontColor
    If Len(Nz(This_FontColor, vbNullString)) > 0 Then
        m_Needed = True
    End If
End Property
Public Property Get FontColor() As String
    FontColor = m_FontColor
End Property

Public Property Let FontFamily(ByVal This_FontFamily As String)
    m_FontFamily = This_FontFamily
    If Len(Nz(This_FontFamily, vbNullString)) > 0 Then
        m_Needed = True
        m_FontNeeded = True
    End If
End Property
Public Property Get FontFamily() As String
    FontFamily = m_FontFamily
End Property

Public Property Let FontSize(ByVal This_FontSize As Integer)
    m_FontSize = This_FontSize
    If This_FontSize > 0 Then
        m_Needed = True
        m_FontNeeded = True
    End If
End Property
Public Property Get FontSize() As Integer
    FontSize = m_FontSize
End Property

Public Property Let FontIsItalic(ByVal This_FontIsItalic As Boolean)
    m_FontIsItalic = This_FontIsItalic
    If m_FontIsItalic = True Then
        m_Needed = True
        m_FontNeeded = True
    End If
End Property
Public Property Get FontIsItalic() As Boolean
    FontIsItalic = m_FontIsItalic
End Property

Public Property Let FontIsBold(ByVal This_FontIsBold As Boolean)
    m_FontIsBold = This_FontIsBold
    If m_FontIsBold = True Then
        m_Needed = True
        m_FontNeeded = True
    End If
End Property
Public Property Get FontIsBold() As Boolean
    FontIsBold = m_FontIsBold
End Property

' ### Generator Function

Friend Function getDataLabelLocation() As String

    'Declarations
    Dim strText As String
    
    Select Case Me.Source
        Case chScriptSource.CDN
            strText = """" & Me.CDN & """"
        Case chScriptSource.LocalFile
            strText = """" & Me.PathFilename & """"
        Case Else
            'Do nothing
    End Select
            
    getDataLabelLocation = strText

End Function

Friend Function getDataLabel() As String

    'Declarations
    Dim strText As New BAC_StringBuilder
    If Not Me.Show Then Exit Function
    If Not m_Needed = True Then Exit Function
    
    With strText
        .AddLine "                datalabels: {"
        If Me.Anchor > 0 Then
            .AddLine "                    anchor: '" & getAnchorString(Me.Anchor) & "',"
        End If
        If Me.FontColor <> vbNullString Then
            .AddLine "                    color: '" & Me.FontColor & "',"
        End If
        If m_FontNeeded = True Then
            .AddLine "                    font: {"
            If Me.FontFamily <> vbNullString Then
                .AddLine "                        family: '" & Me.FontFamily & "',"
            End If
            If Me.FontSize > 0 Then
                .AddLine "                        size: " & Me.FontSize & ","
            End If
            If Me.FontIsBold = True Then
                .AddLine "                        weight: 'bold',"
            End If
            If Me.FontIsItalic = True Then
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


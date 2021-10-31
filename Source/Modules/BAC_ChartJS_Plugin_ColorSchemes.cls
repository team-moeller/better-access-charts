VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_ChartJS_Plugin_ColorSchemes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###########################################################################################
'# Copyright (c) 2020, 2021 Thomas M�ller, supported by K.D.Gundermann                     #
'# MIT License  => https://github.com/team-moeller/better-access-charts/blob/main/LICENSE  #
'# Version 1.31.12  published: 29.10.2021                                                  #
'###########################################################################################

Option Compare Database
Option Explicit

Private m_ColorSchema As String
Private m_Source As chScriptSource
Private m_CDN As String
Private m_PathFilename As String


' ### Properties

Public Property Let ColorSchema(ByVal This_ColorSchema As String)
    m_ColorSchema = This_ColorSchema
End Property
Public Property Get ColorSchema() As String
    ColorSchema = m_ColorSchema
End Property

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


' ### Public Functions

Public Function SavePluginToDisk() As Boolean
    With BAC.Helper
        If .SaveFileToDisk("chartjs-plugin-colorschemes.min.js", CurrentProject.Path) = False Then
            SavePluginToDisk = False
        ElseIf .SaveFileToDisk("colorschemes.brewer.js", CurrentProject.Path) = False Then
            SavePluginToDisk = False
        ElseIf .SaveFileToDisk("colorschemes.office.js", CurrentProject.Path) = False Then
            SavePluginToDisk = False
        ElseIf .SaveFileToDisk("colorschemes.tableau.js", CurrentProject.Path) = False Then
            SavePluginToDisk = False
        Else
            SavePluginToDisk = True
        End If
    End With
End Function


' ### Generator Function

Friend Function getLocation() As String

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
            
    getLocation = strText

End Function

Friend Function getColorSchema() As String

    'Declarations
    Dim strText As New BAC_StringBuilder
    With strText
        .AddLine "                colorschemes: {"
        .AddLine "                    scheme: '" & ColorSchema & "'"
        .AddLine "                    },"
        
        getColorSchema = .ToString
    End With
End Function

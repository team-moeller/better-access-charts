VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Better_Access_Chart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###########################################################################################
'# Copyright (c) 2020, 2021 Thomas M�ller                                                  #
'# MIT License  => https://github.com/team-moeller/better-access-charts/blob/main/LICENSE  #
'# Version 1.31.12  published: 24.10.2021                                                  #
'###########################################################################################

Option Compare Database
Option Explicit


'### Enums


Public Enum chChartType
    line = 1
    Bar = 2
    HorizontalBar = 3
    Radar = 4
    Doughnut = 5
    Pie = 6
    PolarArea = 7
    Bubble = 8
    Scatter = 9
End Enum

Public Enum chDataSourceType
    dstDemo
    dstRecordset
    dstTableName
    dstQueryName
    dstSQLStament
End Enum

Public Enum chPosition
    posTop = 1
    posLeft = 2
    posBottom = 3
    posRight = 4
End Enum

Public Enum chScriptSource
    CDN = 1
    LocalFile = 2
End Enum

Public Enum chDataLabelAnchor
    anStart = 1
    anCenter = 2
    anEnd = 3
End Enum


'### Constants


'### Members

Private WithEvents m_Control As WebBrowserControl
Attribute m_Control.VB_VarHelpID = -1
Private m_IE As Object 'InternetExplorer

Private m_ChartType As chChartType

Private m_ChartJsSource As chScriptSource
Private m_ChartJsCDN As String
Private m_ChartJsPathFilename As String
Private m_StackedChart As Boolean
Private m_ColorSchemaSource As chScriptSource
Private m_ColorSchemaCDN As String
Private m_ColorSchemaPathFilename As String
Private m_ColorSchema As String
Private m_AreaFilled As Boolean
Private m_ChartBackgroundColor As String
Private m_DefaultFontColor As String
Private m_DefaultFontFamily As String
Private m_DefaultFontSize As Integer
Private m_ShowDownloadLink As Boolean
Private m_DataLabelSource As chScriptSource
Private m_DataLabelCDN As String
Private m_DataLabelPathFilename As String
Private m_ShowDataLabel As Boolean
Private m_DataLabelAnchor As chDataLabelAnchor
Private m_DataLabelFontColor As String
Private m_DataLabelFontFamily As String
Private m_DataLabelFontSize As Integer
Private m_DataLabelFontIsItalic As Boolean
Private m_DataLabelFontIsBold As Boolean
Private m_SaveHtmlFileForDebugging As Boolean

Private m_IsSingleBarChart As Boolean
Private m_DataLabelNeeded As Boolean
Private m_DataLabelFontNeeded As Boolean
    
Private m_DataSource As BAC_DataSource
Private m_XAxis     As BAC_Axis
Private m_YAxis     As BAC_Axis
Private m_Legend    As BAC_Legend
Private m_Title     As BAC_Title

    
'### Initializer

Public Sub Init(Optional ByVal TheControl As WebBrowserControl, _
                Optional ByVal ChartType As chChartType = chChartType.Bar, _
                Optional ByVal DataSourceType As chDataSourceType, _
                Optional ByVal DataSource As Variant, _
                Optional ByVal DataFieldNames As Variant, _
                Optional ByVal LabelFieldName As String)
  
  Set Me.Control = TheControl
  Me.ChartType = ChartType
  Me.DataSource.Init DataSourceType, DataSource, DataFieldNames, LabelFieldName
End Sub
    
'### Methods

Public Sub UseIE()
    Set m_IE = CreateObject("InternetExplorer.Application")
    m_IE.Navigate "about:"
    m_IE.Visible = True
    Do While m_IE.ReadyState <> 4: DoEvents: Loop
End Sub
    
'### Properties

Public Property Set Control(ByVal This_Control As WebBrowserControl)
    Set m_Control = This_Control
End Property
Public Property Get Control() As WebBrowserControl
    Set Control = m_Control
End Property

Public Property Let ChartType(ByVal This_ChartType As chChartType)
    m_ChartType = This_ChartType
    Select Case This_ChartType
        Case chChartType.Bar, chChartType.HorizontalBar, chChartType.line, chChartType.Bubble, chChartType.Scatter
            YAxis.Show = True
            XAxis.Show = True
        Case chChartType.Doughnut, chChartType.Pie, chChartType.PolarArea, chChartType.Radar
            YAxis.Show = False
            XAxis.Show = False
        Case Else
            'Do Nothing
    End Select
End Property
Public Property Get ChartType() As chChartType
    ChartType = m_ChartType
End Property


Public Property Let ChartJsSource(ByVal This_ChartJsSource As chScriptSource)
    m_ChartJsSource = This_ChartJsSource
End Property
Public Property Get ChartJsSource() As chScriptSource
    ChartJsSource = m_ChartJsSource
End Property

Public Property Let ChartJsCDN(ByVal This_ChartJsCDN As String)
    m_ChartJsCDN = This_ChartJsCDN
End Property
Public Property Get ChartJsCDN() As String
    ChartJsCDN = m_ChartJsCDN
End Property

Public Property Let ChartJsPathFilename(ByVal This_ChartJsPathFilename As String)
    m_ChartJsPathFilename = This_ChartJsPathFilename
End Property
Public Property Get ChartJsPathFilename() As String
    ChartJsPathFilename = m_ChartJsPathFilename
End Property

Public Property Let StackedChart(ByVal This_StackedChart As Boolean)
    m_StackedChart = This_StackedChart
End Property
Public Property Get StackedChart() As Boolean
    Select Case ChartType
        Case chChartType.line, chChartType.Bar, chChartType.HorizontalBar
            StackedChart = m_StackedChart
        Case Else
            StackedChart = False
    End Select
End Property

Public Property Let ColorSchemaSource(ByVal This_ColorSchemaSource As chScriptSource)
    m_ColorSchemaSource = This_ColorSchemaSource
End Property
Public Property Get ColorSchemaSource() As chScriptSource
    ColorSchemaSource = m_ColorSchemaSource
End Property

Public Property Let ColorSchemaCDN(ByVal This_ColorSchemaCDN As String)
    m_ColorSchemaCDN = This_ColorSchemaCDN
End Property
Public Property Get ColorSchemaCDN() As String
    ColorSchemaCDN = m_ColorSchemaCDN
End Property

Public Property Let ColorSchemaPathFilename(ByVal This_ColorSchemaPathFilename As String)
    m_ColorSchemaPathFilename = This_ColorSchemaPathFilename
End Property
Public Property Get ColorSchemaPathFilename() As String
    ColorSchemaPathFilename = m_ColorSchemaPathFilename
End Property

Public Property Let ColorSchema(ByVal This_ColorSchema As String)
    m_ColorSchema = This_ColorSchema
End Property
Public Property Get ColorSchema() As String
    ColorSchema = m_ColorSchema
End Property

Public Property Let AreaFilled(ByVal This_AreaFilled As Boolean)
    m_AreaFilled = This_AreaFilled
End Property
Public Property Get AreaFilled() As Boolean
    AreaFilled = m_AreaFilled
End Property

Public Property Let ChartBackgroundColor(ByVal This_ChartBackgroundColor As String)
    m_ChartBackgroundColor = This_ChartBackgroundColor
End Property
Public Property Get ChartBackgroundColor() As String
    ChartBackgroundColor = m_ChartBackgroundColor
End Property

Public Property Let DefaultFontColor(ByVal This_DefaultFontColor As String)
    m_DefaultFontColor = This_DefaultFontColor
End Property
Public Property Get DefaultFontColor() As String
    DefaultFontColor = m_DefaultFontColor
End Property

Public Property Let DefaultFontFamily(ByVal This_DefaultFontFamily As String)
    m_DefaultFontFamily = This_DefaultFontFamily
End Property
Public Property Get DefaultFontFamily() As String
    DefaultFontFamily = m_DefaultFontFamily
End Property

Public Property Let DefaultFontSize(ByVal This_DefaultFontSize As Integer)
    m_DefaultFontSize = This_DefaultFontSize
End Property
Public Property Get DefaultFontSize() As Integer
    DefaultFontSize = m_DefaultFontSize
End Property

Public Property Let ShowDownloadLink(ByVal This_ShowDownloadLink As Boolean)
    m_ShowDownloadLink = This_ShowDownloadLink
End Property
Public Property Get ShowDownloadLink() As Boolean
    ShowDownloadLink = m_ShowDownloadLink
End Property

Public Property Let DataLabelSource(ByVal This_DataLabelSource As chScriptSource)
    m_DataLabelSource = This_DataLabelSource
End Property
Public Property Get DataLabelSource() As chScriptSource
    DataLabelSource = m_DataLabelSource
End Property

Public Property Let DataLabelCDN(ByVal This_DataLabelCDN As String)
    m_DataLabelCDN = This_DataLabelCDN
End Property
Public Property Get DataLabelCDN() As String
    DataLabelCDN = m_DataLabelCDN
End Property

Public Property Let DataLabelPathFilename(ByVal This_DataLabelPathFilename As String)
    m_DataLabelPathFilename = This_DataLabelPathFilename
End Property
Public Property Get DataLabelPathFilename() As String
    DataLabelPathFilename = m_DataLabelPathFilename
End Property

Public Property Let ShowDataLabel(ByVal This_ShowDataLabel As Boolean)
    m_ShowDataLabel = This_ShowDataLabel
End Property
Public Property Get ShowDataLabel() As Boolean
    ShowDataLabel = m_ShowDataLabel
End Property

Public Property Let DataLabelAnchor(ByVal This_DataLabelAnchor As chDataLabelAnchor)
    m_DataLabelAnchor = This_DataLabelAnchor
    If This_DataLabelAnchor > 0 Then
        m_DataLabelNeeded = True
    End If
End Property
Public Property Get DataLabelAnchor() As chDataLabelAnchor
    DataLabelAnchor = m_DataLabelAnchor
End Property

Public Property Let DataLabelFontColor(ByVal This_DataLabelFontColor As String)
    m_DataLabelFontColor = This_DataLabelFontColor
    If Len(Nz(This_DataLabelFontColor, vbNullString)) > 0 Then
        m_DataLabelNeeded = True
    End If
End Property
Public Property Get DataLabelFontColor() As String
    DataLabelFontColor = m_DataLabelFontColor
End Property

Public Property Let DataLabelFontFamily(ByVal This_DataLabelFontFamily As String)
    m_DataLabelFontFamily = This_DataLabelFontFamily
    If Len(Nz(This_DataLabelFontFamily, vbNullString)) > 0 Then
        m_DataLabelNeeded = True
        m_DataLabelFontNeeded = True
    End If
End Property
Public Property Get DataLabelFontFamily() As String
    DataLabelFontFamily = m_DataLabelFontFamily
End Property

Public Property Let DataLabelFontSize(ByVal This_DataLabelFontSize As Integer)
    m_DataLabelFontSize = This_DataLabelFontSize
    If This_DataLabelFontSize > 0 Then
        m_DataLabelNeeded = True
        m_DataLabelFontNeeded = True
    End If
End Property
Public Property Get DataLabelFontSize() As Integer
    DataLabelFontSize = m_DataLabelFontSize
End Property

Public Property Let DataLabelFontIsItalic(ByVal This_DataLabelFontIsItalic As Boolean)
    m_DataLabelFontIsItalic = This_DataLabelFontIsItalic
    If m_DataLabelFontIsItalic = True Then
        m_DataLabelNeeded = True
        m_DataLabelFontNeeded = True
    End If
End Property
Public Property Get DataLabelFontIsItalic() As Boolean
    DataLabelFontIsItalic = m_DataLabelFontIsItalic
End Property

Public Property Let DataLabelFontIsBold(ByVal This_DataLabelFontIsBold As Boolean)
    m_DataLabelFontIsBold = This_DataLabelFontIsBold
    If m_DataLabelFontIsBold = True Then
        m_DataLabelNeeded = True
        m_DataLabelFontNeeded = True
    End If
End Property
Public Property Get DataLabelFontIsBold() As Boolean
    DataLabelFontIsBold = m_DataLabelFontIsBold
End Property

Public Property Let SaveHtmlFileForDebugging(ByVal This_SaveHtmlFileForDebugging As Boolean)
    m_SaveHtmlFileForDebugging = This_SaveHtmlFileForDebugging
End Property
Public Property Get SaveHtmlFileForDebugging() As Boolean
    SaveHtmlFileForDebugging = m_SaveHtmlFileForDebugging
End Property

'### Object Properties

Public Property Get DataSource() As BAC_DataSource
    Set DataSource = m_DataSource
End Property

Public Property Get Legend() As BAC_Legend
    Set Legend = m_Legend
End Property

Public Property Get Title() As BAC_Title
    Set Title = m_Title
End Property

Public Property Get XAxis() As BAC_Axis
    Set XAxis = m_XAxis
End Property

Public Property Get YAxis() As BAC_Axis
    Set YAxis = m_YAxis
End Property


'### Instancing

Private Sub Class_Initialize()
    
    'create Sub Classes
    Set m_DataSource = New BAC_DataSource
    Set m_Legend = New BAC_Legend
    Set m_Title = New BAC_Title
    Set m_XAxis = New BAC_Axis
    m_XAxis.Init apXAxis
    Set m_YAxis = New BAC_Axis
    m_YAxis.Init apYAxis
    
    'Set default values
    ChartType = chChartType.Bar
    StackedChart = False
    'Title
    Title.Show = False
    Title.Position = posTop
    Title.Text = "My first Better Access Chart"
    'Legend
    Legend.Show = False
    Legend.Position = posRight
    'Axis
    YAxis.Show = True
    XAxis.Show = True
    'chart.js
    ChartJsSource = chScriptSource.CDN
    ChartJsCDN = "https://cdnjs.cloudflare.com/ajax/libs/Chart.js/2.9.4/Chart.min.js"
    ChartJsPathFilename = CurrentProject.Path & "\Chart.min.js"
    'ColorSchemes Plug-In
    ColorSchemaSource = chScriptSource.CDN
    ColorSchemaCDN = "https://cdn.jsdelivr.net/npm/chartjs-plugin-colorschemes"
    ColorSchemaPathFilename = CurrentProject.Path & "\chartjs-plugin-colorschemes.min.js"
    'ColorSchema
    ColorSchema = "brewer.Paired12"
    'DataLabel Plug-In
    DataLabelSource = chScriptSource.CDN
    DataLabelCDN = "https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@0.7.0"
    DataLabelPathFilename = CurrentProject.Path & "\chartjs-plugin-datalabels.min.js"
End Sub

Private Sub Class_Terminate()
    ' Do Nothing
End Sub


'### Public Methods

Public Sub DrawChart()
    If Not m_IE Is Nothing Then
        With m_IE.Document
            .Open
            .Write getHtmlString
            .Close
        End With
    End If
    With Control.Object.Document
        .Open
        .Write getHtmlString
        .Close
    End With
    
End Sub


'### Private Methods

Private Function getHtmlString() As String

    'Declarations
    Dim strHtml As String
    
    strHtml = strHtml & getHeader
    strHtml = strHtml & getChartType
    strHtml = strHtml & getChartData
    strHtml = strHtml & getChartOptions
    strHtml = strHtml & getFooter
 
    If m_SaveHtmlFileForDebugging = True Then
        Call saveHtmlFile(strHtml)
    End If
 
    getHtmlString = strHtml
  
End Function

Private Function getHeader() As String

    'Declarations
    Dim strText As String
    
    strText = strText & "<!DOCTYPE html>" & vbCrLf
    strText = strText & "<!-- saved from url=(0014)about:internet -->" & vbCrLf
    strText = strText & "<!-- saved from url=(0016)http://localhost -->" & vbCrLf
    strText = strText & "<html>" & vbCrLf
    strText = strText & "<head>" & vbCrLf
    strText = strText & "    <meta charset=""utf-8"">" & vbCrLf
    strText = strText & "    <meta http-equiv=""X-UA-Compatible"" content=""IE=Edge""/>" & vbCrLf
    strText = strText & "    <title>Chart</title>" & vbCrLf
    strText = strText & "    <style>body{margin:0;padding:0;}</style>" & vbCrLf
    strText = strText & "    <script src=" & getChartJsLocation & "></script>" & vbCrLf
    strText = strText & "    <script src=" & getColorSchemaLocation & "></script>" & vbCrLf
    If ShowDataLabel = True Then
        strText = strText & "    <script src=" & getDataLabelLocation & "></script>" & vbCrLf
    End If
    strText = strText & "</head>" & vbCrLf
    strText = strText & "<body>" & vbCrLf
    strText = strText & "<div id=""container"" style=""width: 100%;"">" & vbCrLf
    If Len(ChartBackgroundColor) > 0 Then
        strText = strText & "    <canvas id=""myChart"" style=""background-color: " & ChartBackgroundColor & """ ></canvas>" & vbCrLf
    Else
        strText = strText & "    <canvas id=""myChart""></canvas>" & vbCrLf
    End If
    strText = strText & "</div>" & vbCrLf
    If ShowDownloadLink = True Then
        strText = strText & "<a href=""#"" id=""download"" onclick=""download()"">Download</a>"
    End If
    strText = strText & "<script>" & vbCrLf
    If Len(DefaultFontColor) > 0 Then
        strText = strText & "    Chart.defaults.global.defaultFontColor = '" & DefaultFontColor & "';" & vbCrLf
    End If
    If Len(DefaultFontFamily) > 0 Then
        strText = strText & "    Chart.defaults.global.defaultFontFamily = '" & DefaultFontFamily & "';" & vbCrLf
    End If
    If DefaultFontSize > 0 Then
        strText = strText & "    Chart.defaults.global.defaultFontSize = " & DefaultFontSize & ";" & vbCrLf
    End If
    strText = strText & "    var ctx = document.getElementById('myChart').getContext('2d');" & vbCrLf
    strText = strText & "    var myChart = new Chart(ctx, {" & vbCrLf
    
    getHeader = strText
    
End Function

Private Function getChartType() As String

    'Declarations
    Dim strText As String
    Dim strType As String
    
    Select Case ChartType
        Case chChartType.line
            strType = "line"
        Case chChartType.Bar
            strType = "bar"
        Case chChartType.HorizontalBar
            strType = "horizontalBar"
        Case chChartType.Radar
            strType = "radar"
        Case chChartType.Doughnut
            strType = "doughnut"
        Case chChartType.Pie
            strType = "pie"
        Case chChartType.PolarArea
            strType = "polarArea"
        Case chChartType.Bubble
            strType = "bubble"
        Case chChartType.Scatter
            strType = "scatter"
        Case Else
            'Do Nothing
    End Select
    
    strText = strText & "        type: '" & strType & "'," & vbCrLf
    
    getChartType = strText

End Function

Private Function getChartData() As String

    'Declarations
    Dim strText As String
    Dim dataSet As BAC_DataSet
    Dim i As Integer

    Set dataSet = Me.DataSource.getData(Me.ChartType)

    strText = strText & "        data: {" & vbCrLf
    If ChartType <> chChartType.Bubble Then
        If ChartType <> chChartType.Scatter Then
            strText = strText & "            labels: [" & dataSet.Label(0) & "]," & vbCrLf
        End If
    End If
    
    strText = strText & "            datasets: [" & vbCrLf
    For i = 0 To dataSet.DataCount - 1
        Select Case ChartType
            Case chChartType.Bubble, chChartType.Scatter
                strText = strText & "               {label: '" & dataSet.Label(i) & "' ," & vbCrLf
            Case Else
                strText = strText & "               {label: '" & dataSet.DatasetName(i) & "' ," & vbCrLf
        End Select
        Select Case ChartType
            Case chChartType.line, chChartType.Radar
                If AreaFilled = False Then
                    strText = strText & "                fill: false ," & vbCrLf
                End If
            Case Else
                'Do Nothing
        End Select
                
        strText = strText & "                data: [" & dataSet.Data(i) & "]," & vbCrLf
        If m_IsSingleBarChart = True Then
            strText = strText & "                backgroundColor: Chart['colorschemes']." & ColorSchema & "," & vbCrLf
        End If
        strText = strText & "                borderWidth: 1}," & vbCrLf
    Next i
    strText = BAC.Helper.CutTrailingComma(strText)
    strText = strText & "             ]" & vbCrLf
    strText = strText & "        }," & vbCrLf
    
    getChartData = strText

End Function

Private Function getChartOptions() As String

    'Declarations
    Dim strText As String
    
    strText = strText & "        options: {" & vbCrLf
    strText = strText & Title.getTitle
    strText = strText & Legend.getLegend
    strText = strText & getScales
    strText = strText & getPlugins
    strText = BAC.Helper.CutTrailingComma(strText)
    strText = strText & "        }" & vbCrLf
    
    getChartOptions = strText

End Function

Private Function getFooter() As String

    'Declarations
    Dim strText As String
    
    strText = strText & "    });" & vbCrLf
    If ShowDownloadLink = True Then
        strText = strText & "    function download() {" & vbCrLf
        strText = strText & "        var myChart = document.getElementById('myChart');" & vbCrLf
        strText = strText & "        var blob = myChart.msToBlob()" & vbCrLf
        strText = strText & "        window.navigator.msSaveBlob(blob, 'Better-Access-Charts.png');" & vbCrLf
        strText = strText & "    }" & vbCrLf
    End If
    strText = strText & "</script>" & vbCrLf
    strText = strText & "</body>" & vbCrLf
    strText = strText & "</html>"
    
    getFooter = strText
    
End Function

Private Function getChartJsLocation() As String

    'Declarations
    Dim strText As String
    
    Select Case ChartJsSource
        Case chScriptSource.CDN
            strText = """" & ChartJsCDN & """"
        Case chScriptSource.LocalFile
            strText = """" & ChartJsPathFilename & """"
        Case Else
            'Do nothing
    End Select
            
    getChartJsLocation = strText

End Function

Private Function getColorSchemaLocation() As String

    'Declarations
    Dim strText As String
    
    Select Case ColorSchemaSource
        Case chScriptSource.CDN
            strText = """" & ColorSchemaCDN & """"
        Case chScriptSource.LocalFile
            strText = """" & ColorSchemaPathFilename & """"
        Case Else
            'Do nothing
    End Select
            
    getColorSchemaLocation = strText

End Function

Private Function getDataLabelLocation() As String

    'Declarations
    Dim strText As String
    
    Select Case DataLabelSource
        Case chScriptSource.CDN
            strText = """" & DataLabelCDN & """"
        Case chScriptSource.LocalFile
            strText = """" & DataLabelPathFilename & """"
        Case Else
            'Do nothing
    End Select
            
    getDataLabelLocation = strText

End Function


Private Sub DetermineSingleBarChart()

    Select Case m_ChartType
        Case chChartType.Bar, chChartType.HorizontalBar
            If DataSource.IsSingleDataset = True Then
                m_IsSingleBarChart = True
            Else
                m_IsSingleBarChart = False
            End If
        Case Else
            m_IsSingleBarChart = False
    End Select

End Sub



Private Function getScales() As String

    'Declarations
    Dim strText As String

    strText = strText & "            scales: {" & vbCrLf
    strText = strText & Me.YAxis.getAxis(Me.StackedChart)
    strText = strText & Me.XAxis.getAxis(Me.StackedChart)
    strText = BAC.Helper.CutTrailingComma(strText)
    strText = strText & "            }," & vbCrLf
    
    getScales = strText
    
End Function

Private Function getPlugins() As String

    'Declarations
    Dim strText As String

    strText = strText & "            plugins: {" & vbCrLf
    strText = strText & getColorSchema
    strText = strText & getDataLabel
    strText = BAC.Helper.CutTrailingComma(strText)

    strText = strText & "                }," & vbCrLf
    
    getPlugins = strText

End Function

Private Function getColorSchema() As String

    'Declarations
    Dim strText As String

    strText = strText & "                colorschemes: {" & vbCrLf
    strText = strText & "                    scheme: '" & ColorSchema & "'" & vbCrLf
    strText = strText & "                    }," & vbCrLf
    
    getColorSchema = strText

End Function

Private Function getDataLabel() As String

    'Declarations
    Dim strText As String

    If ShowDataLabel = True Then
        If m_DataLabelNeeded = True Then
            strText = strText & "                datalabels: {" & vbCrLf
            If DataLabelAnchor > 0 Then
                strText = strText & "                    anchor: '" & getDataLabelAnchorString(DataLabelAnchor) & "'," & vbCrLf
            End If
            If Len(Nz(DataLabelFontColor, vbNullString)) > 0 Then
                strText = strText & "                    color: '" & DataLabelFontColor & "'," & vbCrLf
            End If
            If m_DataLabelFontNeeded = True Then
                strText = strText & "                    font: {" & vbCrLf
                If Len(Nz(DataLabelFontFamily, vbNullString)) > 0 Then
                    strText = strText & "                        family: '" & DataLabelFontFamily & "'," & vbCrLf
                End If
                If DataLabelFontSize > 0 Then
                    strText = strText & "                        size: " & DataLabelFontSize & "," & vbCrLf
                End If
                If DataLabelFontIsBold = True Then
                    strText = strText & "                        weight: 'bold'," & vbCrLf
                End If
                If DataLabelFontIsItalic = True Then
                    strText = strText & "                        style: 'italic'," & vbCrLf
                End If
                strText = BAC.Helper.CutTrailingComma(strText)
                strText = strText & "                        }," & vbCrLf
            End If
            strText = BAC.Helper.CutTrailingComma(strText)
            strText = strText & "                    }," & vbCrLf
        End If
    End If
    
    getDataLabel = strText

End Function


Private Function getDataLabelAnchorString(ByVal ThisAnchor As chDataLabelAnchor) As String

    Select Case ThisAnchor
        Case chDataLabelAnchor.anStart
            getDataLabelAnchorString = "start"
        Case chDataLabelAnchor.anCenter
            getDataLabelAnchorString = "center"
        Case chDataLabelAnchor.anEnd
            getDataLabelAnchorString = "end"
        Case Else
            getDataLabelAnchorString = "center"
    End Select
    
End Function


'### Private Helper Methods


Private Sub saveHtmlFile(ByVal strContent As String)

    'Declarations
    Dim strPath As String
    Dim strFileName As String
    Dim f As Integer
    Const HtmlFileName As String = "BetterAccessChart_{}.html"
    
    strPath = CurrentProject.Path & "\"
    strFileName = Replace(HtmlFileName, "{}", m_Control.Name)
    
    f = FreeFile()
    Open strPath & strFileName For Output As f
    Print #f, strContent;
    Close f
    
End Sub

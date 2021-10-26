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
'# Copyright (c) 2020, 2021 Thomas Möller                                                  #
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

Private m_StackedChart As Boolean
Private m_AreaFilled As Boolean
Private m_ChartBackgroundColor As String
Private m_DefaultFontColor As String
Private m_DefaultFontFamily As String
Private m_DefaultFontSize As Integer
Private m_ShowDownloadLink As Boolean
Private m_SaveHtmlFileForDebugging As Boolean

Private m_IsSingleBarChart As Boolean
    
Private m_ChartJS     As BAC_ChartJS
Private m_DataSource  As BAC_DataSource
Private m_XAxis       As BAC_Axis
Private m_YAxis       As BAC_Axis
Private m_Legend      As BAC_Legend
Private m_Title       As BAC_Title

    
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

Public Property Let SaveHtmlFileForDebugging(ByVal This_SaveHtmlFileForDebugging As Boolean)
    m_SaveHtmlFileForDebugging = This_SaveHtmlFileForDebugging
End Property
Public Property Get SaveHtmlFileForDebugging() As Boolean
    SaveHtmlFileForDebugging = m_SaveHtmlFileForDebugging
End Property

'### Object Properties

Public Property Get ChartJS() As BAC_ChartJS
  Set ChartJS = m_ChartJS
End Property

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
    Set m_ChartJS = New BAC_ChartJS
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
    ChartJS.Source = chScriptSource.CDN
    ChartJS.CDN = "https://cdnjs.cloudflare.com/ajax/libs/Chart.js/2.9.4/Chart.min.js"
    ChartJS.PathFilename = CurrentProject.Path & "\Chart.min.js"
    'ColorSchemes Plug-In
    With ChartJS.Plugin_ColorSchema
      .Source = chScriptSource.CDN
      .CDN = "https://cdn.jsdelivr.net/npm/chartjs-plugin-colorschemes"
      .PathFilename = CurrentProject.Path & "\chartjs-plugin-colorschemes.min.js"
      'ColorSchema
      .ColorSchema = "brewer.Paired12"
    End With
    'DataLabel Plug-In
    With ChartJS.Plugin_DataLabel
      .Source = chScriptSource.CDN
      .CDN = "https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@0.7.0"
      .PathFilename = CurrentProject.Path & "\chartjs-plugin-datalabels.min.js"
    End With
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
    Dim strHtml As New BAC_StringBuilder
    With strHtml
        .Add getHeader
        .Add getChartType
        .Add getChartData
        .Add getChartOptions
        .Add getFooter
    End With
 
    If m_SaveHtmlFileForDebugging = True Then
        Call saveHtmlFile(strHtml.ToString)
    End If
 
    getHtmlString = strHtml.ToString
  
End Function

Private Function getHeader() As String

    'Declarations
    Dim strText As New BAC_StringBuilder
    With strText
        .AddLine "<!DOCTYPE html>"
        .AddLine "<!-- saved from url=(0014)about:internet -->"
        .AddLine "<!-- saved from url=(0016)http://localhost -->"
        .AddLine "<html>"
        .AddLine "<head>"
        .AddLine "    <meta charset=""utf-8"">"
        .AddLine "    <meta http-equiv=""X-UA-Compatible"" content=""IE=Edge""/>"
        .AddLine "    <title>Chart</title>"
        .AddLine "    <style>body{margin:0;padding:0;}</style>"
        .AddLine "    <script src=" & ChartJS.getLocation & "></script>"
        .AddLine "    <script src=" & ChartJS.Plugin_ColorSchema.getLocation & "></script>"
        If ChartJS.Plugin_DataLabel.Show = True Then
            .AddLine "    <script src=" & ChartJS.Plugin_DataLabel.getDataLabelLocation & "></script>"
        End If
        .AddLine "</head>"
        .AddLine "<body>"
        .AddLine "<div id=""container"" style=""width: 100%;"">"
        If Len(ChartBackgroundColor) > 0 Then
            .AddLine "    <canvas id=""myChart"" style=""background-color: " & ChartBackgroundColor & """ ></canvas>"
        Else
            .AddLine "    <canvas id=""myChart""></canvas>"
        End If
        .AddLine "</div>"
        If ShowDownloadLink = True Then
            .AddLine "<a href=""#"" id=""download"" onclick=""download()"">Download</a>"
        End If
        .AddLine "<script>"
        If Len(DefaultFontColor) > 0 Then
            .AddLine "    Chart.defaults.global.defaultFontColor = '" & DefaultFontColor & "';"
        End If
        If Len(DefaultFontFamily) > 0 Then
            .AddLine "    Chart.defaults.global.defaultFontFamily = '" & DefaultFontFamily & "';"
        End If
        If DefaultFontSize > 0 Then
            .AddLine "    Chart.defaults.global.defaultFontSize = " & DefaultFontSize & ";"
        End If
        .AddLine "    var ctx = document.getElementById('myChart').getContext('2d');"
        .AddLine "    var myChart = new Chart(ctx, {"
        
        getHeader = .ToString
    End With
    
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
    Dim strText As New BAC_StringBuilder
    Dim dataSet As BAC_DataSet
    Dim i As Integer

    Set dataSet = Me.DataSource.getData(Me.ChartType)
    With strText
        .AddLine "        data: {"
        If ChartType <> chChartType.Bubble Then
            If ChartType <> chChartType.Scatter Then
                .AddLine "            labels: [" & dataSet.Label(0) & "],"
            End If
        End If
        
        .AddLine "            datasets: ["
        For i = 0 To dataSet.DataCount - 1
            Select Case ChartType
                Case chChartType.Bubble, chChartType.Scatter
                    .AddLine "               {label: '" & dataSet.Label(i) & "' ,"
                Case Else
                    .AddLine "               {label: '" & dataSet.DatasetName(i) & "' ,"
            End Select
            Select Case ChartType
                Case chChartType.line, chChartType.Radar
                    If AreaFilled = False Then
                        .AddLine "                fill: false ,"
                    End If
                Case Else
                    'Do Nothing
            End Select
                    
            .AddLine "                data: [" & dataSet.Data(i) & "],"
            If m_IsSingleBarChart = True Then
                .AddLine "                backgroundColor: Chart['colorschemes']." & ChartJS.Plugin_ColorSchema.ColorSchema & ","
            End If
            .AddLine "                borderWidth: 1},"
        Next i
        .CutTrailingComma
        .AddLine "             ]"
        .AddLine "        },"
        
        getChartData = .ToString
    End With
End Function
    
Private Function getChartOptions() As String
    
    'Declarations
    Dim strText As New BAC_StringBuilder
    With strText
        .AddLine "        options: {"
        .Add Title.getTitle
        .Add Legend.getLegend
        .Add getScales
        .Add getPlugins
        .CutTrailingComma
        .AddLine "        }"
        
        getChartOptions = .ToString
    End With
End Function

Private Function getFooter() As String

    'Declarations
    Dim strText As New BAC_StringBuilder
    With strText
        .AddLine "    });"  '???
        
        If ShowDownloadLink = True Then
            .AddLine "    function download() {"
            .AddLine "        var myChart = document.getElementById('myChart');"
            .AddLine "        var blob = myChart.msToBlob()"
            .AddLine "        window.navigator.msSaveBlob(blob, 'Better-Access-Charts.png');"
            .AddLine "    }"
        End If
        .AddLine "</script>"
        .AddLine "</body>"
        .AddLine "</html>"
        
        getFooter = strText.ToString
    End With
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
    Dim strText As New BAC_StringBuilder
    With strText
        .AddLine "            scales: {"
        .Add Me.YAxis.getAxis(Me.StackedChart)
        .Add Me.XAxis.getAxis(Me.StackedChart)
        .CutTrailingComma
        .AddLine "            },"
        
        getScales = .ToString
    End With
End Function

Private Function getPlugins() As String

    'Declarations
    Dim strText As New BAC_StringBuilder
    With strText
        .AddLine "            plugins: {"
        .Add ChartJS.Plugin_ColorSchema.getColorSchema
        .Add ChartJS.Plugin_DataLabel.getDataLabel
        .CutTrailingComma
        .AddLine "                },"
        
        getPlugins = .ToString
    End With
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

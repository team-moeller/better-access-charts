VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_Chart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###########################################################################################
'# Copyright (c) 2020 - 2022 Thomas Moeller, supported by K.D.Gundermann                   #
'# MIT License  => https://github.com/team-moeller/better-access-charts/blob/main/LICENSE  #
'# Version 2.05.07  published: 16.01.2022                                                  #
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

Private Const HtmlFileName As String = "BetterAccessChart_{}.html"


'### Members

Private WithEvents m_Control As WebBrowserControl
Attribute m_Control.VB_VarHelpID = -1
Private m_ChartType As chChartType
Private m_StackedChart As Boolean
Private m_AreaFilled As Boolean
Private m_ChartBackgroundColor As String
Private m_ShowDownloadLink As Boolean
Private m_SaveHtmlFileForDebugging As Boolean
Private m_AspectRatio As Double

Private m_IsSingleBarChart As Boolean
    
Private m_ColorSchema As BAC_ColorSchema
Private m_DataSource  As BAC_DataSource
Private m_DataLabels  As BAC_DataLabels
Private m_Defaults    As BAC_Defaults
Private m_XAxis       As BAC_Axis
Private m_YAxis       As BAC_Axis
Private m_Legend      As BAC_Legend
Private m_Title       As BAC_Title
    
    
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

' Proportion of Width / Height
Public Property Let AspectRatio(ByVal This_Ratio As Double)
    m_AspectRatio = This_Ratio
End Property
Public Property Get AspectRatio() As Double
    AspectRatio = m_AspectRatio
End Property

Public Property Let ChartBackgroundColor(ByVal This_ChartBackgroundColor As String)
    m_ChartBackgroundColor = This_ChartBackgroundColor
End Property
Public Property Get ChartBackgroundColor() As String
    ChartBackgroundColor = m_ChartBackgroundColor
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

Public Property Get ColorSchema() As BAC_ColorSchema
    Set ColorSchema = m_ColorSchema
End Property

Public Property Get DataSource() As BAC_DataSource
    Set DataSource = m_DataSource
End Property

Public Property Get Default() As BAC_Defaults
  Set Default = m_Defaults
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

Public Property Get DataLabels() As BAC_DataLabels
    Set DataLabels = m_DataLabels
End Property

'### Instancing

Private Sub Class_Initialize()
    
    'create Sub Classes
    Set m_ColorSchema = New BAC_ColorSchema
    Set m_DataSource = New BAC_DataSource
    Set m_DataLabels = New BAC_DataLabels
    Set m_Defaults = New BAC_Defaults
    Set m_Legend = New BAC_Legend
    Set m_Title = New BAC_Title
    Set m_XAxis = New BAC_Axis
    m_XAxis.Init apXAxis
    Set m_YAxis = New BAC_Axis
    m_YAxis.Init apYAxis
    
    'Set default values
    ChartType = chChartType.Bar
    StackedChart = False
    AspectRatio = 2
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
        
    'ColorSchema
    ColorSchema.ColorSchema = "brewer.Paired12"

End Sub

Private Sub Class_Terminate()
    ' Do Nothing
End Sub


'### Public Methods

Public Sub DrawChart()
    
    CreateChartFile
    
    Control.Object.Navigate getHTMLFileName
    
End Sub

    
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


'### Private Methods

Private Sub CreateChartFile()

    'Declarations
    Dim objFS As Object
    Dim strHtml As New BAC_StringBuilder
    With strHtml
        .Add getHeader
        .Add getChartType
        .Add getChartData
        .Add getChartOptions
        .Add getFooter
    End With
       
    Set objFS = CreateObject("ADODB.Stream")
    objFS.Charset = "utf-8"
    objFS.Open
    objFS.WriteText strHtml.ToString
    objFS.SaveToFile getHTMLFileName, 2   '2: Create Or Update
    objFS.Close
    
End Sub

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
        .AddLine "    <script src=" & BAC.Scripts.ChartJS.getLocation & "></script>"
        .AddLine "    <script src=" & BAC.Scripts.Plugin_ColorSchemes.getLocation & "></script>"
        If DataLabels.Show = True Then
            .AddLine "    <script src=" & BAC.Scripts.Plugin_DataLabels.getDataLabelLocation & "></script>"
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
        If Len(Default.Font.Color) > 0 Then
            .AddLine "    Chart.defaults.global.defaultFontColor = '" & Default.Font.Color & "';"
        End If
        If Len(Default.Font.Family) > 0 Then
            .AddLine "    Chart.defaults.global.defaultFontFamily = '" & Default.Font.Family & "';"
        End If
        If Default.Font.Size > 0 Then
            .AddLine "    Chart.defaults.global.defaultFontSize = " & Default.Font.Size & ";"
        End If
        If Default.Font.IsItalic = True Then
            .AddLine "    Chart.defaults.global.defaultFontStyle = 'italic';"
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
                .AddLine "                backgroundColor: Chart['colorschemes']." & ColorSchema.ColorSchema & ","
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
        If Me.AspectRatio <> 2# Then
            .AddLine " aspectRatio: " & Str(AspectRatio) & ","
        End If
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
        .Add ColorSchema.getColorSchema
        .Add DataLabels.getDataLabel
        .CutTrailingComma
        .AddLine "                },"
        
        getPlugins = .ToString
    End With
    
End Function


'### Private Helper Methods

Private Function getHTMLFileName() As String
    
    'Declarations
    Dim strPath As String
    Dim strFilename As String
    
    strPath = BAC.Helper.getTempFolder
    strFilename = Replace(HtmlFileName, "{}", m_Control.Name)

    getHTMLFileName = strPath & strFilename
    
End Function

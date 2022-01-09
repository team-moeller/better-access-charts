VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_ChartGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'###########################################################################################
'# Copyright (c) 2020 - 2022 Thomas Möller, supported by K.D.Gundermann                    #
'# MIT License  => https://github.com/team-moeller/better-access-charts/blob/main/LICENSE  #
'# Version 2.xx.xx  published: xx.01.2022                                                  #
'###########################################################################################
Option Compare Database
Option Explicit

Private c_Options  As BAC_Chart
Private c_Data     As BAC_DataSet

'### Initializer
Public Sub Init(ByVal ChartOptions As BAC_Chart, ByVal ChartData As BAC_DataSet)
    Set c_Options = ChartOptions
    Set c_Data = ChartData
End Sub

'### Public Methods
Public Function GenerateHTML() As String
    
    Dim strHTML As New BAC_StringBuilder
    With strHTML
        .Add getHeader
        .Add getChartType
        .Add getChartData
        .Add getChartOptions
        .Add getFooter
    End With
    GenerateHTML = strHTML.ToString()
End Function

'### Private Properties
Private Property Get Options() As BAC_Chart
    Set Options = c_Options
End Property
Private Property Get Data() As BAC_DataSet
    Set Data = c_Data
End Property


'### Private Methods / Functions
Private Function getHeader() As String
    Dim Font As BAC_Font
    
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
        If Options.DataLabels.Show = True Then
            .AddLine "    <script src=" & BAC.Scripts.Plugin_DataLabels.getDataLabelLocation & "></script>"
        End If
        .AddLine "</head>"
        
        .AddLine "<body>"
        .AddLine "<div id=""container"" style=""width: 100%;"">"
        If Len(Options.ChartBackgroundColor) > 0 Then
            .AddLine "    <canvas id=""myChart"" style=""background-color: " & Options.ChartBackgroundColor & """ ></canvas>"
        Else
            .AddLine "    <canvas id=""myChart""></canvas>"
        End If
        .AddLine "</div>"
        
        If Options.ShowDownloadLink = True Then
            .AddLine "<a href=""#"" id=""download"" onclick=""download()"">Download</a>"
        End If
        
        .AddLine "<script>"
        
        Set Font = Options.Default.Font
        If Len(Font.Color) > 0 Then
            .AddLine "    Chart.defaults.global.defaultFontColor = '" & Font.Color & "';"
        End If
        If Len(Font.Family) > 0 Then
            .AddLine "    Chart.defaults.global.defaultFontFamily = '" & Font.Family & "';"
        End If
        If Font.Size > 0 Then
            .AddLine "    Chart.defaults.global.defaultFontSize = " & Font.Size & ";"
        End If
        
        .AddLine "    var ctx = document.getElementById('myChart').getContext('2d');"
        .AddLine "    var myChart = new Chart(ctx, {"
        
        getHeader = .ToString
    End With
    
End Function

Private Function getChartType() As String

    'Declarations
    Dim strType As String
    
    Select Case Options.ChartType
        Case chChartType.Line
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
        
    getChartType = "        type: '" & strType & "'," & vbCrLf

End Function

Private Function getChartData() As String

    'Declarations
    Dim strText As New BAC_StringBuilder
    Dim i As Integer

    With strText
        .AddLine "        data: {"
        If Options.ChartType <> chChartType.Bubble Then
            If Options.ChartType <> chChartType.Scatter Then
                .AddLine "            labels: [" & Data.Label(0) & "],"
            End If
        End If
        
        .AddLine "            datasets: ["
        For i = 0 To Data.DataCount - 1
        
            Select Case Options.ChartType
                Case chChartType.Bubble, chChartType.Scatter
                    .AddLine "               {label: '" & Data.Label(i) & "' ,"
                Case Else
                    .AddLine "               {label: '" & Data.DatasetName(i) & "' ,"
            End Select
            
            Select Case Options.ChartType
                Case chChartType.Line, chChartType.Radar
                    If Options.AreaFilled = False Then
                        .AddLine "                fill: false ,"
                    End If
                Case Else
                    'Do Nothing
            End Select
                    
            .AddLine "                data: [" & Data.Data(i) & "],"
            
            If Options.IsSingleBarChart Then
                .AddLine "                backgroundColor: Chart['colorschemes']." & Options.ColorSchema.ColorSchema & ","
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
        If Options.AspectRatio <> 2# Then
            .AddLine " aspectRatio: " & Str(Options.AspectRatio) & ","
        End If
        .Add Options.Title.getTitle
        .Add Options.Legend.getLegend
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
        .AddLine "    });"  ' closing: var myChart = new Chart(ctx, {"
        
        If Options.ShowDownloadLink = True Then
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

Private Function getScales() As String

    'Declarations
    Dim strText As New BAC_StringBuilder
    With strText
        .AddLine "            scales: {"
        .Add Options.YAxis.getAxis(Options.StackedChart)
        .Add Options.XAxis.getAxis(Options.StackedChart)
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
        .Add Options.ColorSchema.getColorSchema
        .Add Options.DataLabels.getDataLabel
        .CutTrailingComma
        .AddLine "                },"
        
        getPlugins = .ToString
    End With
    
End Function


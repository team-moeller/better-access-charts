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
'# Version 1.28.04  published: 10.04.2021                                                  #
'###########################################################################################

Option Compare Database
Option Explicit


'### Enums

Public Enum chChartType
    Line = 1
    Bar = 2
    HorizontalBar = 3
    Radar = 4
    Doughnut = 5
    Pie = 6
    PolarArea = 7
    Bubble = 8
    Scatter = 9
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

Private Const HtmlFileName As String = "BetterAccessChart.html"


'### Members

Private WithEvents m_Control As WebbrowserControl
Attribute m_Control.VB_VarHelpID = -1
Private m_ChartType As chChartType
Private m_DataSourceObjectName As String
Private m_DataSourceSQLStatement As String
Private m_DataSourceRecordset As DAO.Recordset
Private m_DataSourceLabelFieldName As String
Private m_DataSourceDataFieldName As Variant
Private m_ShowTitle As Boolean
Private m_TitlePostition As chPosition
Private m_TitleText As String
Private m_ShowLegend As Boolean
Private m_LegendPostition As chPosition
Private m_ShowYAxis As Boolean
Private m_ShowXAxis As Boolean
Private m_YAxisLabelText As String
Private m_XAxisLabelText As String
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
Private m_YAxisBeginAtZero As Boolean
Private m_XAxisBeginAtZero As Boolean
Private m_YAxisTicksMin As String
Private m_YAxisTicksMax As String
Private m_XAxisTicksMin As String
Private m_XAxisTicksMax As String
Private m_ShowDownloadLink As Boolean
Private m_DataLabelSource As chScriptSource
Private m_DataLabelCDN As String
Private m_DataLabelPathFilename As String
Private m_ShowDataLabel As Boolean
Private m_DataLabelAnchor As chDataLabelAnchor
Private m_DataLabelFontColor As String

Private m_DataLabels As Variant
Private m_DataSetLabel As Variant
Private m_DataSetData As Variant
Private m_IsSingleDataset As Boolean
Private m_IsSingleBarChart As Boolean
Private m_YAxisTicksNeeded As Boolean
Private m_XAxisTicksNeeded As Boolean
Private m_DataLabelNeeded As Boolean


'### Properties

Public Property Set Control(ByVal This_Control As WebbrowserControl)
    Set m_Control = This_Control
End Property
Public Property Get Control() As WebbrowserControl
    Set Control = m_Control
End Property

Public Property Let ChartType(ByVal This_ChartType As chChartType)
    m_ChartType = This_ChartType
    Select Case This_ChartType
        Case chChartType.Bar, chChartType.HorizontalBar, chChartType.Line, chChartType.Bubble, chChartType.Scatter
            ShowYAxis = True
            ShowXAxis = True
        Case chChartType.Doughnut, chChartType.Pie, chChartType.PolarArea, chChartType.Radar
            ShowYAxis = False
            ShowXAxis = False
        Case Else
            'Do Nothing
    End Select
End Property
Public Property Get ChartType() As chChartType
    ChartType = m_ChartType
End Property

Public Property Let DataSourceObjectName(ByVal This_DataSourceObjectName As String)
    m_DataSourceObjectName = This_DataSourceObjectName
End Property
Public Property Get DataSourceObjectName() As String
    DataSourceObjectName = m_DataSourceObjectName
End Property

Public Property Let DataSourceSQLStatement(ByVal This_DataSourceSQLStatement As String)
    m_DataSourceSQLStatement = This_DataSourceSQLStatement
End Property
Public Property Get DataSourceSQLStatement() As String
    DataSourceSQLStatement = m_DataSourceSQLStatement
End Property

Public Property Set DataSourceRecordset(ByVal This_DataSourceRecordset As DAO.Recordset)
    Set m_DataSourceRecordset = This_DataSourceRecordset
End Property
Public Property Get DataSourceRecordset() As DAO.Recordset
    Set DataSourceRecordset = m_DataSourceRecordset
End Property

Public Property Let DataSourceLabelFieldName(ByVal This_DataSourceLabelFieldName As String)
    m_DataSourceLabelFieldName = This_DataSourceLabelFieldName
End Property
Public Property Get DataSourceLabelFieldName() As String
    DataSourceLabelFieldName = m_DataSourceLabelFieldName
End Property

Public Property Let DataSourceDataFieldName(ByRef This_DataSourceDataFieldName As Variant)
    If IsArray(This_DataSourceDataFieldName) = True Then
        m_DataSourceDataFieldName = This_DataSourceDataFieldName
    Else
        m_DataSourceDataFieldName = Array(This_DataSourceDataFieldName)
    End If
    If LBound(m_DataSourceDataFieldName) = UBound(m_DataSourceDataFieldName) Then
        m_IsSingleDataset = True
    Else
        m_IsSingleDataset = False
    End If
End Property
Public Property Get DataSourceDataFieldName() As Variant
    DataSourceDataFieldName = m_DataSourceDataFieldName
End Property

Public Property Let ShowTitle(ByVal This_ShowTitle As Boolean)
    m_ShowTitle = This_ShowTitle
End Property
Public Property Get ShowTitle() As Boolean
    ShowTitle = m_ShowTitle
End Property

Public Property Let TitlePostition(ByVal This_TitlePostition As chPosition)
    m_TitlePostition = This_TitlePostition
End Property
Public Property Get TitlePostition() As chPosition
    TitlePostition = m_TitlePostition
End Property

Public Property Let TitleText(ByVal This_TitleText As String)
    m_TitleText = This_TitleText
End Property
Public Property Get TitleText() As String
    TitleText = m_TitleText
End Property

Public Property Let ShowLegend(ByVal This_ShowLegend As Boolean)
    m_ShowLegend = This_ShowLegend
End Property
Public Property Get ShowLegend() As Boolean
    ShowLegend = m_ShowLegend
End Property

Public Property Let LegendPostition(ByVal This_LegendPostition As chPosition)
    m_LegendPostition = This_LegendPostition
End Property
Public Property Get LegendPostition() As chPosition
    LegendPostition = m_LegendPostition
End Property

Public Property Let ShowYAxis(ByVal This_ShowYAxis As Boolean)
    m_ShowYAxis = This_ShowYAxis
End Property
Public Property Get ShowYAxis() As Boolean
    ShowYAxis = m_ShowYAxis
End Property

Public Property Let ShowXAxis(ByVal This_ShowXAxis As Boolean)
    m_ShowXAxis = This_ShowXAxis
End Property
Public Property Get ShowXAxis() As Boolean
    ShowXAxis = m_ShowXAxis
End Property

Public Property Let YAxisLabelText(ByVal This_YAxisLabelText As String)
    m_YAxisLabelText = This_YAxisLabelText
End Property
Public Property Get YAxisLabelText() As String
    YAxisLabelText = m_YAxisLabelText
End Property

Public Property Let XAxisLabelText(ByVal This_XAxisLabelText As String)
    m_XAxisLabelText = This_XAxisLabelText
End Property
Public Property Get XAxisLabelText() As String
    XAxisLabelText = m_XAxisLabelText
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
    StackedChart = m_StackedChart
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

Public Property Let YAxisBeginAtZero(ByVal This_YAxisBeginAtZero As Boolean)
    m_YAxisBeginAtZero = This_YAxisBeginAtZero
    If This_YAxisBeginAtZero = True Then
        m_YAxisTicksNeeded = True
    End If
End Property
Public Property Get YAxisBeginAtZero() As Boolean
    YAxisBeginAtZero = m_YAxisBeginAtZero
End Property

Public Property Let XAxisBeginAtZero(ByVal This_XAxisBeginAtZero As Boolean)
    m_XAxisBeginAtZero = This_XAxisBeginAtZero
    If This_XAxisBeginAtZero = True Then
        m_XAxisTicksNeeded = True
    End If
End Property
Public Property Get XAxisBeginAtZero() As Boolean
    XAxisBeginAtZero = m_XAxisBeginAtZero
End Property

Public Property Let YAxisTicksMin(ByVal This_YAxisTicksMin As String)
    m_YAxisTicksMin = This_YAxisTicksMin
    If Len(Nz(This_YAxisTicksMin, vbNullString)) > 0 Then
        m_YAxisTicksNeeded = True
    End If
End Property
Public Property Get YAxisTicksMin() As String
    YAxisTicksMin = m_YAxisTicksMin
End Property

Public Property Let YAxisTicksMax(ByVal This_YAxisTicksMax As String)
    m_YAxisTicksMax = This_YAxisTicksMax
    If Len(Nz(This_YAxisTicksMax, vbNullString)) > 0 Then
        m_YAxisTicksNeeded = True
    End If
End Property
Public Property Get YAxisTicksMax() As String
    YAxisTicksMax = m_YAxisTicksMax
End Property

Public Property Let XAxisTicksMin(ByVal This_XAxisTicksMin As String)
    m_XAxisTicksMin = This_XAxisTicksMin
    If Len(Nz(This_XAxisTicksMin, vbNullString)) > 0 Then
        m_XAxisTicksNeeded = True
    End If
End Property
Public Property Get XAxisTicksMin() As String
    XAxisTicksMin = m_XAxisTicksMin
End Property

Public Property Let XAxisTicksMax(ByVal This_XAxisTicksMax As String)
    m_XAxisTicksMax = This_XAxisTicksMax
    If Len(Nz(This_XAxisTicksMax, vbNullString)) > 0 Then
        m_XAxisTicksNeeded = True
    End If
End Property
Public Property Get XAxisTicksMax() As String
    XAxisTicksMax = m_XAxisTicksMax
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


'### Instancing

Private Sub Class_Initialize()
    
    'Set default values
    ChartType = chChartType.Bar
    StackedChart = False
    'Title
    ShowTitle = False
    TitlePostition = posTop
    TitleText = "My first Better Access Chart"
    'Legend
    ShowLegend = False
    LegendPostition = posRight
    'Axis
    ShowYAxis = True
    ShowXAxis = True
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

    'Declarations
    Dim FileName As String
    
    CreateChartFile
       
    FileName = CurrentProject.Path & "\" & HtmlFileName
    
    Control.Object.Navigate FileName
    
End Sub


'### Private Methods

Private Sub CreateChartFile()

    'Declarations
    Dim f As Integer
    Dim strPath As String
    Dim strFilename As String
    Dim strContent As String
    
    strPath = CurrentProject.Path
    strFilename = "\" & HtmlFileName
    
    strContent = strContent & getHeader
    strContent = strContent & getChartType
    strContent = strContent & getChartData
    strContent = strContent & getChartOptions
    strContent = strContent & getFooter
 
    f = FreeFile()
    Open strPath & strFilename For Output As f
    Print #f, strContent;
 
    Close f
  
End Sub

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
    
    strText = strText & "        type: '" & strType & "'," & vbCrLf
    
    getChartType = strText

End Function

Private Function getChartData() As String

    'Declarations
    Dim strText As String
    Dim i As Integer

    Call getData

    strText = strText & "        data: {" & vbCrLf
    If ChartType <> chChartType.Bubble Then
        If ChartType <> chChartType.Scatter Then
            strText = strText & "            labels: [" & m_DataLabels & "]," & vbCrLf
        End If
    End If
    strText = strText & "            datasets: [" & vbCrLf
    For i = LBound(m_DataSetData) To UBound(m_DataSetData)
        Select Case ChartType
            Case chChartType.Bubble, chChartType.Scatter
                strText = strText & "               {label: '" & m_DataLabels(i) & "' ," & vbCrLf
            Case Else
                strText = strText & "               {label: '" & m_DataSetLabel(i) & "' ," & vbCrLf
        End Select
        Select Case ChartType
            Case chChartType.Line, chChartType.Radar
                If AreaFilled = False Then
                    strText = strText & "                fill: false ," & vbCrLf
                End If
            Case Else
                'Do Nothing
        End Select
        strText = strText & "                data: [" & m_DataSetData(i) & "]," & vbCrLf
        If m_IsSingleBarChart = True Then
            strText = strText & "                backgroundColor: Chart['colorschemes']." & ColorSchema & "," & vbCrLf
        End If
        strText = strText & "                borderWidth: 1}," & vbCrLf
    Next i
    strText = CutTrailingComma(strText)
    strText = strText & "             ]" & vbCrLf
    strText = strText & "        }," & vbCrLf
    
    getChartData = strText

End Function

Private Function getChartOptions() As String

    'Declarations
    Dim strText As String
    
    strText = strText & "        options: {" & vbCrLf
    strText = strText & getTitle
    strText = strText & getLegend
    strText = strText & getScales
    strText = strText & getPlugins
    strText = CutTrailingComma(strText)
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

Private Sub getData()
On Error GoTo Handle_Error

    'Declarations
    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    Dim strLabels() As String
    Dim strData() As String
    Dim Field As Variant
    Dim i As Integer
    
    Set db = CurrentDb()
    
    'Detect data source
    If Len(DataSourceObjectName) > 0 Then
        Set rst = db.OpenRecordset(DataSourceObjectName)
    ElseIf Len(DataSourceSQLStatement) > 0 Then
        Set rst = db.OpenRecordset(DataSourceSQLStatement)
    ElseIf Not DataSourceRecordset Is Nothing Then
        Set rst = DataSourceRecordset
    Else
        DataSourceDataFieldName = "Demo"
    End If
    
    'Extract data and labels
    Select Case ChartType
        Case chChartType.Bubble
            If Not rst Is Nothing Then
                ReDim strLabels(rst.RecordCount - 1)
                ReDim strData(rst.RecordCount - 1)
                Do While Not rst.EOF
                    strLabels(i) = rst(DataSourceLabelFieldName)
                    strData(i) = "{x:" & Str$(rst(DataSourceDataFieldName(0))) & _
                                 ", y:" & Str$(rst(DataSourceDataFieldName(1))) & _
                                 ", r:" & Str$(rst(DataSourceDataFieldName(2))) & "}"
                    i = i + 1
                    rst.MoveNext
                Loop
            Else
                ReDim strLabels(0)
                ReDim strData(0)
                strLabels(0) = "Demo"
                strData(0) = "{x:2,y:3,r:17}"
            End If
                'Cleanup result
                m_DataLabels = strLabels()
                m_DataSetData = strData()
        Case chChartType.Scatter
            If Not rst Is Nothing Then
                ReDim strLabels(rst.RecordCount - 1)
                ReDim strData(rst.RecordCount - 1)
                Do While Not rst.EOF
                    strLabels(i) = rst(DataSourceLabelFieldName)
                    strData(i) = "{x:" & Str$(rst(DataSourceDataFieldName(0))) & _
                                 ", y:" & Str$(rst(DataSourceDataFieldName(1))) & "}"
                    i = i + 1
                    rst.MoveNext
                Loop
            Else
                ReDim strLabels(0)
                ReDim strData(0)
                strLabels(0) = "Demo"
                strData(0) = "{x:2,y:3,r:17}"
            End If
                'Cleanup result
                m_DataLabels = strLabels()
                m_DataSetData = strData()
        Case Else
            ReDim strLabels(1)
            ReDim strData(UBound(DataSourceDataFieldName))
            If Not rst Is Nothing Then
                Do While Not rst.EOF
                    strLabels(0) = strLabels(0) & "'" & rst(DataSourceLabelFieldName) & "',"
                    i = 0
                    For Each Field In DataSourceDataFieldName
                        strData(i) = strData(i) & Str$(rst(DataSourceDataFieldName(i))) & ","
                        i = i + 1
                    Next Field
                    rst.MoveNext
                Loop
            Else
                strLabels(0) = "'Berlin', 'London', 'Paris', 'Rome', 'Lisbon', 'Madrid',"
                strData(0) = "2, 13, 7, 9, 12, 5,"
            End If
            'Cleanup result
            DetermineSingleBarChart
            ReDim m_DataSet1Colors(UBound(DataSourceDataFieldName))
            m_DataLabels = CutTrailingComma(strLabels(0))
            m_DataSetLabel = DataSourceDataFieldName()
            i = 0
            For Each Field In DataSourceDataFieldName()
                strData(i) = CutTrailingComma(strData(i))
                i = i + 1
            Next Field
            m_DataSetData = strData()
    End Select
    
Exit_Here:
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    Set db = Nothing
    Exit Sub

Handle_Error:
    Select Case Err.Number
        Case Else
            MsgBox Err.Description, vbExclamation, Err.Number
    End Select
    Resume Exit_Here
    Resume Next

End Sub

Private Sub DetermineSingleBarChart()

    Select Case m_ChartType
        Case chChartType.Bar, chChartType.HorizontalBar
            If m_IsSingleDataset = True Then
                m_IsSingleBarChart = True
            Else
                m_IsSingleBarChart = False
            End If
        Case Else
            m_IsSingleBarChart = False
    End Select

End Sub

Private Function getTitle() As String

    'Declarations
    Dim strText As String

    If ShowTitle = False Then
        strText = strText & "            title: {" & vbCrLf
        strText = strText & "                display: false," & vbCrLf
        strText = strText & "            }," & vbCrLf
        getTitle = strText
        Exit Function
    End If
    
    strText = strText & "            title: {" & vbCrLf
    strText = strText & "                display: true," & vbCrLf
    strText = strText & "                position: " & getPositionString(TitlePostition) & "," & vbCrLf
    strText = strText & "                text: '" & TitleText & "'" & vbCrLf
    strText = strText & "            }," & vbCrLf
    
    getTitle = strText

End Function

Private Function getLegend() As String

    'Declarations
    Dim strText As String

    If ShowLegend = False Then
        strText = strText & "            legend: {" & vbCrLf
        strText = strText & "                display: false" & vbCrLf
        strText = strText & "            }," & vbCrLf
        getLegend = strText
        Exit Function
    End If

    strText = strText & "            legend: {" & vbCrLf
    strText = strText & "                display: true," & vbCrLf
    strText = strText & "                position: " & getPositionString(LegendPostition) & vbCrLf
    strText = strText & "            }," & vbCrLf

    getLegend = strText

End Function

Private Function getScales() As String

    'Declarations
    Dim strText As String

    strText = strText & "            scales: {" & vbCrLf
    strText = strText & "                yAxes: [{" & vbCrLf
    strText = strText & "                    id: 'first-y-Axis'," & vbCrLf
    If StackedChart = True Then
        Select Case ChartType
            Case chChartType.Line, chChartType.Bar, chChartType.HorizontalBar
                strText = strText & "                    stacked: true," & vbCrLf
            Case Else
                'Do Nothing
        End Select
    End If
    If ShowYAxis = False Then
        strText = strText & "                    display: false" & vbCrLf
    Else
        strText = strText & "                    display: true," & vbCrLf
        If Len(YAxisLabelText) > 0 Then
            strText = strText & "                    scaleLabel: {" & vbCrLf
            strText = strText & "                       display: true," & vbCrLf
            strText = strText & "                       labelString: '" & YAxisLabelText & "'" & vbCrLf
            strText = strText & "                    }," & vbCrLf
        End If
        If m_YAxisTicksNeeded = True Then
            strText = strText & "                    ticks: {" & vbCrLf
        End If
        If YAxisBeginAtZero = True Then
            strText = strText & "                       beginAtZero: true," & vbCrLf
        End If
        If Len(YAxisTicksMin) > 0 Then
            If IsNumeric(YAxisTicksMin) = True Then
                strText = strText & "                       min: " & Str$(YAxisTicksMin) & " ," & vbCrLf
            Else
                strText = strText & "                       min: '" & YAxisTicksMin & "' ," & vbCrLf
            End If
        End If
        If Len(YAxisTicksMax) > 0 Then
            If IsNumeric(YAxisTicksMax) = True Then
                strText = strText & "                       max: " & Str$(YAxisTicksMax) & " ," & vbCrLf
            Else
                strText = strText & "                       max: '" & YAxisTicksMax & "' ," & vbCrLf
            End If
        End If
        strText = CutTrailingComma(strText)
        If m_YAxisTicksNeeded = True Then
            strText = strText & "                    }," & vbCrLf
        End If
    End If
    strText = CutTrailingComma(strText)
    strText = strText & "                    }]," & vbCrLf
    strText = strText & "                xAxes: [{" & vbCrLf
    strText = strText & "                    id: 'first-x-Axis'," & vbCrLf
    If StackedChart = True Then
        Select Case ChartType
            Case chChartType.Line, chChartType.Bar, chChartType.HorizontalBar
                strText = strText & "                    stacked: true," & vbCrLf
            Case Else
                'Do Nothing
        End Select
    End If
    If ShowXAxis = False Then
        strText = strText & "                    display: false" & vbCrLf
    Else
        strText = strText & "                    display: true," & vbCrLf
        If Len(XAxisLabelText) > 0 Then
            strText = strText & "                    scaleLabel: {" & vbCrLf
            strText = strText & "                        display: true," & vbCrLf
            strText = strText & "                        labelString: '" & XAxisLabelText & "'" & vbCrLf
            strText = strText & "                    }," & vbCrLf
        End If
        If m_XAxisTicksNeeded = True Then
            strText = strText & "                    ticks: {" & vbCrLf
        End If
        If XAxisBeginAtZero = True Then
            strText = strText & "                       beginAtZero: true," & vbCrLf
        End If
        If Len(XAxisTicksMin) > 0 Then
            If IsNumeric(XAxisTicksMin) = True Then
                strText = strText & "                       min: " & Str$(XAxisTicksMin) & " ," & vbCrLf
            Else
                strText = strText & "                       min: '" & XAxisTicksMin & "' ," & vbCrLf
            End If
        End If
        If Len(XAxisTicksMax) > 0 Then
            If IsNumeric(XAxisTicksMax) = True Then
                strText = strText & "                       max: " & Str$(XAxisTicksMax) & " ," & vbCrLf
            Else
                strText = strText & "                       max: '" & XAxisTicksMax & "' ," & vbCrLf
            End If
        End If
        strText = CutTrailingComma(strText)
        If m_XAxisTicksNeeded = True Then
            strText = strText & "                    }," & vbCrLf
        End If
    End If
    strText = CutTrailingComma(strText)
    strText = strText & "                    }]" & vbCrLf
    strText = strText & "            }," & vbCrLf
    
    getScales = strText
    
End Function

Private Function getPlugins() As String

    'Declarations
    Dim strText As String

    strText = strText & "            plugins: {" & vbCrLf
    strText = strText & getColorSchema
    strText = strText & getDataLabel
    strText = CutTrailingComma(strText)

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

    If m_DataLabelNeeded = True Then
        strText = strText & "                datalabels: {" & vbCrLf
    End If
    If DataLabelAnchor > 0 Then
        strText = strText & "                    anchor: '" & getDataLabelAnchorString(DataLabelAnchor) & "'," & vbCrLf
    End If
    If Len(Nz(DataLabelFontColor, vbNullString)) > 0 Then
        strText = strText & "                    color: '" & DataLabelFontColor & "'," & vbCrLf
    End If
    strText = CutTrailingComma(strText)
    If m_DataLabelNeeded = True Then
        strText = strText & "                    }," & vbCrLf
    End If
    
    getDataLabel = strText

End Function

Private Function getPositionString(ByVal ThisPosition As chPosition) As String

    Select Case ThisPosition
        Case chPosition.posTop
            getPositionString = "'top'"
        Case chPosition.posLeft
            getPositionString = "'left'"
        Case chPosition.posBottom
            getPositionString = "'bottom'"
        Case chPosition.posRight
            getPositionString = "'right'"
        Case Else
            getPositionString = "'top'"
    End Select

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

Private Function CutTrailingComma(ByVal strText As String) As String

    If Right$(strText, 1) = "," Then
        CutTrailingComma = Left$(strText, Len(strText) - 1)
    ElseIf Right$(strText, 3) = "," & vbCrLf Then
        CutTrailingComma = Left$(strText, Len(strText) - 3) & vbCrLf
    Else
        CutTrailingComma = strText
    End If

End Function


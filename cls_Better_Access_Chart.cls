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
'# Copyright (c) 2020 Thomas Möller                                                        #
'# MIT License  => https://github.com/team-moeller/better-access-charts/blob/main/LICENSE  #
'# Version 1.04.14  published: 05.12.2020                                                  #
'###########################################################################################

Option Compare Database
Option Explicit


'### Enums

Public Enum chChartType
    Line = 1
    Bar = 2
    Radar = 3
    Doughnut = 4
    Pie = 5
    PolarArea = 6
    Bubble = 7
    Scatter = 8
End Enum

Public Enum chPosition
    posTop = 1
    posLeft = 2
    posBottom = 3
    posRight = 4
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
Private m_DataSourceDataFieldName As String
Private m_ShowTitle As Boolean
Private m_TitlePostition As chPosition
Private m_TitleText As String
Private m_ShowLegend As Boolean
Private m_LegendPostition As chPosition
Private m_ShowYAxis As Boolean
Private m_ShowXAxis As Boolean
Private m_YAxisLabelText As String
Private m_XAxisLabelText As String

Private m_DataLabels As String
Private m_DataSet1Label As String
Private m_DataSet1Data As String


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
        Case chChartType.Bar, chChartType.Line, chChartType.Bubble, chChartType.Scatter
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

Public Property Let DataSourceDataFieldName(ByVal This_DataSourceDataFieldName As String)
    m_DataSourceDataFieldName = This_DataSourceDataFieldName
End Property
Public Property Get DataSourceDataFieldName() As String
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


'### Instancing

Private Sub Class_Initialize()
    
    'Set default values
    ChartType = chChartType.Bar
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
'    strText = strText & "    <script src=""chartjs/dist/Chart.min.js""></script>" & vbCrLf
    strText = strText & "    <script src=""https://cdnjs.cloudflare.com/ajax/libs/Chart.js/2.9.4/Chart.min.js""></script>" & vbCrLf
    strText = strText & "</head>" & vbCrLf
    strText = strText & "<body>" & vbCrLf
    strText = strText & "<div id=""container"" style=""width: 100%;"">" & vbCrLf
    strText = strText & "    <canvas id=""myChart""></canvas>" & vbCrLf
    strText = strText & "</div>" & vbCrLf
    strText = strText & "<script>" & vbCrLf
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

    Call getData

    strText = strText & "        data: {" & vbCrLf
    strText = strText & "            labels: [" & m_DataLabels & "]," & vbCrLf
    strText = strText & "            datasets: [{" & vbCrLf
    strText = strText & "                label: '" & m_DataSet1Label & "' ," & vbCrLf
    strText = strText & "                data: [" & m_DataSet1Data & "]," & vbCrLf
    strText = strText & "                backgroundColor: ['red', 'blue', 'yellow', 'green', 'purple', 'orange']," & vbCrLf
    strText = strText & "                borderWidth: 1" & vbCrLf
'    strText = strText & "            },{" & vbCrLf
'    strText = strText & "                label: '# of Votes 2'," & vbCrLf
'    strText = strText & "                data: [9, 9, 12, 5, 2, 13]," & vbCrLf
'    strText = strText & "                backgroundColor: ['red', 'blue', 'yellow', 'green', 'purple', 'orange']," & vbCrLf
'    strText = strText & "                borderWidth: 1" & vbCrLf
    strText = strText & "            }]" & vbCrLf
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
    strText = CutTrailingComma(strText)
    strText = strText & "        }" & vbCrLf
    
    getChartOptions = strText

End Function

Private Function getFooter() As String

    'Declarations
    Dim strText As String
    
    strText = strText & "    });" & vbCrLf
    strText = strText & "</script>" & vbCrLf
    strText = strText & "</body>" & vbCrLf
    strText = strText & "</html>"
    
    getFooter = strText
    
End Function

Private Sub getData()
On Error GoTo Handle_Error

    'Declarations
    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    Dim strLabels As String
    Dim strData As String
    
    Set db = CurrentDb()
    
    'Detect data source
    If Len(DataSourceObjectName) > 0 Then
        Set rst = db.OpenRecordset(DataSourceObjectName)
    ElseIf Len(DataSourceSQLStatement) > 0 Then
        Set rst = db.OpenRecordset(DataSourceSQLStatement)
    ElseIf Not DataSourceRecordset Is Nothing Then
        Set rst = DataSourceRecordset
    End If
    
    'Extract data and labels
    If Not rst Is Nothing Then
        Do While Not rst.EOF
            strLabels = strLabels & "'" & rst(DataSourceLabelFieldName) & "',"
            strData = strData & rst(DataSourceDataFieldName) & ","
            rst.MoveNext
        Loop
    Else
        strLabels = "'red', 'blue', 'yellow', 'green', 'purple', 'orange',"
        strData = "2, 13, 7, 9, 12, 5,"
    End If
    
    'Cleanup result
    m_DataLabels = Left$(strLabels, Len(strLabels) - 1)
    m_DataSet1Label = DataSourceDataFieldName
    m_DataSet1Data = Left$(strData, Len(strData) - 1)

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
    If ShowYAxis = False Then
        strText = strText & "                    display: false" & vbCrLf
    Else
        strText = strText & "                    display: true," & vbCrLf
        If Len(YAxisLabelText) > 0 Then
            strText = strText & "                    scaleLabel: {" & vbCrLf
            strText = strText & "                    display: true," & vbCrLf
            strText = strText & "                    labelString: '" & YAxisLabelText & "'" & vbCrLf
            strText = strText & "                    }" & vbCrLf
        Else
            strText = CutTrailingComma(strText)
        End If
    End If
    strText = strText & "                    }]," & vbCrLf
    strText = strText & "                xAxes: [{" & vbCrLf
    strText = strText & "                    id: 'first-x-Axis'," & vbCrLf
    If ShowXAxis = False Then
        strText = strText & "                    display: false" & vbCrLf
    Else
        strText = strText & "                    display: true," & vbCrLf
        If Len(XAxisLabelText) > 0 Then
            strText = strText & "                    scaleLabel: {" & vbCrLf
            strText = strText & "                    display: true," & vbCrLf
            strText = strText & "                    labelString: '" & XAxisLabelText & "'" & vbCrLf
            strText = strText & "                    }" & vbCrLf
        Else
            strText = CutTrailingComma(strText)
        End If
    End If
    strText = strText & "                    }]" & vbCrLf
    strText = strText & "            }" & vbCrLf
    
    getScales = strText
    
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

Private Function CutTrailingComma(ByVal strText As String) As String

    If Right$(strText, 1) = "," Then
        CutTrailingComma = Left$(strText, Len(strText) - 1)
    ElseIf Right$(strText, 3) = "," & vbCrLf Then
        CutTrailingComma = Left$(strText, Len(strText) - 3) & vbCrLf
    Else
        CutTrailingComma = strText
    End If

End Function

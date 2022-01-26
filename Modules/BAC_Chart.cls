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
'# Version 2.09.05  published: 22.01.2022                                                  #
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

Public Enum chDataSourceType
    dstDemo
    dstTableName
    dstQueryName
    dstSQLStament
    dstRecordset
    dstDataset
    dstEmpty
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
Private m_BackgroundColor As String
Private m_ShowDownloadLink As Boolean
Private m_SaveHtmlFileForDebugging As Boolean
Private m_AspectRatio As Double
    
Private m_ColorSchema As BAC_ColorSchema
Private m_DataSource  As BAC_DataSource
Private m_DataLabels  As BAC_DataLabels
Private m_Defaults    As BAC_Defaults
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
        Case chChartType.Bar, chChartType.HorizontalBar, chChartType.Line, chChartType.Bubble, chChartType.Scatter
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
        Case chChartType.Line, chChartType.Bar, chChartType.HorizontalBar
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

Public Property Let BackgroundColor(ByVal This_BackgroundColor As String)
    m_BackgroundColor = This_BackgroundColor
End Property
Public Property Get BackgroundColor() As String
    BackgroundColor = m_BackgroundColor
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

Public Property Get DataSet() As BAC_DataSet
    Set DataSet = m_DataSource.DataSet
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
    ColorSchema.Name = "brewer.Paired12"

End Sub

Private Sub Class_Terminate()
    ' Do Nothing
End Sub


'### Public Methods

Public Sub DrawChart()

    Dim strHTML As String
    With BAC.Generator(Me, Me.DataSource.getData(Me.ChartType))
        strHTML = .GenerateHTML
    End With
    
    If CreateChartFile(getHTMLFileName, strHTML) Then
        Control.Object.Navigate getHTMLFileName
    End If
    
End Sub

Public Sub ShowInBrowser()
    
    Dim strHTML As String
    With BAC.Generator(Me, Me.DataSource.getData(Me.ChartType))
        strHTML = .GenerateHTML
    End With
    
    If CreateChartFile(getHTMLFileName, strHTML) Then
        BAC.Helper.ShellExecute getHTMLFileName
    End If

End Sub

    

'### Private Methods

 
Friend Function IsSingleBarChart()

    Select Case m_ChartType
        Case chChartType.Bar, chChartType.HorizontalBar
            If DataSource.IsSingleDataset = True Then
                IsSingleBarChart = True
            Else
                IsSingleBarChart = False
            End If
        Case Else
            IsSingleBarChart = False
    End Select

End Function

Friend Function IsBubbleOrScatterChart() As Boolean
    If m_ChartType = Bubble Then
        IsBubbleOrScatterChart = True
    ElseIf m_ChartType = Scatter Then
        IsBubbleOrScatterChart = True
    Else
        IsBubbleOrScatterChart = False
    End If
End Function


'### Private Helper Methods

Private Sub saveHtmlFile(ByVal strContent As String)

    'Declarations
    Dim strPath As String
    Dim strFilename As String
    Dim f As Integer
    Const HtmlFileName As String = "BetterAccessChart_{}.html"
    
    strPath = CurrentProject.Path & "\"
    strFilename = Replace(HtmlFileName, "{}", m_Control.Name)
    
    f = FreeFile()
    Open strPath & strFilename For Output As f
    Print #f, strContent;
    Close f
    
End Sub

Private Function getHTMLFileName() As String
    
    'Declarations
    Dim strPath As String
    Dim strFilename As String
    
    strPath = BAC.Helper.getTempFolder
    strFilename = Replace(HtmlFileName, "{}", m_Control.Name)

    getHTMLFileName = strPath & strFilename
    
End Function

Private Function CreateChartFile(ByVal FileName As String, ByVal strHTML As String) As Boolean

    'Declarations
    Dim objFS As Object
       
    Set objFS = CreateObject("ADODB.Stream")
    objFS.Charset = "utf-8"
    objFS.Open
    objFS.WriteText strHTML
    objFS.SaveToFile FileName, 2   '2: Create Or Update
    objFS.Close
    CreateChartFile = True
    
End Function


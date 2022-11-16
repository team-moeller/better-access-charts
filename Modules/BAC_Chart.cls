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
'# Version 2.25.03  published: 16.11.2022                                                  #
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
    dstData
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

Public Enum chAlign
    alStart = 1
    alCenter = 2
    alEnd = 3
End Enum

Public Enum chDataLabelAnchor
    anStart = 1
    anCenter = 2
    anEnd = 3
End Enum

Public Enum chDisplayIn
    chWebBrowserControl = 1
    chWebBrowserActiveX = 2
    chImageControl = 3
    chSystemBrowser = 4
End Enum

Public Enum chEasing
    linear = 0
    easeInQuad = 1
    easeOutQuad = 2
    easeInOutQuad = 3
    easeInCubic = 4
    easeOutCubic = 5
    easeInOutCubic = 6
    easeInQuart = 7
    easeOutQuart = 8
    easeInOutQuart = 9
    easeInQuint = 10
    easeOutQuint = 11
    easeInOutQuint = 12
    easeInSine = 13
    easeOutSine = 14
    easeInOutSine = 15
    easeInExpo = 16
    easeOutExpo = 17
    easeInOutExpo = 18
    easeInCirc = 19
    easeOutCirc = 20
    easeInOutCirc = 21
    easeInElastic = 22
    easeOutElastic = 23
    easeInOutElastic = 24
    easeInBack = 25
    easeOutBack = 26
    easeInOutBack = 27
    easeInBounce = 28
    easeOutBounce = 29
    easeInOutBounce = 30
End Enum

'### Members

Private m_WebBrowserControl As Access.WebBrowserControl
Private m_WebBrowserActiveX As SHDocVw.WebBrowser
Private m_ImageControl As Access.Image
Private m_ChartType As chChartType
Private m_StackedChart As Boolean
Private m_AreaFilled As Boolean
Private m_BackgroundColor As String
Private m_ShowDownloadLink As Boolean
Private m_SaveHtmlFileForDebugging As Boolean
Private m_AspectRatio As Double
Private m_DisplayIn As chDisplayIn
    
Private m_Animation   As BAC_Animation
Private m_ColorSchema As BAC_ColorSchema
Private m_DataSource  As BAC_DataSource
Private m_DataLabels  As BAC_DataLabels
Private m_Defaults    As BAC_Defaults
Private m_XAxis       As BAC_Axis
Private m_YAxis       As BAC_Axis
Private m_RadialAxis  As BAC_Axis
Private m_Legend      As BAC_Legend
Private m_Title       As BAC_Title
    
'### Initializer

Public Sub Init(Optional ByVal TheControl As Object, _
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

Public Property Set Control(ByVal New_Control As Object)
    If TypeName(New_Control) = "WebBrowserControl" Then
        Set Me.WebBrowserControl = New_Control
    ElseIf TypeName(New_Control) = "Image" Then
        Set Me.ImageControl = New_Control
    ElseIf TypeName(New_Control.Object) = "WebBrowser" Then
        Set Me.WebBrowserActiveX = New_Control.Object
    Else
        Err.Raise vbObjectError, , "Type of control not supported!"
    End If
End Property

Public Property Set WebBrowserControl(ByVal New_Control As Access.WebBrowserControl)
    Set m_WebBrowserControl = New_Control
    m_DisplayIn = chWebBrowserControl
End Property
Public Property Get WebBrowserControl() As WebBrowserControl
    Set WebBrowserControl = m_WebBrowserControl
End Property

Public Property Set WebBrowserActiveX(ByVal New_Control As SHDocVw.WebBrowser)
    Set m_WebBrowserActiveX = New_Control
    m_DisplayIn = chWebBrowserActiveX
End Property
Public Property Get WebBrowserActiveX() As SHDocVw.WebBrowser
    Set WebBrowserActiveX = m_WebBrowserActiveX
End Property

Public Property Set ImageControl(ByVal New_Control As Access.Image)
    Set m_ImageControl = New_Control
    m_DisplayIn = chImageControl
End Property
Public Property Get ImageControl() As Access.Image
    Set ImageControl = m_ImageControl
End Property

Public Property Let ChartType(ByVal This_ChartType As chChartType)
    m_ChartType = This_ChartType
    Select Case This_ChartType
        Case chChartType.Bar, chChartType.HorizontalBar, chChartType.Line, chChartType.Bubble, chChartType.Scatter
            YAxis.Show = True
            XAxis.Show = True
            RadialAxis.Show = False
        Case chChartType.Doughnut, chChartType.Pie, chChartType.PolarArea, chChartType.Radar
            YAxis.Show = False
            XAxis.Show = False
            RadialAxis.Show = True
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

Public Property Let DisplayIn(ByVal newValue As chDisplayIn)
    m_DisplayIn = newValue
End Property
Public Property Get DisplayIn() As chDisplayIn
    DisplayIn = m_DisplayIn
End Property



'### Object Properties

Public Property Get Animation() As BAC_Animation
    Set Animation = m_Animation
End Property

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

Public Property Get RadialAxis() As BAC_Axis
    Set RadialAxis = m_RadialAxis
End Property

Public Property Get DataLabels() As BAC_DataLabels
    Set DataLabels = m_DataLabels
End Property

Public Property Get Data() As BAC_Data
    Set Data = m_DataSource.Data
End Property


'### Instancing

Private Sub Class_Initialize()
    
    'Create Sub Classes
    Set m_Animation = New BAC_Animation
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
    Set m_RadialAxis = New BAC_Axis
    
    'Set default values
    ChartType = chChartType.Bar
    StackedChart = False
    AspectRatio = 2
    DisplayIn = chWebBrowserControl
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

Public Sub CreateChart()

    'Declarations
    Dim ImageFileName As String

    CreateChartFile
    
    Select Case Me.DisplayIn
        Case chDisplayIn.chWebBrowserControl
            Me.WebBrowserControl.Object.Navigate getHTMLFileName
        Case chDisplayIn.chWebBrowserActiveX
            Me.WebBrowserActiveX.Navigate getHTMLFileName
        Case chDisplayIn.chImageControl
            Call SaveChartAsImage(ImageFileName)
            Me.ImageControl.Picture = ImageFileName
        Case chDisplayIn.chSystemBrowser
            BAC.Helper.ShellExecute getHTMLFileName
        Case Else
            'Do nothing
    End Select
    
End Sub

    
'### Private Methods

Private Sub CreateChartFile()

    'Declarations
    Dim strHTML As String
    
    With BAC.Generator(Me, Me.DataSource.getData(Me.ChartType))
        strHTML = .GenerateHTML
    End With
 
    If Me.SaveHtmlFileForDebugging = True Then
        Call saveHtmlFile(CurrentProject.Path & "\", strHTML)
    End If
    
    Call saveHtmlFile(BAC.Helper.getTempFolder, strHTML)
    
End Sub

Private Sub SaveChartAsImage(ByRef ImageFileName As String)

    'Declarations
    Dim objIE As Object 'As SHDocVw.InternetExplorer 'InternetExplorer
    Dim objCanvas As Object 'As MSHTML.HTMLCanvasElement 'As HTMLCanvasElement
    Dim objXML As Object 'As MSXML2.DOMDocument
    Dim objDocElem As Object 'As MSXML2.IXMLDOMElement
    Dim objStream As Object 'As ADODB.Stream
    Dim strImg As String
    Dim strData As String
    Dim arr64decode() As Byte

    'Start IE and load chart
    Set objIE = CreateObject("InternetExplorer.Application") 'New InternetExplorer
''    objIE.Visible = True
    objIE.Navigate getHTMLFileName
    Do While objIE.ReadyState <> 4
        DoEvents
    Loop

    'Get canvas and grab image
    Set objCanvas = objIE.Document.getElementById("myChart")
    strImg = objCanvas.toDataURL("image/png")
    
    'Use exclamation mark if no image available
    If Left$(strImg, 22) <> "data:image/png;base64," Then
        strImg = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAsAAAASCAIAAAACF7MiAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAABpSURBVChTYzxw4AADfgBU8R83AMoyQRXiBqSquDPRmpExfTuUBwHk2GKlpQplQQCpZqio6UBZCIBpho6aCpQFASS7VFXLCsqCAzQVKvlHZ3pC2VCAqoJwiAGN+P8fzRCSXYoFEEpBDAwAPNYyBnTMkl4AAAAASUVORK5CYII="
    End If
    
    'Convert image to byte array
    strData = Right$(strImg, Len(strImg) - 22)
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objDocElem = objXML.createElement("tmp")
    objDocElem.DataType = "bin.base64"
    objDocElem.Text = strData
    arr64decode = objDocElem.NodeTypedValue
    
    'Save to file
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 1 ' adTypeBinary = 1
    objStream.Open
    objStream.Write arr64decode
    ImageFileName = BAC.Helper.getTempFolder & "image.png"
    objStream.SaveToFile ImageFileName, 2 ' adSaveCreateOverWrite = 2
    objIE.Quit

End Sub


'### Private Helper Methods

Private Sub saveHtmlFile(ByVal Folder As String, ByVal strHTML As String)

    'Declarations
    Dim objFS As Object
    Dim FileName As String
       
    FileName = getHTMLFileName(Folder)
    
    Set objFS = CreateObject("ADODB.Stream")
    objFS.Charset = "utf-8"
    objFS.Open
    objFS.WriteText strHTML
    objFS.SaveToFile FileName, 2   '2: Create Or Update
    objFS.Close
    
End Sub

Private Function getHTMLFileName(Optional Path As String = vbNullString) As String
    
    'Declarations
    Dim strFilename As String
    Const HtmlFileName As String = "BetterAccessChart_{}.html"

    If Path = vbNullString Then
        Path = BAC.Helper.getTempFolder
    End If
    strFilename = Replace(HtmlFileName, "{}", ObjPtr(Me))

    getHTMLFileName = Path & strFilename
    
End Function

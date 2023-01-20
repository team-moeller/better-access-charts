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
'# Version 2.30.02  published: 30.12.2022                                                  #
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
    dstEmpty = 0
    dstTableName = 1      ' what is the difference between a table and a query used in a data source?
    dstQueryName = 2
    dstSQLStament = 3
    dstRecordset = 4
    dstData = 5
    dstDemo = 255         ' is /demo/ a data source type?
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
Private m_Properties As BAC_Properties

Private m_WebBrowserControl As Access.WebBrowserControl
Private m_WebBrowserActiveX As SHDocVw.WebBrowser
Private m_ImageControl As Access.Image

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
Public Property Get Properties() As BAC_Properties
   Set Properties = m_Properties
End Property

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
Public Property Get Control() As Object
    If Not m_WebBrowserControl Is Nothing Then
        Set Control = m_WebBrowserControl
    ElseIf Not m_WebBrowserActiveX Is Nothing Then
        Set Control = m_WebBrowserActiveX
    ElseIf Not m_ImageControl Is Nothing Then
        Set Control = m_ImageControl
    Else
        Set Control = Nothing
    End If
End Property

Public Property Set WebBrowserControl(ByVal New_Control As Access.WebBrowserControl)
    Set m_WebBrowserControl = New_Control
    DisplayIn = chWebBrowserControl
End Property
Public Property Get WebBrowserControl() As WebBrowserControl
    Set WebBrowserControl = m_WebBrowserControl
End Property

Public Property Set WebBrowserActiveX(ByVal New_Control As SHDocVw.WebBrowser)
    Set m_WebBrowserActiveX = New_Control
    DisplayIn = chWebBrowserActiveX
End Property
Public Property Get WebBrowserActiveX() As SHDocVw.WebBrowser
    Set WebBrowserActiveX = m_WebBrowserActiveX
End Property

Public Property Set ImageControl(ByVal New_Control As Access.Image)
    Set m_ImageControl = New_Control
    DisplayIn = chImageControl
End Property
Public Property Get ImageControl() As Access.Image
    Set ImageControl = m_ImageControl
End Property

Public Property Let ChartType(ByVal This_ChartType As chChartType)
    m_Properties.Item("ChartType").Value = This_ChartType
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
    ChartType = Nz(m_Properties.Item("ChartType").Value, 0)
End Property

Public Property Let StackedChart(ByVal This_StackedChart As Boolean)
    m_Properties.Item("StackedChart").Value = This_StackedChart
End Property
Public Property Get StackedChart() As Boolean
    Select Case ChartType
        Case chChartType.Line, chChartType.Bar, chChartType.HorizontalBar
            StackedChart = Nz(m_Properties.Item("StackedChart").Value, False)
        Case Else
            StackedChart = False
    End Select
End Property

Public Property Let AreaFilled(ByVal This_AreaFilled As Boolean)
    m_Properties.Item("AreaFilled").Value = This_AreaFilled
End Property
Public Property Get AreaFilled() As Boolean
    AreaFilled = Nz(m_Properties.Item("AreaFilled").Value, False)
End Property

' Proportion of Width / Height
Public Property Let AspectRatio(ByVal This_Ratio As Double)
    m_Properties.Item("AspectRatio").Value = This_Ratio
End Property
Public Property Get AspectRatio() As Double
    AspectRatio = Nz(m_Properties.Item("AspectRatio").Value, 0)
    If AspectRatio = 0 Then
      If Me.Control Is Nothing Then
         AspectRatio = 2
      ElseIf Control.Height > 0 Then
         AspectRatio = Control.Width / Control.Height
      Else
         AspectRatio = 2
      End If
    End If
End Property

Public Property Let BackgroundColor(ByVal This_BackgroundColor As String)
    m_Properties.Item("BackgroundColor").Value = This_BackgroundColor
End Property
Public Property Get BackgroundColor() As String
    BackgroundColor = Nz(m_Properties.Item("BackgroundColor").Value, vbNullString)
End Property

Public Property Let ShowDownloadLink(ByVal This_ShowDownloadLink As Boolean)
    m_Properties.Item("ShowDownloadLink").Value = This_ShowDownloadLink
End Property
Public Property Get ShowDownloadLink() As Boolean
    ShowDownloadLink = Nz(m_Properties.Item("ShowDownloadLink").Value, False)
End Property

Public Property Let SaveHtmlFileForDebugging(ByVal This_SaveHtmlFileForDebugging As Boolean)
    m_Properties.Item("SaveHtmlFileForDebugging").Value = This_SaveHtmlFileForDebugging
End Property
Public Property Get SaveHtmlFileForDebugging() As Boolean
    SaveHtmlFileForDebugging = Nz(m_Properties.Item("SaveHtmlFileForDebugging").Value, False)
End Property

Public Property Let DisplayIn(ByVal NewValue As chDisplayIn)
    m_Properties.Item("DisplayIn").Value = NewValue
End Property
Public Property Get DisplayIn() As chDisplayIn
    DisplayIn = Nz(m_Properties.Item("DisplayIn").Value, 0)
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
    
    Set m_Properties = New BAC_Properties
    InitProperties
    
    'Create Sub Classes
    Set m_Animation = New BAC_Animation
    Set m_Properties.NewProperty("Animation").Properties = m_Animation.Properties
    
    Set m_ColorSchema = New BAC_ColorSchema
    Set m_Properties.NewProperty("ColorSchema").Properties = m_ColorSchema.Properties
    
    Set m_DataSource = New BAC_DataSource
    
    Set m_DataLabels = New BAC_DataLabels
    Set m_Properties.NewProperty("DataLabels").Properties = m_DataLabels.Properties
    
    Set m_Defaults = New BAC_Defaults
    Set m_Properties.NewProperty("Default").Properties = m_Defaults.Properties
    
    Set m_Legend = New BAC_Legend
    Set m_Properties.NewProperty("Legend").Properties = m_Legend.Properties
    
    Set m_Title = New BAC_Title
    Set m_Properties.NewProperty("Title").Properties = m_Title.Properties
    
    Set m_XAxis = New BAC_Axis
    Set m_Properties.NewProperty("XAxis").Properties = m_XAxis.Properties
    m_XAxis.Init apXAxis
    
    Set m_YAxis = New BAC_Axis
    Set m_Properties.NewProperty("YAxis").Properties = m_YAxis.Properties
    m_YAxis.Init apYAxis
    
    Set m_RadialAxis = New BAC_Axis
    Set m_Properties.NewProperty("RadialAxis").Properties = m_RadialAxis.Properties
    
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

Private Sub InitProperties()
   m_Properties.NewProperty("ChartType").Value = 0
   m_Properties.NewProperty "StackedChart"
   m_Properties.NewProperty "AreaFilled"
   m_Properties.NewProperty "BackgroundColor"
   m_Properties.NewProperty "ShowDownloadLink"
   m_Properties.NewProperty "SaveHtmlFileForDebugging"
   m_Properties.NewProperty "AspectRatio"
   m_Properties.NewProperty "DisplayIn"
End Sub

Private Sub Class_Terminate()
    ' Do Nothing
End Sub


'### Public Methods

Public Sub CreateChart()

    'Declarations
    Dim ImageFileName As String
    Dim strHtml As String
    
    Select Case Me.DisplayIn
        Case chDisplayIn.chWebBrowserControl
            CreateEmptyHtmlFile
            Me.WebBrowserControl.Object.Navigate getHTMLFileName(, True)
            Do While Me.WebBrowserControl.Object.Busy Or Me.WebBrowserControl.Object.ReadyState <> 4
                DoEvents
            Loop
            With Me.WebBrowserControl.Object.Document
                .Open
                strHtml = BAC.Generator(Me, Me.DataSource.getData(Me.ChartType)).GenerateHTML
                .Write strHtml
                .Close
            End With
            If Me.SaveHtmlFileForDebugging = True Then
                Call saveHtmlFile(CurrentProject.Path & "\", strHtml)
            End If
        Case chDisplayIn.chWebBrowserActiveX
            CreateEmptyHtmlFile
            Me.WebBrowserActiveX.Navigate getHTMLFileName(, True)
            Do While Me.WebBrowserActiveX.Busy Or Me.WebBrowserActiveX.ReadyState <> 4
                DoEvents
            Loop
            With Me.WebBrowserActiveX.Document
                .Open
                strHtml = BAC.Generator(Me, Me.DataSource.getData(Me.ChartType)).GenerateHTML
                .Write strHtml
                .Close
            End With
            If Me.SaveHtmlFileForDebugging = True Then
                Call saveHtmlFile(CurrentProject.Path & "\", strHtml)
            End If
        Case chDisplayIn.chImageControl
            Call SaveChartAsImage(ImageFileName)
            Me.ImageControl.Picture = ImageFileName
        Case chDisplayIn.chSystemBrowser
            CreateChartFile
            BAC.Helper.ShellExecute getHTMLFileName
        Case Else
            'Do nothing
    End Select
    
End Sub

    
'### Private Methods

Private Sub CreateEmptyHtmlFile()

    'Decalarations
    Dim strHtml As New BAC_StringBuilder
    
    With strHtml
        .AddLine "<!DOCTYPE html>"
        .AddLine "<!-- saved from url=(0014)about:internet -->"
        .AddLine "<html>"
        .AddLine "<head>"
        .AddLine "    <meta http-equiv='X-UA-Compatible' content='IE=Edge'>"
        .AddLine "</head>"
        .AddLine "<body>"
        .AddLine "</body>"
        .AddLine "</html>"
    End With
    
    Call saveHtmlFile(vbNullString, strHtml.ToString, True)
    
End Sub

Private Sub CreateChartFile()

    'Declarations
    Dim strHtml As String
    
    With BAC.Generator(Me, Me.DataSource.getData(Me.ChartType))
        strHtml = .GenerateHTML
    End With
 
    If Me.SaveHtmlFileForDebugging = True Then
        Call saveHtmlFile(CurrentProject.Path & "\", strHtml)
    End If
    
    Call saveHtmlFile(BAC.Helper.getTempFolder, strHtml)
    
End Sub

Private Sub SaveChartAsImage(ByRef ImageFileName As String)

    'Declarations
    Dim objIE As Object      'As SHDocVw.InternetExplorer 'InternetExplorer
    Dim objCanvas As Object  'As MSHTML.HTMLCanvasElement 'As HTMLCanvasElement
    Dim objXML As Object     'As MSXML2.DOMDocument
    Dim objDocElem As Object 'As MSXML2.IXMLDOMElement
    Dim objStream As Object  'As ADODB.Stream
    Dim strImg As String
    Dim strData As String
    Dim arr64decode() As Byte

    'Start IE and load chart
    Set objIE = CreateObject("InternetExplorer.Application") 'New InternetExplorer
''    objIE.Visible = True
    CreateEmptyHtmlFile
    objIE.Navigate getHTMLFileName(, True)
    Do While objIE.Busy Or objIE.ReadyState <> 4
        DoEvents
    Loop
    With objIE.Document
        .Open
        Dim strHtml As String
        strHtml = BAC.Generator(Me, Me.DataSource.getData(Me.ChartType)).GenerateHTML
        .Write strHtml
        .Close
    End With
    If Me.SaveHtmlFileForDebugging = True Then
        Call saveHtmlFile(CurrentProject.Path & "\", strHtml)
    End If
    
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

Private Sub saveHtmlFile(ByVal Folder As String, ByVal strHtml As String, Optional ByVal fEmpty As Boolean = False)

    'Declarations
    Dim objFS As Object
    Dim FileName As String
       
    FileName = getHTMLFileName(Folder, fEmpty)
    
    Set objFS = CreateObject("ADODB.Stream")
    objFS.Charset = "utf-8"
    objFS.Open
    objFS.WriteText strHtml
    objFS.SaveToFile FileName, 2   '2: Create Or Update
    objFS.Close
    
End Sub

Private Function getHTMLFileName(Optional Path As String = vbNullString, Optional fEmpty As Boolean = False) As String
    
    'Declarations
    Dim strFilename As String
    Const HtmlFileName As String = "BetterAccessChart_{}.html"

    If Path = vbNullString Then
        Path = BAC.Helper.getTempFolder
    End If
    If fEmpty = True Then
        strFilename = "Empty.html"
    Else
        strFilename = Replace(HtmlFileName, "{}", ObjPtr(Me))
    End If

    getHTMLFileName = Path & strFilename
    
End Function

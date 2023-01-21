Attribute VB_Name = "BetterAccessChartsFactory"
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>%AppFolder%/source/BetterAccessChartsFactory.bas</file>
'</codelib>
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Private m_BAC As Object
Private m_BACx As Object

#If BAC_EarlyBinding Then
#Else

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
   
#End If

' Factory
#If BAC_EarlyBinding Then
Public Function BAC() As BetterAccessChartsLoader.BAC__Factory
#Else
Public Function BAC() As Object
#End If
    If m_BAC Is Nothing Then
        Set m_BAC = BetterAccessChartsLoader.GetBetterAccessChartsFactory
    End If
    Set BAC = m_BAC
End Function

#If BAC_EarlyBinding Then
Public Function BACx() As BetterAccessChartsLoader.BacAddInTools
#Else
Public Function BACx() As Object
#End If
    If m_BACx Is Nothing Then
        Set m_BACx = BetterAccessChartsLoader.GetBetterAccessChartsAddInTools
    End If
    Set BACx = m_BACx
End Function

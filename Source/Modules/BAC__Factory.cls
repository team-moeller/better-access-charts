VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC__Factory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Chart(Optional ByVal TheControl As WebBrowserControl, _
                Optional ByVal ChartType As chChartType = chChartType.Bar, _
                Optional ByVal DataSourceType As chDataSourceType = dstDemo, _
                Optional ByVal DataSource As Variant, _
                Optional ByVal DataFieldName As String, _
                Optional ByVal LabelFieldName As String) As cls_Better_Access_Chart
  
  Set Chart = New cls_Better_Access_Chart
  Chart.Init TheControl, ChartType, DataSourceType, DataSource, DataFieldName, LabelFieldName
End Function

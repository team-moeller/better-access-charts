VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC__Factory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'###########################################################################################
'# Copyright (c) 2020, 2021 Thomas Möller, supported by K.D.Gundermann                     #
'# MIT License  => https://github.com/team-moeller/better-access-charts/blob/main/LICENSE  #
'# Version 1.31.12  published: 29.10.2021                                                  #
'###########################################################################################

Option Compare Database
Option Explicit

Private m_Helper As BAC_Helper

Public Function Chart(Optional ByVal TheControl As WebBrowserControl, _
                Optional ByVal ChartType As chChartType = chChartType.Bar, _
                Optional ByVal DataSourceType As chDataSourceType = dstDemo, _
                Optional ByVal DataSource As Variant, _
                Optional ByVal DataFieldNames As Variant, _
                Optional ByVal LabelFieldName As String) As cls_Better_Access_Chart
  
  Set Chart = New cls_Better_Access_Chart
  Chart.Init TheControl, ChartType, DataSourceType, DataSource, DataFieldNames, LabelFieldName
End Function

Friend Function Helper() As BAC_Helper
    Set Helper = m_Helper
End Function

Private Sub Class_Initialize()
    Set m_Helper = New BAC_Helper
End Sub


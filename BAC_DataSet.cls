VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_DataSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'###########################################################################################
'# Copyright (c) 2020 - 2022 Thomas Moeller, supported by K.D.Gundermann                   #
'# MIT License  => https://github.com/team-moeller/better-access-charts/blob/main/LICENSE  #
'# Version 2.08.04  published: 19.01.2022                                                  #
'###########################################################################################

Option Compare Database
Option Explicit


' Private m_DataSetLabel As Variant

Private m_LabelCount As Integer
Private m_Labels() As String

Private m_DataCount As Integer
Private m_Data() As String
Private m_DatasetNames() As String


' ### Properties

Public Property Get DataCount() As Integer
    DataCount = m_DataCount
End Property
Public Property Let DataCount(ByVal NewValue As Integer)
    m_DataCount = NewValue
    ReDim m_Data(NewValue - 1)
    ReDim m_DatasetNames(NewValue - 1)
End Property

Public Property Get Data(ByVal Index As Integer) As String
    Data = m_Data(Index)
End Property
Public Property Let Data(ByVal Index As Integer, ByVal NewValue As String)
    m_Data(Index) = NewValue
End Property

Public Property Get DatasetName(ByVal Index As Integer) As String
    DatasetName = m_DatasetNames(Index)
End Property
Public Property Let DatasetName(ByVal Index As Integer, ByVal NewValue As String)
    m_DatasetNames(Index) = NewValue
End Property

Public Property Get LabelCount() As Integer
    LabelCount = m_LabelCount
End Property
Public Property Let LabelCount(ByVal NewValue As Integer)
    m_LabelCount = NewValue
    ReDim m_Labels(NewValue - 1)
End Property

Public Property Get Label(ByVal Index As Integer) As String
    Label = m_Labels(Index)
End Property
Public Property Let Label(ByVal Index As Integer, ByVal NewValue As String)
    m_Labels(Index) = NewValue
End Property


' ### Public Function

Public Function Result() As BAC_DataSet
    ' Cleanup the data and return me as Result
    Dim i  As Integer
    
    For i = 0 To DataCount - 1
        Data(i) = BAC.Helper.CutTrailingComma(Data(i))
    Next
    For i = 0 To LabelCount - 1
        Label(i) = BAC.Helper.CutTrailingComma(Label(i))
    Next
    Set Result = Me
End Function


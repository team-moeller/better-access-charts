VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_DataSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###########################################################################################
'# Copyright (c) 2020 - 2022 Thomas M�ller, supported by K.D.Gundermann                    #
'# MIT License  => https://github.com/team-moeller/better-access-charts/blob/main/LICENSE  #
'# Version 2.03.05  published: 12.01.2022                                                  #
'###########################################################################################

Option Compare Database
Option Explicit


Private m_Type As chDataSourceType
Private m_ObjectName As String
Private m_SQLStatement As String
Private m_Recordset As DAO.Recordset
Private m_LabelFieldName As String
Private m_DataFieldNames() As String
Private m_Dataset As BAC_DataSet

Private m_IsSingleDataset As Boolean


' ### Initializer

Public Sub Init(Optional ByVal DataSourceType As chDataSourceType = dstDemo, _
                Optional ByVal DataSource As Variant, _
                Optional ByVal DataFieldNames As Variant, _
                Optional ByVal LabelFieldName As String)
  
  Select Case DataSourceType
        Case dstDemo:       Me.ObjectName = ""
        Case dstTableName:  Me.ObjectName = DataSource
        Case dstQueryName:  Me.ObjectName = DataSource
        Case dstSQLStament: Me.SQLStatement = DataSource
        Case dstRecordset:  Set Me.Recordset = DataSource
        Case dstDataset:    Set Me.Dataset = DataSource
  End Select
  Me.DataFieldNames = DataFieldNames
  Me.LabelFieldName = LabelFieldName
  
End Sub

' ### Methods

Public Sub UseDemoData()
    m_Type = dstDemo
End Sub

' ### Properties

Public Property Let ObjectName(ByVal This_ObjectName As String)
    m_Type = dstTableName
    m_ObjectName = This_ObjectName
End Property
Public Property Get ObjectName() As String
    ObjectName = m_ObjectName
End Property

Public Property Let SQLStatement(ByVal This_SQLStatement As String)
    m_Type = dstSQLStament
    m_SQLStatement = This_SQLStatement
End Property
Public Property Get SQLStatement() As String
    SQLStatement = m_SQLStatement
End Property

Public Property Set Recordset(ByVal This_Recordset As Object)  ' ADODB/DAO.Recordset
    m_Type = dstRecordset
    Set m_Recordset = This_Recordset
End Property
Public Property Get Recordset() As Object   ' ADODB/DAO.Recordset
    Set Recordset = m_Recordset
End Property

Public Property Let Dataset(ByVal This_Dataset As BAC_DataSet)
    m_Type = dstDataset
    Set m_Dataset = This_Dataset
End Property
Public Property Get Dataset() As BAC_DataSet
    Set Dataset = m_Dataset
End Property


Public Property Let LabelFieldName(ByVal This_LabelFieldName As String)
    m_LabelFieldName = This_LabelFieldName
End Property
Public Property Get LabelFieldName() As String
    LabelFieldName = m_LabelFieldName
End Property

Public Property Let DataFieldNames(ByRef This_DataFieldNames As Variant)
    Dim i As Integer
    
    If IsMissing(This_DataFieldNames) Then Exit Property
    
    If IsArray(This_DataFieldNames) = True Then
        ReDim m_DataFieldNames(UBound(This_DataFieldNames))
        For i = LBound(This_DataFieldNames) To UBound(This_DataFieldNames)
            m_DataFieldNames(i) = This_DataFieldNames(i)
        Next
    Else
        m_DataFieldNames = Split(This_DataFieldNames, ",")
    End If
    
    If LBound(m_DataFieldNames) = UBound(m_DataFieldNames) Then
        m_IsSingleDataset = True
    Else
        m_IsSingleDataset = False
    End If
End Property
Public Property Get DataFieldNames() As Variant
    DataFieldNames = m_DataFieldNames
End Property

Public Property Get IsSingleDataset() As Boolean
    IsSingleDataset = m_IsSingleDataset
End Property


' ### Generator Function

Friend Function getData(ByVal ChartType As chChartType) As BAC_DataSet
On Error GoTo Handle_Error

    'Declarations
    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    Dim dst As New BAC_DataSet
    
    Set db = CurrentDb()
    
    'Detect data source
    Select Case m_Type
        Case dstTableName, dstQueryName:
            Set rst = db.OpenRecordset(ObjectName)
        Case dstSQLStament:
            Set rst = db.OpenRecordset(SQLStatement)
        Case dstRecordset:
            Set rst = Me.Recordset
        Case dstDataset:
            Set dst = Me.Dataset
        Case dstDemo:
            If (ChartType = chChartType.Bubble) Then
                With dst
                    .AddLabel "Demo"
                    .AddDatasetName "Data"
                    .AddData "Data", "{x:2,y:3}"
                End With
            ElseIf (ChartType = chChartType.Scatter) Then
                With dst
                    .AddLabel "Demo"
                    .AddDatasetName "Data"
                    .AddData "Data", "{x:2,y:3,r:17}"
                End With
            Else
                With dst
                    .AddLabels "'Berlin', 'London', 'Paris', 'Rome', 'Lisbon', 'Madrid',"
                    .AddDatasetName "Data"
                    .AddData "Data", "2, 13, 7, 9, 12, 5,"
                End With
            End If
            
            DataFieldNames = Array("Demo")
    End Select
    
    If Not (rst Is Nothing) Then
        Set dst = getDataFromRS(rst, ChartType)
    End If
    
    
Exit_Here:
    If Not (rst Is Nothing) Then
        rst.Close
        Set rst = Nothing
    End If
    Set db = Nothing
    Set getData = dst.Result
    Exit Function

Handle_Error:
    Select Case Err.Number
        Case Else
            MsgBox Err.Description, vbExclamation, Err.Number
    End Select
    Resume Exit_Here
    Resume Next

End Function

Private Function getDataFromRS(ByVal rst As Object, ByVal ChartType As chChartType) As BAC_DataSet
    Dim dst As New BAC_DataSet
    Dim Field As Variant ' As String
    Dim i As Integer
    
    If rst Is Nothing Then Exit Function
    
    'Extract data and labels
    Select Case ChartType
        Case chChartType.Bubble
            With dst
                Do While Not rst.EOF
                    .AddLabel rst(LabelFieldName)
                    .AddDatasetName "Record" & Str(i)
                    .AddData i, "{x:" & Str$(Nz(rst(DataFieldNames(0)))) & _
                                 ", y:" & Str$(Nz(rst(DataFieldNames(1)))) & _
                                 ", r:" & Str$(Nz(rst(DataFieldNames(2)))) & "}"
                    i = i + 1
                    rst.MoveNext
                Loop
            End With
                
        Case chChartType.Scatter
            With dst
                Do While Not rst.EOF
                    .AddLabel rst(LabelFieldName)
                    .AddDatasetName "Record" & Str(i)
                    .AddData i, "{x:" & Str$(Nz(rst(DataFieldNames(0)))) & _
                                 ", y:" & Str$(Nz(rst(DataFieldNames(1)))) & "}"
                    i = i + 1
                    rst.MoveNext
                Loop
            End With
                
        Case Else
            With dst
                For Each Field In DataFieldNames
                    .AddDatasetName Field
                Next

                Do While Not rst.EOF
                    .AddLabel rst(LabelFieldName)
                    For Each Field In DataFieldNames
                        .AddData Field, Str$(Nz(rst(Field)))
                    Next Field
                    rst.MoveNext
                Loop
            End With
                        
    End Select
    Set getDataFromRS = dst
End Function


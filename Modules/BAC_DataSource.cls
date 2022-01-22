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
'# Copyright (c) 2020 - 2022 Thomas Moeller, supported by K.D.Gundermann                   #
'# MIT License  => https://github.com/team-moeller/better-access-charts/blob/main/LICENSE  #
'# Version 2.09.05  published: 22.01.2022                                                  #
'###########################################################################################

Option Compare Database
Option Explicit


Private m_Type As chDataSourceType
Private m_ObjectName As String
Private m_SQLStatement As String
Private m_Recordset As DAO.Recordset
Private m_LabelFieldName As String
Private m_DataFieldNames() As String

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
  End Select
  Me.DataFieldNames = DataFieldNames
  Me.LabelFieldName = LabelFieldName
  
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
        m_DataFieldNames = Array(This_DataFieldNames)
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
    Dim rst As Object ' ADODB/DAO.Recordset
    Dim dataSet As New BAC_DataSet
    Dim Field As Variant
    Dim i As Integer
    
    Set db = CurrentDb()
    
    'Detect data source
    If Len(ObjectName) > 0 Then
        Set rst = db.OpenRecordset(ObjectName)
    ElseIf Len(SQLStatement) > 0 Then
        Set rst = db.OpenRecordset(SQLStatement)
    ElseIf Not Recordset Is Nothing Then
        Set rst = Recordset
    Else
        DataFieldNames = Array("Demo")
    End If
    
    'Extract data and labels
    Select Case ChartType
        Case chChartType.Bubble
            If Not rst Is Nothing Then
                With dataSet
                    .LabelCount = rst.RecordCount
                    .DataCount = rst.RecordCount
                    Do While Not rst.EOF
                        .Label(i) = rst(LabelFieldName)
                        .Data(i) = "{x:" & Str$(Nz(rst(DataFieldNames(0)))) & _
                                     ", y:" & Str$(Nz(rst(DataFieldNames(1)))) & _
                                     ", r:" & Str$(Nz(rst(DataFieldNames(2)))) & "}"
                        i = i + 1
                        rst.MoveNext
                    Loop
                End With
            Else
                With dataSet
                    .DataCount = 1
                    .LabelCount = 1
                    .Label(0) = "Demo"
                    .Data(0) = "{x:2,y:3,r:17}"
                End With
            End If
                
        Case chChartType.Scatter
            If Not rst Is Nothing Then
                With dataSet
                    .LabelCount = rst.RecordCount
                    .DataCount = rst.RecordCount
                    Do While Not rst.EOF
                        .Label(i) = rst(LabelFieldName)
                        .Data(i) = "{x:" & Str$(Nz(rst(DataFieldNames(0)))) & _
                                     ", y:" & Str$(Nz(rst(DataFieldNames(1)))) & "}"
                        i = i + 1
                        rst.MoveNext
                    Loop
                End With
            Else
                With dataSet
                    .DataCount = 1
                    .LabelCount = 1
                    .Label(0) = "Demo"
                    .Data(0) = "{x:2,y:3,r:17}"
                End With
            End If
                
        Case Else
            With dataSet
                .LabelCount = 1
                .DataCount = UBound(DataFieldNames) + 1
                i = 0
                For Each Field In DataFieldNames
                    .DatasetName(i) = Field
                    i = i + 1
                Next

                If Not rst Is Nothing Then
                    Do While Not rst.EOF
                        .Label(0) = .Label(0) & "'" & rst(LabelFieldName) & "',"
                        i = 0
                        For Each Field In DataFieldNames
                            .Data(i) = .Data(i) & Str$(Nz(rst(DataFieldNames(i)))) & ","
                            i = i + 1
                        Next Field
                        rst.MoveNext
                    Loop
                Else
                    .Label(0) = "'Berlin', 'London', 'Paris', 'Rome', 'Lisbon', 'Madrid',"
                    .Data(0) = "2, 13, 7, 9, 12, 5,"
                End If
            End With
            
            'Cleanup result
            'DetermineSingleBarChart
            
            ReDim m_DataSet1Colors(UBound(DataFieldNames))
            
    End Select
    
Exit_Here:
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    Set db = Nothing
    Set getData = dataSet.Result
    Exit Function

Handle_Error:
    Select Case Err.Number
        Case Else
            MsgBox Err.Description, vbExclamation, Err.Number
    End Select
    Resume Exit_Here
    Resume Next

End Function


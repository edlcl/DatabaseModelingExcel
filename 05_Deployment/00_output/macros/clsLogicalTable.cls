VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogicalTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================
'-- Database Modeling Excel
'===========================================================
'-- Copyright (c) 2012, Yang Ning (Steven)
'-- All rights reserved.
'-- Email: steven.n.yang@gmail.com
'===========================================================
Option Explicit

Private mTableName As String
Private mDescription As String
Private mPrimaryKey As clsLogicalPrimaryKey
Private mForeignKeys As Collection
Private mIndexes As Collection
Private mColumns As Collection

''' --------------------------------------------------------
''' <summary>
'''     Property: TableName (String)
''' </summary>
''' --------------------------------------------------------
Public Property Get tableName() As String
        tableName = mTableName
End Property
Public Property Let tableName(value As String)
        mTableName = value
End Property

''' --------------------------------------------------------
''' <summary>
'''     Property: Description (String)
''' </summary>
''' --------------------------------------------------------
Public Property Get Description() As String
        Description = mDescription
End Property
Public Property Let Description(value As String)
        mDescription = value
End Property

''' --------------------------------------------------------
''' <summary>
'''     Property: PrimaryKey (clsLogicalPrimaryKey)
''' </summary>
''' --------------------------------------------------------
Public Property Get PrimaryKey() As clsLogicalPrimaryKey
    Set PrimaryKey = mPrimaryKey
End Property
Public Property Set PrimaryKey(value As clsLogicalPrimaryKey)
    Set mPrimaryKey = value
End Property

''' --------------------------------------------------------
''' <summary>
'''     Property: ForeignKeys (Collection)
''' </summary>
''' --------------------------------------------------------
Public Property Get ForeignKeys() As Collection
    Set ForeignKeys = mForeignKeys
End Property
Public Property Set ForeignKeys(value As Collection)
    Set mForeignKeys = value
End Property

''' --------------------------------------------------------
''' <summary>
'''     Property: Indexes (Collection)
''' </summary>
''' --------------------------------------------------------
Public Property Get Indexes() As Collection
    Set Indexes = mIndexes
End Property
Public Property Set Indexes(value As Collection)
    Set mIndexes = value
End Property

''' --------------------------------------------------------
''' <summary>
'''     Property: Columns (Collection)
''' </summary>
''' --------------------------------------------------------
Public Property Get Columns() As Collection
    Set Columns = mColumns
End Property
Public Property Set Columns(value As Collection)
    Set mColumns = value
End Property

Public Function GetPrimaryKeyText() As String
    Dim pkText As String
    pkText = ""
    If Me.PrimaryKey Is Nothing Then GoTo Flag_Exit
    pkText = Me.PrimaryKey.PKcolumns
    
    '-- Return
Flag_Exit:
    GetPrimaryKeyText = pkText
End Function


Public Sub GetPrimaryKeyInfoText(ByRef primaryKeyText As String, _
                    ByRef primaryKeyClustered As String)
    primaryKeyText = ""
    primaryKeyClustered = ""
    
    If Me.PrimaryKey Is Nothing Then Exit Sub
    
    With Me.PrimaryKey
        primaryKeyText = .PKcolumns
        primaryKeyClustered = IIf(.IsClustered, _
                                Table_Sheet_PK_Clustered, _
                                Table_Sheet_PK_NonClustered)
    End With
End Sub

Public Function GetIndexesText() As String
    Dim indexText       As String
    Dim index           As clsLogicalIndex
    indexText = ""
    
    If Me.Indexes Is Nothing Then GoTo Flag_Exit
    
    For Each index In Me.Indexes
        If Len(indexText) > 0 Then
            indexText = indexText & ";" & vbLf
        End If
        indexText = indexText & index.IKColumns
    Next

    '-- Return
Flag_Exit:
    GetIndexesText = indexText
End Function

Public Sub GetIndexexInfoText(ByRef indexText As String, _
                    ByRef indexClustered As String, _
                    ByRef indexUnique As String)
    Dim index           As clsLogicalIndex
    indexText = ""
    indexClustered = ""
    indexUnique = ""
    
    If Me.Indexes Is Nothing Then Exit Sub
    
    For Each index In Me.Indexes
        If Len(indexText) > 0 Then
            indexText = indexText & ";" & vbLf
            indexClustered = indexClustered & ";" & vbLf
            indexUnique = indexUnique & ";" & vbLf
        End If
        indexText = indexText & index.IKColumns
        indexClustered = indexClustered _
                        & IIf(index.IsClustered, _
                                Table_Sheet_Index_Clustered, _
                                Table_Sheet_Index_NonClustered)
        indexUnique = indexUnique _
                        & IIf(index.isUnique, _
                                Table_Sheet_Index_Unique, _
                                Table_Sheet_Index_NonUnique)
    Next
End Sub

Public Function GetForeignKeysText() As String
    Dim fkText       As String
    Dim fk           As clsLogicalForeignKey
    fkText = ""
    
    If Me.ForeignKeys Is Nothing Then GoTo Flag_Exit
    
    For Each fk In Me.ForeignKeys
        If Len(fkText) > 0 Then
            fkText = fkText & ";" & vbLf
        End If
        fkText = fkText & fk.FKcolumns & " " & fk.RefTableAndColumns
        If Len(fk.OnDelete) > 0 Then
            fkText = fkText & " " & fk.OnDelete
        End If
        If Len(fk.OnUpdate) > 0 Then
            fkText = fkText & " " & fk.OnUpdate
        End If
    Next

    '-- Return
Flag_Exit:
    GetForeignKeysText = fkText
End Function

Public Function IsPKColumn(ByVal columnName As String) As String
    Dim retBoo          As Boolean
    Dim PKcolumns()     As String
    Dim index           As Integer
    
    retBoo = False
    If Me.PrimaryKey Is Nothing Then GoTo Flag_Exit
    
    PKcolumns = VBA.Split(Me.PrimaryKey.PKcolumns, ",")
    columnName = LCase(Trim(columnName))
    For index = LBound(PKcolumns) To UBound(PKcolumns)
        If LCase(Trim(PKcolumns(index))) = columnName Then
            retBoo = True
            GoTo Flag_Exit
        End If
    Next

Flag_Exit:
    '-- Return
    IsPKColumn = retBoo
End Function

Public Function IsFKColumn(ByVal columnName As String) As String
    Dim retBoo          As Boolean
    Dim foreignKey      As clsLogicalForeignKey
    Dim FKcolumns()     As String
    Dim index           As Integer
    
    retBoo = False
    If Me.ForeignKeys Is Nothing Then GoTo Flag_Exit
    
    columnName = LCase(Trim(columnName))
    For Each foreignKey In Me.ForeignKeys
        FKcolumns = VBA.Split(foreignKey.FKcolumns, ",")
        For index = LBound(FKcolumns) To UBound(FKcolumns)
            If LCase(Trim(FKcolumns(index))) = columnName Then
                retBoo = True
                GoTo Flag_Exit
            End If
        Next
    Next

Flag_Exit:
    '-- Return
    IsFKColumn = retBoo
End Function


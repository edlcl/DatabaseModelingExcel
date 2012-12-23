VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogicalColumn"
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

Private mColumnLabel    As String
Private mColumnName     As String
Private mDataType       As String
Private mNullable       As Boolean
Private mDefault        As String
Private mComment        As String

''' --------------------------------------------------------
''' <summary>
'''     Property: ColumnLabel (String)
''' </summary>
''' --------------------------------------------------------
Public Property Get columnLabel() As String
        columnLabel = mColumnLabel
End Property
Public Property Let columnLabel(value As String)
        mColumnLabel = value
End Property

''' --------------------------------------------------------
''' <summary>
'''     Property: ColumnName (String)
''' </summary>
''' --------------------------------------------------------
Public Property Get columnName() As String
        columnName = mColumnName
End Property
Public Property Let columnName(value As String)
        mColumnName = value
End Property

''' --------------------------------------------------------
''' <summary>
'''     Property: DataType (String)
''' </summary>
''' --------------------------------------------------------
Public Property Get dataType() As String
        dataType = mDataType
End Property
Public Property Let dataType(value As String)
        mDataType = value
End Property

''' --------------------------------------------------------
''' <summary>
'''     Property: Nullable (Boolean)
''' </summary>
''' --------------------------------------------------------
Public Property Get Nullable() As Boolean
        Nullable = mNullable
End Property
Public Property Let Nullable(value As Boolean)
        mNullable = value
End Property
    
''' --------------------------------------------------------
''' <summary>
'''     Property: Default (String)
''' </summary>
''' --------------------------------------------------------
Public Property Get Default() As String
        Default = mDefault
End Property
Public Property Let Default(value As String)
        mDefault = value
End Property

''' --------------------------------------------------------
''' <summary>
'''     Property: Comment (String)
''' </summary>
''' --------------------------------------------------------
Public Property Get Comment() As String
        Comment = mComment
End Property
Public Property Let Comment(value As String)
        mComment = value
End Property

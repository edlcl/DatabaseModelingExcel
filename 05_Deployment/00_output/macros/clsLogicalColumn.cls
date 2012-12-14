VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogicalColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mColumnLabel    As String
Private mColumnName     As String
Private mDataType       As String
Private mNullable       As Boolean
Private mDefault        As String
Private mNote           As String

''' -----------------------------------------------------------------------------
''' <summary>
'''     Property: ColumnLabel (String)
''' </summary>
''' -----------------------------------------------------------------------------
Public Property Get columnLabel() As String
        columnLabel = mColumnLabel
End Property
Public Property Let columnLabel(value As String)
        mColumnLabel = value
End Property

''' -----------------------------------------------------------------------------
''' <summary>
'''     Property: ColumnName (String)
''' </summary>
''' -----------------------------------------------------------------------------
Public Property Get columnName() As String
        columnName = mColumnName
End Property
Public Property Let columnName(value As String)
        mColumnName = value
End Property

''' -----------------------------------------------------------------------------
''' <summary>
'''     Property: DataType (String)
''' </summary>
''' -----------------------------------------------------------------------------
Public Property Get dataType() As String
        dataType = mDataType
End Property
Public Property Let dataType(value As String)
        mDataType = value
End Property

''' -----------------------------------------------------------------------------
''' <summary>
'''     Property: Nullable (Boolean)
''' </summary>
''' -----------------------------------------------------------------------------
Public Property Get Nullable() As Boolean
        Nullable = mNullable
End Property
Public Property Let Nullable(value As Boolean)
        mNullable = value
End Property
    
''' -----------------------------------------------------------------------------
''' <summary>
'''     Property: Description (String)
''' </summary>
''' -----------------------------------------------------------------------------
Public Property Get Default() As String
        Default = mDefault
End Property
Public Property Let Default(value As String)
        mDefault = value
End Property

''' -----------------------------------------------------------------------------
''' <summary>
'''     Property: Note (String)
''' </summary>
''' -----------------------------------------------------------------------------
Public Property Get Note() As String
        Note = mNote
End Property
Public Property Let Note(value As String)
        mNote = value
End Property

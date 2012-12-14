VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogicalForeignKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mFKName             As String
Private mFKColumns          As String
Private mRefTableName       As String
Private mRefTableColumns    As String
Private mFKOption           As String
Private mOnUpdate           As String
Private mOnDelete           As String

''' -----------------------------------------------------------------------------
''' <summary>
'''     Property: FKName (String)
''' </summary>
''' -----------------------------------------------------------------------------
Public Property Get FKName() As String
        FKName = mFKName
End Property
Public Property Let FKName(value As String)
        mFKName = value
End Property

''' -----------------------------------------------------------------------------
''' <summary>
'''     Property: FKColumns (String)
''' </summary>
''' -----------------------------------------------------------------------------
Public Property Get FKcolumns() As String
        FKcolumns = mFKColumns
End Property
Public Property Let FKcolumns(value As String)
        mFKColumns = value
End Property

''' -----------------------------------------------------------------------------
''' <summary>
'''     Property: RefTableAndColumns (String)
''' </summary>
''' -----------------------------------------------------------------------------
Public Property Get RefTableAndColumns() As String
        RefTableAndColumns = mRefTableName & "(" & mRefTableColumns & ")"
End Property

''' -----------------------------------------------------------------------------
''' <summary>
'''     Property: RefTableName (String)
''' </summary>
''' -----------------------------------------------------------------------------
Public Property Get refTableName() As String
        refTableName = mRefTableName
End Property
Public Property Let refTableName(value As String)
        mRefTableName = value
End Property

''' -----------------------------------------------------------------------------
''' <summary>
'''     Property: RefTableColumns (String)
''' </summary>
''' -----------------------------------------------------------------------------
Public Property Get RefTableColumns() As String
        RefTableColumns = mRefTableColumns
End Property
Public Property Let RefTableColumns(value As String)
        mRefTableColumns = value
End Property

''' -----------------------------------------------------------------------------
''' <summary>
'''     Property: FKOption (String)
'''     like "ON DELETE CASCADE"
''' </summary>
''' -----------------------------------------------------------------------------
Public Property Get fkOption() As String
        fkOption = mFKOption
End Property
Public Property Let fkOption(value As String)
        mFKOption = TrimEx(value)
End Property

''' -----------------------------------------------------------------------------
''' <summary>
'''     Property: OnUpdate (String)
''' </summary>
''' -----------------------------------------------------------------------------
Public Property Get OnUpdate() As String
        OnUpdate = mOnUpdate
End Property
Public Property Let OnUpdate(value As String)
        mOnUpdate = value
End Property

''' -----------------------------------------------------------------------------
''' <summary>
'''     Property: OnDelete (String)
''' </summary>
''' -----------------------------------------------------------------------------
Public Property Get OnDelete() As String
        OnDelete = mOnDelete
End Property
Public Property Let OnDelete(value As String)
        mOnDelete = value
End Property



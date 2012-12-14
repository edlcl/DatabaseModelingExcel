VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogicalPrimaryKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mPKName As String
Private mPKColumns As String
Private mIsClustered As Boolean

''' -----------------------------------------------------------------------------
''' <summary>
'''     Property: PKName (String)
''' </summary>
''' -----------------------------------------------------------------------------
Public Property Get PKName() As String
        PKName = mPKName
End Property
Public Property Let PKName(value As String)
        mPKName = value
End Property

''' -----------------------------------------------------------------------------
''' <summary>
'''     Property: PKColumns (String)
''' </summary>
''' -----------------------------------------------------------------------------
Public Property Get PKcolumns() As String
        PKcolumns = mPKColumns
End Property
Public Property Let PKcolumns(value As String)
        mPKColumns = value
End Property

''' -----------------------------------------------------------------------------
''' <summary>
'''     Property: IsClustered (Boolean)
''' </summary>
''' -----------------------------------------------------------------------------
Public Property Get IsClustered() As Boolean
        IsClustered = mIsClustered
End Property
Public Property Let IsClustered(value As Boolean)
        mIsClustered = value
End Property

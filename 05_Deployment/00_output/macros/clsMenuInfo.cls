VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMenuInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mParent         As String
Private mCaption        As String
Private mTooltipText    As String
Private mOnAction       As String
Private mStyle          As MsoControlType


Private mInstance       As Object
Private mInstanceIndex  As Integer
Private mChildCount     As Integer

''' -----------------------------------------------------------------------------
''' <summary>
'''     Property: Parent (String)
''' </summary>
''' -----------------------------------------------------------------------------
Public Property Get Parent() As String
        Parent = mParent
End Property
Public Property Let Parent(value As String)
        mParent = value
End Property

''' -----------------------------------------------------------------------------
''' <summary>
'''     Property: Caption (String)
''' </summary>
''' -----------------------------------------------------------------------------
Public Property Get Caption() As String
        Caption = mCaption
End Property
Public Property Let Caption(value As String)
        mCaption = value
End Property

''' -----------------------------------------------------------------------------
''' <summary>
'''     Property: TooltipText (String)
''' </summary>
''' -----------------------------------------------------------------------------
Public Property Get TooltipText() As String
        TooltipText = mTooltipText
End Property
Public Property Let TooltipText(value As String)
        mTooltipText = value
End Property

''' -----------------------------------------------------------------------------
''' <summary>
'''     Property: OnAction (String)
''' </summary>
''' -----------------------------------------------------------------------------
Public Property Get OnAction() As String
        OnAction = mOnAction
End Property
Public Property Let OnAction(value As String)
        mOnAction = value
End Property

''' -----------------------------------------------------------------------------
''' <summary>
'''     Property: Style (MsoControlType)
''' </summary>
''' -----------------------------------------------------------------------------
Public Property Get Style() As MsoControlType
        Style = mStyle
End Property
Public Property Let Style(value As MsoControlType)
        mStyle = value
End Property

''' -----------------------------------------------------------------------------
''' <summary>
'''     Property: Instance (Object)
''' </summary>
''' -----------------------------------------------------------------------------
Public Property Get Instance() As Object
        Set Instance = mInstance
End Property
Public Property Set Instance(value As Object)
        Set mInstance = value
End Property

''' -----------------------------------------------------------------------------
''' <summary>
'''     Property: InstanceIndex (Integer)
''' </summary>
''' -----------------------------------------------------------------------------
Public Property Get InstanceIndex() As Integer
        InstanceIndex = mInstanceIndex
End Property
Public Property Let InstanceIndex(value As Integer)
        mInstanceIndex = value
End Property

''' -----------------------------------------------------------------------------
''' <summary>
'''     Property: ChildCount (Integer)
''' </summary>
''' -----------------------------------------------------------------------------
Public Property Get ChildCount() As Integer
        ChildCount = mChildCount
End Property
Public Property Let ChildCount(value As Integer)
        mChildCount = value
End Property


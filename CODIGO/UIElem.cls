VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UIElem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get top() As Integer
End Property
Public Property Get left() As Integer
End Property
Public Property Get width() As Integer
End Property
Public Property Get height() As Integer
End Property
Public Property Let width(ByVal new_width As Integer)
End Property
Public Property Let height(ByVal new_height As Integer)
End Property
Public Property Let top(ByVal new_top As Integer)
End Property
Public Property Let left(ByVal new_left As Integer)
End Property
Public Property Get visible() As Boolean
End Property
Public Property Let visible(ByVal visible As Boolean)
End Property
Public Property Get name() As String
End Property
Public Property Let name(ByVal new_name As String)
End Property
Public Property Get slot() As Long
End Property
Public Property Let slot(ByVal new_slot As Long)
End Property
Public Property Get classname() As String
End Property
Public Property Get parent() As UIElem
End Property
Public Property Let parent(ByRef new_parent As UIElem)
End Property
Public Function addelem(ByRef object As UIElem) As Long
End Function
Public Sub render()
End Sub

Public Property Get eventhandler() As UIEventHandler
End Property

Public Property Let eventhandler(ByRef eh As UIEventHandler)
End Property

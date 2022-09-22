VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UIWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Option Base 0


Public base As New UIBaseObject

Implements UIElem


Private Sub Class_Initialize()
UIElem_name = UIElem_classname()
End Sub

Public Property Get UIElem_top() As Integer
    UIElem_top = base.top
End Property
Public Property Get UIElem_left() As Integer
    UIElem_left = base.left
End Property

Public Property Get UIElem_height() As Integer
    UIElem_height = base.height
End Property
Public Property Let UIElem_width(ByVal new_width As Integer)
    base.width = new_width
End Property

Public Property Get UIElem_width() As Integer
    UIElem_width = base.width
End Property
Public Property Let UIElem_height(ByVal new_height As Integer)
    base.height = new_height
End Property
Public Property Let UIElem_top(ByVal new_top As Integer)
    base.top = new_top
End Property
Public Property Let UIElem_left(ByVal new_left As Integer)
    base.left = new_left
End Property
Public Property Let UIElem_visible(ByVal new_visible As Boolean)
    base.visible = new_visible
End Property
Public Property Get UIElem_visible() As Boolean
    UIElem_visible = base.visible
End Property

Public Function UIElem_addelem(ByRef object As UIElem) As Long
    object.parent = Me
    UIElem_addelem = base.add_ui_object(object)
End Function

Public Property Get UIElem_name() As String
    UIElem_name = base.name
End Property
Public Property Let UIElem_name(ByVal new_name As String)
    base.name = new_name
End Property

Public Property Get UIElem_classname() As String
    UIElem_classname = TypeName(Me)
End Property

Public Property Get UIElem_slot() As Long
    UIElem_slot = base.slot
End Property

Public Property Let UIElem_slot(ByVal new_slot As Long)
    base.slot = new_slot
End Property

Public Property Get UIElem_parent() As UIElem
    UIElem_parent = base.parent
End Property

Public Property Let UIElem_parent(ByRef new_parent As UIElem)
    base.parent = new_parent
End Property

Public Property Get UIElem_eventhandler() As UIEventHandler
    UIElem_eventhandler = base.eventhandler
End Property

Public Property Let UIElem_eventhandler(ByRef eh As UIEventHandler)
    base.eventhandler = eh
End Property

Public Sub UIElem_render()
    If base.visible Then
        With base
        
            Call engine.Engine_Draw_Box_Border(.left, .top, .width, .height _
            , RGBA_From_Comp(20, 196, 255, 200), RGBA_From_Comp(20, 0, 255, 255))
                
            Call .render
        End With
    End If
    
    'Call base.eventhandler.onmousemove(0, 1)
End Sub

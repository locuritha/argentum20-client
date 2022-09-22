VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UILabel"
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
    Set UIElem_parent = base.parent
End Property

Public Property Let UIElem_parent(ByRef new_parent As UIElem)
    base.parent = new_parent
End Property

Public Sub UIElem_render()
    With base
        
        If .visible Then
        If Not .parent Is Nothing And .parent.visible Then
            Dim temp_array(3) As RGBA
            Call RGBAList(temp_array, 0, 0, 100, 250)
            'Call Engine_Text_Render(base.id(), .parent.left, .parent.top, temp_array)
            Call Engine_Text_Render_LetraChica(base.id(), .parent.left, .parent.top, temp_array, 5)
            'Texto As String, x As Integer, y As Integer, ByRef text_color() As RGBA, Optional ByVal font_index As Integer = 1, Optional multi_line As Boolean = False, Optional charindex As Integer = 0, Optional ByVal Alpha As Byte = 255)
    '       Public Sub Engine_Text_Render_LetraChica(Texto As String, x As Integer, y As Integer, ByRef text_color() As RGBA, Optional ByVal font_index As Integer = 1, Optional multi_line As Boolean = False, Optional charindex As Integer = 0, Optional ByVal Alpha As Byte = 255)
        End If
        End If
    End With
End Sub

Public Property Get UIElem_eventhandler() As UIEventHandler
    UIElem_eventhandler = base.eventhandler
End Property

Public Property Let UIElem_eventhandler(ByRef eh As UIEventHandler)
    base.eventhandler = eh
End Property
   

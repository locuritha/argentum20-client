Attribute VB_Name = "engine_ui"
Option Explicit
Option Base 0

Private Const MAX_UI_OBJECTS = 1500
Private uiobjects(MAX_UI_OBJECTS) As UIElem

Private Function get_free_slot() As Long
    get_free_slot = -1
    Dim i As Long
    For i = LBound(uiobjects) To UBound(uiobjects)
        If uiobjects(i) Is Nothing Then
            get_free_slot = i
            Exit Function
        End If
    Next i
End Function

Public Function add_ui_object(ByRef object As UIElem, ByRef event_handler As UIEventHandler) As Long
    Dim slot As Integer: slot = get_free_slot()
    Set uiobjects(slot) = object
    object.slot = slot
    object.eventhandler = event_handler
End Function

Public Sub render()
    Dim i As Long
    For i = LBound(uiobjects) To UBound(uiobjects)
        If Not uiobjects(i) Is Nothing Then
            Call uiobjects(i).render
        End If
    Next i
End Sub



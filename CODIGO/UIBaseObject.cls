VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UIBaseObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Option Base 0


Private m_top As Integer
Private m_left As Integer
Private m_width As Integer
Private m_height As Integer
Private m_id As Long
Private m_z As Integer
Private m_visible As Boolean
Private m_name As String
Private m_slot As Long
Private m_parent As UIElem
Private Const MAX_UI_CHILDREN = 500
Private uiobjects(MAX_UI_CHILDREN) As UIElem
Private m_eventhandler As UIEventHandler

Public Property Get eventhandler() As UIEventHandler
    Set eventhandler = m_eventhandler
End Property

Public Property Let eventhandler(ByRef eh As UIEventHandler)
    Debug.Assert m_eventhandler Is Nothing
    Debug.Assert Not eh Is Nothing
    Set m_eventhandler = eh
End Property

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

Public Function add_ui_object(ByRef object As UIElem) As Long
    Dim slot As Integer: slot = get_free_slot()
    Set uiobjects(slot) = object
    add_ui_object = slot
    Debug.Assert Not object.parent Is Nothing
End Function

Public Sub render_children()
    Dim i As Long
    For i = LBound(uiobjects) To UBound(uiobjects)
        If Not uiobjects(i) Is Nothing Then
            If uiobjects(i).visible Then
                Call uiobjects(i).render
            End If
        Else
            Exit Sub
        End If
    Next i
End Sub

'
'Private Function CrearGrafico(ByVal Archivo As Integer, ByRef texwidth As Long, ByRef textheight As Long) As Direct3DTexture8
'
'    On Error GoTo ErrHandler
'
'    Dim surface_desc As D3DSURFACE_DESC
'    Dim Index        As Integer
'    Dim DirArchico   As String
'    Dim bytArr()     As Byte
'
'    Index = Archivo Mod HASH_TABLE_SIZE
'
'    With TexList(Index)
'        .surfaceCount = .surfaceCount + 1
'        ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
'
'        With .SurfaceEntry(.surfaceCount)
'            .FileName = Archivo
'            .UltimoAcceso = GetTickCount()
'            #If Compresion = 1 Then
'
'                If Not Extract_File_To_Memory(Graphics, App.path & "\..\Recursos\OUTPUT\", LTrim(str(Archivo)) & ".png", bytArr, ResourcesPassword) Then
'                    Debug.Print "¡No se puede cargar el grafico numero " & Archivo & "!"
'                    Exit Function
'
'                End If
'                Set .Texture = mD3D.CreateTextureFromFileInMemoryEx(device, bytArr(0), UBound(bytArr) + 1, D3DX_DEFAULT, D3DX_DEFAULT, 1, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, &HFF000000, ByVal 0, ByVal 0)
'            #Else
'
'                Dim PathToFile As String
'
'                PathToFile = App.path & "\..\Recursos\Graficos\" & LTrim(str(Archivo)) & ".png"
'
'                If Not FileExist(PathToFile, vbArchive) Then
'                    Debug.Print "¡No se puede cargar el grafico numero " & Archivo & "!"
'                    Exit Function
'
'                End If
'
'                Set .Texture = mD3D.CreateTextureFromFileEx(device, PathToFile, D3DX_DEFAULT, D3DX_DEFAULT, 1, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, &HFF000000, ByVal 0, ByVal 0)
'            #End If
'
'            .Texture.GetLevelDesc 0, surface_desc
'            .texture_width = surface_desc.width
'            .texture_height = surface_desc.height
'            .size = surface_desc.size
'            texwidth = .texture_width
'            textheight = .texture_height
'
'            Set CrearGrafico = TexList(Index).SurfaceEntry(TexList(Index).surfaceCount).Texture
'        End With
'
'    End With
'
'    Exit Function
'
'ErrHandler:
'    Debug.Print "ERROR EN GRHLOAD>" & Archivo & ".png"
'
'End Function

Public Property Get top() As Integer
    top = m_top
End Property
Public Property Get left() As Integer
    left = m_left
End Property
Public Property Get width() As Integer
    width = m_width
End Property
Public Property Get height() As Integer
    height = m_height
End Property

Public Property Let width(ByVal new_width As Integer)
    m_width = new_width
    Exit Property
End Property
Public Property Let height(ByVal new_height As Integer)
    m_height = new_height
    Exit Property
End Property

Public Property Let top(ByVal new_top As Integer)
    m_top = new_top
    Exit Property
End Property
Public Property Let left(ByVal new_left As Integer)
    m_left = new_left
    Exit Property
End Property

Public Property Get visible() As Boolean
    visible = m_visible
End Property
Public Property Let visible(ByVal new_visible As Boolean)
    m_visible = new_visible
End Property

Public Property Get name() As String
    name = m_name
End Property
Public Property Let name(ByVal new_name As String)
    m_name = new_name
End Property

Public Property Get id() As String
    id = name & "-" & slot
End Property

Public Property Get slot() As Long
    slot = m_slot
End Property

Public Property Let slot(ByVal new_slot As Long)
    m_slot = new_slot
End Property

Public Property Get parent() As UIElem
    Set parent = m_parent
End Property

Public Property Let parent(ByRef new_parent As UIElem)
    Debug.Assert m_parent Is Nothing
    Set m_parent = new_parent
End Property

Private Sub init_members(ByVal value As Integer)
m_top = value
m_left = value
m_width = value
m_height = value
m_id = value
m_visible = value
m_slot = -1
Set m_parent = Nothing
End Sub


Private Sub Class_Initialize()
Call init_members(0)
m_visible = True
End Sub

Private Sub Class_Terminate()
Call init_members(1)
m_visible = False
End Sub


Public Sub render()

#If DEBUGGING = 1 Then
            Dim temp_array(3) As RGBA
            Call RGBAList(temp_array, 0, 0, 100, 250)
            Call Engine_Text_Render(id(), left, top, temp_array)
#End If
Call render_children
    
End Sub

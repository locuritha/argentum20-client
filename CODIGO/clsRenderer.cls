VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'    Argentum 20 - Game Client Program
'    Copyright (C) 2023 - Noland Studios
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.

Option Base 0

Option Explicit


Private mDirectX As New DirectX8
Private mDirectD3D8 As D3DX8
Private mDirectD3D As Direct3D8

Private mDirectDevice As Direct3DDevice8

Private mD3DPP As D3DPRESENT_PARAMETERS

Private mProjection As D3DMATRIX
Private mIdentityMatrix As D3DMATRIX

Private mDisplayMode As D3DDISPLAYMODE

Const ScreenWidth  As Integer = 1024
Const ScreenHeight As Integer = 768

Const FVF_UIVERTEX = (D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR)
Const FVF_VERTEX = (D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1)
Const QuadDim = 1
Const MapSize As Integer = 10

Private Type UIVERTEX
    x As Single
    y As Single
    z As Single
    rhw As Single
    color As Long
    Specular As Long
    u As Single
    v As Single
End Type

Private Type VERTEX
    x As Single
    y As Single
    z As Single
    nX As Single
    nY As Single
    nz As Single
    u As Single
    v As Single
End Type

Dim textures(1 To 10) As Direct3DTexture8

Dim QuadVerts(0 To 5) As UIVERTEX ' 2d quad vertexs
Dim PlaneVerts(0 To 1000) As VERTEX ' 3d plane vertexs
Dim MapVerts(0 To 1000) As VERTEX ' 3d plane vertexs

Dim objectVertexBuffer As Direct3DVertexBuffer8 ' Object vertex buffer
Dim uiVertexBuffer As Direct3DVertexBuffer8 ' UI vertex buffer
Dim mapVertexBuffer As Direct3DVertexBuffer8 ' Map vertex buffer
Dim mapIndexBuffer As Direct3DIndexBuffer8 ' Map index buffer
Dim offsetObject As Integer ' object vertex pos to write


Dim matProj As D3DMATRIX
Dim matView As D3DMATRIX
Dim matWorld As D3DMATRIX

Private Sub get_game_resolution(ByRef mode As D3DDISPLAYMODE)
    'For the time being we hard code it to 1024x768x32
    mode.Width = 1024
    mode.Height = 768
    mode.RefreshRate = 60
    mode.format = D3DFMT_A8R8G8B8
End Sub

Public Sub list_modes(ByRef d3d As Direct3D8)
    Dim tmpDispMode As D3DDISPLAYMODE
    Dim i As Long
    For i = 0 To d3d.GetAdapterModeCount(0) - 1 'primary adapter
        Call d3d.EnumAdapterModes(0, i, tmpDispMode)
        Debug.Print tmpDispMode.Width & "x" & tmpDispMode.Height & " fmt:" & tmpDispMode.format
        
    Next i
End Sub
Private Function init_dx8_and_d3d() As Long
On Error Resume Next
    Err.Clear
    Set mDirectX = New DirectX8
    If Err.Number <> 0 Then
                Call MsgBox("Fatal error creating DirectX8 objetct", vbCritical, App.title)
                Debug.Print "Error Number Returned: " & Err.Number
                Exit Function
    End If
    
    Set mDirectD3D = DirectX.Direct3DCreate()
    If Err.Number <> 0 Then
                Call MsgBox("Fatal error creating DirectD3D objetct", vbCritical, App.title)
                Debug.Print "Error Number Returned: " & Err.Number
                Exit Function
    End If
    
    Set mDirectD3D8 = New D3DX8
    If Err.Number <> 0 Then
                Call MsgBox("Fatal error creating DirectD3D8 objetct", vbCritical, App.title)
                Debug.Print "Error Number Returned: " & Err.Number
                Exit Function
    End If
    init_dx8_and_d3d = Err.Number
End Function

Public Function Init(ByVal window As Long) As Long
    Init = init_dx8_and_d3d()
    If Init <> 0 Then
        Debug.Print "init_dx8_and_d3d() failed with code " & Init
        Exit Function
    End If
    Init = init_dx_device(window, mDirectD3D, mDisplayMode)
    If Init <> 0 Then
        Debug.Print "init_dx_device() failed with code " & Init
        Exit Function
    End If
    Call InitScene
End Function

Private Function init_dx_device(ByVal window_handle As Long, _
                                ByRef d3d_obj As Direct3D8, _
                                ByRef disp_mode As D3DDISPLAYMODE) As Long
On Error Resume Next
    Debug.Assert Not d3d_obj Is Nothing
    Dim Caps As D3DCAPS8
    Dim DevType As CONST_D3DDEVTYPE
    Dim DevBehaviorFlags As Long
    Dim d3dDispMode  As D3DDISPLAYMODE
    Err.Clear
    d3d_obj.GetDeviceCaps D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Caps
    If Err.Number = D3DERR_NOTAVAILABLE Then
        Debug.Print "HAL Is Not available, using; software; vertex; processing"
        DevType = D3DDEVTYPE_REF
        DevBehaviorFlags = D3DCREATE_SOFTWARE_VERTEXPROCESSING
    Else
        DevType = D3DDEVTYPE_HAL
        Debug.Print "VertexProcessingCaps = " & Caps.VertexProcessingCaps
        If Caps.VertexProcessingCaps = 0 Then
            Debug.Print "HAL Is available, " & "Using; software; vertex; processing"
            DevBehaviorFlags = D3DCREATE_SOFTWARE_VERTEXPROCESSING
        ElseIf Caps.VertexProcessingCaps = &H4B Then
            Debug.Print "HAL Is available, " & "Using; hardware; vertex; processing; "
            DevBehaviorFlags = D3DCREATE_HARDWARE_VERTEXPROCESSING
        Else
            Debug.Print "HAL Is available, " & "Using; mixed; vertex; processing; "
            DevBehaviorFlags = D3DCREATE_MIXED_VERTEXPROCESSING
        End If
    End If
    d3d_obj.GetAdapterDisplayMode D3DADAPTER_DEFAULT, d3dDispMode
    
    Debug.Print "Using; Windowed; mode"
    Debug.Assert disp_mode.Width > 0
    Debug.Assert disp_mode.Height > 0
    mD3DPP.Windowed = 1
    mD3DPP.BackBufferWidth = disp_mode.Width
    mD3DPP.BackBufferHeight = disp_mode.Height
    mD3DPP.BackBufferFormat = disp_mode.format
    mD3DPP.SwapEffect = D3DSWAPEFFECT_DISCARD
    mD3DPP.BackBufferCount = 1
    mD3DPP.hDeviceWindow = window_handle
    mD3DPP.EnableAutoDepthStencil = 1
    mD3DPP.AutoDepthStencilFormat = D3DFMT_D16
    
    Err.Clear
 
    Set mDirectDevice = d3d_obj.CreateDevice(D3DADAPTER_DEFAULT, DevType, window_handle, DevBehaviorFlags, mD3DPP)
    
    Debug.Print "Create; Direct3D; device: ", Err
    If (Err.Number <> 0) Then
        'if we failed to create the device with D3DFMT_A8R8G8B8 we try to do it with current display fmt
        D3DWindow.BackBufferFormat = d3dDispMode.format
        Err.Clear
        Set mDirectDevice = d3d_obj.CreateDevice(D3DADAPTER_DEFAULT, DevType, window_handle, DevBehaviorFlags, mD3DPP)
        Debug.Print "Create; Direct3D; device: ", Err
    End If
    init_dx_device = Err.Number
End Function

Private Sub Class_Initialize()
    Set mDirectX = Nothing
    Set mDirectD3D8 = Nothing
    Set mDirectD3D = Nothing
    Set mDirectDevice = Nothing
    mDisplayMode.Width = ScreenWidth
    mDisplayMode.Height = ScreenHeight
    mDisplayMode.RefreshRate = 60
    mDisplayMode.format = D3DFMT_A8R8G8B8
End Sub

Private Function CreateUIVertex(x As Single, y As Single, rhw As Single, _
                                                color As Long, Specular As Long, u As Single, _
                                                v As Single) As UIVERTEX
    CreateUIVertex.x = x
    CreateUIVertex.y = y
    CreateUIVertex.z = 0
    CreateUIVertex.rhw = rhw
    CreateUIVertex.color = color
    CreateUIVertex.Specular = Specular
    CreateUIVertex.u = u
    CreateUIVertex.v = v
End Function

Private Function CreateVertex(x As Single, y As Single, z As Single, _
                                            u As Single, v As Single) As VERTEX
    CreateVertex.x = x
    CreateVertex.y = y
    CreateVertex.z = z
    CreateVertex.u = u
    CreateVertex.v = v
End Function

Private Sub initUI(PosX As Single, PosY As Single, Width As Single, Height As Single)
    QuadVerts(0) = CreateUIVertex(PosX, PosY, 1, &HFF0000, 0, 0, 0)
    QuadVerts(1) = CreateUIVertex(PosX + Width, PosY, 1, &HFF0000, 0, 1, 0)
    QuadVerts(2) = CreateUIVertex(PosX, PosY + Height, 1, &HFF0000, 0, 0, 1)
    QuadVerts(3) = CreateUIVertex(PosX, PosY + Height, 1, &HFF&, 0, 0, 1)
    QuadVerts(4) = CreateUIVertex(PosX + Width, PosY, 1, &HFF&, 0, 1, 0)
    QuadVerts(5) = CreateUIVertex(PosX + Width, PosY + Height, 1, &HFF&, 0, 1, 1)
        
    ' set up the vertex buffer
    Set uiVertexBuffer = mDirectDevice.CreateVertexBuffer(Len(QuadVerts(0)) * 6, D3DUSAGE_WRITEONLY, FVF_UIVERTEX, D3DPOOL_MANAGED)
    If uiVertexBuffer Is Nothing Then Debug.Print "ERROR: can't create the vertex buffer": Exit Sub
    
    D3DVertexBuffer8SetData uiVertexBuffer, 0, Len(QuadVerts(0)) * 6, 0, QuadVerts(0)
End Sub

Private Sub initMap2()
    Dim x As Integer
    Dim y As Integer
    Dim cel As Integer
    Dim vList(0 To 24) As VERTEX

    cel = 0
    For y = 0 To ((MapSize - 1) * QuadDim) Step QuadDim
        For x = 0 To ((MapSize * QuadDim - 1)) Step QuadDim
            MapVerts(cel) = CreateVertex(CSng(x), CSng(y), 1, 0, 0)
            MapVerts(cel + 1) = CreateVertex(CSng(x) + QuadDim, CSng(y), 1, 1, 0)
            MapVerts(cel + 2) = CreateVertex(CSng(x), CSng(y) + QuadDim, 1, 0, 1)
        
            MapVerts(cel + 3) = CreateVertex(CSng(x), CSng(y) + QuadDim, 1, 0, 1)
            MapVerts(cel + 4) = CreateVertex(CSng(x) + QuadDim, CSng(y), 1, 1, 0)
            MapVerts(cel + 5) = CreateVertex(CSng(x) + QuadDim, CSng(y) + QuadDim, 1, 1, 1)
            cel = cel + 6
        Next x
    Next y
    
    Set mapVertexBuffer = mDirectDevice.CreateVertexBuffer(Len(MapVerts(0)) * (MapSize * MapSize * 6), 0, FVF_VERTEX, D3DPOOL_MANAGED)
    D3DVertexBuffer8SetData mapVertexBuffer, 0, Len(MapVerts(0)) * (MapSize * MapSize * 6), 0, MapVerts(0)
    
End Sub

Private Sub initMap()
Dim x As Integer
Dim y As Integer
Dim Index As Integer
Dim cel As Integer
Dim aux As Integer

Dim numVerts As Integer
numVerts = (MapSize + 1) * (MapSize + 1)
Dim vList(0 To ((MapSize + 1) * (MapSize + 1) - 1)) As LITVERTEX
Dim iList(0 To (MapSize * MapSize * 6 - 1)) As Integer

cel = 0
For y = 0 To (MapSize * 32) Step 32
    For x = 0 To (MapSize * 32) Step 32
        'vList(y * mapSize + x) = CreateLitVertex(CSng(x), CSng(y), 0, &HFF0000, 0, 0, 0)
        vList(cel) = CreateLitVertex(CSng(x), CSng(y), 0, &HFF0000, 0, 0, 0)
        cel = cel + 1
    Next x
Next y

' creating vertex buffer
Set mapVertexBuffer = mDirectDevice.CreateVertexBuffer(Len(vList(0)) * numVerts, 0, FVF_LVERTEX, D3DPOOL_MANAGED)
If mapVertexBuffer Is Nothing Then Debug.Print "ERROR: Can't creating map vertex buffer": Exit Sub
D3DVertexBuffer8SetData mapVertexBuffer, 0, Len(vList(0)) * numVerts, 0, vList(0)
cel = 0
y = 0
aux = 1
'For y = 0 To (mapSize * mapSize - 1)
Do While (y < (MapSize * MapSize + MapSize))
        iList(cel) = y: iList(cel + 1) = y + 1: iList(cel + 2) = y + MapSize + 1
        iList(cel + 3) = y + MapSize + 1: iList(cel + 4) = y + 1: iList(cel + 5) = y + MapSize + 2
        y = y + 1
        If (aux = MapSize) Then
            y = y + 1
            aux = 0
        End If
        aux = aux + 1
        cel = cel + 6
Loop
'Next y

' create index buffer
Set mapIndexBuffer = mDirectDevice.CreateIndexBuffer(Len(iList(0)) * (MapSize * MapSize * 6), 0, D3DFMT_INDEX16, D3DPOOL_MANAGED)
If mapIndexBuffer Is Nothing Then Debug.Print "ERROR: creating map index buffer": Exit Sub
D3DIndexBuffer8SetData mapIndexBuffer, 0, Len(iList(0)) * (MapSize * MapSize * 6), 0, iList(0)

End Sub

Private Sub initObjects()
    Dim x As Integer
    Dim y As Integer
    Dim texQuad As Single
    Dim depth As Single
    
    x = 0
    y = 0
    texQuad = 0.09
    offsetObject = 0
    depth = 0.9
    ' Esqueleto
    'PlaneVerts(0) = CreateVertex(CSng(x), CSng(y), depth, 0, 0)
    'PlaneVerts(1) = CreateVertex(CSng(x) + QuadDim, CSng(y), depth, texQuad, 0)
    'PlaneVerts(2) = CreateVertex(CSng(x), CSng(y) + QuadDim, depth, 0, texQuad)
        
    'PlaneVerts(3) = CreateVertex(CSng(x), CSng(y) + QuadDim, depth, 0, texQuad)
    'PlaneVerts(4) = CreateVertex(CSng(x) + QuadDim, CSng(y), depth, texQuad, 0)
    'PlaneVerts(5) = CreateVertex(CSng(x) + QuadDim, CSng(y) + QuadDim, depth, texQuad, texQuad)
    
    
    PlaneVerts(0) = CreateVertex(CSng(x), CSng(y), depth, texQuad, texQuad)
    PlaneVerts(1) = CreateVertex(CSng(x) + QuadDim, CSng(y), depth, 0, texQuad)
    PlaneVerts(2) = CreateVertex(CSng(x), CSng(y) + QuadDim, depth, texQuad, 0)
    
    PlaneVerts(3) = CreateVertex(CSng(x), CSng(y) + QuadDim, depth, texQuad, 0)
    PlaneVerts(4) = CreateVertex(CSng(x) + QuadDim, CSng(y), depth, 0, texQuad)
    PlaneVerts(5) = CreateVertex(CSng(x) + QuadDim, CSng(y) + QuadDim, depth, 0, 0)
    ' ------------------------------------------------------------------------
    PlaneVerts(6) = CreateVertex(CSng(x), CSng(y), depth, texQuad * 1, 0)
    PlaneVerts(7) = CreateVertex(CSng(x) + QuadDim, CSng(y), depth, texQuad * 2, 0)
    PlaneVerts(8) = CreateVertex(CSng(x), CSng(y) + QuadDim, depth, 0, texQuad)
        
    PlaneVerts(9) = CreateVertex(CSng(x), CSng(y) + QuadDim, depth, texQuad * 1, texQuad)
    PlaneVerts(10) = CreateVertex(CSng(x) + QuadDim, CSng(y), depth, texQuad * 2, 0)
    PlaneVerts(11) = CreateVertex(CSng(x) + QuadDim, CSng(y) + QuadDim, depth, texQuad * 2, texQuad)
    ' ------------------------------------------------------------------------
    
    PlaneVerts(12) = CreateVertex(CSng(x), CSng(y), depth, texQuad * 2, 0)
    PlaneVerts(13) = CreateVertex(CSng(x) + QuadDim, CSng(y), depth, texQuad * 3, 0)
    PlaneVerts(14) = CreateVertex(CSng(x), CSng(y) + QuadDim, depth, 0, texQuad)
        
    PlaneVerts(15) = CreateVertex(CSng(x), CSng(y) + QuadDim, depth, texQuad * 1, texQuad)
    PlaneVerts(16) = CreateVertex(CSng(x) + QuadDim, CSng(y), depth, texQuad * 3, 0)
    PlaneVerts(17) = CreateVertex(CSng(x) + QuadDim, CSng(y) + QuadDim, depth, texQuad * 3, texQuad)
    ' ------------------------------------------------------------------------
    PlaneVerts(18) = CreateVertex(CSng(x), CSng(y), depth, texQuad * 3, 0)
    PlaneVerts(19) = CreateVertex(CSng(x) + QuadDim, CSng(y), depth, texQuad * 4, 0)
    PlaneVerts(20) = CreateVertex(CSng(x), CSng(y) + QuadDim, depth, 0, texQuad)
        
    PlaneVerts(21) = CreateVertex(CSng(x), CSng(y) + QuadDim, depth, texQuad * 3, texQuad)
    PlaneVerts(22) = CreateVertex(CSng(x) + QuadDim, CSng(y), depth, texQuad * 4, 0)
    PlaneVerts(23) = CreateVertex(CSng(x) + QuadDim, CSng(y) + QuadDim, depth, texQuad * 4, texQuad)
    ' ------------------------------------------------------------------------
    PlaneVerts(24) = CreateVertex(CSng(x), CSng(y), depth, texQuad * 4, 0)
    PlaneVerts(25) = CreateVertex(CSng(x) + QuadDim, CSng(y), depth, texQuad * 5, 0)
    PlaneVerts(26) = CreateVertex(CSng(x), CSng(y) + QuadDim, depth, 0, texQuad)
        
    PlaneVerts(27) = CreateVertex(CSng(x), CSng(y) + QuadDim, depth, texQuad * 4, texQuad)
    PlaneVerts(28) = CreateVertex(CSng(x) + QuadDim, CSng(y), depth, texQuad * 5, 0)
    PlaneVerts(29) = CreateVertex(CSng(x) + QuadDim, CSng(y) + QuadDim, depth, texQuad * 5, texQuad)
    
    Set objectVertexBuffer = mDirectDevice.CreateVertexBuffer(Len(PlaneVerts(0)) * 30, 0, FVF_VERTEX, D3DPOOL_MANAGED)
    D3DVertexBuffer8SetData objectVertexBuffer, 0, Len(PlaneVerts(0)) * 30, 0, PlaneVerts(0)
End Sub

Private Sub InitScene()
            
    Call loadTextures
    ' World Matrix
    D3DXMatrixIdentity matWorld
    mDirectDevice.SetTransform D3DTS_WORLD, matWorld

    ' View Matrix
    'D3DXMatrixLookAtLH matView, MakeVector(0, 5, 0), MakeVector(0, 0, 0), MakeVector(0, 0, 1)
    D3DXMatrixLookAtLH matView, MakeVector(5, 5, -10), MakeVector(5, 5, 1), MakeVector(0, 1, 0)
    mDirectDevice.SetTransform D3DTS_VIEW, matView

    ' Projection Matrix
    D3DXMatrixPerspectiveFovLH matProj, PI / 3, 1, 0.1, 100
    mDirectDevice.SetTransform D3DTS_PROJECTION, matProj
    Call initUI(300, 300, 300, 300)
    Call initMap2
    Call initObjects

End Sub

Public Sub render()
    On Error GoTo clsRender_Err
    
    mDirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, D3DColorRGBA(0, 150, 150, 255), 1#, 0
    mDirectDevice.BeginScene
    
'    Call prepareRender
'    Call renderMap2
    
'    Call prepareObjects
'    Call renderObjects
    
'    Call prepareUI
'    Call renderUI
    Call RenderScreen
        
    Call mDirectDevice.EndScene
      
    Call mDirectDevice.Present(ByVal 0, ByVal 0, 0, ByVal 0)
    Exit Sub
clsRender_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsRenderer.render", Erl)
    Resume Next
End Sub

Private Sub prepareUI()
    mDirectDevice.SetRenderState D3DRS_LIGHTING, 0  ' lighting off
    mDirectDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE ' No culling mode
    mDirectDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
    mDirectDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
    mDirectDevice.SetTexture 0, textures(1)
    mDirectDevice.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_SELECTARG1
    mDirectDevice.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
    mDirectDevice.SetVertexShader FVF_UIVERTEX ' render 2D vertex

End Sub

Private Sub renderUI()
    mDirectDevice.SetStreamSource 0, uiVertexBuffer, Len(QuadVerts(0))
    mDirectDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2

End Sub

Private Sub prepareRender()
    mDirectDevice.SetRenderState D3DRS_ZENABLE, True ' enable Z buffer
    mDirectDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE ' No culling mode
    mDirectDevice.SetRenderState D3DRS_LIGHTING, False  ' lighting off
    mDirectDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
    mDirectDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
    mDirectDevice.SetTexture 0, textures(2)
    mDirectDevice.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_SELECTARG1
    mDirectDevice.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
    mDirectDevice.SetVertexShader FVF_VERTEX
  
End Sub

Private Sub renderMap2()
    mDirectDevice.SetStreamSource 0, mapVertexBuffer, Len(MapVerts(0))
    mDirectDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, (MapSize * MapSize * 6)

End Sub

Private Sub renderMap()
    mDirectDevice.SetStreamSource 0, mapVertexBuffer, Len(MapVerts(0))
    mDirectDevice.SetIndices mapIndexBuffer, 0
    mDirectDevice.DrawIndexedPrimitive D3DPT_TRIANGLELIST, 0, (MapSize * MapSize * 6), 0, (MapSize * MapSize * 6) / 3

End Sub

Private Sub prepareObjects()
    mDirectDevice.SetRenderState D3DRS_LIGHTING, 0  ' lighting off
    mDirectDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE ' No culling mode
    mDirectDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
    mDirectDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
    mDirectDevice.SetTexture 0, textures(3)
    mDirectDevice.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_SELECTARG1
    mDirectDevice.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
    mDirectDevice.SetVertexShader FVF_VERTEX

End Sub

Private Sub renderObjects()
    mDirectDevice.SetStreamSource 0, objectVertexBuffer, Len(PlaneVerts(0))
    mDirectDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2

End Sub

Private Sub addQuad()

End Sub

Private Sub addPlane()


End Sub

Private Sub loadTextures()
    Set textures(1) = mDirectD3D8.CreateTextureFromFile(mDirectDevice, App.path + "\textura.bmp")
    If textures(1) Is Nothing Then Exit Sub
    Set textures(2) = mDirectD3D8.CreateTextureFromFile(mDirectDevice, App.path + "\grass.png")
    If textures(2) Is Nothing Then Exit Sub
    Set textures(3) = mDirectD3D8.CreateTextureFromFile(mDirectDevice, App.path + "\esqueleto.png")
    If textures(3) Is Nothing Then Exit Sub
End Sub

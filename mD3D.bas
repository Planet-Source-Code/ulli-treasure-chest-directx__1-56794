Attribute VB_Name = "mD3D"
Option Explicit

Public Dx               As DirectX8
Public D3D              As Direct3D8
Public D3DDevice        As Direct3DDevice8
Private D3DX            As D3DX8
Private vBuffWaves      As Direct3DVertexBuffer8
Private vBuffChest      As Direct3DVertexBuffer8
Private TexWaves        As Direct3DTexture8
Private TexChest        As Direct3DTexture8
Private TextRect        As RECT
Private MainFont        As D3DXFont
Private MainFontDesc    As IFont
Private Fnt             As New StdFont
Private RotateAngle     As Single
Private IncRot          As Single
Private matWorld        As D3DMATRIX
Private matView         As D3DMATRIX
Private matProj         As D3DMATRIX
Private matTemp         As D3DMATRIX

Private Type LITVERTEX
    x       As Single
    y       As Single
    z       As Single
    Color   As Long
    Specular As Long
    tu      As Single
    tv      As Single
End Type
Private Waves(0 To 61)  As LITVERTEX
Private Chest(0 To 36)  As LITVERTEX
Private Const LitFVF    As Long = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX1)
Public Pi               As Single
Public FPSCurrent       As Long
Public Busy             As Boolean

Private Sub CreateChest(Move As Single, Size As Single, x As Single, y As Single, z As Single)

  Dim sy As Single

    RotateAngle = RotateAngle + IncRot
    If RotateAngle >= 0.6 Or RotateAngle <= -0.6 Then
        IncRot = -IncRot
    End If

    sy = Sin(Move) * 3 + y - 2

    'Front
    Chest(0) = CreateLitVertex(x - Size, sy - Size, z + Size, vbWhite, 0, 0, 0)
    Chest(1) = CreateLitVertex(x + Size, sy + Size, z + Size, vbWhite, 0, 1, 1)
    Chest(2) = CreateLitVertex(x - Size, sy + Size, z + Size, vbWhite, 0, 0, 1)
    Chest(3) = CreateLitVertex(x + Size, sy + Size, z + Size, vbWhite, 0, 1, 1)
    Chest(4) = CreateLitVertex(x - Size, sy - Size, z + Size, vbWhite, 0, 0, 0)
    Chest(5) = CreateLitVertex(x + Size, sy - Size, z + Size, vbWhite, 0, 1, 0)

    'Back
    Chest(6) = CreateLitVertex(x - Size, sy + Size, z - Size, vbWhite, 0, 0, 1)
    Chest(7) = CreateLitVertex(x + Size, sy + Size, z - Size, vbWhite, 0, 1, 0)
    Chest(8) = CreateLitVertex(x - Size, sy - Size, z - Size, vbWhite, 0, 0, 0)
    Chest(9) = CreateLitVertex(x + Size, sy + Size, z - Size, vbWhite, 0, 0, 1)
    Chest(10) = CreateLitVertex(x - Size, sy - Size, z - Size, vbWhite, 0, 1, 0)
    Chest(11) = CreateLitVertex(x + Size, sy - Size, z - Size, vbWhite, 0, 0, 0)

    'Right
    Chest(12) = CreateLitVertex(x - Size, sy + Size, z - Size, vbWhite, 0, 0, 1)
    Chest(13) = CreateLitVertex(x - Size, sy + Size, z + Size, vbWhite, 0, 1, 0)
    Chest(14) = CreateLitVertex(x - Size, sy - Size, z - Size, vbWhite, 0, 0, 0)
    Chest(15) = CreateLitVertex(x - Size, sy + Size, z + Size, vbWhite, 0, 0, 1)
    Chest(16) = CreateLitVertex(x - Size, sy - Size, z - Size, vbWhite, 0, 1, 0)
    Chest(17) = CreateLitVertex(x - Size, sy - Size, z + Size, vbWhite, 0, 0, 0)

    'Left
    Chest(18) = CreateLitVertex(x + Size, sy + Size, z - Size, vbWhite, 0, 0, 1)
    Chest(19) = CreateLitVertex(x + Size, sy + Size, z + Size, vbWhite, 0, 1, 0)
    Chest(20) = CreateLitVertex(x + Size, sy - Size, z - Size, vbWhite, 0, 0, 0)
    Chest(21) = CreateLitVertex(x + Size, sy + Size, z + Size, vbWhite, 0, 0, 1)
    Chest(22) = CreateLitVertex(x + Size, sy - Size, z - Size, vbWhite, 0, 1, 0)
    Chest(23) = CreateLitVertex(x + Size, sy - Size, z + Size, vbWhite, 0, 0, 0)

    'Top
    Chest(24) = CreateLitVertex(x - Size, sy + Size, z + Size, vbWhite, 0, 0, 1)
    Chest(25) = CreateLitVertex(x + Size, sy + Size, z + Size, vbWhite, 0, 1, 0)
    Chest(26) = CreateLitVertex(x - Size, sy + Size, z - Size, vbWhite, 0, 0, 0)
    Chest(27) = CreateLitVertex(x + Size, sy + Size, z + Size, vbWhite, 0, 0, 1)
    Chest(28) = CreateLitVertex(x - Size, sy + Size, z - Size, vbWhite, 0, 1, 0)
    Chest(29) = CreateLitVertex(x + Size, sy + Size, z - Size, vbWhite, 0, 0, 0)

    'Bottom
    Chest(30) = CreateLitVertex(x - Size, sy - Size, z + Size, vbWhite, 0, 0, 1)
    Chest(31) = CreateLitVertex(x + Size, sy - Size, z + Size, vbWhite, 0, 1, 0)
    Chest(32) = CreateLitVertex(x - Size, sy - Size, z - Size, vbWhite, 0, 0, 0)
    Chest(33) = CreateLitVertex(x + Size, sy - Size, z + Size, vbWhite, 0, 0, 1)
    Chest(34) = CreateLitVertex(x - Size, sy - Size, z - Size, vbWhite, 0, 1, 0)
    Chest(35) = CreateLitVertex(x + Size, sy - Size, z - Size, vbWhite, 0, 0, 0)

    Set vBuffChest = D3DDevice.CreateVertexBuffer(Len(Chest(0)) * 36, 0, LitFVF, D3DPOOL_DEFAULT)
    D3DVertexBuffer8SetData vBuffChest, 0, Len(Chest(0)) * 36, 0, Chest(0)

End Sub

Private Function CreateLitVertex(x As Single, y As Single, z As Single, Diffuse As Long, Specular As Long, tu As Single, tv As Single) As LITVERTEX

    With CreateLitVertex
        .x = x
        .y = y
        .z = z
        .Color = Diffuse
        .Specular = Specular
        .tu = tu
        .tv = tv
    End With 'CREATELITVERTEX

End Function

Private Sub CreateWaves(Move As Single, WaveHeight As Single, WaveSpeed As Single, Länge As Single, Breite As Single)

  Dim x As Single

    For x = -15 To 15
        Waves((x + 15) * 2 + 0) = CreateLitVertex(x * Länge, Sin(Move + x) * WaveHeight, Breite, vbWhite, 0, ((x + 15) / 30) + Move / WaveSpeed, 1)
        Waves((x + 15) * 2 + 1) = CreateLitVertex(x * Länge, Sin(Move + x) * WaveHeight, -Breite, vbYellow, 0, ((x + 15) / 30) + Move / WaveSpeed, 0)
    Next x

    Set vBuffWaves = D3DDevice.CreateVertexBuffer(Len(Waves(0)) * 120, 0, LitFVF, D3DPOOL_DEFAULT)
    D3DVertexBuffer8SetData vBuffWaves, 0, Len(Waves(0)) * 120, 0, Waves(0)

End Sub

Public Function Initialise() As Boolean

  Dim DispMode As D3DDISPLAYMODE
  Dim D3DWindow As D3DPRESENT_PARAMETERS

    Pi = 4 * Atn(1)
    IncRot = 0.25 * Pi / 180

    Set Dx = New DirectX8
    Set D3D = Dx.Direct3DCreate()
    Set D3DX = New D3DX8
    
    DispMode.Format = D3DFMT_X8R8G8B8
    
    With D3DWindow
        .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC '= D3DSWAPEFFECT_FLIP
        .BackBufferCount = 1
        .Windowed = True
        .BackBufferFormat = DispMode.Format
        .hDeviceWindow = fTreasureChest.hWnd
        .EnableAutoDepthStencil = 1
        If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, D3DRTYPE_SURFACE, D3DFMT_D16) = D3D_OK Then
            .AutoDepthStencilFormat = D3DFMT_D16
        End If
    End With 'D3DWINDOW
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, fTreasureChest.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, D3DWindow)
    With D3DDevice
        .SetVertexShader LitFVF
        .SetRenderState D3DRS_LIGHTING, 0
        .SetRenderState D3DRS_CULLMODE, 1
        .SetRenderState D3DRS_ZENABLE, 1
        .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
        .SetRenderState D3DRS_SRCBLEND, 2 '2 (light) or 3 (litle darker)
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        Set TexWaves = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\water.jpg", 128, 128, D3DX_DEFAULT, 0, DispMode.Format, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
        Set TexChest = D3DX.CreateTextureFromFile(D3DDevice, App.Path & "\chest.jpg")

        TextRect.Top = 0
        TextRect.bottom = 20
        TextRect.Right = 75
        Fnt.Name = "Arial"
        Fnt.Size = 12
        Fnt.Bold = True
        Set MainFontDesc = Fnt
        Set MainFont = D3DX.CreateFont(D3DDevice, MainFontDesc.hFont)

        D3DXMatrixIdentity matWorld
        .SetTransform D3DTS_WORLD, matWorld

        D3DXMatrixLookAtLH matView, MakeVector(-10, 35, 50), MakeVector(0, 0, 0), MakeVector(0, 1, 0)
        .SetTransform D3DTS_VIEW, matView '0,80,100

        D3DXMatrixPerspectiveFovLH matProj, 0.4, 1, 0.1, 5000
        .SetTransform D3DTS_PROJECTION, matProj
    End With 'D3DDEVICE

    Initialise = True

End Function

Private Function MakeVector(x As Single, y As Single, z As Single) As D3DVECTOR

    With MakeVector
        .x = x
        .y = y
        .z = z
    End With 'MAKEVECTOR

End Function

Public Sub Render(Move As Single)

  Dim Aspect As Single
  Const AspectRatio      As Single = 0.45

    CreateWaves Move, 3, 50, 3, 50
    CreateChest Move, 4, 0, 0, 0
    Aspect = fTreasureChest.ScaleHeight / fTreasureChest.ScaleWidth
    D3DXMatrixPerspectiveFovLH matProj, AspectRatio + IIf(Aspect > 1, (Aspect - 1) * AspectRatio, 0), Aspect, 0.1, 20000
    With D3DDevice
        .SetTransform D3DTS_PROJECTION, matProj
        .BeginScene
        D3DXMatrixIdentity matWorld
        D3DXMatrixIdentity matTemp
        D3DXMatrixRotationX matTemp, RotateAngle
        D3DXMatrixMultiply matWorld, matWorld, matTemp
        D3DXMatrixIdentity matTemp
        D3DXMatrixRotationZ matTemp, RotateAngle
        D3DXMatrixMultiply matWorld, matWorld, matTemp
        .SetTransform D3DTS_WORLD, matWorld
        .SetRenderState D3DRS_ALPHABLENDENABLE, False
        .SetTexture 0, TexChest
        .SetStreamSource 0, vBuffChest, Len(Chest(0))
        .Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0
        .DrawPrimitive D3DPT_TRIANGLELIST, 0, 12
        D3DXMatrixIdentity matWorld
        .SetTransform D3DTS_WORLD, matWorld
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetTexture 0, TexWaves
        .SetStreamSource 0, vBuffWaves, Len(Waves(0))
        .DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 60
        If FPSCurrent Then
            D3DX.DrawText MainFont, &HFFFFC000, CStr(FPSCurrent) & "fps", TextRect, DT_TOP Or DT_LEFT
        End If
        .EndScene
        On Error Resume Next
            .Present ByVal 0, ByVal 0, 0, ByVal 0
        On Error GoTo 0
    End With 'D3DDEVICE

End Sub

':) Ulli's VB Code Formatter V2.17.4 (2004-Okt-17 21:53)  Decl: 36  Code: 206  Total: 242 Lines
':) CommentOnly: 6 (2,5%)  Commented: 8 (3,3%)  Empty: 39 (16,1%)  Max Logic Depth: 3

Option Explicit
' **********************
' ** Renders graphics **
' **********************
' DirectX8 Object
Private DirectX8 As DirectX8 'The master DirectX object.
Private Direct3D As Direct3D8 'Controls all things 3D.
Public Direct3D_Device As Direct3DDevice8 'Represents the hardware rendering.
Private Direct3DX As D3DX8

'The 2D (Transformed and Lit) vertex format.
Private Const FVF_TLVERTEX As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE

'The 2D (Transformed and Lit) vertex format type.
Public Type TLVERTEX
    x As Single
    y As Single
    Z As Single
    RHW As Single
    Color As Long
    TU As Single
    TV As Single
End Type

Private Vertex_List(3) As TLVERTEX '4 vertices will make a square.

'Some color depth constants to help make the DX constants more readable.
Private Const COLOR_DEPTH_16_BIT As Long = D3DFMT_R5G6B5
Private Const COLOR_DEPTH_24_BIT As Long = D3DFMT_A8R8G8B8
Private Const COLOR_DEPTH_32_BIT As Long = D3DFMT_X8R8G8B8

Public RenderingMode As Long

Private Direct3D_Window As D3DPRESENT_PARAMETERS 'Backbuffer and viewport description.
Private Display_Mode As D3DDISPLAYMODE

'Graphic Textures
Public Tex_Item() As DX8TextureRec ' arrays
Public Tex_Character() As DX8TextureRec
Public Tex_Paperdoll() As DX8TextureRec
Public Tex_Tileset() As DX8TextureRec
Public Tex_Resource() As DX8TextureRec
Public Tex_Animation() As DX8TextureRec
Public Tex_SpellIcon() As DX8TextureRec
Public Tex_Face() As DX8TextureRec
Public Tex_Fog() As DX8TextureRec
Public Tex_Door As DX8TextureRec ' singes
Public Tex_Blood As DX8TextureRec
Public Tex_Misc As DX8TextureRec
Public Tex_Direction As DX8TextureRec
Public Tex_Target As DX8TextureRec
Public Tex_Bars As DX8TextureRec
Public Tex_Selection As DX8TextureRec
Public Tex_White As DX8TextureRec
Public Tex_Weather As DX8TextureRec
Public Tex_ChatBubble As DX8TextureRec
Public Tex_Fade As DX8TextureRec
Public Tex_Hair() As DX8TextureRec
Public Tex_Menu_Win() As DX8TextureRec
Public Tex_Menu_Buttons() As DX8TextureRec
Public Tex_Main_Win() As DX8TextureRec
Public Tex_Main_Buttons() As DX8TextureRec
Public Tex_ProjectTiles() As DX8TextureRec
Public Tex_ProjectTilesTarget() As DX8TextureRec
Public Tex_Head() As DX8TextureRec
Public Tex_Ghost As DX8TextureRec
Public Tex_LightMap As DX8TextureRec
Public Tex_Light As DX8TextureRec
Public Tex_BackGround() As DX8TextureRec
Public Tex_MiniMap As DX8TextureRec
Public Tex_Map() As DX8TextureRec
Public Tex_Skull As DX8TextureRec
Public Tex_Buff() As DX8TextureRec
Public Tex_Addon() As DX8TextureRec
Public Tex_Talentos() As DX8TextureRec
Public Tex_Corpses() As DX8TextureRec
Public Tex_WFrames() As DX8TextureRec

' Number of graphic files
Public NumTileSets As Long
Public NumCharacters As Long
Public NumPaperdolls As Long
Public numitems As Long
Public NumResources As Long
Public NumAnimations As Long
Public NumSpellIcons As Long
Public NumFaces As Long
Public NumFogs As Long
Public NumHairs As Long
Public NumMenuWin As Long
Public NumMenuButtons As Long
Public NumMainWin As Long
Public NumMainButtons As Long
Public NumProjectTiles As Long
Public NumProjectilesTarget As Long
Public NumHeads As Long
Public NumBackGround As Long
Public NumMap As Long
Public NumBuff As Long
Public NumAddon As Long
Public NumTalentos As Long
Public NumCorpses As Long
Public NumWFrames As Long

Public Type DX8TextureRec
    Texture As Long
    Width As Long
    Height As Long
    filepath As String
    TexWidth As Long
    TexHeight As Long
    ImageData() As Byte
    HasData As Boolean
End Type

Public Type GlobalTextureRec
    Texture As Direct3DTexture8
    TexWidth As Long
    TexHeight As Long
End Type

Public Type RECT
    Top As Long
    Left As Long
    Bottom As Long
    Right As Long
End Type

Public gTexture() As GlobalTextureRec
Public NumTextures As Long

' ********************
' ** Initialization **
' ********************
Public Function InitDX8() As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set DirectX8 = New DirectX8 'Creates the DirectX object.
    Set Direct3D = DirectX8.Direct3DCreate() 'Creates the Direct3D object using the DirectX object.
    Set Direct3DX = New D3DX8
    
    Direct3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Display_Mode 'Use the current display mode that you
                                                                    'are already on. Incase you are confused, I'm
                                                                    'talking about your current screen resolution. ;)
    Direct3D_Window.Windowed = True 'The app will be in windowed mode.
    
    Direct3D_Window.SwapEffect = D3DSWAPEFFECT_DISCARD 'Refresh when the monitor does.
    Direct3D_Window.BackBufferFormat = Display_Mode.Format 'Sets the format that was retrieved into the backbuffer.
    'Creates the rendering device with some useful info, along with the info
    'DispMode.Format = D3DFMT_X8R8G8B8
    Direct3D_Window.SwapEffect = D3DSWAPEFFECT_COPY
    Direct3D_Window.BackBufferCount = 1 '1 backbuffer only
    Direct3D_Window.BackBufferWidth = 1024 ' frmMain.picScreen.ScaleWidth 'Match the backbuffer width with the display width
    Direct3D_Window.BackBufferHeight = 726 'frmMain.picScreen.ScaleHeight 'Match the backbuffer height with the display height
    Direct3D_Window.hDeviceWindow = frmMain.picScreen.hWnd 'Use frmMain as the device window.
    
    'we've already setup for Direct3D_Window.
    If TryCreateDirectX8Device = False Then
        MsgBox "Unable to initialize DirectX8. You may be missing dx8vb.dll or have incompatible hardware to use DirectX8."
        DestroyGame
    End If

    With Direct3D_Device
        .SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE
    
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        .SetRenderState D3DRS_ZENABLE, False
        .SetRenderState D3DRS_ZWRITEENABLE, False
        
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
        .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
    End With
    
    ' Initialise the surfaces
    LoadTextures
    
    ' We're done
    InitDX8 = True
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "InitDX8", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function TryCreateDirectX8Device() As Boolean
Dim i As Long

On Error GoTo nexti

    For i = 1 To 4
        Select Case i
            Case 1
                Set Direct3D_Device = Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.picScreen.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, Direct3D_Window)
                TryCreateDirectX8Device = True
                Exit Function
            Case 2
                Set Direct3D_Device = Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.picScreen.hWnd, D3DCREATE_MIXED_VERTEXPROCESSING, Direct3D_Window)
                TryCreateDirectX8Device = True
                Exit Function
            Case 3
                Set Direct3D_Device = Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.picScreen.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, Direct3D_Window)
                TryCreateDirectX8Device = True
                Exit Function
            Case 4
                TryCreateDirectX8Device = False
                Exit Function
        End Select
nexti:
    Next

End Function

Function GetNearestPOT(value As Long) As Long
Dim i As Long
    Do While 2 ^ i < value
        i = i + 1
    Loop
    GetNearestPOT = 2 ^ i
End Function

Public Sub LoadTexture(ByRef TextureRec As DX8TextureRec)
Dim SourceBitmap As cGDIpImage, ConvertedBitmap As cGDIpImage, GDIGraphics As cGDIpRenderer, GDIToken As cGDIpToken, i As Long
Dim newWidth As Long, newHeight As Long, ImageData() As Byte, fn As Long
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If TextureRec.HasData = False Then
        Set GDIToken = New cGDIpToken
        If GDIToken.Token = 0& Then MsgBox "GDI+ failed to load, exiting game!": DestroyGame
        Set SourceBitmap = New cGDIpImage
        Call SourceBitmap.LoadPicture_FileName(TextureRec.filepath, GDIToken)
        
        TextureRec.Width = SourceBitmap.Width
        TextureRec.Height = SourceBitmap.Height
        
        newWidth = GetNearestPOT(TextureRec.Width)
        newHeight = GetNearestPOT(TextureRec.Height)
        If newWidth <> SourceBitmap.Width Or newHeight <> SourceBitmap.Height Then
            Set ConvertedBitmap = New cGDIpImage
            Set GDIGraphics = New cGDIpRenderer
            i = GDIGraphics.CreateGraphicsFromImageClass(SourceBitmap)
            Call ConvertedBitmap.LoadPicture_FromNothing(newWidth, newHeight, i, GDIToken) 'This is no longer backwards and it now works.
            Call GDIGraphics.DestroyHGraphics(i)
            i = GDIGraphics.CreateGraphicsFromImageClass(ConvertedBitmap)
            Call GDIGraphics.AttachTokenClass(GDIToken)
            Call GDIGraphics.RenderImageClassToHGraphics(SourceBitmap, i)
            Call ConvertedBitmap.SaveAsPNG(ImageData)
            GDIGraphics.DestroyHGraphics (i)
            TextureRec.ImageData = ImageData
            Set ConvertedBitmap = Nothing
            Set GDIGraphics = Nothing
            Set SourceBitmap = Nothing
        Else
            Call SourceBitmap.SaveAsPNG(ImageData)
            TextureRec.ImageData = ImageData
            Set SourceBitmap = Nothing
        End If
    Else
        ImageData = TextureRec.ImageData
    End If
    
    
    Set gTexture(TextureRec.Texture).Texture = Direct3DX.CreateTextureFromFileInMemoryEx(Direct3D_Device, _
                                                    ImageData(0), _
                                                    UBound(ImageData) + 1, _
                                                    newWidth, _
                                                    newHeight, _
                                                    D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, ByVal (0), ByVal 0, ByVal 0)
    
    gTexture(TextureRec.Texture).TexWidth = newWidth
    gTexture(TextureRec.Texture).TexHeight = newHeight
    Exit Sub
errorhandler:
    HandleError "LoadTexture", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub LoadTextures()
Dim i As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call CheckMenuWindows
    Call CheckMenuButtons
    Call CheckMainWindows
    Call CheckMainButtons
    Call CheckTilesets
    Call CheckCharacters
    Call CheckPaperdolls
    Call CheckAnimations
    Call CheckItems
    Call CheckResources
    Call CheckSpellIcons
    Call CheckFaces
    Call CheckFogs
    Call CheckHairs
    Call CheckProjectTiles
    Call CheckProjectilesTarget
    Call CheckHeads
    Call CheckBackGround
    Call CheckMaps
    Call CheckBuffs
    Call CheckAddons
    Call CheckTalentos
    Call CheckCorpses
    Call CheckWFrames
    
    NumTextures = NumTextures + 16
    
    ReDim Preserve gTexture(NumTextures)
    
    Tex_Skull.filepath = App.Path & "\data files\graphics\misc\skulls.png"
    Tex_Skull.Texture = NumTextures - 15
    LoadTexture Tex_Skull
    
    Tex_MiniMap.filepath = App.Path & "\data files\graphics\misc\MiniMap.png"
    Tex_MiniMap.Texture = NumTextures - 14
    LoadTexture Tex_MiniMap
    
    Tex_Light.filepath = App.Path & "\data files\graphics\misc\light.png"
    Tex_Light.Texture = NumTextures - 13
    LoadTexture Tex_Light
    
    Tex_LightMap.filepath = App.Path & "\data files\graphics\misc\lightmap.png"
    Tex_LightMap.Texture = NumTextures - 12
    LoadTexture Tex_LightMap
    
    Tex_Ghost.filepath = App.Path & "\data files\graphics\misc\ghost.png"
    Tex_Ghost.Texture = NumTextures - 11
    LoadTexture Tex_Ghost
    
    Tex_Fade.filepath = App.Path & "\data files\graphics\misc\fader.png"
    Tex_Fade.Texture = NumTextures - 10
    LoadTexture Tex_Fade
    
    Tex_ChatBubble.filepath = App.Path & "\data files\graphics\misc\chatbubble.png"
    Tex_ChatBubble.Texture = NumTextures - 9
    LoadTexture Tex_ChatBubble
    
    Tex_Weather.filepath = App.Path & "\data files\graphics\misc\weather.png"
    Tex_Weather.Texture = NumTextures - 8
    LoadTexture Tex_Weather
    
    Tex_White.filepath = App.Path & "\data files\graphics\misc\white.png"
    Tex_White.Texture = NumTextures - 7
    LoadTexture Tex_White
    
    Tex_Door.filepath = App.Path & "\data files\graphics\misc\door.png"
    Tex_Door.Texture = NumTextures - 6
    LoadTexture Tex_Door
    
    Tex_Direction.filepath = App.Path & "\data files\graphics\misc\direction.png"
    Tex_Direction.Texture = NumTextures - 5
    LoadTexture Tex_Direction
    
    Tex_Target.filepath = App.Path & "\data files\graphics\misc\target.png"
    Tex_Target.Texture = NumTextures - 4
    LoadTexture Tex_Target
    
    Tex_Misc.filepath = App.Path & "\data files\graphics\misc\misc.png"
    Tex_Misc.Texture = NumTextures - 3
    LoadTexture Tex_Misc
    
    Tex_Blood.filepath = App.Path & "\data files\graphics\misc\blood.png"
    Tex_Blood.Texture = NumTextures - 2
    LoadTexture Tex_Blood
    
    Tex_Bars.filepath = App.Path & "\data files\graphics\misc\bars.png"
    Tex_Bars.Texture = NumTextures - 1
    LoadTexture Tex_Bars
    
    Tex_Selection.filepath = App.Path & "\data files\graphics\misc\select.png"
    Tex_Selection.Texture = NumTextures
    LoadTexture Tex_Selection
    
    EngineInitFontTextures
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadTextures", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UnloadTextures()
Dim i As Long
    
    ' If debug mode, handle error then exit out
    On Error Resume Next
    
    For i = 1 To NumTextures
        Set gTexture(i).Texture = Nothing
        ZeroMemory ByVal VarPtr(gTexture(i)), LenB(gTexture(i))
    Next
    
    ReDim gTexture(1)

    
    For i = 1 To NumTileSets
        Tex_Tileset(i).Texture = 0
    Next

    For i = 1 To numitems
        Tex_Item(i).Texture = 0
    Next

    For i = 1 To NumCharacters
        Tex_Character(i).Texture = 0
    Next
    
    For i = 1 To NumPaperdolls
        Tex_Paperdoll(i).Texture = 0
    Next
    
    For i = 1 To NumResources
        Tex_Resource(i).Texture = 0
    Next
    
    For i = 1 To NumAnimations
        Tex_Animation(i).Texture = 0
    Next
    
    For i = 1 To NumSpellIcons
        Tex_SpellIcon(i).Texture = 0
    Next
    
    For i = 1 To NumFaces
        Tex_Face(i).Texture = 0
    Next
    
    For i = 1 To NumHairs
        Tex_Hair(i).Texture = 0
    Next
    
    For i = 1 To NumMenuWin
        Tex_Menu_Win(i).Texture = 0
    Next
    
    For i = 1 To NumMenuButtons
        Tex_Menu_Buttons(i).Texture = 0
    Next
    
    For i = 1 To NumMainWin
        Tex_Main_Win(i).Texture = 0
    Next
    
    For i = 1 To NumMainButtons
        Tex_Main_Buttons(i).Texture = 0
    Next
    
    For i = 1 To NumProjectTiles
        Tex_ProjectTiles(i).Texture = 0
    Next
    
    For i = 1 To NumProjectilesTarget
        Tex_ProjectTilesTarget(i).Texture = 0
    Next
    
    For i = 1 To NumHeads
        Tex_Head(i).Texture = 0
    Next
    
    For i = 1 To NumBackGround
        Tex_BackGround(i).Texture = 0
    Next
    
    For i = 1 To NumMap
        Tex_Map(i).Texture = 0
    Next
    
    For i = 1 To NumBuff
        Tex_Buff(i).Texture = 0
    Next
        
    For i = 1 To NumAddon
        Tex_Addon(i).Texture = 0
    Next
    
    For i = 1 To NumTalentos
        Tex_Talentos(i).Texture = 0
    Next
    
    For i = 1 To NumCorpses
        Tex_Corpses(i).Texture = 0
    Next
    
    For i = 1 To NumWFrames
        Tex_WFrames(i).Texture = 0
    Next
    
    Tex_Misc.Texture = 0
    Tex_Blood.Texture = 0
    Tex_Door.Texture = 0
    Tex_Direction.Texture = 0
    Tex_Target.Texture = 0
    Tex_Selection.Texture = 0
    Tex_Fade.Texture = 0
    Tex_Ghost.Texture = 0
    Tex_Weather.Texture = 0
    Tex_MiniMap.Texture = 0
    Tex_Skull.Texture = 0
    
    UnloadFontTextures
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UnloadTextures", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' **************
' ** Drawing **
' **************
Public Sub RenderTexture(ByRef TextureRec As DX8TextureRec, ByVal dX As Single, ByVal dY As Single, ByVal sX As Single, ByVal sY As Single, ByVal dWidth As Single, ByVal dHeight As Single, ByVal sWidth As Single, ByVal sHeight As Single, Optional Color As Long = -1, Optional ByVal Degrees As Single = 0)
    Dim TextureNum As Long
    Dim textureWidth As Long, textureHeight As Long, sourceX As Single, sourceY As Single, sourceWidth As Single, sourceHeight As Single
    
    Dim RadAngle As Single 'The angle in Radians
    Dim CenterX As Single
    Dim CenterY As Single
    Dim NewX As Single
    Dim NewY As Single
    Dim SinRad As Single
    Dim CosRad As Single
    Dim i As Long
    
    TextureNum = TextureRec.Texture
    
    textureWidth = gTexture(TextureNum).TexWidth
    textureHeight = gTexture(TextureNum).TexHeight
    
    If sY + sHeight > textureHeight Then Exit Sub
    If sX + sWidth > textureWidth Then Exit Sub
    If sX < 0 Then Exit Sub
    If sY < 0 Then Exit Sub

    sX = sX - 0.5
    sY = sY - 0.5
    dY = dY - 0.5
    dX = dX - 0.5
    sWidth = sWidth
    sHeight = sHeight
    dWidth = dWidth
    dHeight = dHeight
    sourceX = (sX / textureWidth)
    sourceY = (sY / textureHeight)
    sourceWidth = ((sX + sWidth) / textureWidth)
    sourceHeight = ((sY + sHeight) / textureHeight)
    
    Vertex_List(0) = Create_TLVertex(dX, dY, 0, 1, Color, 0, sourceX + 0.000003, sourceY + 0.000003)
    Vertex_List(1) = Create_TLVertex(dX + dWidth, dY, 0, 1, Color, 0, sourceWidth + 0.000003, sourceY + 0.000003)
    Vertex_List(2) = Create_TLVertex(dX, dY + dHeight, 0, 1, Color, 0, sourceX + 0.000003, sourceHeight + 0.000003)
    Vertex_List(3) = Create_TLVertex(dX + dWidth, dY + dHeight, 0, 1, Color, 0, sourceWidth + 0.000003, sourceHeight + 0.000003)
    
        'Check if a rotation is required
    If Degrees <> 0 And Degrees <> 360 Then

        'Converts the angle to rotate by into radians
        RadAngle = Degrees * DegreeToRadian

        'Set the CenterX and CenterY values
        CenterX = dX + (dWidth * 0.5)
        CenterY = dY + (dHeight * 0.5)

        'Pre-calculate the cosine and sine of the radiant
        SinRad = Sin(RadAngle)
        CosRad = Cos(RadAngle)

        'Loops through the passed vertex buffer
        For i = 0 To 3

            'Calculates the new X and Y co-ordinates of the vertices for the given angle around the center co-ordinates
            NewX = CenterX + (Vertex_List(i).x - CenterX) * CosRad - (Vertex_List(i).y - CenterY) * SinRad
            NewY = CenterY + (Vertex_List(i).y - CenterY) * CosRad + (Vertex_List(i).x - CenterX) * SinRad

            'Applies the new co-ordinates to the buffer
            Vertex_List(i).x = NewX
            Vertex_List(i).y = NewY
        Next
    End If
    
    Direct3D_Device.SetTexture 0, gTexture(TextureNum).Texture
    Direct3D_Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex_List(0), Len(Vertex_List(0))
End Sub

Public Sub RenderTextureByRects(TextureRec As DX8TextureRec, sRECT As RECT, drect As RECT, Optional Colour As Long = -1)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    RenderTexture TextureRec, drect.Left, drect.Top, sRECT.Left, sRECT.Top, drect.Right - drect.Left, drect.Bottom - drect.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, Colour

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "RenderTextureByRects", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawDirection(ByVal x As Long, ByVal y As Long)
Dim Rec As RECT
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' render grid
    Rec.Top = 24
    Rec.Left = 0
    Rec.Right = Rec.Left + 32
    Rec.Bottom = Rec.Top + 32
    RenderTexture Tex_Direction, ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, D3DColorRGBA(255, 255, 255, 255)
    
    ' render dir blobs
    For i = 1 To 4
        Rec.Left = (i - 1) * 8
        Rec.Right = Rec.Left + 8
        ' find out whether render blocked or not
        If Not isDirBlocked(Map.Tile(x, y).DirBlock, CByte(i)) Then
            Rec.Top = 8
        Else
            Rec.Top = 16
        End If
        Rec.Bottom = Rec.Top + 8
        'render!
        RenderTexture Tex_Direction, ConvertMapX(x * PIC_X) + DirArrowX(i), ConvertMapY(y * PIC_Y) + DirArrowY(i), Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, D3DColorRGBA(255, 255, 255, 255)
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawDirection", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawTarget(ByVal x As Long, ByVal y As Long)
Dim sRECT As RECT
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Tex_Target.Texture = 0 Then Exit Sub
    
    Width = Tex_Target.Width / 2
    Height = Tex_Target.Height

    With sRECT
        .Top = 0
        .Bottom = Height
        .Left = 0
        .Right = Width
    End With
    
    x = x - ((Width - 32) / 2)
    y = y
    
    x = ConvertMapX(x)
    y = ConvertMapY(y)
    
    ' clipping
    If y < 0 Then
        With sRECT
            .Top = .Top - y
        End With
        y = 0
    End If

    If x < 0 Then
        With sRECT
            .Left = .Left - x
        End With
        x = 0
    End If
    
    RenderTexture Tex_Target, x, y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawTarget", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawHover(ByVal tType As Long, ByVal target As Long, ByVal x As Long, ByVal y As Long)
Dim sRECT As RECT
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Tex_Target.Texture = 0 Then Exit Sub
    
    Width = Tex_Target.Width / 2
    Height = Tex_Target.Height

    With sRECT
        .Top = 0
        .Bottom = Height
        .Left = Width
        .Right = .Left + Width
    End With
    
    x = x - ((Width - 32) / 2)
    y = y

    x = ConvertMapX(x)
    y = ConvertMapY(y)
    
    ' clipping
    If y < 0 Then
        With sRECT
            .Top = .Top - y
        End With
        y = 0
    End If

    If x < 0 Then
        With sRECT
            .Left = .Left - x
        End With
        x = 0
    End If
    
    RenderTexture Tex_Target, x, y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawHover", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawChestGraphic(ByVal x As Long, ByVal y As Long)
    Dim chestNum As Long
    Dim sprite(1 To 2) As Long
    Dim i As Byte
    Dim texX As Long, texY As Long
    
    ' Obter número do baú na posição do mapa
    chestNum = Map.Tile(x, y).Data1
    If chestNum <= 0 Then Exit Sub
    
    ' Validar gráficos do baú
    For i = 1 To 2
        sprite(i) = Chest(chestNum).Graphic(i)
        If sprite(i) <= 0 Or sprite(i) > NumResources Then Exit Sub
    Next i
    
    ' Coordenadas de renderização
    texX = ConvertMapX(x * 32)
    texY = ConvertMapY(y * 32)
    
    ' Renderizar gráfico do baú baseado no estado de abertura
    If Player(MyIndex).Chest(chestNum).Get Then
        RenderTexture Tex_Resource(sprite(2)), texX, texY, 0, 0, 32, 32, 32, 32
    Else
        RenderTexture Tex_Resource(sprite(1)), texX, texY, 0, 0, 32, 32, 32, 32
    End If
End Sub


Public Sub DrawMapTile(ByVal x As Long, ByVal y As Long)
Dim Rec As RECT
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With Map.Tile(x, y)
            i = MapLayer.Ground
            If .Layer(i).Tileset <> 0 Then
                If Autotile(x, y).Layer(i).renderState = RENDER_STATE_NORMAL Then
                
                RenderTexture Tex_Tileset(.Layer(i).Tileset), ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), .Layer(i).x * 32, .Layer(i).y * 32, 32, 32, 32, 32, -1
                    
                ElseIf Autotile(x, y).Layer(i).renderState = RENDER_STATE_AUTOTILE Then
                    ' Draw autotiles
                    DrawAutoTile i, ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), 1, x, y
                    DrawAutoTile i, ConvertMapX((x * PIC_X) + 16), ConvertMapY(y * PIC_Y), 2, x, y
                    DrawAutoTile i, ConvertMapX(x * PIC_X), ConvertMapY((y * PIC_Y) + 16), 3, x, y
                    DrawAutoTile i, ConvertMapX((x * PIC_X) + 16), ConvertMapY((y * PIC_Y) + 16), 4, x, y
                End If
            End If
    End With
    
    With Map.Tile(x, y)
            i = 3
            If .Layer(i).Tileset <> 0 Then
                If MapAnim = 0 Then
                    RenderTexture Tex_Tileset(.Layer(i).Tileset), ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), .Layer(i).x * 32, .Layer(i).y * 32, 32, 32, 32, 32, -1
                End If
            End If
    End With
    
    ' Error handler
    Exit Sub
    
errorhandler:
    HandleError "DrawMapTile", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawMapFringeTile(ByVal x As Long, ByVal y As Long)
Dim Rec As RECT
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With Map.Tile(x, y)
        i = MapLayer.Fringe
            If .Layer(i).Tileset <> 0 Then
                If Autotile(x, y).Layer(i).renderState = RENDER_STATE_NORMAL Then
                    
                    
                    'If isInRange(2, GetPlayerX(MyIndex), GetPlayerY(MyIndex), x, Y) Then
                        ' Draw normally
                        'RenderTexture Tex_Tileset(.Layer(i).Tileset), ConvertMapX(x * PIC_X), ConvertMapY(Y * PIC_Y), .Layer(i).x * 32, .Layer(i).Y * 32, 32, 32, 32, 32, D3DColorARGB(150, 255, 255, 255)
                    'Else
                        RenderTexture Tex_Tileset(.Layer(i).Tileset), ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), .Layer(i).x * 32, .Layer(i).y * 32, 32, 32, 32, 32, -1
                    'End If
                
                ElseIf Autotile(x, y).Layer(i).renderState = RENDER_STATE_AUTOTILE Then
                    ' Draw autotiles
                    DrawAutoTile i, ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), 1, x, y
                    DrawAutoTile i, ConvertMapX((x * PIC_X) + 16), ConvertMapY(y * PIC_Y), 2, x, y
                    DrawAutoTile i, ConvertMapX(x * PIC_X), ConvertMapY((y * PIC_Y) + 16), 3, x, y
                    DrawAutoTile i, ConvertMapX((x * PIC_X) + 16), ConvertMapY((y * PIC_Y) + 16), 4, x, y
                End If
            End If
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawMapFringeTile", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawDoor(ByVal x As Long, ByVal y As Long)
Dim Rec As RECT
Dim X2 As Long, Y2 As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' sort out animation
    With TempTile(x, y)
        If .DoorAnimate = 1 Then ' opening
            If .DoorTimer + 100 < GetTickCount Then
                If .DoorFrame < 4 Then
                    .DoorFrame = .DoorFrame + 1
                Else
                    .DoorAnimate = 2 ' set to closing
                End If
                .DoorTimer = GetTickCount
            End If
        ElseIf .DoorAnimate = 2 Then ' closing
            If .DoorTimer + 100 < GetTickCount Then
                If .DoorFrame > 1 Then
                    .DoorFrame = .DoorFrame - 1
                Else
                    .DoorAnimate = 0 ' end animation
                End If
                .DoorTimer = GetTickCount
            End If
        End If
        
        If .DoorFrame = 0 Then .DoorFrame = 1
    End With

    With Rec
        .Top = 0
        .Bottom = Tex_Door.Height
        .Left = ((TempTile(x, y).DoorFrame - 1) * (Tex_Door.Width / 4))
        .Right = .Left + (Tex_Door.Width / 4)
    End With

    X2 = (x * PIC_X)
    Y2 = (y * PIC_Y) - (Tex_Door.Height / 2) + 4
    RenderTexture Tex_Door, ConvertMapX(X2), ConvertMapY(Y2), Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, D3DColorRGBA(255, 255, 255, 255)
    'Call DDS_BackBuffer.DrawFast(ConvertMapX(X2), ConvertMapY(Y2), DDS_Door, rec, DDDrawFAST_WAIT Or DDDrawFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawDoor", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawCorpse(ByVal Index As Long)
Dim Rec As RECT
Dim x As Long, y As Long, Width As Long, Height As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Controle Para Aparecer depois de 1,5seg
    If MapDrop(Index).SpawnTick > 0 Then
        If MapDrop(Index).SpawnTick < GetTickCount Then
            MapDrop(Index).SpawnTick = 0
        End If
        Exit Sub
    End If
    
    With MapDrop(Index)
        Width = Tex_Corpses(.Graphic).Width
        Height = Tex_Corpses(.Graphic).Height
        x = ConvertMapX(.x * 32) - ((Width - 32) / 2) '(Width / 4)
        y = ConvertMapY(.y * 32) - ((Height - 32) / 2)
        
    Select Case .Dir
    Case DIR_UP, DIR_RIGHT, DIR_UP_RIGHT, DIR_DOWN_RIGHT
        RenderTexture Tex_Corpses(.Graphic), x, y, 0, 0, Width, Height, Width, Height, D3DColorRGBA(255, 255, 255, 255)
    Case DIR_DOWN, DIR_LEFT, DIR_UP_LEFT, DIR_DOWN_LEFT
        RenderTexture Tex_Corpses(.Graphic), x + Width, y, 0, 0, -Width, Height, Width, Height, D3DColorRGBA(255, 255, 255, 255)
    End Select
    
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawCorpse", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawBlood(ByVal Index As Long)
Dim Rec As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    'load blood then
    BloodCount = Tex_Blood.Width / 32
    
    With Blood(Index)
        ' check if we should be seeing it
        If .Timer + 20000 < GetTickCount Then Exit Sub
        
        Rec.Top = ((Tex_Blood.Height / 5) * .Color) - 44
        Rec.Bottom = Rec.Top + PIC_Y
        Rec.Left = (.sprite - 1) * PIC_X
        Rec.Right = Rec.Left + PIC_X
        RenderTexture Tex_Blood, ConvertMapX(.x * PIC_X), ConvertMapY(.y * PIC_Y) - 7, Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, D3DColorRGBA(255, 255, 255, 255)
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawBlood", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawAnimation(ByVal Index As Long, ByVal Layer As Long)
Dim sprite As Long
Dim sRECT As RECT
Dim drect As RECT
Dim i As Long
Dim Width As Long, Height As Long
Dim looptime As Long
Dim FrameCount As Long
Dim x As Long, y As Long
Dim LockIndex As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If AnimInstance(Index).Animation = 0 Then
        ClearAnimInstance Index
        Exit Sub
    End If
    
    sprite = Animation(AnimInstance(Index).Animation).sprite(Layer)
    
    If sprite < 1 Or sprite > NumAnimations Then Exit Sub
    
    FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)
    
    ' total width divided by frame count
    Width = Tex_Animation(sprite).Width / FrameCount
    Height = Tex_Animation(sprite).Height
    
    sRECT.Top = 0
    sRECT.Bottom = Height
    sRECT.Left = (AnimInstance(Index).frameIndex(Layer) - 1) * Width
    sRECT.Right = sRECT.Left + Width
    
    ' change x or y if locked
    If AnimInstance(Index).LockType > TARGET_TYPE_NONE Then ' if <> none
        ' is a player
        If AnimInstance(Index).LockType = TARGET_TYPE_PLAYER Then
            ' quick save the index
            LockIndex = AnimInstance(Index).LockIndex
            ' check if is ingame
            If IsPlaying(LockIndex) Then
                ' check if on same map
                If GetPlayerMap(LockIndex) = GetPlayerMap(MyIndex) Then
                    ' is on map, is playing, set x & y
                    x = (GetPlayerX(LockIndex) * PIC_X) + 16 - (Width / 2) + Player(LockIndex).xOffset
                    y = (GetPlayerY(LockIndex) * PIC_Y) + 16 - (Height / 2) + Player(LockIndex).yOffset
                End If
            End If
        ElseIf AnimInstance(Index).LockType = TARGET_TYPE_NPC Then
            ' quick save the index
            LockIndex = AnimInstance(Index).LockIndex
            ' check if NPC exists
            If MapNpc(LockIndex).num > 0 Then
                ' check if alive
                If MapNpc(LockIndex).Vital(Vitals.HP) > 0 Then
                    ' exists, is alive, set x & y
                    x = (MapNpc(LockIndex).x * PIC_X) + 16 - (Width / 2) + MapNpc(LockIndex).xOffset
                    y = (MapNpc(LockIndex).y * PIC_Y) + 16 - (Height / 2) + MapNpc(LockIndex).yOffset
                Else
                    ' npc not alive anymore, kill the animation
                    ClearAnimInstance Index
                    Exit Sub
                End If
            Else
                ' npc not alive anymore, kill the animation
                ClearAnimInstance Index
                Exit Sub
            End If
        End If
    Else
        ' no lock, default x + y
        x = (AnimInstance(Index).x * 32) + 16 - (Width / 2)
        y = (AnimInstance(Index).y * 32) + 16 - (Height / 2)
    End If
    
    x = ConvertMapX(x)
    y = ConvertMapY(y)

    ' Clip to screen
    If y < 0 Then

        With sRECT
            .Top = .Top - y
        End With

        y = 0
    End If

    If x < 0 Then

        With sRECT
            .Left = .Left - x
        End With

        x = 0
    End If
    
    RenderTexture Tex_Animation(sprite), x, y - 12, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawAnimation", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawItem(ByVal ItemNum As Long)
Dim PicNum As Long
Dim Rec As RECT
Dim MaxFrames As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' get the picture
    PicNum = Item(MapItem(ItemNum).num).pic

    If PicNum < 1 Or PicNum > numitems Then Exit Sub

    If Tex_Item(PicNum).Width > 32 Then ' has more than 1 frame
        With Rec
            .Top = 0
            .Bottom = 32
            .Left = (MapItem(ItemNum).Frame * 32)
            .Right = .Left + 32
        End With
    Else
        With Rec
            .Top = 0
            .Bottom = 32
            .Left = 0
            .Right = 32
        End With
    End If
    
    RenderTexture Tex_Item(PicNum), ConvertMapX(MapItem(ItemNum).x * PIC_X), ConvertMapY(MapItem(ItemNum).y * PIC_Y), Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, D3DColorRGBA(255, 255, 255, 255)
    
'    If Item(MapItem(ItemNum).num).Type = ITEM_TYPE_TORCH Then
'        Call DrawLight(MapItem(ItemNum).x * 32, MapItem(ItemNum).Y * 32, 50, 255, 255, 0)
'    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawItem", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ScreenshotMap()
Dim x As Long, y As Long, i As Long, Rec As RECT, drec As RECT
Dim num As Long
Dim sFilePath As String, wasFound As Boolean

 For num = 1 To 255

     sFilePath = App.Path & "\ScreenShots\Screen" & num & ".jpg"

     If Not FileExist(sFilePath, True) Then

         frmMain.picSSMap.Picture = CaptureForm(frmMain)

         SavePicture frmMain.picSSMap.Picture, App.Path & "\ScreenShots\Screen" & num & ".jpg"

         Exit Sub

     End If

 Next

 If wasFound Then

     MsgBox "The maximum screenshots for this map have been reached! Delete some if you wish to take more...", vbOKOnly

 End If

End Sub

Public Sub DrawMapResource(ByVal Resource_num As Long, Optional ByVal screenShot As Boolean = False)
Dim Resource_master As Long
Dim Resource_state As Long
Dim Resource_sprite As Long
Dim Rec As RECT
Dim x As Long, y As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' make sure it's not out of map
    If MapResource(Resource_num).x > Map.MaxX Then Exit Sub
    If MapResource(Resource_num).y > Map.MaxY Then Exit Sub
    
    ' Get the Resource type
    Resource_master = Map.Tile(MapResource(Resource_num).x, MapResource(Resource_num).y).Data1
    
    If Resource_master = 0 Then Exit Sub

    If Resource(Resource_master).ResourceImage = 0 Then Exit Sub
    ' Get the Resource state
    Resource_state = MapResource(Resource_num).ResourceState

    If Resource_state = 0 Then ' normal
        Resource_sprite = Resource(Resource_master).ResourceImage
    ElseIf Resource_state = 1 Then ' used
        Resource_sprite = Resource(Resource_master).ExhaustedImage
    End If
    
    ' cut down everything if we're editing
    If InMapEditor Then
        Resource_sprite = Resource(Resource_master).ExhaustedImage
    End If

    ' src rect
    With Rec
        .Top = 0
        .Bottom = Tex_Resource(Resource_sprite).Height
        .Left = 0
        .Right = Tex_Resource(Resource_sprite).Width
    End With

    ' Set base x + y, then the offset due to size
    x = (MapResource(Resource_num).x * PIC_X) - (Tex_Resource(Resource_sprite).Width / 2) + 16
    y = (MapResource(Resource_num).y * PIC_Y) - Tex_Resource(Resource_sprite).Height + 32
    
    ' render it
    If Not screenShot Then
        Call DrawResource(Resource_sprite, x, y, Rec)
    Else
        Call ScreenshotResource(Resource_sprite, x, y, Rec)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawMapResource", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawResource(ByVal Resource As Long, ByVal dX As Long, dY As Long, Rec As RECT)
Dim x As Long
Dim y As Long
Dim Width As Long
Dim Height As Long
Dim destRect As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Resource < 1 Or Resource > NumResources Then Exit Sub

    x = ConvertMapX(dX)
    y = ConvertMapY(dY)
    
    Width = (Rec.Right - Rec.Left)
    Height = (Rec.Bottom - Rec.Top)
    
    RenderTexture Tex_Resource(Resource), x, y, Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, D3DColorRGBA(255, 255, 255, 255)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawResource", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub ScreenshotResource(ByVal Resource As Long, ByVal x As Long, y As Long, Rec As RECT)
Dim Width As Long
Dim Height As Long
Dim destRect As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Resource < 1 Or Resource > NumResources Then Exit Sub
    
    Width = (Rec.Right - Rec.Left)
    Height = (Rec.Bottom - Rec.Top)

    If y < 0 Then
        With Rec
            .Top = .Top - y
        End With
        y = 0
    End If

    If x < 0 Then
        With Rec
            .Left = .Left - x
        End With
        x = 0
    End If
    RenderTexture Tex_Resource(Resource), x, y, Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, D3DColorRGBA(255, 255, 255, 255)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ScreenshotResource", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawNpcBar(ByVal i As Long)
Dim tmpY As Long, tmpX As Long
Dim sWidth As Long, sHeight As Long
Dim sRECT As RECT
Dim BarWidth As Long
Dim npcNum As Long, partyIndex As Long
Dim BarColor As Long
Dim sprite As Long
    
    ' dynamic bar calculations
    sWidth = Tex_Bars.Width
    sHeight = Tex_Bars.Height / 12
    
    npcNum = MapNpc(i).num
    sprite = NPC(MapNpc(i).num).sprite
    
        If npcNum > 0 And sprite > 0 Then
        
        ' Npc Friendly
        If NPC(npcNum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Then Exit Sub
        
            ' alive?
            If MapNpc(i).Vital(Vitals.HP) > 0 Then
            
                ' lock to npc
                tmpX = MapNpc(i).x * PIC_X + MapNpc(i).xOffset + 15 - (sWidth / 2)
                tmpY = (MapNpc(i).y * PIC_Y) + MapNpc(i).yOffset + 24
                
                ' calculate the width to fill
                If NPC(npcNum).HP = 0 Then Exit Sub
                If NPC(npcNum).HP > 9999999 Then Exit Sub
                BarWidth = ((MapNpc(i).Vital(Vitals.HP) / sWidth) / (NPC(npcNum).HP / sWidth)) * sWidth
                
                ' draw bar background
                With sRECT
                    .Top = sHeight * 1 ' HP bar background
                    .Left = 0
                    .Right = .Left + sWidth
                    .Bottom = .Top + sHeight
                End With
                RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                
                ' draw the bar proper
                If MapNpc(i).Vital(1) >= NPC(npcNum).HP * 0.99 Then
                    BarColor = 0
                ElseIf MapNpc(i).Vital(1) >= NPC(npcNum).HP * 0.75 Then
                    BarColor = 2
                ElseIf MapNpc(i).Vital(1) >= NPC(npcNum).HP * 0.25 Then
                    BarColor = 4
                Else
                    BarColor = 6
                End If
                
                With sRECT
                    .Top = sHeight * BarColor ' HP bar background
                    .Left = 0
                    .Right = .Left + BarWidth
                    .Bottom = .Top + sHeight
                End With
                
                
                RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
            
            End If
        End If
    
End Sub

Private Sub DrawPlayerBar(ByVal i As Long)
Dim tmpY As Long, tmpX As Long
Dim sWidth As Long, sHeight As Long
Dim sRECT As RECT
Dim BarWidth As Long
Dim npcNum As Long, partyIndex As Long
Dim BarColor As Long
    
    ' dynamic bar calculations
    sWidth = Tex_Bars.Width
    sHeight = Tex_Bars.Height / 12
    
    If GetPlayerVital(i, Vitals.HP) > 0 Then
    
        ' lock to Player
        tmpX = GetPlayerX(i) * PIC_X + Player(i).xOffset + 15 - (sWidth / 2)
        tmpY = GetPlayerY(i) * PIC_Y + Player(i).yOffset + 24
        
        ' calculate the width to fill
        BarWidth = ((GetPlayerVital(i, Vitals.HP) / sWidth) / (GetPlayerMaxVital(i, Vitals.HP) / sWidth)) * sWidth
       
        If GetPlayerVital(i, Vitals.HP) > GetPlayerMaxVital(i, Vitals.HP) * 0.99 Then
            BarColor = 0
        ElseIf GetPlayerVital(i, Vitals.HP) > GetPlayerMaxVital(i, Vitals.HP) * 0.75 Then
            BarColor = 2
        ElseIf GetPlayerVital(i, Vitals.HP) > GetPlayerMaxVital(i, Vitals.HP) * 0.25 Then
            BarColor = 4
        Else
            BarColor = 6
        End If
       
        ' draw bar background
        With sRECT
            .Top = sHeight * 1 ' HP bar background
            .Left = 0
            .Right = .Left + sWidth
            .Bottom = .Top + sHeight
        End With
        RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
       
        ' draw the bar proper
        With sRECT
            .Top = sHeight * BarColor ' HP bar
            .Left = 0
            .Right = .Left + BarWidth
            .Bottom = .Top + sHeight
        End With
        RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
    End If

    ' check for casting time bar
    If SpellBuffer > 0 Then
        If Spell(Hotbar(SpellBuffer).Slot).Level(Hotbar(SpellBuffer).LvlSpell).CastTime > 0 Then
        
            ' lock to player
            tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).xOffset + 15 - (sWidth / 2)
            tmpY = GetPlayerY(MyIndex) * PIC_Y + Player(MyIndex).yOffset + 29
    
            ' calculate the width to fill
            If GetPlayerCasting(MyIndex) > 0 Then
                BarWidth = (GetTickCount - SpellBufferTimer) / (GetPlayerCasting(MyIndex)) * sWidth
            End If
    
            ' draw bar background
            With sRECT
                .Top = sHeight * 3 ' cooldown bar background
                .Left = 0
                .Right = sWidth
                .Bottom = .Top + sHeight
            End With
            RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
    
            ' draw the bar proper
            With sRECT
                .Top = sHeight * 8 ' cooldown bar
                .Left = 0
                .Right = BarWidth
                .Bottom = .Top + sHeight
            End With
            RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
        End If
    End If

End Sub

Public Sub DrawPlayer(ByVal Index As Long)
Dim Anim As Byte, i As Long, x As Long, y As Long
Dim sprite As Long, SpriteTop As Long
Dim Rec As RECT
Dim AttackSpeed As Long
Dim AnimStep As Byte, AnimAtk(1 To 5) As Byte
Dim StepX(1 To 2) As Byte, WAnim As Byte
Dim RingNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Vamos Pular Tudo Caso já esteja Morto!
    If Player(Index).Dead = True Then GoTo Morto

    ' Sprite Comum
    sprite = GetPlayerSprite(Index)
    
    ' Transformação!?
    If Player(Index).BuffSprite > 0 Then sprite = Player(Index).BuffSprite

    ' Personagens "Negros?"
    'If Player(Index).Cor = 1 Then
    '    Sprite = Sprite + 1
    'End If

    ' Anti OverFlow
    If sprite < 1 Or sprite > NumCharacters Then Exit Sub
    
    ' speed from weapon
    If GetPlayerEquipment(Index, Weapon) > 0 Then
        AttackSpeed = MyAttackSpeed
    Else
        AttackSpeed = 1000
    End If
    
    ' Check for attacking animation
    If Player(Index).AttackTimer + (AttackSpeed / 2) > GetTickCount Then
        If Player(Index).Attacking = 1 Then
            If GetPlayerEquipment(Index, Weapon) = 0 Then
                Anim = 1
            Else
                Anim = 0
            End If
        End If
    Else
        
        Select Case Player(Index).Step
        Case 1: AnimStep = 1
        Case 3: AnimStep = 1
        End Select
        
        Select Case GetPlayerDir(Index)
            Case DIR_UP
                If (Player(Index).yOffset > 8) Then
                Anim = AnimStep
                Player(Index).StandDelay = 300 + GetTickCount
                End If
            Case DIR_DOWN
                If (Player(Index).yOffset < -8) Then
                Anim = AnimStep
                Player(Index).StandDelay = 300 + GetTickCount
                End If
            Case DIR_LEFT, DIR_UP_LEFT, DIR_DOWN_LEFT
                If (Player(Index).xOffset > 8) Then
                Anim = AnimStep
                Player(Index).StandDelay = 300 + GetTickCount
                End If
            Case DIR_RIGHT, DIR_UP_RIGHT, DIR_DOWN_RIGHT
                If (Player(Index).xOffset < -8) Then
                Anim = AnimStep
                Player(Index).StandDelay = 300 + GetTickCount
                End If
        End Select
    End If

    ' Check to see if we want to stop making him attack
    With Player(Index)
        If .AttackTimer + AttackSpeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With
    
    ' Dead
    If Player(Index).Dead = True Then Anim = 4

    ' Set the left
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            SpriteTop = 2
        Case DIR_RIGHT, DIR_UP_RIGHT, DIR_DOWN_RIGHT
            SpriteTop = 0
        Case DIR_DOWN
            SpriteTop = 1
        Case DIR_LEFT, DIR_UP_LEFT, DIR_DOWN_LEFT
            SpriteTop = 3
    End Select

    With Rec
        .Top = Anim * (Tex_Character(sprite).Height / FraHeight)
        .Bottom = .Top + (Tex_Character(sprite).Height / FraHeight)
        .Left = SpriteTop * (Tex_Character(sprite).Width / FraWidth)
        .Right = .Left + (Tex_Character(sprite).Width / FraWidth)
    End With

    ' Calculate the X
    x = GetPlayerX(Index) * PIC_X + Player(Index).xOffset - ((Tex_Character(sprite).Width / FraWidth - 32) / 2) + 1

    ' Is the player's height more than 32..?
    If (Tex_Character(sprite).Height) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        y = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset - ((Tex_Character(sprite).Height / FraHeight) - 32) - 12
    Else
        ' Proceed as normal
        y = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset + 12
    End If
    
    ' Render the actual sprite
    If GetPlayerDir(Index) = DIR_UP Then
        Call DrawSprite(sprite, x, y, Rec)
        DrawHair x, y, Player(Index).Hair, Anim, SpriteTop, Index
    End If
    
    If Player(Index).Attacking = 0 Then
        For i = 1 To UBound(PaperdollOrder)
            If GetPlayerEquipment(Index, PaperdollOrder(i)) > 0 Then
                If Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll > 0 Then
                    If Player(Index).Dead = False Then
                        Call DrawPaperdoll(x, y, Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll, Anim, SpriteTop)
                    End If
                End If
            End If
        Next
    End If
    
    ' Render the actual sprite
    If GetPlayerDir(Index) <> DIR_UP Then
        Call DrawSprite(sprite, x, y, Rec)
        Call DrawHair(x, y, Player(Index).Hair, Anim, SpriteTop, Index)
    End If
    
    If GetPlayerDir(Index) = DIR_DOWN Then
        Call DrawSprite(sprite, x, y, Rec)
        DrawHair x, y, Player(Index).Hair, Anim, SpriteTop, Index
    End If
    
    ' Addons
    If Player(Index).BuffSprite = 0 Then
        If Player(Index).Addon(1) = True Then
            If Player(Index).Sex = 0 Then
                DrawAddon x, y, Outfit(Player(Index).Outfit).MAddon1, Anim, SpriteTop, Index
            Else
                DrawAddon x, y, Outfit(Player(Index).Outfit).FAddon1, Anim, SpriteTop, Index
            End If
        End If
        
        If Player(Index).Addon(2) = True Then
            If Player(Index).Sex = 0 Then
                DrawAddon x, y, Outfit(Player(Index).Outfit).MAddon2, Anim, SpriteTop, Index
            Else
                DrawAddon x, y, Outfit(Player(Index).Outfit).FAddon2, Anim, SpriteTop, Index
            End If
        End If
    End If
    
Morto:
    
    ' Fantasma
    If Player(Index).Dead = True Then
        Call DrawLight(x, y, 91, 0, 157, 255)
        DrawGhostPlayer Index
    End If
    
    ' Skulls
    If Player(Index).PlayerKiller(1) Then DrawSkullsPlayer Index
    
    ' Tocha
    'If PlayerLight = True Then
    '    If Index = MyIndex Then
    '        Call DrawLight(GetPlayerX(MyIndex) * 32 + Player(MyIndex).xOffset, GetPlayerY(MyIndex) * 32 + Player(MyIndex).yOffset, 50, 255, 255, 0)
    '    End If
    'End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawPlayer", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawNpc(ByVal MapNpcNum As Long)
Dim Anim As Byte, i As Long, x As Long, y As Long, sprite As Long, SpriteTop As Long
Dim Rec As RECT
Dim AttackSpeed As Long
Dim WalkAnim As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If MapNpc(MapNpcNum).num = 0 Then Exit Sub ' no npc set
    sprite = NPC(MapNpc(MapNpcNum).num).sprite

    If sprite < 1 Or sprite > NumCharacters Then Exit Sub

    AttackSpeed = 1000

    ' Reset frame
    Anim = 1
    
    ' Walk Frame
    Select Case MapNpc(MapNpcNum).Step
    Case 1: WalkAnim = 0
    Case 3: WalkAnim = 1
    End Select
    
    ' Npc Anim
    If NPC(MapNpc(MapNpcNum).num).FrameAnim > 0 Then
        Anim = MapNpc(MapNpcNum).AnimTimer
        GoTo Continue
    End If
    
    ' Check for attacking animation
    If MapNpc(MapNpcNum).AttackTimer + (AttackSpeed / 2) > GetTickCount Then
        If MapNpc(MapNpcNum).Attacking = 1 Then
            Anim = 1
        End If
    Else
        ' If not attacking, walk normally
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                If (MapNpc(MapNpcNum).yOffset > 8) Then Anim = WalkAnim
            Case DIR_DOWN
                If (MapNpc(MapNpcNum).yOffset < -8) Then Anim = WalkAnim
            Case DIR_LEFT
                If (MapNpc(MapNpcNum).xOffset > 8) Then Anim = WalkAnim
            Case DIR_RIGHT
                If (MapNpc(MapNpcNum).xOffset < -8) Then Anim = WalkAnim
        End Select
    End If

    ' Check to see if we want to stop making him attack
    With MapNpc(MapNpcNum)
        If .AttackTimer + AttackSpeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With
    
Continue:

    ' Set the left
    Select Case MapNpc(MapNpcNum).Dir
        Case DIR_UP
            SpriteTop = 2
        Case DIR_RIGHT
            SpriteTop = 0
        Case DIR_DOWN
            SpriteTop = 1
        Case DIR_LEFT
            SpriteTop = 3
    End Select

    With Rec
        .Top = Anim * (Tex_Character(sprite).Height / FraNpcHeight)
        .Bottom = .Top + Tex_Character(sprite).Height / FraNpcHeight
        .Left = SpriteTop * (Tex_Character(sprite).Width / FraNpcWidth)
        .Right = .Left + (Tex_Character(sprite).Width / FraNpcWidth)
    End With

    ' Calculate the X
    x = MapNpc(MapNpcNum).x * PIC_X + MapNpc(MapNpcNum).xOffset - ((Tex_Character(sprite).Width / FraNpcWidth - 32) / 2)

    ' Is the player's height more than 32..?
    If (Tex_Character(sprite).Height / FraHeight) > 32 Or (Tex_Character(sprite).Width / FraWidth) > 32 Then
    
        ' Create a 32 pixel offset for larger sprites
         y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).yOffset - ((Tex_Character(sprite).Height / FraHeight) - 20)
        
    Else
        ' Proceed as normal
        y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).yOffset - 12
    End If

    Call DrawSprite(sprite, x, y, Rec)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawNpc", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub DrawPaperdoll(ByVal X2 As Long, ByVal Y2 As Long, ByVal sprite As Long, ByVal Anim As Long, ByVal SpriteTop As Long, Optional ByVal AlfaColor As Byte, Optional ByVal Cor1 As Byte, Optional ByVal Cor2 As Byte, Optional ByVal Cor3 As Byte)
Dim Rec As RECT
Dim x As Long, y As Long
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If sprite < 1 Or sprite > NumPaperdolls Then Exit Sub
    
    With Rec
        .Top = SpriteTop * (Tex_Paperdoll(sprite).Height / 4)
        .Bottom = .Top + (Tex_Paperdoll(sprite).Height / 4)
        .Left = Anim * (Tex_Paperdoll(sprite).Width / 9)
        .Right = .Left + (Tex_Paperdoll(sprite).Width / 9)
    End With
    
    ' clipping
    x = ConvertMapX(X2)
    y = ConvertMapY(Y2)
    Width = (Rec.Right - Rec.Left)
    Height = (Rec.Bottom - Rec.Top)

    ' Clip to screen
    If y < 0 Then
        With Rec
            .Top = .Top - y
        End With
        y = 0
    End If

    If x < 0 Then
        With Rec
            .Left = .Left - x
        End With
        x = 0
    End If
    
    If AlfaColor = 0 Then
        RenderTexture Tex_Paperdoll(sprite), x, y, Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, D3DColorRGBA(255, 255, 255, 255)
    Else
        RenderTexture Tex_Paperdoll(sprite), x, y, Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, D3DColorRGBA(Cor1, Cor2, Cor3, AlfaColor)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawPaperdoll", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawHair(ByVal X2 As Long, ByVal Y2 As Long, ByVal sprite As Long, ByVal Anim As Long, ByVal SpriteTop As Long, ByVal Index As Long, Optional ByVal Dir As Byte)
Dim Rec As RECT
Dim x As Long, y As Long
Dim Width As Long, Height As Long
Dim CharSprite As Long
Dim HairNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    CharSprite = GetPlayerSprite(Index)
    If Player(Index).BuffSprite > 0 Then CharSprite = Player(Index).BuffSprite
    
    ' Evitar Problemas
    If sprite > sprite Then Exit Sub
    
    'If Sprite < 1 Or Sprite > NumHairs Then Exit Sub
    If sprite > 0 And sprite <= NumHairs Then
    
    ' Retirar Cabelo com Addon Especifico
    If Player(Index).BuffSprite = 0 Then
        If Player(Index).Addon(1) = True Then
            If Player(Index).Sex = 0 Then
                If Outfit(Player(Index).Outfit).HMAddOff1 = True Then Exit Sub
            Else
                If Outfit(Player(Index).Outfit).HFAddOff1 = True Then Exit Sub
            End If
        End If
        
        If Player(Index).Addon(2) = True Then
            If Player(Index).Sex = 0 Then
                If Outfit(Player(Index).Outfit).HMAddOff2 = True Then Exit Sub
            Else
                If Outfit(Player(Index).Outfit).HFAddOff2 = True Then Exit Sub
            End If
        End If
    End If
    
    If Player(Index).BuffSprite > 0 Then Exit Sub
    
    With Rec
        .Top = Anim * (Tex_Hair(sprite).Height / 2)
        .Bottom = .Top + (Tex_Hair(sprite).Height / 2)
        .Left = SpriteTop * (Tex_Hair(sprite).Width / 4)
        .Right = .Left + (Tex_Hair(sprite).Width / 4)
    End With
    
    ' clipping
    x = ConvertMapX(X2)
    y = ConvertMapY(Y2)
    Width = (Rec.Right - Rec.Left)
    Height = (Rec.Bottom - Rec.Top)
        
End If
    
    Select Case Dir
    Case DIR_DOWN, DIR_LEFT, DIR_UP_LEFT, DIR_DOWN_LEFT
        Rec.Left = (Anim + 1) * (Tex_Hair(sprite).Width / 4)
        Rec.Right = Rec.Left + (Tex_Hair(sprite).Width / 4)
        RenderTexture Tex_Hair(sprite), x, y, Rec.Left, Rec.Top, (Rec.Right - Rec.Left), Rec.Bottom - Rec.Top, (Rec.Right - Rec.Left) * -1, Rec.Bottom - Rec.Top, D3DColorRGBA(Player(Index).HairRGB(1), Player(Index).HairRGB(2), Player(Index).HairRGB(3), 255)
    Case Else
        RenderTexture Tex_Hair(sprite), x, y, Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, D3DColorRGBA(Player(Index).HairRGB(1), Player(Index).HairRGB(2), Player(Index).HairRGB(3), 255)
    End Select
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawCabelo", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawAddon(ByVal X2 As Long, ByVal Y2 As Long, ByVal sprite As Long, ByVal Anim As Long, ByVal SpriteTop As Long, ByVal Index As Long)
Dim Rec As RECT
Dim x As Long, y As Long
Dim Width As Long, Height As Long
Dim CharSprite As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Evitar Overflow
    If sprite > 0 And sprite <= NumAddon Then
    If Tex_Addon(sprite).Height = 0 Then Exit Sub
    If Tex_Addon(sprite).Width = 0 Then Exit Sub
    
    With Rec
        .Top = Anim * (Tex_Addon(sprite).Height / 2)
        .Bottom = .Top + (Tex_Addon(sprite).Height / 2)
        .Left = SpriteTop * (Tex_Addon(sprite).Width / 4)
        .Right = .Left + (Tex_Addon(sprite).Width / 4)
    End With
    
    ' clipping
    x = ConvertMapX(X2)
    y = ConvertMapY(Y2)
    Width = (Rec.Right - Rec.Left)
    Height = (Rec.Bottom - Rec.Top)

    ' Renderizar
    RenderTexture Tex_Addon(sprite), x, y, Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top

End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawAddon", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawSpriteRight(ByVal sprite As Long, ByVal X2 As Long, Y2 As Long, Rec As RECT, Optional ByVal color1 As Byte, Optional ByVal color2 As Byte, Optional ByVal Color3 As Byte, Optional ByVal Alpha As Byte)
Dim x As Long
Dim y As Long
Dim Width As Long
Dim Height As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If sprite < 1 Or sprite > NumCharacters Then Exit Sub
    x = ConvertMapX(X2)
    y = ConvertMapY(Y2)
    Width = (Rec.Right - Rec.Left)
    Height = (Rec.Bottom - Rec.Top)
    
    If Alpha > 0 Then
        RenderTexture Tex_Character(sprite), x, y, Rec.Left, Rec.Top, Width, Height, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, D3DColorRGBA(color1, color2, Color3, Alpha)
    Else
        RenderTexture Tex_Character(sprite), x, y, Rec.Left, Rec.Top, Width, Height, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, D3DColorRGBA(255, 255, 255, 255)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawSprite", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawSpriteLeft(ByVal sprite As Long, ByVal X2 As Long, Y2 As Long, Rec As RECT, Optional ByVal color1 As Byte, Optional ByVal color2 As Byte, Optional ByVal Color3 As Byte, Optional ByVal Alpha As Byte)
Dim x As Long
Dim y As Long
Dim Width As Long
Dim Height As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If sprite < 1 Or sprite > NumCharacters Then Exit Sub
    x = ConvertMapX(X2)
    y = ConvertMapY(Y2)
    Width = (Rec.Right - Rec.Left)
    Height = (Rec.Bottom - Rec.Top)
    
    If Alpha > 0 Then
        RenderTexture Tex_Character(sprite), x, y, Rec.Left, Rec.Top, Width, Height, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, D3DColorRGBA(color1, color2, Color3, Alpha)
    Else
        RenderTexture Tex_Character(sprite), x, y, Rec.Left, Rec.Top, Width, Height, (Rec.Right - Rec.Left) * -1, Rec.Bottom - Rec.Top, D3DColorRGBA(255, 255, 255, 255)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawSprite", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Private Sub DrawSprite(ByVal sprite As Long, ByVal X2 As Long, Y2 As Long, Rec As RECT, Optional ByVal color1 As Byte, Optional ByVal color2 As Byte, Optional ByVal Color3 As Byte, Optional ByVal Alpha As Byte)
Dim x As Long
Dim y As Long
Dim Width As Long
Dim Height As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If sprite < 1 Or sprite > NumCharacters Then Exit Sub
    x = ConvertMapX(X2)
    y = ConvertMapY(Y2)
    Width = (Rec.Right - Rec.Left)
    Height = (Rec.Bottom - Rec.Top)
    
    If Alpha > 0 Then
        RenderTexture Tex_Character(sprite), x, y, Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, D3DColorRGBA(color1, color2, Color3, Alpha)
    Else
        RenderTexture Tex_Character(sprite), x, y, Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, D3DColorRGBA(255, 255, 255, 255)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawSprite", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawFog()
Dim fogNum As Long, Color As Long, x As Long, y As Long, renderState As Long

    fogNum = CurrentFog
    If fogNum <= 0 Or fogNum > NumFogs Then Exit Sub
    Color = D3DColorRGBA(255, 255, 255, 255 - CurrentFogOpacity)

    renderState = 0
    ' render state
    Select Case renderState
        Case 1 ' Additive
            Direct3D_Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
            Direct3D_Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
        Case 2 ' Subtractive
            Direct3D_Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_SUBTRACT
            Direct3D_Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ZERO
            Direct3D_Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCCOLOR
    End Select
    
    For x = 0 To ((Map.MaxX * 32) / 256) + 1
        For y = 0 To ((Map.MaxY * 32) / 256) + 1
            RenderTexture Tex_Fog(fogNum), ConvertMapX((x * 256) + fogOffsetX), ConvertMapY((y * 256) + fogOffsetY), 0, 0, 256, 256, 256, 256, Color
        Next
    Next
    
    ' reset render state
    If renderState > 0 Then
        Direct3D_Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        Direct3D_Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        Direct3D_Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
    End If
End Sub

Public Sub DrawTint()
Dim Color As Long
    Color = D3DColorRGBA(CurrentTintR, CurrentTintG, CurrentTintB, CurrentTintA)
    RenderTexture Tex_White, 0, 0, 0, 0, frmMain.picScreen.ScaleWidth, frmMain.picScreen.ScaleHeight, 32, 32, Color
End Sub

Public Sub DrawWeather()
Dim Color As Long, i As Long, SpriteLeft As Long
    For i = 1 To MAX_WEATHER_PARTICLES
        If WeatherParticle(i).InUse Then
            If WeatherParticle(i).Type = WEATHER_TYPE_STORM Then
                SpriteLeft = 0
            Else
                SpriteLeft = WeatherParticle(i).Type - 1
            End If
            RenderTexture Tex_Weather, ConvertMapX(WeatherParticle(i).x), ConvertMapY(WeatherParticle(i).y), SpriteLeft * 32, 0, 32, 32, 32, 32, -1
        End If
    Next
End Sub

' ******************
' ** Game Editors **
' ******************
Public Sub EditorMap_DrawTileset()
Dim Height As Long, srcRect As D3DRECT, destRect As D3DRECT
Dim Width As Long
Dim Tileset As Long
Dim sRECT As RECT
Dim drect As RECT, scrlX As Long, scrlY As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' find tileset number
    Tileset = frmMain.scrlTileSet.value
    
    ' exit out if doesn't exist
    If Tileset < 0 Or Tileset > NumTileSets Then Exit Sub
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    
    scrlX = frmMain.scrlPictureX.value * PIC_X
    scrlY = frmMain.scrlPictureY.value * PIC_Y
    
    Height = Tex_Tileset(Tileset).Height - scrlY
    Width = Tex_Tileset(Tileset).Width - scrlX
    
    sRECT.Left = frmMain.scrlPictureX.value * PIC_X
    sRECT.Top = frmMain.scrlPictureY.value * PIC_Y
    sRECT.Right = sRECT.Left + Width
    sRECT.Bottom = sRECT.Top + Height
    
    drect.Top = 0
    drect.Bottom = Height
    drect.Left = 0
    drect.Right = Width
    
    RenderTextureByRects Tex_Tileset(Tileset), sRECT, drect
    
    ' change selected shape for autotiles
    If frmMain.scrlAutotile.value > 0 Then
        Select Case frmMain.scrlAutotile.value
            Case 1 ' autotile
                EditorTileWidth = 2
                EditorTileHeight = 3
            Case 2 ' fake autotile
                EditorTileWidth = 1
                EditorTileHeight = 1
            Case 3 ' animated
                EditorTileWidth = 6
                EditorTileHeight = 3
            Case 4 ' cliff
                EditorTileWidth = 2
                EditorTileHeight = 2
            Case 5 ' waterfall
                EditorTileWidth = 2
                EditorTileHeight = 3
        End Select
    End If
    
    With destRect
        .X1 = (EditorTileX * 32) - sRECT.Left
        .X2 = (EditorTileWidth * 32) + .X1
        .Y1 = (EditorTileY * 32) - sRECT.Top
        .Y2 = (EditorTileHeight * 32) + .Y1
    End With
    
    DrawSelectionBox destRect
        
    With srcRect
        .X1 = 0
        .X2 = Width
        .Y1 = 0
        .Y2 = Height
    End With
                    
    With destRect
        .X1 = 0
        .X2 = frmMain.picBack.ScaleWidth
        .Y1 = 0
        .Y2 = frmMain.picBack.ScaleHeight
    End With
    
    'Now render the selection tiles and we are done!
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmMain.picBack.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_DrawTileset", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub DrawSelectionBox(drect As D3DRECT)
Dim Width As Long, Height As Long, x As Long, y As Long
    Width = drect.X2 - drect.X1
    Height = drect.Y2 - drect.Y1
    x = drect.X1
    y = drect.Y1
    If Width > 6 And Height > 6 Then
        'Draw Box 32 by 32 at graphicselx and graphicsely
        RenderTexture Tex_Selection, x, y, 1, 1, 2, 2, 2, 2, -1 'top left corner
        RenderTexture Tex_Selection, x + 2, y, 3, 1, Width - 4, 2, 32 - 6, 2, -1 'top line
        RenderTexture Tex_Selection, x + 2 + (Width - 4), y, 29, 1, 2, 2, 2, 2, -1 'top right corner
        RenderTexture Tex_Selection, x, y + 2, 1, 3, 2, Height - 4, 2, 32 - 6, -1 'Left Line
        RenderTexture Tex_Selection, x + 2 + (Width - 4), y + 2, 32 - 3, 3, 2, Height - 4, 2, 32 - 6, -1 'right line
        RenderTexture Tex_Selection, x, y + 2 + (Height - 4), 1, 32 - 3, 2, 2, 2, 2, -1 'bottom left corner
        RenderTexture Tex_Selection, x + 2 + (Width - 4), y + 2 + (Height - 4), 32 - 3, 32 - 3, 2, 2, 2, 2, -1 'bottom right corner
        RenderTexture Tex_Selection, x + 2, y + 2 + (Height - 4), 3, 32 - 3, Width - 4, 2, 32 - 6, 2, -1 'bottom line
    End If
End Sub

Public Sub DrawTileOutline()
Dim Rec As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmMain.optBlock.value Then Exit Sub

    With Rec
        .Top = 0
        .Bottom = .Top + PIC_Y
        .Left = 0
        .Right = .Left + PIC_X
    End With

    RenderTexture Tex_Misc, ConvertMapX(CurX * PIC_X), ConvertMapY(CurY * PIC_Y), Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, D3DColorRGBA(255, 255, 255, 255)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawTileOutline", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub NewCharacterDrawSprite()
Dim sprite As Long, srcRect As D3DRECT, destRect As D3DRECT
Dim sRECT As RECT
Dim drect As RECT
Dim Width As Long, Height As Long
Dim Anim As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmMenu.cmbClass.ListIndex = -1 Then Exit Sub
    
    If sprite < 1 Or sprite > NumCharacters Then
        frmMenu.picSprite.Cls
        Exit Sub
    End If
    
    If NewCharHair = 0 Then NewCharHair = 1
    
    Width = Tex_Character(sprite).Width / FraWidth
    Height = Tex_Character(sprite).Height / FraHeight
    
    frmMenu.picSprite.Width = Width
    frmMenu.picSprite.Height = Height
    
    Anim = NewCharStep
    
    sRECT.Top = NewCharDir * (Tex_Character(sprite).Height / FraHeight)
    sRECT.Bottom = sRECT.Top + Height
    sRECT.Left = Anim * (Tex_Character(sprite).Width / FraWidth)
    sRECT.Right = sRECT.Left + Width
    
    drect.Top = 0
    drect.Bottom = Height
    drect.Left = 0
    drect.Right = Width
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    
    RenderTextureByRects Tex_Character(sprite), sRECT, drect
    RenderTextureByRects Tex_Hair(NewCharHair), sRECT, drect, D3DColorRGBA(NewCharHairRGB(1), NewCharHairRGB(2), NewCharHairRGB(3), 255)
    
    With srcRect
        .X1 = 0
        .X2 = Width
        .Y1 = 0
        .Y2 = Height
    End With
                    
    With destRect
        .X1 = 0
        .X2 = Width
        .Y1 = 0
        .Y2 = Height
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, destRect, frmMenu.picSprite.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NewCharacterDrawSprite", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorMap_DrawMapItem()
Dim ItemNum As Long
Dim sRECT As RECT, destRect As D3DRECT
Dim drect As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ItemNum = Item(frmMain.scrlMapItem.value).pic

    If ItemNum < 1 Or ItemNum > numitems Then
        frmMain.picMapItem.Cls
        Exit Sub
    End If

    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    drect.Top = 0
    drect.Bottom = PIC_Y
    drect.Left = 0
    drect.Right = PIC_X
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Item(ItemNum), sRECT, drect
    With destRect
        .X1 = 0
        .X2 = PIC_X
        .Y1 = 0
        .Y2 = PIC_Y
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmMain.picMapItem.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_DrawMapItem", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorMap_DrawKey()
Dim ItemNum As Long
Dim sRECT As RECT, destRect As D3DRECT
Dim drect As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ItemNum = Item(frmMain.scrlMapKey.value).pic

    If ItemNum < 1 Or ItemNum > numitems Then
        frmMain.picMapKey.Cls
        Exit Sub
    End If
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene

    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    drect.Top = 0
    drect.Bottom = PIC_Y
    drect.Left = 0
    drect.Right = PIC_X
    
    RenderTextureByRects Tex_Item(ItemNum), sRECT, drect
    
    With destRect
        .X1 = 0
        .X2 = PIC_X
        .Y1 = 0
        .Y2 = PIC_Y
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmMain.picMapKey.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_DrawKey", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorItem_DrawItem()
Dim ItemNum As Long
Dim sRECT As RECT, destRect As D3DRECT
Dim drect As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ItemNum = frmEditor_Item.scrlPic.value

    If ItemNum < 1 Or ItemNum > numitems Then
        frmEditor_Item.picItem.Cls
        Exit Sub
    End If


    ' rect for source
    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    
    ' same for destination as source
    drect = sRECT
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Item(ItemNum), sRECT, drect
    With destRect
        .X1 = 0
        .X2 = PIC_X
        .Y1 = 0
        .Y2 = PIC_Y
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Item.picItem.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorItem_DrawItem", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorItem_DrawPaperdoll()
Dim sprite As Long, srcRect As D3DRECT, destRect As D3DRECT
Dim sRECT As RECT
Dim drect As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    'frmEditor_Item.picPaperdoll.Cls
    
    sprite = frmEditor_Item.scrlPaperdoll.value

    If sprite < 1 Or sprite > NumPaperdolls Then
        frmEditor_Item.picPaperdoll.Cls
        Exit Sub
    End If

    ' rect for source
    sRECT.Top = 0
    sRECT.Bottom = Tex_Paperdoll(sprite).Height / 4
    sRECT.Left = 0
    sRECT.Right = Tex_Paperdoll(sprite).Width / 4
    ' same for destination as source
    drect = sRECT
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Paperdoll(sprite), sRECT, drect
                    
    With destRect
        .X1 = 0
        .X2 = Tex_Paperdoll(sprite).Width / 4
        .Y1 = 0
        .Y2 = Tex_Paperdoll(sprite).Height / 4
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Item.picPaperdoll.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorItem_DrawPaperdoll", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorSpell_DrawBuffSprite()
Dim IconNum As Long, destRect As D3DRECT
Dim sRECT As RECT
Dim drect As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IconNum = frmEditor_Spell.ScrlBuff(13).value
    
    If IconNum < 1 Or IconNum > NumCharacters Then
        frmEditor_Spell.BuffSprite.Cls
        Exit Sub
    End If
    
    With Tex_Character(IconNum)
        sRECT.Top = 0
        sRECT.Bottom = .Height / 4
        sRECT.Left = 0
        sRECT.Right = .Width / 3
        
        drect.Top = 0
        drect.Bottom = .Height / 4
        drect.Left = 0
        drect.Right = .Width / 3
    End With
    
    With destRect
        .X1 = 0
        .X2 = Tex_Character(IconNum).Width / 3
        .Y1 = 0
        .Y2 = Tex_Character(IconNum).Height / 4
    End With
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(255, 255, 255, 255), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Character(IconNum), sRECT, drect
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Spell.BuffSprite.hWnd, ByVal (0)
    

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorSpell_DrawBuffIcon", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorSpell_DrawBuffIcon()
Dim IconNum As Long, destRect As D3DRECT
Dim sRECT As RECT
Dim drect As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    IconNum = frmEditor_Spell.ScrlBuff(0).value
    
    If IconNum < 1 Or IconNum > NumBuff Then
        frmEditor_Spell.BuffIcon.Cls
        Exit Sub
    End If
    
    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    drect.Top = 0
    drect.Bottom = PIC_Y
    drect.Left = 0
    drect.Right = PIC_X
    
    With destRect
        .X1 = 0
        .X2 = PIC_X
        .Y1 = 0
        .Y2 = PIC_Y
    End With
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Buff(IconNum), sRECT, drect
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Spell.BuffIcon.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorSpell_DrawBuffIcon", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorSkillTree_DrawSpellIcon()
Dim IconNum As Long, destRect As D3DRECT
Dim sRECT As RECT
Dim drect As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    IconNum = frmEditor_Skilltree.scrlSTIcon.value
    
    If IconNum < 1 Or IconNum > NumSpellIcons Then
        frmEditor_Skilltree.STIcon.Cls
        Exit Sub
    End If
    
    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    drect.Top = 0
    drect.Bottom = PIC_Y
    drect.Left = 0
    drect.Right = PIC_X
    
    With destRect
        .X1 = 0
        .X2 = PIC_X
        .Y1 = 0
        .Y2 = PIC_Y
    End With
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_SpellIcon(IconNum), sRECT, drect
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Skilltree.STIcon.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorSkillTree_DrawSpellIcon", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorSpell_DrawIcon()
Dim IconNum As Long, destRect As D3DRECT
Dim sRECT As RECT
Dim drect As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IconNum = frmEditor_Spell.scrlIcon.value
    
    If IconNum < 1 Or IconNum > NumSpellIcons Then
        frmEditor_Spell.picSprite.Cls
        Exit Sub
    End If
    
    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    drect.Top = 0
    drect.Bottom = PIC_Y
    drect.Left = 0
    drect.Right = PIC_X
    
    With destRect
        .X1 = 0
        .X2 = PIC_X
        .Y1 = 0
        .Y2 = PIC_Y
    End With
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_SpellIcon(IconNum), sRECT, drect
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Spell.picSprite.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorSpell_DrawIcon", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorAnim_DrawAnim()
Dim AnimationNum As Long
Dim sRECT As RECT
Dim drect As RECT
Dim i As Long
Dim Width As Long, Height As Long, srcRect As D3DRECT, destRect As D3DRECT
Dim looptime As Long
Dim FrameCount As Long
Dim ShouldRender As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 0 To 1
        AnimationNum = frmEditor_Animation.scrlSprite(i).value
        
        If AnimationNum < 1 Or AnimationNum > NumAnimations Then
            frmEditor_Animation.picSprite(i).Cls
        Else
            looptime = frmEditor_Animation.scrlLoopTime(i)
            FrameCount = frmEditor_Animation.scrlFrameCount(i)
            
            ShouldRender = False
            
            ' check if we need to render new frame
            If AnimEditorTimer(i) + looptime <= GetTickCount Then
                ' check if out of range
                If AnimEditorFrame(i) >= FrameCount Then
                    AnimEditorFrame(i) = 1
                Else
                    AnimEditorFrame(i) = AnimEditorFrame(i) + 1
                End If
                AnimEditorTimer(i) = GetTickCount
                ShouldRender = True
            End If
        
            If ShouldRender Then
                'frmEditor_Animation.picSprite(i).Cls
                
                If frmEditor_Animation.scrlFrameCount(i).value > 0 Then
                    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                    Direct3D_Device.BeginScene
                    ' total width divided by frame count
                    Width = Tex_Animation(AnimationNum).Width / frmEditor_Animation.scrlFrameCount(i).value
                    Height = Tex_Animation(AnimationNum).Height
                    
                    sRECT.Top = 0
                    sRECT.Bottom = Height
                    sRECT.Left = (AnimEditorFrame(i) - 1) * Width
                    sRECT.Right = sRECT.Left + Width
                    
                    drect.Top = 0
                    drect.Bottom = Height
                    drect.Left = 0
                    drect.Right = Width
                    
                    RenderTextureByRects Tex_Animation(AnimationNum), sRECT, drect
                    
                    With srcRect
                        .X1 = 0
                        .X2 = frmEditor_Animation.picSprite(i).Width
                        .Y1 = 0
                        .Y2 = frmEditor_Animation.picSprite(i).Height
                    End With
                                
                    With destRect
                        .X1 = 0
                        .X2 = frmEditor_Animation.picSprite(i).Width
                        .Y1 = 0
                        .Y2 = frmEditor_Animation.picSprite(i).Height
                    End With
                                
                    Direct3D_Device.EndScene
                    Direct3D_Device.Present srcRect, destRect, frmEditor_Animation.picSprite(i).hWnd, ByVal (0)
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorAnim_DrawAnim", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorItem_DrawWFrames()
Dim sprite As Long, destRect As D3DRECT
Dim sRECT As RECT
Dim drect As RECT
Dim WitX As Long, WitY As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If frmEditor_Item.Visible = False Then Exit Sub
    
    sprite = frmEditor_Item.scrlWFrame.value

    ' Evitar OverFlow
    If sprite < 1 Or sprite > NumWFrames Then
        frmEditor_Item.PicWFrame.Cls
        Exit Sub
    End If

    WitX = Tex_WFrames(sprite).Width / 5
    WitY = Tex_WFrames(sprite).Height

    sRECT.Top = 0
    sRECT.Bottom = WitX
    sRECT.Left = 0
    sRECT.Right = WitY
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(255, 255, 255, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_WFrames(sprite), sRECT, sRECT
    
    With destRect
        .X1 = 0
        .X2 = WitX
        .Y1 = 0
        .Y2 = WitY
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Item.PicWFrame.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorItem_DrawWFrames", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Public Sub EditorNpc_DrawCorpse()
Dim sprite As Long, destRect As D3DRECT
Dim sRECT As RECT
Dim drect As RECT
Dim WitX As Long, WitY As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If frmEditor_NPC.Visible = False Then Exit Sub
    
    sprite = frmEditor_NPC.ScrlCorpse.value

    If sprite < 1 Or sprite > NumCorpses Then
        frmEditor_NPC.PicCorpse.Cls
        Exit Sub
    End If

    WitX = Tex_Corpses(sprite).Width
    WitY = Tex_Corpses(sprite).Height

    sRECT.Top = 0
    sRECT.Bottom = 64
    sRECT.Left = 0
    sRECT.Right = 64
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(255, 255, 255, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Corpses(sprite), sRECT, sRECT
    
    With destRect
        .X1 = 0
        .X2 = WitX
        .Y1 = 0
        .Y2 = WitY
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_NPC.PicCorpse.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorNpc_DrawCorpse", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorNpc_DrawSprite()
Dim sprite As Long, destRect As D3DRECT
Dim sRECT As RECT
Dim drect As RECT
Dim WitX As Long, WitY As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If frmEditor_NPC.Visible = False Then Exit Sub
    
    sprite = frmEditor_NPC.scrlSprite.value

    If sprite < 1 Or sprite > NumCharacters Then
        frmEditor_NPC.picSprite.Cls
        Exit Sub
    End If

    WitX = Tex_Character(sprite).Width / 4
    WitY = Tex_Character(sprite).Height / 2

    sRECT.Top = 0
    sRECT.Bottom = WitY
    sRECT.Left = WitX * Player(MyIndex).Resp
    sRECT.Right = sRECT.Left + WitX
    drect.Top = 0
    drect.Bottom = WitY
    drect.Left = 0
    drect.Right = WitX
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(255, 255, 255, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Character(sprite), sRECT, drect
    
    With destRect
        .X1 = 0
        .X2 = WitX
        .Y1 = 0
        .Y2 = WitY
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_NPC.picSprite.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorNpc_DrawSprite", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorChest_DrawSprite()
Dim sprite As Long
Dim sRECT As RECT, destRect As D3DRECT, srcRect As D3DRECT
Dim drect As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' normal sprite
    sprite = frmEditor_Chest.scrlPic(0).value

    If sprite < 1 Or sprite > NumResources Then
        frmEditor_Resource.picNormalPic.Cls
    Else
        sRECT.Top = 0
        sRECT.Bottom = Tex_Resource(sprite).Height
        sRECT.Left = 0
        sRECT.Right = Tex_Resource(sprite).Width
        drect.Top = 0
        drect.Bottom = Tex_Resource(sprite).Height
        drect.Left = 0
        drect.Right = Tex_Resource(sprite).Width
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene
        RenderTextureByRects Tex_Resource(sprite), sRECT, drect
        With srcRect
            .X1 = 0
            .X2 = Tex_Resource(sprite).Width
            .Y1 = 0
            .Y2 = Tex_Resource(sprite).Height
        End With
        
        With destRect
            .X1 = 0
            .X2 = frmEditor_Resource.picNormalPic.ScaleWidth
            .Y1 = 0
            .Y2 = frmEditor_Resource.picNormalPic.ScaleHeight
        End With
                    
        Direct3D_Device.EndScene
        Direct3D_Device.Present srcRect, srcRect, frmEditor_Chest.picSprite(0).hWnd, ByVal (0)
    End If

    ' exhausted sprite
    sprite = frmEditor_Chest.scrlPic(1).value

    If sprite < 1 Or sprite > NumResources Then
        frmEditor_Resource.picExhaustedPic.Cls
    Else
        sRECT.Top = 0
        sRECT.Bottom = Tex_Resource(sprite).Height
        sRECT.Left = 0
        sRECT.Right = Tex_Resource(sprite).Width
        
        drect.Top = 0
        drect.Bottom = Tex_Resource(sprite).Height
        drect.Left = 0
        drect.Right = Tex_Resource(sprite).Width
        
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene
        RenderTextureByRects Tex_Resource(sprite), sRECT, drect
        
        With destRect
            .X1 = 0
            .X2 = frmEditor_Resource.picExhaustedPic.ScaleWidth
            .Y1 = 0
            .Y2 = frmEditor_Resource.picExhaustedPic.ScaleHeight
        End With
        
        With srcRect
            .X1 = 0
            .X2 = Tex_Resource(sprite).Width
            .Y1 = 0
            .Y2 = Tex_Resource(sprite).Height
        End With
                    
        Direct3D_Device.EndScene
        Direct3D_Device.Present srcRect, srcRect, frmEditor_Chest.picSprite(1).hWnd, ByVal (0)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorResource_DrawSprite", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorResource_DrawSprite()
Dim sprite As Long
Dim sRECT As RECT, destRect As D3DRECT, srcRect As D3DRECT
Dim drect As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' normal sprite
    sprite = frmEditor_Resource.scrlNormalPic.value

    If sprite < 1 Or sprite > NumResources Then
        frmEditor_Resource.picNormalPic.Cls
    Else
        sRECT.Top = 0
        sRECT.Bottom = Tex_Resource(sprite).Height
        sRECT.Left = 0
        sRECT.Right = Tex_Resource(sprite).Width
        drect.Top = 0
        drect.Bottom = Tex_Resource(sprite).Height
        drect.Left = 0
        drect.Right = Tex_Resource(sprite).Width
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene
        RenderTextureByRects Tex_Resource(sprite), sRECT, drect
        With srcRect
            .X1 = 0
            .X2 = Tex_Resource(sprite).Width
            .Y1 = 0
            .Y2 = Tex_Resource(sprite).Height
        End With
        
        With destRect
            .X1 = 0
            .X2 = frmEditor_Resource.picNormalPic.ScaleWidth
            .Y1 = 0
            .Y2 = frmEditor_Resource.picNormalPic.ScaleHeight
        End With
                    
        Direct3D_Device.EndScene
        Direct3D_Device.Present srcRect, destRect, frmEditor_Resource.picNormalPic.hWnd, ByVal (0)
    End If

    ' exhausted sprite
    sprite = frmEditor_Resource.scrlExhaustedPic.value

    If sprite < 1 Or sprite > NumResources Then
        frmEditor_Resource.picExhaustedPic.Cls
    Else
        sRECT.Top = 0
        sRECT.Bottom = Tex_Resource(sprite).Height
        sRECT.Left = 0
        sRECT.Right = Tex_Resource(sprite).Width
        drect.Top = 0
        drect.Bottom = Tex_Resource(sprite).Height
        drect.Left = 0
        drect.Right = Tex_Resource(sprite).Width
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene
        RenderTextureByRects Tex_Resource(sprite), sRECT, drect
        
        With destRect
            .X1 = 0
            .X2 = frmEditor_Resource.picExhaustedPic.ScaleWidth
            .Y1 = 0
            .Y2 = frmEditor_Resource.picExhaustedPic.ScaleHeight
        End With
        
        With srcRect
            .X1 = 0
            .X2 = Tex_Resource(sprite).Width
            .Y1 = 0
            .Y2 = Tex_Resource(sprite).Height
        End With
                    
        Direct3D_Device.EndScene
        Direct3D_Device.Present srcRect, destRect, frmEditor_Resource.picExhaustedPic.hWnd, ByVal (0)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorResource_DrawSprite", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Render_Graphics()
    On Error GoTo errorhandler
    
    ' Verifica se podemos renderizar
    If Not CanRender() Then Exit Sub
    
    ' Atualiza câmera
    UpdateCamera
    
    ' Limpa e inicia cena
    ClearScene
    Direct3D_Device.BeginScene
    
    ' --- Camadas de render ---
    RenderBackground
    RenderLowerTiles
    RenderDecals
    RenderCorpses
    RenderTargetsAndHover
    RenderItems
    RenderMapEvents Position:=0
    RenderAnimations Layer:=0
    RenderYBasedObjects
    RenderProjectiles
    RenderAnimations Layer:=1
    RenderUpperTiles
    RenderMapEvents Position:=2
    RenderBars
    RenderLights
    RenderWeatherEffects
    RenderEditorOverlay
    RenderUI
    RenderPlayerBuffs
    
    RenderHUD
    
    ' Finaliza e apresenta cena
    Direct3D_Device.EndScene
    PresentScene
    Exit Sub
    
' -------------------
errorhandler:
    HandleRenderError
End Sub


Sub HandleDeviceLost()
'Do a loop while device is lost
   Do While Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST
       Exit Sub
   Loop
   
   UnloadTextures
   
   'Reset the device
   Direct3D_Device.Reset Direct3D_Window
   
   DirectX_ReInit
    
   LoadTextures
   
End Sub

Private Function DirectX_ReInit() As Boolean

    On Error GoTo Error_Handler

    Direct3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Display_Mode 'Use the current display mode that you
                                                                    'are already on. Incase you are confused, I'm
                                                                    'talking about your current screen resolution. ;)
        
    Direct3D_Window.Windowed = True 'The app will be in windowed mode.

    Direct3D_Window.SwapEffect = D3DSWAPEFFECT_COPY 'Refresh when the monitor does.
    Direct3D_Window.BackBufferFormat = Display_Mode.Format 'Sets the format that was retrieved into the backbuffer.
    'Creates the rendering device with some useful info, along with the info
    'we've already setup for Direct3D_Window.
    'Creates the rendering device with some useful info, along with the info
    Direct3D_Window.BackBufferCount = 1 '1 backbuffer only
    Direct3D_Window.BackBufferWidth = 1024 ' frmMain.picScreen.ScaleWidth 'Match the backbuffer width with the display width
    Direct3D_Window.BackBufferHeight = 726 'frmMain.picScreen.ScaleHeight 'Match the backbuffer height with the display height
    Direct3D_Window.hDeviceWindow = frmMain.picScreen.hWnd 'Use frmMain as the device window.
    
    With Direct3D_Device
        .SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE
    
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        .SetRenderState D3DRS_ZENABLE, False
        .SetRenderState D3DRS_ZWRITEENABLE, False
        
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
        .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
    End With
    
    DirectX_ReInit = True

    Exit Function
    
Error_Handler:
    MsgBox "An error occured while initializing DirectX", vbCritical
    
    DestroyGame
    
    DirectX_ReInit = False
End Function

Public Sub UpdateCamera()
Dim offsetX As Long
Dim offsetY As Long
Dim StartX As Long
Dim StartY As Long
Dim EndX As Long
Dim EndY As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    offsetX = Player(MyIndex).xOffset + PIC_X
    offsetY = Player(MyIndex).yOffset + PIC_Y
    StartX = GetPlayerX(MyIndex) - ((MAX_MAPX + 1) \ 2) - 1
    StartY = GetPlayerY(MyIndex) - ((MAX_MAPY) \ 2) - 1

    If StartX < 0 Then
        offsetX = 0

        If StartX = -1 Then
            If Player(MyIndex).xOffset > 0 Then
                offsetX = Player(MyIndex).xOffset
            End If
        End If

        StartX = 0
    End If

    If StartY < 0 Then
        offsetY = 0

        If StartY = -1 Then
            If Player(MyIndex).yOffset > 0 Then
                offsetY = Player(MyIndex).yOffset
            End If
        End If

        StartY = 0
    End If

    EndX = StartX + (MAX_MAPX + 1) + 1
    EndY = StartY + (MAX_MAPY + 1) + 1

    If EndX > Map.MaxX Then
        offsetX = 32

        If EndX = Map.MaxX + 1 Then
            If Player(MyIndex).xOffset < 0 Then
                offsetX = Player(MyIndex).xOffset + PIC_X
            End If
        End If

        EndX = Map.MaxX
        StartX = EndX - MAX_MAPX - 1
    End If

    If EndY > Map.MaxY Then
        offsetY = 32

        If EndY = Map.MaxY + 1 Then
            If Player(MyIndex).yOffset < 0 Then
                offsetY = Player(MyIndex).yOffset + PIC_Y
            End If
        End If

        EndY = Map.MaxY
        StartY = EndY - MAX_MAPY - 1
    End If

    With TileView
        .Top = StartY
        .Bottom = EndY
        .Left = StartX
        .Right = EndX
    End With

    With Camera
        .Top = offsetY
        .Bottom = .Top + ScreenY
        .Left = offsetX
        .Right = .Left + ScreenX
    End With
    
    UpdateDrawMapName

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdateCamera", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function ConvertMapX(ByVal x As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ConvertMapX = x - (TileView.Left * PIC_X) - Camera.Left
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConvertMapX", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function ConvertMapY(ByVal y As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ConvertMapY = y - (TileView.Top * PIC_Y) - Camera.Top
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConvertMapY", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function InViewPort(ByVal x As Long, ByVal y As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    InViewPort = False

    If x < TileView.Left Then Exit Function
    If y < TileView.Top Then Exit Function
    If x > TileView.Right Then Exit Function
    If y > TileView.Bottom Then Exit Function
    InViewPort = True
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "InViewPort", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsValidMapPoint(ByVal x As Long, ByVal y As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsValidMapPoint = False

    If x < 0 Then Exit Function
    If y < 0 Then Exit Function
    If x > Map.MaxX Then Exit Function
    If y > Map.MaxY Then Exit Function
    IsValidMapPoint = True
        
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsValidMapPoint", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub LoadTilesets()
Dim x As Long
Dim y As Long
Dim i As Long
Dim tilesetInUse() As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ReDim tilesetInUse(0 To NumTileSets)
    
    For x = 0 To Map.MaxX
        For y = 0 To Map.MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                ' check exists
                If Map.Tile(x, y).Layer(i).Tileset > 0 And Map.Tile(x, y).Layer(i).Tileset <= NumTileSets Then
                    tilesetInUse(Map.Tile(x, y).Layer(i).Tileset) = True
                End If
            Next
        Next
    Next
    
    For i = 1 To NumTileSets
        If tilesetInUse(i) Then
        
        Else
            ' unload tileset
            'Call ZeroMemory(ByVal VarPtr(DDSD_Tileset(i)), LenB(DDSD_Tileset(i)))
            'Set Tex_Tileset(i) = Nothing
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadTilesets", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawEvents()
Dim sRECT As RECT
Dim Width As Long, Height As Long, i As Long, x As Long, y As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Map.EventCount <= 0 Then Exit Sub
    
    For i = 1 To Map.EventCount
        If Map.Events(i).pageCount <= 0 Then
                sRECT.Top = 0
                sRECT.Bottom = 32
                sRECT.Left = 0
                sRECT.Right = 32
                RenderTexture Tex_Selection, ConvertMapX(x), ConvertMapY(y), sRECT.Left, sRECT.Right, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
            GoTo nextevent
        End If
        
        Width = 32
        Height = 32
    
        x = Map.Events(i).x * 32
        y = Map.Events(i).y * 32
        x = ConvertMapX(x)
        y = ConvertMapY(y)
    
        
        If i > Map.EventCount Then Exit Sub
        If 1 > Map.Events(i).pageCount Then Exit Sub
        Select Case Map.Events(i).Pages(1).GraphicType
            Case 0
                sRECT.Top = 0
                sRECT.Bottom = 32
                sRECT.Left = 0
                sRECT.Right = 32
                RenderTexture Tex_Selection, x, y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
            Case 1
                If Map.Events(i).Pages(1).Graphic > 0 And Map.Events(i).Pages(1).Graphic <= NumCharacters Then
                    
                    sRECT.Top = (Map.Events(i).Pages(1).GraphicY * (Tex_Character(Map.Events(i).Pages(1).Graphic).Height / 4))
                    sRECT.Left = (Map.Events(i).Pages(1).GraphicX * (Tex_Character(Map.Events(i).Pages(1).Graphic).Width / 4))
                    sRECT.Bottom = sRECT.Top + 32
                    sRECT.Right = sRECT.Left + 32
                    RenderTexture Tex_Character(Map.Events(i).Pages(1).Graphic), x, y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                    
                    sRECT.Top = 0
                    sRECT.Bottom = 32
                    sRECT.Left = 0
                    sRECT.Right = 32
                    RenderTexture Tex_Selection, x, y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                Else
                    sRECT.Top = 0
                    sRECT.Bottom = 32
                    sRECT.Left = 0
                    sRECT.Right = 32
                    RenderTexture Tex_Selection, x, y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                End If
            Case 2
                If Map.Events(i).Pages(1).Graphic > 0 And Map.Events(i).Pages(1).Graphic < NumTileSets Then
                    sRECT.Top = Map.Events(i).Pages(1).GraphicY * 32
                    sRECT.Left = Map.Events(i).Pages(1).GraphicX * 32
                    sRECT.Bottom = sRECT.Top + 32
                    sRECT.Right = sRECT.Left + 32
                    RenderTexture Tex_Tileset(Map.Events(i).Pages(1).Graphic), x, y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                    
                    sRECT.Top = 0
                    sRECT.Bottom = 32
                    sRECT.Left = 0
                    sRECT.Right = 32
                    RenderTexture Tex_Selection, x, y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                Else
                    sRECT.Top = 0
                    sRECT.Bottom = 32
                    sRECT.Left = 0
                    sRECT.Right = 32
                    RenderTexture Tex_Selection, x, y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                End If
        End Select
nextevent:
    Next
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawEvents", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorEvent_DrawGraphic()
Dim sRECT As RECT, destRect As D3DRECT, srcRect As D3DRECT
Dim drect As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Events.picGraphicSel.Visible Then
        Select Case frmEditor_Events.cmbGraphic.ListIndex
            Case 0
                'None
                frmEditor_Events.picGraphicSel.Cls
                Exit Sub
            Case 1
                If frmEditor_Events.scrlGraphic.value > 0 And frmEditor_Events.scrlGraphic.value <= NumCharacters Then
                    
                    If Tex_Character(frmEditor_Events.scrlGraphic.value).Width > 793 Then
                        sRECT.Left = frmEditor_Events.hScrlGraphicSel.value
                        sRECT.Right = sRECT.Left + (Tex_Character(frmEditor_Events.scrlGraphic.value).Width - sRECT.Left)
                    Else
                        sRECT.Left = 0
                        sRECT.Right = Tex_Character(frmEditor_Events.scrlGraphic.value).Width
                    End If
                    
                    If Tex_Character(frmEditor_Events.scrlGraphic.value).Height > 472 Then
                        sRECT.Top = frmEditor_Events.hScrlGraphicSel.value
                        sRECT.Bottom = sRECT.Top + (Tex_Character(frmEditor_Events.scrlGraphic.value).Height - sRECT.Top)
                    Else
                        sRECT.Top = 0
                        sRECT.Bottom = Tex_Character(frmEditor_Events.scrlGraphic.value).Height
                    End If
                    
                    With drect
                        .Top = 0
                        .Bottom = sRECT.Bottom - sRECT.Top
                        .Left = 0
                        .Right = sRECT.Right - sRECT.Left
                    End With
                    
                    With destRect
                        .X1 = drect.Left
                        .X2 = drect.Right
                        .Y1 = drect.Top
                        .Y2 = drect.Bottom
                    End With
                    
                    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                    Direct3D_Device.BeginScene
                    RenderTextureByRects Tex_Character(frmEditor_Events.scrlGraphic.value), sRECT, drect
                    If (GraphicSelX2 < GraphicSelX Or GraphicSelY2 < GraphicSelY) Or (GraphicSelX2 = 0 And GraphicSelY2 = 0) Then
                        With destRect
                            .X1 = (GraphicSelX * (Tex_Character(frmEditor_Events.scrlGraphic.value).Width / 3)) - sRECT.Left
                            .X2 = (Tex_Character(frmEditor_Events.scrlGraphic.value).Width / 4) + .X1
                            .Y1 = (GraphicSelY * (Tex_Character(frmEditor_Events.scrlGraphic.value).Height / 4)) - sRECT.Top
                            .Y2 = (Tex_Character(frmEditor_Events.scrlGraphic.value).Height / 4) + .Y1
                        End With

                    Else
                        With destRect
                            .X1 = (GraphicSelX * 32) - sRECT.Left
                            .X2 = ((GraphicSelX2 - GraphicSelX) * 32) + .X1
                            .Y1 = (GraphicSelY * 32) - sRECT.Top
                            .Y2 = ((GraphicSelY2 - GraphicSelY) * 32) + .Y1
                        End With
                    End If
                    DrawSelectionBox destRect
                    
                    With srcRect
                        .X1 = drect.Left
                        .X2 = frmEditor_Events.picGraphicSel.ScaleWidth
                        .Y1 = drect.Top
                        .Y2 = frmEditor_Events.picGraphicSel.ScaleHeight
                    End With
                    With destRect
                        .X1 = 0
                        .Y1 = 0
                        .X2 = frmEditor_Events.picGraphicSel.ScaleWidth
                        .Y2 = frmEditor_Events.picGraphicSel.ScaleHeight
                    End With
                    Direct3D_Device.EndScene
                    Direct3D_Device.Present srcRect, destRect, frmEditor_Events.picGraphicSel.hWnd, ByVal (0)
                    
                    If GraphicSelX <= 3 And GraphicSelY <= 3 Then
                    Else
                        GraphicSelX = 0
                        GraphicSelY = 0
                    End If
                Else
                    frmEditor_Events.picGraphicSel.Cls
                    Exit Sub
                End If
            Case 2
                If frmEditor_Events.scrlGraphic.value > 0 And frmEditor_Events.scrlGraphic.value <= NumTileSets Then
                    
                    If Tex_Tileset(frmEditor_Events.scrlGraphic.value).Width > 793 Then
                        sRECT.Left = frmEditor_Events.hScrlGraphicSel.value
                        sRECT.Right = sRECT.Left + 1024
                    Else
                        sRECT.Left = 0
                        sRECT.Right = Tex_Tileset(frmEditor_Events.scrlGraphic.value).Width
                        sRECT.Left = frmEditor_Events.hScrlGraphicSel.value = 0
                    End If
                    
                    If Tex_Tileset(frmEditor_Events.scrlGraphic.value).Height > 472 Then
                        sRECT.Top = frmEditor_Events.vScrlGraphicSel.value
                        sRECT.Bottom = sRECT.Top + 512
                    Else
                        sRECT.Top = 0
                        sRECT.Bottom = Tex_Tileset(frmEditor_Events.scrlGraphic.value).Height
                        frmEditor_Events.vScrlGraphicSel.value = 0
                    End If
                    
                    If sRECT.Left = -1 Then sRECT.Left = 0
                    If sRECT.Top = -1 Then sRECT.Top = 0
                    
                    With drect
                        .Top = 0
                        .Bottom = sRECT.Bottom - sRECT.Top
                        .Left = 0
                        .Right = sRECT.Right - sRECT.Left
                    End With
                    
                    
                    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                    Direct3D_Device.BeginScene
                    RenderTextureByRects Tex_Tileset(frmEditor_Events.scrlGraphic.value), sRECT, drect
                    If (GraphicSelX2 < GraphicSelX Or GraphicSelY2 < GraphicSelY) Or (GraphicSelX2 = 0 And GraphicSelY2 = 0) Then
                        With destRect
                            .X1 = (GraphicSelX * 32) - sRECT.Left
                            .X2 = PIC_X + .X1
                            .Y1 = (GraphicSelY * 32) - sRECT.Top
                            .Y2 = PIC_Y + .Y1
                        End With

                    Else
                        With destRect
                            .X1 = (GraphicSelX * 32) - sRECT.Left
                            .X2 = ((GraphicSelX2 - GraphicSelX) * 32) + .X1
                            .Y1 = (GraphicSelY * 32) - sRECT.Top
                            .Y2 = ((GraphicSelY2 - GraphicSelY) * 32) + .Y1
                        End With
                    End If
                    
                    DrawSelectionBox destRect
                    
                    With srcRect
                        .X1 = drect.Left
                        .X2 = frmEditor_Events.picGraphicSel.ScaleWidth
                        .Y1 = drect.Top
                        .Y2 = frmEditor_Events.picGraphicSel.ScaleHeight
                    End With
                    With destRect
                        .X1 = 0
                        .Y1 = 0
                        .X2 = frmEditor_Events.picGraphicSel.ScaleWidth
                        .Y2 = frmEditor_Events.picGraphicSel.ScaleHeight
                    End With
                    Direct3D_Device.EndScene
                    Direct3D_Device.Present srcRect, destRect, frmEditor_Events.picGraphicSel.hWnd, ByVal (0)
                Else
                    frmEditor_Events.picGraphicSel.Cls
                    Exit Sub
                End If
        End Select
    Else
        Select Case tmpEvent.Pages(curPageNum).GraphicType
            Case 0
                frmEditor_Events.picGraphic.Cls
            Case 1
                If tmpEvent.Pages(curPageNum).Graphic > 0 And tmpEvent.Pages(curPageNum).Graphic <= NumCharacters Then
                    sRECT.Top = tmpEvent.Pages(curPageNum).GraphicY * (Tex_Character(tmpEvent.Pages(curPageNum).Graphic).Height / 4)
                    sRECT.Left = tmpEvent.Pages(curPageNum).GraphicX * (Tex_Character(tmpEvent.Pages(curPageNum).Graphic).Width / 4)
                    sRECT.Bottom = sRECT.Top + (Tex_Character(tmpEvent.Pages(curPageNum).Graphic).Height / 4)
                    sRECT.Right = sRECT.Left + (Tex_Character(tmpEvent.Pages(curPageNum).Graphic).Width / 3)
                    With drect
                        drect.Top = (193 / 2) - ((sRECT.Bottom - sRECT.Top) / 2)
                        drect.Bottom = drect.Top + (sRECT.Bottom - sRECT.Top)
                        drect.Left = (121 / 2) - ((sRECT.Right - sRECT.Left) / 2)
                        drect.Right = drect.Left + (sRECT.Right - sRECT.Left)
                    End With
                    With destRect
                        .X1 = drect.Left
                        .X2 = drect.Right
                        .Y1 = drect.Top
                        .Y2 = drect.Bottom
                    End With
                    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                    Direct3D_Device.BeginScene
                    RenderTextureByRects Tex_Character(frmEditor_Events.scrlGraphic.value), sRECT, drect
                    Direct3D_Device.EndScene
                    Direct3D_Device.Present destRect, destRect, frmEditor_Events.picGraphic.hWnd, ByVal (0)
                End If
            Case 2
                If tmpEvent.Pages(curPageNum).Graphic > 0 And tmpEvent.Pages(curPageNum).Graphic <= NumTileSets Then
                    If tmpEvent.Pages(curPageNum).GraphicX2 = 0 Or tmpEvent.Pages(curPageNum).GraphicY2 = 0 Then
                        sRECT.Top = tmpEvent.Pages(curPageNum).GraphicY * 32
                        sRECT.Left = tmpEvent.Pages(curPageNum).GraphicX * 32
                        sRECT.Bottom = sRECT.Top + 32
                        sRECT.Right = sRECT.Left + 32
                        With drect
                            drect.Top = (193 / 2) - ((sRECT.Bottom - sRECT.Top) / 2)
                            drect.Bottom = drect.Top + (sRECT.Bottom - sRECT.Top)
                            drect.Left = (120 / 2) - ((sRECT.Right - sRECT.Left) / 2)
                            drect.Right = drect.Left + (sRECT.Right - sRECT.Left)
                        End With
                        With destRect
                            .X1 = drect.Left
                            .X2 = drect.Right
                            .Y1 = drect.Top
                            .Y2 = drect.Bottom
                        End With
                        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                        Direct3D_Device.BeginScene
                        RenderTextureByRects Tex_Tileset(frmEditor_Events.scrlGraphic.value), sRECT, drect
                        Direct3D_Device.EndScene
                        Direct3D_Device.Present destRect, destRect, frmEditor_Events.picGraphic.hWnd, ByVal (0)
                    Else
                        sRECT.Top = tmpEvent.Pages(curPageNum).GraphicY * 32
                        sRECT.Left = tmpEvent.Pages(curPageNum).GraphicX * 32
                        sRECT.Bottom = sRECT.Top + ((tmpEvent.Pages(curPageNum).GraphicY2 - tmpEvent.Pages(curPageNum).GraphicY) * 32)
                        sRECT.Right = sRECT.Left + ((tmpEvent.Pages(curPageNum).GraphicX2 - tmpEvent.Pages(curPageNum).GraphicX) * 32)
                        With drect
                            drect.Top = (193 / 2) - ((sRECT.Bottom - sRECT.Top) / 2)
                            drect.Bottom = drect.Top + (sRECT.Bottom - sRECT.Top)
                            drect.Left = (120 / 2) - ((sRECT.Right - sRECT.Left) / 2)
                            drect.Right = drect.Left + (sRECT.Right - sRECT.Left)
                        End With
                        With destRect
                            .X1 = drect.Left
                            .X2 = drect.Right
                            .Y1 = drect.Top
                            .Y2 = drect.Bottom
                        End With
                        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                        Direct3D_Device.BeginScene
                        RenderTextureByRects Tex_Tileset(frmEditor_Events.scrlGraphic.value), sRECT, drect
                        Direct3D_Device.EndScene
                        Direct3D_Device.Present destRect, destRect, frmEditor_Events.picGraphic.hWnd, ByVal (0)
                    End If
                End If
        End Select
    End If
    
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_DrawKey", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawEvent(id As Long)
    Dim x As Long, y As Long, Width As Long, Height As Long, sRECT As RECT, drect As RECT, Anim As Long, SpriteTop As Long
    If Map.MapEvents(id).Visible = 0 Then Exit Sub
    If InMapEditor Then Exit Sub
    Select Case Map.MapEvents(id).GraphicType
        Case 0
            Exit Sub
            
        Case 1
            If Map.MapEvents(id).GraphicNum <= 0 Or Map.MapEvents(id).GraphicNum > NumCharacters Then Exit Sub
            Width = Tex_Character(Map.MapEvents(id).GraphicNum).Width / 3
            Height = Tex_Character(Map.MapEvents(id).GraphicNum).Height / 4
            ' Reset frame
            If Map.MapEvents(id).Step = 3 Then
                Anim = 0
            ElseIf Map.MapEvents(id).Step = 1 Then
                Anim = 2
            End If
            
            Select Case Map.MapEvents(id).Dir
                Case DIR_UP
                    If (Map.MapEvents(id).yOffset > 8) Then Anim = Map.MapEvents(id).Step
                Case DIR_DOWN
                    If (Map.MapEvents(id).yOffset < -8) Then Anim = Map.MapEvents(id).Step
                Case DIR_LEFT
                    If (Map.MapEvents(id).xOffset > 8) Then Anim = Map.MapEvents(id).Step
                Case DIR_RIGHT
                    If (Map.MapEvents(id).xOffset < -8) Then Anim = Map.MapEvents(id).Step
            End Select
            
            ' Set the left
            Select Case Map.MapEvents(id).ShowDir
                Case DIR_UP
                    SpriteTop = 3
                Case DIR_RIGHT
                    SpriteTop = 2
                Case DIR_DOWN
                    SpriteTop = 0
                Case DIR_LEFT
                    SpriteTop = 1
            End Select
            
            If Map.MapEvents(id).WalkAnim = 1 Then Anim = 1
            
            If Map.MapEvents(id).Moving = 0 Then Anim = Map.MapEvents(id).GraphicX
            
            With sRECT
                .Top = SpriteTop * Height
                .Bottom = .Top + Height
                .Left = Anim * Width
                .Right = .Left + Width
            End With
        
            ' Calculate the X
            x = Map.MapEvents(id).x * PIC_X + Map.MapEvents(id).xOffset - ((Width - 32) / 2)
        
            ' Is the player's height more than 32..?
            If (Height * 4) > 32 Then
                ' Create a 32 pixel offset for larger sprites
                y = Map.MapEvents(id).y * PIC_Y + Map.MapEvents(id).yOffset - ((Height) - 32)
            Else
                ' Proceed as normal
                y = Map.MapEvents(id).y * PIC_Y + Map.MapEvents(id).yOffset
            End If
        
            ' render the actual sprite
            Call DrawSprite(Map.MapEvents(id).GraphicNum, x, y, sRECT)
            
        Case 2
            If Map.MapEvents(id).GraphicNum < 1 Or Map.MapEvents(id).GraphicNum > NumTileSets Then Exit Sub
            
            If Map.MapEvents(id).GraphicY2 > 0 Or Map.MapEvents(id).GraphicX2 > 0 Then
                With sRECT
                    .Top = Map.MapEvents(id).GraphicY * 32
                    .Bottom = .Top + ((Map.MapEvents(id).GraphicY2 - Map.MapEvents(id).GraphicY) * 32)
                    .Left = Map.MapEvents(id).GraphicX * 32
                    .Right = .Left + ((Map.MapEvents(id).GraphicX2 - Map.MapEvents(id).GraphicX) * 32)
                End With
            Else
                With sRECT
                    .Top = Map.MapEvents(id).GraphicY * 32
                    .Bottom = .Top + 32
                    .Left = Map.MapEvents(id).GraphicX * 32
                    .Right = .Left + 32
                End With
            End If
            
            x = Map.MapEvents(id).x * 32
            y = Map.MapEvents(id).y * 32
            
            x = x - ((sRECT.Right - sRECT.Left) / 2)
            y = y - (sRECT.Bottom - sRECT.Top) + 32
            
            
            If Map.MapEvents(id).GraphicY2 > 0 Then
                RenderTexture Tex_Tileset(Map.MapEvents(id).GraphicNum), ConvertMapX(Map.MapEvents(id).x * 32), ConvertMapY((Map.MapEvents(id).y - ((Map.MapEvents(id).GraphicY2 - Map.MapEvents(id).GraphicY) - 1)) * 32), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
            Else
                RenderTexture Tex_Tileset(Map.MapEvents(id).GraphicNum), ConvertMapX(Map.MapEvents(id).x * 32), ConvertMapY(Map.MapEvents(id).y * 32), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
            End If
    End Select
End Sub

'This function will make it much easier to setup the vertices with the info it needs.
Private Function Create_TLVertex(x As Single, y As Single, Z As Single, RHW As Single, Color As Long, Specular As Long, TU As Single, TV As Single) As TLVERTEX

    Create_TLVertex.x = x
    Create_TLVertex.y = y
    Create_TLVertex.Z = Z
    Create_TLVertex.RHW = RHW
    Create_TLVertex.Color = Color
    'Create_TLVertex.Specular = Specular
    Create_TLVertex.TU = TU
    Create_TLVertex.TV = TV
    
End Function

Public Function Ceiling(dblValIn As Double, dblCeilIn As Double) As Double
' round it
Ceiling = Round(dblValIn / dblCeilIn, 0) * dblCeilIn
' if it rounded down, force it up
If Ceiling < dblValIn Then Ceiling = Ceiling + dblCeilIn
End Function

Public Sub DestroyDX8()
    UnloadTextures
    Set Direct3DX = Nothing
    Set Direct3D_Device = Nothing
    Set Direct3D = Nothing
    Set DirectX8 = Nothing
End Sub

Public Sub DrawGDI()
    'Cycle Through in-game stuff before cycling through editors
    If frmMenu.Visible Then If frmMenu.picCharacter.Visible Then NewCharacterDrawSprite
    
    'Menu DX8
    If frmMenu.PicMenuDx8.Visible = True Then Call RenderDx8Menu
    
    If frmEditor_Animation.Visible Then EditorAnim_DrawAnim
    
    If frmEditor_Item.Visible Then
        EditorItem_DrawItem
        EditorItem_DrawPaperdoll
        EditorItem_DrawProjectile
        EditorItem_DrawProjectileTarget
        EditorItem_DrawWFrames
    End If
    
    If InMapEditor = True Then
        EditorMap_DrawTileset
        If frmMain.fraMapItem.Visible Then EditorMap_DrawMapItem
        If frmMain.fraMapKey.Visible Then EditorMap_DrawKey
        If frmMain.fraLight.Visible Then EditorMap_DrawLight
    End If
    
    If frmEditor_NPC.Visible = True Then
        EditorNpc_DrawSprite
        EditorNpc_DrawCorpse
    End If
    
    If frmEditor_Resource.Visible Then
        EditorResource_DrawSprite
    End If
    
    If frmEditor_Chest.Visible Then
        EditorChest_DrawSprite
    End If
    
    ' Spell Editor Icones
    If frmEditor_Spell.Visible Then
        EditorSpell_DrawIcon
        EditorSpell_DrawBuffIcon
        EditorSpell_DrawBuffSprite
    End If
    
    If frmEditor_Events.Visible Then
        EditorEvent_DrawGraphic
    End If
    
    If frmEditor_Skilltree.Visible = True Then EditorSkillTree_DrawSpellIcon
End Sub


'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'   All of this code is for auto tiles and the math behind generating them.
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Public Sub placeAutotile(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long, ByVal tileQuarter As Byte, ByVal autoTileLetter As String)
    With Autotile(x, y).Layer(layerNum).QuarterTile(tileQuarter)
        Select Case autoTileLetter
            Case "a"
                .x = autoInner(1).x
                .y = autoInner(1).y
            Case "b"
                .x = autoInner(2).x
                .y = autoInner(2).y
            Case "c"
                .x = autoInner(3).x
                .y = autoInner(3).y
            Case "d"
                .x = autoInner(4).x
                .y = autoInner(4).y
            Case "e"
                .x = autoNW(1).x
                .y = autoNW(1).y
            Case "f"
                .x = autoNW(2).x
                .y = autoNW(2).y
            Case "g"
                .x = autoNW(3).x
                .y = autoNW(3).y
            Case "h"
                .x = autoNW(4).x
                .y = autoNW(4).y
            Case "i"
                .x = autoNE(1).x
                .y = autoNE(1).y
            Case "j"
                .x = autoNE(2).x
                .y = autoNE(2).y
            Case "k"
                .x = autoNE(3).x
                .y = autoNE(3).y
            Case "l"
                .x = autoNE(4).x
                .y = autoNE(4).y
            Case "m"
                .x = autoSW(1).x
                .y = autoSW(1).y
            Case "n"
                .x = autoSW(2).x
                .y = autoSW(2).y
            Case "o"
                .x = autoSW(3).x
                .y = autoSW(3).y
            Case "p"
                .x = autoSW(4).x
                .y = autoSW(4).y
            Case "q"
                .x = autoSE(1).x
                .y = autoSE(1).y
            Case "r"
                .x = autoSE(2).x
                .y = autoSE(2).y
            Case "s"
                .x = autoSE(3).x
                .y = autoSE(3).y
            Case "t"
                .x = autoSE(4).x
                .y = autoSE(4).y
        End Select
    End With
End Sub

Public Sub initAutotiles()
Dim x As Long, y As Long, layerNum As Long
    ' Procedure used to cache autotile positions. All positioning is
    ' independant from the tileset. Calculations are convoluted and annoying.
    ' Maths is not my strong point. Luckily we're caching them so it's a one-off
    ' thing when the map is originally loaded. As such optimisation isn't an issue.
    
    ' For simplicity's sake we cache all subtile SOURCE positions in to an array.
    ' We also give letters to each subtile for easy rendering tweaks. ;]
    
    ' First, we need to re-size the array
    ReDim Autotile(0 To Map.MaxX, 0 To Map.MaxY)
    
    ' Inner tiles (Top right subtile region)
    ' NW - a
    autoInner(1).x = 32
    autoInner(1).y = 0
    
    ' NE - b
    autoInner(2).x = 48
    autoInner(2).y = 0
    
    ' SW - c
    autoInner(3).x = 32
    autoInner(3).y = 16
    
    ' SE - d
    autoInner(4).x = 48
    autoInner(4).y = 16
    
    ' Outer Tiles - NW (bottom subtile region)
    ' NW - e
    autoNW(1).x = 0
    autoNW(1).y = 32
    
    ' NE - f
    autoNW(2).x = 16
    autoNW(2).y = 32
    
    ' SW - g
    autoNW(3).x = 0
    autoNW(3).y = 48
    
    ' SE - h
    autoNW(4).x = 16
    autoNW(4).y = 48
    
    ' Outer Tiles - NE (bottom subtile region)
    ' NW - i
    autoNE(1).x = 32
    autoNE(1).y = 32
    
    ' NE - g
    autoNE(2).x = 48
    autoNE(2).y = 32
    
    ' SW - k
    autoNE(3).x = 32
    autoNE(3).y = 48
    
    ' SE - l
    autoNE(4).x = 48
    autoNE(4).y = 48
    
    ' Outer Tiles - SW (bottom subtile region)
    ' NW - m
    autoSW(1).x = 0
    autoSW(1).y = 64
    
    ' NE - n
    autoSW(2).x = 16
    autoSW(2).y = 64
    
    ' SW - o
    autoSW(3).x = 0
    autoSW(3).y = 80
    
    ' SE - p
    autoSW(4).x = 16
    autoSW(4).y = 80
    
    ' Outer Tiles - SE (bottom subtile region)
    ' NW - q
    autoSE(1).x = 32
    autoSE(1).y = 64
    
    ' NE - r
    autoSE(2).x = 48
    autoSE(2).y = 64
    
    ' SW - s
    autoSE(3).x = 32
    autoSE(3).y = 80
    
    ' SE - t
    autoSE(4).x = 48
    autoSE(4).y = 80
    
    For x = 0 To Map.MaxX
        For y = 0 To Map.MaxY
            For layerNum = 1 To MapLayer.Layer_Count - 1
                ' calculate the subtile positions and place them
                CalculateAutotile x, y, layerNum
                ' cache the rendering state of the tiles and set them
                CacheRenderState x, y, layerNum
            Next
        Next
    Next
End Sub

Public Sub CacheRenderState(ByVal x As Long, ByVal y As Long, ByVal layerNum As Long)
Dim quarterNum As Long

    ' exit out early
    If x < 0 Or x > Map.MaxX Or y < 0 Or y > Map.MaxY Then Exit Sub

    With Map.Tile(x, y)
        ' check if the tile can be rendered
        If .Layer(layerNum).Tileset <= 0 Or .Layer(layerNum).Tileset > NumTileSets Then
            Autotile(x, y).Layer(layerNum).renderState = RENDER_STATE_NONE
            Exit Sub
        End If
        
        ' check if it needs to be rendered as an autotile
        If .Autotile(layerNum) = AUTOTILE_NONE Or .Autotile(layerNum) = AUTOTILE_FAKE Then
            ' default to... default
            Autotile(x, y).Layer(layerNum).renderState = RENDER_STATE_NORMAL
        Else
            Autotile(x, y).Layer(layerNum).renderState = RENDER_STATE_AUTOTILE
            ' cache tileset positioning
            For quarterNum = 1 To 4
                Autotile(x, y).Layer(layerNum).srcX(quarterNum) = (Map.Tile(x, y).Layer(layerNum).x * 32) + Autotile(x, y).Layer(layerNum).QuarterTile(quarterNum).x
                Autotile(x, y).Layer(layerNum).srcY(quarterNum) = (Map.Tile(x, y).Layer(layerNum).y * 32) + Autotile(x, y).Layer(layerNum).QuarterTile(quarterNum).y
            Next
        End If
    End With
End Sub

Public Sub CalculateAutotile(ByVal x As Long, ByVal y As Long, ByVal layerNum As Long)
    ' Right, so we've split the tile block in to an easy to remember
    ' collection of letters. We now need to do the calculations to find
    ' out which little lettered block needs to be rendered. We do this
    ' by reading the surrounding tiles to check for matches.
    
    ' First we check to make sure an autotile situation is actually there.
    ' Then we calculate exactly which situation has arisen.
    ' The situations are "inner", "outer", "horizontal", "vertical" and "fill".
    
    ' Exit out if we don't have an auatotile
    If Map.Tile(x, y).Autotile(layerNum) = 0 Then Exit Sub
    
    ' Okay, we have autotiling but which one?
    Select Case Map.Tile(x, y).Autotile(layerNum)
    
        ' Normal or animated - same difference
        Case AUTOTILE_NORMAL, AUTOTILE_ANIM
            ' North West Quarter
            CalculateNW_Normal layerNum, x, y
            
            ' North East Quarter
            CalculateNE_Normal layerNum, x, y
            
            ' South West Quarter
            CalculateSW_Normal layerNum, x, y
            
            ' South East Quarter
            CalculateSE_Normal layerNum, x, y
            
        ' Cliff
        Case AUTOTILE_CLIFF
            ' North West Quarter
            CalculateNW_Cliff layerNum, x, y
            
            ' North East Quarter
            CalculateNE_Cliff layerNum, x, y
            
            ' South West Quarter
            CalculateSW_Cliff layerNum, x, y
            
            ' South East Quarter
            CalculateSE_Cliff layerNum, x, y
            
        ' Waterfalls
        Case AUTOTILE_WATERFALL
            ' North West Quarter
            CalculateNW_Waterfall layerNum, x, y
            
            ' North East Quarter
            CalculateNE_Waterfall layerNum, x, y
            
            ' South West Quarter
            CalculateSW_Waterfall layerNum, x, y
            
            ' South East Quarter
            CalculateSE_Waterfall layerNum, x, y
        
        ' Anything else
        Case Else
            ' Don't need to render anything... it's fake or not an autotile
    End Select
End Sub

' Normal autotiling
Public Sub CalculateNW_Normal(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North West
    If checkTileMatch(layerNum, x, y, x - 1, y - 1) Then tmpTile(1) = True
    
    ' North
    If checkTileMatch(layerNum, x, y, x, y - 1) Then tmpTile(2) = True
    
    ' West
    If checkTileMatch(layerNum, x, y, x - 1, y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(2) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(2) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(2) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If Not tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 1, "e"
        Case AUTO_OUTER
            placeAutotile layerNum, x, y, 1, "a"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 1, "i"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 1, "m"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 1, "q"
    End Select
End Sub

Public Sub CalculateNE_Normal(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North
    If checkTileMatch(layerNum, x, y, x, y - 1) Then tmpTile(1) = True
    
    ' North East
    If checkTileMatch(layerNum, x, y, x + 1, y - 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, x, y, x + 1, y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 2, "j"
        Case AUTO_OUTER
            placeAutotile layerNum, x, y, 2, "b"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 2, "f"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 2, "r"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 2, "n"
    End Select
End Sub

Public Sub CalculateSW_Normal(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' West
    If checkTileMatch(layerNum, x, y, x - 1, y) Then tmpTile(1) = True
    
    ' South West
    If checkTileMatch(layerNum, x, y, x - 1, y + 1) Then tmpTile(2) = True
    
    ' South
    If checkTileMatch(layerNum, x, y, x, y + 1) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 3, "o"
        Case AUTO_OUTER
            placeAutotile layerNum, x, y, 3, "c"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 3, "s"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 3, "g"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 3, "k"
    End Select
End Sub

Public Sub CalculateSE_Normal(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' South
    If checkTileMatch(layerNum, x, y, x, y + 1) Then tmpTile(1) = True
    
    ' South East
    If checkTileMatch(layerNum, x, y, x + 1, y + 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, x, y, x + 1, y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 4, "t"
        Case AUTO_OUTER
            placeAutotile layerNum, x, y, 4, "d"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 4, "p"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 4, "l"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 4, "h"
    End Select
End Sub

' Waterfall autotiling
Public Sub CalculateNW_Waterfall(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile As Boolean
    
    ' West
    If checkTileMatch(layerNum, x, y, x - 1, y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, x, y, 1, "i"
    Else
        ' Edge
        placeAutotile layerNum, x, y, 1, "e"
    End If
End Sub

Public Sub CalculateNE_Waterfall(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile As Boolean
    
    ' East
    If checkTileMatch(layerNum, x, y, x + 1, y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, x, y, 2, "f"
    Else
        ' Edge
        placeAutotile layerNum, x, y, 2, "j"
    End If
End Sub

Public Sub CalculateSW_Waterfall(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile As Boolean
    
    ' West
    If checkTileMatch(layerNum, x, y, x - 1, y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, x, y, 3, "k"
    Else
        ' Edge
        placeAutotile layerNum, x, y, 3, "g"
    End If
End Sub

Public Sub CalculateSE_Waterfall(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile As Boolean
    
    ' East
    If checkTileMatch(layerNum, x, y, x + 1, y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, x, y, 4, "h"
    Else
        ' Edge
        placeAutotile layerNum, x, y, 4, "l"
    End If
End Sub

' Cliff autotiling
Public Sub CalculateNW_Cliff(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North West
    If checkTileMatch(layerNum, x, y, x - 1, y - 1) Then tmpTile(1) = True
    
    ' North
    If checkTileMatch(layerNum, x, y, x, y - 1) Then tmpTile(2) = True
    
    ' West
    If checkTileMatch(layerNum, x, y, x - 1, y) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If Not tmpTile(2) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(2) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(2) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 1, "e"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 1, "i"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 1, "m"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 1, "q"
    End Select
End Sub

Public Sub CalculateNE_Cliff(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North
    If checkTileMatch(layerNum, x, y, x, y - 1) Then tmpTile(1) = True
    
    ' North East
    If checkTileMatch(layerNum, x, y, x + 1, y - 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, x, y, x + 1, y) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 2, "j"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 2, "f"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 2, "r"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 2, "n"
    End Select
End Sub

Public Sub CalculateSW_Cliff(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' West
    If checkTileMatch(layerNum, x, y, x - 1, y) Then tmpTile(1) = True
    
    ' South West
    If checkTileMatch(layerNum, x, y, x - 1, y + 1) Then tmpTile(2) = True
    
    ' South
    If checkTileMatch(layerNum, x, y, x, y + 1) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 3, "o"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 3, "s"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 3, "g"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 3, "k"
    End Select
End Sub

Public Sub CalculateSE_Cliff(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' South
    If checkTileMatch(layerNum, x, y, x, y + 1) Then tmpTile(1) = True
    
    ' South East
    If checkTileMatch(layerNum, x, y, x + 1, y + 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, x, y, x + 1, y) Then tmpTile(3) = True
    
    ' Calculate Situation -  Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 4, "t"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 4, "p"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 4, "l"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 4, "h"
    End Select
End Sub

Public Function checkTileMatch(ByVal layerNum As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Boolean
    ' we'll exit out early if true
    checkTileMatch = True
    
    ' if it's off the map then set it as autotile and exit out early
    If X2 < 0 Or X2 > Map.MaxX Or Y2 < 0 Or Y2 > Map.MaxY Then
        checkTileMatch = True
        Exit Function
    End If
    
    ' fakes ALWAYS return true
    If Map.Tile(X2, Y2).Autotile(layerNum) = AUTOTILE_FAKE Then
        checkTileMatch = True
        Exit Function
    End If
    
    ' check neighbour is an autotile
    If Map.Tile(X2, Y2).Autotile(layerNum) = 0 Then
        checkTileMatch = False
        Exit Function
    End If
    
    ' check we're a matching
    If Map.Tile(X1, Y1).Layer(layerNum).Tileset <> Map.Tile(X2, Y2).Layer(layerNum).Tileset Then
        checkTileMatch = False
        Exit Function
    End If
    
    ' check tiles match
    If Map.Tile(X1, Y1).Layer(layerNum).x <> Map.Tile(X2, Y2).Layer(layerNum).x Then
        checkTileMatch = False
        Exit Function
    End If
        
    If Map.Tile(X1, Y1).Layer(layerNum).y <> Map.Tile(X2, Y2).Layer(layerNum).y Then
        checkTileMatch = False
        Exit Function
    End If
End Function

Public Sub DrawAutoTile(ByVal layerNum As Long, ByVal destX As Long, ByVal destY As Long, ByVal quarterNum As Long, ByVal x As Long, ByVal y As Long)
Dim yOffset As Long, xOffset As Long

    ' calculate the offset
    Select Case Map.Tile(x, y).Autotile(layerNum)
        Case AUTOTILE_WATERFALL
            yOffset = (waterfallFrame - 1) * 32
        Case AUTOTILE_ANIM
            xOffset = autoTileFrame * 64
        Case AUTOTILE_CLIFF
            yOffset = -32
    End Select
    
    ' Draw the quarter
    RenderTexture Tex_Tileset(Map.Tile(x, y).Layer(layerNum).Tileset), destX, destY, Autotile(x, y).Layer(layerNum).srcX(quarterNum) + xOffset, Autotile(x, y).Layer(layerNum).srcY(quarterNum) + yOffset, 16, 16, 16, 16
End Sub

Public Sub RenderDx8Menu()
Dim srcRect As D3DRECT, destRect As D3DRECT
Dim sRECT As RECT, drect As RECT
Dim Width As Long, Height As Long
Dim i As Long, x As Long, y As Long
Dim n As Long, Anim As Byte
Dim sprite As Integer
Dim LenPass As Byte, OPass As String
Dim ClassPaperdoll As Integer
Dim NewStatus(1 To 7) As String
Dim HairView As Boolean
Dim LocX As Long, LocY As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    i = MENU_BACKGROUND
    
    Width = Tex_Menu_Win(i).Width
    Height = Tex_Menu_Win(i).Height
    
    sRECT.Top = 0
    sRECT.Bottom = sRECT.Top + Height
    sRECT.Left = 0
    sRECT.Right = sRECT.Left + Width
    
    drect.Top = 0
    drect.Bottom = Height
    drect.Left = 0
    drect.Right = Width
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    
    ' BackGround
    RenderTextureByRects Tex_Menu_Win(i), sRECT, drect
    
    ' Logotipo
    i = MENU_LOGO
    If Menu_Win(i).Visible = True Then
        RenderTexture Tex_Menu_Win(i), 154, 9, 0, 0, Tex_Menu_Win(i).Width, Tex_Menu_Win(i).Height, Tex_Menu_Win(i).Width, Tex_Menu_Win(i).Height
    End If
    
    ' News
    i = MENU_NEWS
    If Menu_Win(i).Visible = True Then
        'Janela
        RenderTexture Tex_Menu_Win(i), 229, 141, 0, 0, Tex_Menu_Win(i).Width, Tex_Menu_Win(i).Height, Tex_Menu_Win(i).Width, Tex_Menu_Win(i).Height
        
        ' Botões
        For i = 1 To 3
            With Menu_Buttons(i)
                If .Visible = True Then
                    RenderTexture Tex_Menu_Buttons(.Picture), .Left, .Top, 0, .state * (.Height / 3), .Witdh, .Height / 3, .Witdh, .Height / 3
                End If
            End With
        Next
    End If
    
    ' Login
    i = MENU_LOGIN
    If Menu_Win(i).Visible = True Then
        'Janela
        RenderTexture Tex_Menu_Win(i), 270, 141, 0, 0, Tex_Menu_Win(i).Width, Tex_Menu_Win(i).Height, Tex_Menu_Win(i).Width, Tex_Menu_Win(i).Height
        
        ' Botões
        For i = 4 To 7
            With Menu_Buttons(i)
                If .Visible = True Then
                    RenderTexture Tex_Menu_Buttons(.Picture), .Left, .Top, 0, .state * (.Height / 3), .Witdh, .Height / 3, .Witdh, .Height / 3
                End If
            End With
        Next
        
        ' Textos
        i = 1
        With Menu_Text(i)
            If Text_Focus = i Then
                RenderText Font_Default, .Text & ChatShowLine, .Left, .Top, .Color, 0, False
            Else
                RenderText Font_Default, .Text, .Left, .Top, .Color, 0, False
            End If
        End With
        
        i = 2
        With Menu_Text(i)
        
        OPass = vbNullString
        For n = 1 To Len(.Text)
            OPass = OPass & "•"
        Next
        
            If Text_Focus = i Then
                RenderText Font_Default, OPass & ChatShowLine, .Left, .Top, .Color, 0, False
            Else
                RenderText Font_Default, OPass, .Left, .Top, .Color, 0, False
            End If
        End With
        
    End If
    
    ' Registro
    i = MENU_REGISTER
    If Menu_Win(i).Visible = True Then
        'Janela
        RenderTexture Tex_Menu_Win(i), 230, 93, 0, 0, Tex_Menu_Win(i).Width, Tex_Menu_Win(i).Height, Tex_Menu_Win(i).Width, Tex_Menu_Win(i).Height
        
        ' Botões
        For i = 10 To 11
            With Menu_Buttons(i)
                If .Visible = True Then
                    RenderTexture Tex_Menu_Buttons(.Picture), .Left, .Top, 0, .state * (.Height / 3), .Witdh, .Height / 3, .Witdh, .Height / 3
                End If
            End With
        Next
        
        ' Textos
        For i = 3 To 10
        With Menu_Text(i)
            If Text_Focus = i Then
                RenderText Font_Default, .Text & ChatShowLine, .Left, .Top, .Color, 0, False
            Else
                If i = 10 Then .Left = 398 - (EngineGetTextWidth(Font_Default, .Text) / 2)
                RenderText Font_Default, .Text, .Left, .Top, .Color, 0, False
            End If
        End With
        Next
        
    End If
    
    ' Recovery
    i = MENU_RECOVERY
    If Menu_Win(i).Visible = True Then
        'Janela
        RenderTexture Tex_Menu_Win(i), 269, 246, 0, 0, Tex_Menu_Win(i).Width, Tex_Menu_Win(i).Height, Tex_Menu_Win(i).Width, Tex_Menu_Win(i).Height
        
        ' Botões
        For i = 8 To 9
            With Menu_Buttons(i)
                If .Visible = True Then
                    If .LockState = True Then .state = 1
                    RenderTexture Tex_Menu_Buttons(.Picture), .Left, .Top, 0, .state * (.Height / 3), .Witdh, .Height / 3, .Witdh, .Height / 3
                End If
            End With
        Next
        
    End If
    
    ' Select Character
    i = MENU_SELECT
    If Menu_Win(i).Visible = True Then

        'Janela
        RenderTexture Tex_Menu_Win(i), 204, 130, 0, 0, Tex_Menu_Win(i).Width, Tex_Menu_Win(i).Height, Tex_Menu_Win(i).Width, Tex_Menu_Win(i).Height
           

        For i = 1 To 3
        HairView = True
        
        With SelectCharInfo(i)
            
            Width = Tex_Character(.sprite).Width / 4
            Height = Tex_Character(.sprite).Height / 2
            
            If i = SelectCharSlot Then
                Anim = NewCharStep
                sRECT.Top = Anim * Height
                sRECT.Bottom = sRECT.Top + Height
                sRECT.Right = sRECT.Left + Width
                sRECT.Left = 0 * Width
                drect.Top = 0
                drect.Bottom = Height
                drect.Left = 0
                drect.Right = Width
            Else
                Anim = NewCharStep
                sRECT.Top = 0 * Height
                sRECT.Bottom = sRECT.Top + Height
                sRECT.Right = sRECT.Left + Width
                sRECT.Left = 0 * Width
                drect.Top = 0
                drect.Bottom = Height
                drect.Left = 0
                drect.Right = Width
            End If
            
            ' Addon Config
            If .Sex = 0 Then
                If .Addon(1) = True Then
                    If Outfit(.Outfit).HMAddOff1 = True Then HairView = False
                End If
                
                If .Addon(2) = True Then
                    If Outfit(.Outfit).HMAddOff2 = True Then HairView = False
                End If
            Else
                If .Addon(1) = True Then
                    If Outfit(.Outfit).HFAddOff1 = True Then HairView = False
                End If
                
                If .Addon(2) = True Then
                    If Outfit(.Outfit).HFAddOff2 = True Then HairView = False
                End If
            End If
            
                If .sprite > 0 Then
                    ' Nome do Personagem
                    If Len(Trim$(SelectCharInfo(i).Name)) > 0 Then
                        If i = SelectCharSlot Then
                            RenderText Font_Georgia, Trim$(SelectCharInfo(i).Name), 310 + ((i - 1) * 89) - (EngineGetTextWidth(Font_Georgia, Trim$(SelectCharInfo(i).Name)) / 2), 167, Yellow, 0, False
                        Else
                            RenderText Font_Georgia, Trim$(SelectCharInfo(i).Name), 310 + ((i - 1) * 89) - (EngineGetTextWidth(Font_Georgia, Trim$(SelectCharInfo(i).Name)) / 2), 167, White, 0, False
                        End If
                    End If

                    'Localização
                    LocX = 286: LocY = 190

                    If i <> SelectCharSlot Then
                        
                        ' Sprite Base
                        RenderTexture Tex_Character(.sprite), LocX + ((i - 1) * 89), LocY, sRECT.Left, sRECT.Top, Width, Height, Width, Height, D3DColorRGBA(255, 255, 255, 255)
                        
                        ' Cabelos
                        If .Hair > 0 And .Hair <= NumHairs Then
                            If HairView = True Then
                                RenderTexture Tex_Hair(.Hair), LocX + ((i - 1) * 89), LocY, sRECT.Left, sRECT.Top, Width, Height, Width, Height, D3DColorRGBA(.HairRGB(1), .HairRGB(2), .HairRGB(3), 255)
                            End If
                        End If
                        
                        'Addons
                        If .Sex = 0 Then
                            If .Addon(1) = True Then RenderTexture Tex_Addon(Outfit(.Outfit).MAddon1), LocX + ((i - 1) * 89), LocY, sRECT.Left, sRECT.Top, Width, Height, Width, Height
                        Else
                            If .Addon(1) = True Then RenderTexture Tex_Addon(Outfit(.Outfit).FAddon1), LocX + ((i - 1) * 89), LocY, sRECT.Left, sRECT.Top, Width, Height, Width, Height
                        End If
                    Else
                    
                        ' Sprite Base
                        RenderTexture Tex_Character(.sprite), LocX + ((i - 1) * 89), LocY, sRECT.Left, sRECT.Top, Width, Height, Width, Height
                        
                        ' Cabelos
                        If .Hair > 0 And .Hair <= NumHairs Then
                            If HairView = True Then
                                RenderTexture Tex_Hair(.Hair), LocX + ((i - 1) * 89), LocY, sRECT.Left, sRECT.Top, Width, Height, Width, Height, D3DColorRGBA(.HairRGB(1), .HairRGB(2), .HairRGB(3), 255)
                            End If
                        End If
                        
                        'Addons
                        If .Sex = 0 Then
                            If .Addon(1) = True Then RenderTexture Tex_Addon(Outfit(.Outfit).MAddon1), LocX + ((i - 1) * 89), LocY, sRECT.Left, sRECT.Top, Width, Height, Width, Height
                        Else
                            If .Addon(1) = True Then RenderTexture Tex_Addon(Outfit(.Outfit).FAddon1), LocX + ((i - 1) * 89), LocY, sRECT.Left, sRECT.Top, Width, Height, Width, Height
                        End If
                    
                    End If
                End If
            End With
        Next
        
        ' Botões
        For i = 12 To 14
            With Menu_Buttons(i)
                If .Visible = True Then
                    RenderTexture Tex_Menu_Buttons(.Picture), .Left, .Top, 0, .state * (.Height / 3), .Witdh, .Height / 3, .Witdh, .Height / 3
                End If
            End With
        Next
        
        If Menu_Buttons(26).Visible = True Then
            RenderTexture Tex_Menu_Buttons(Menu_Buttons(26).Picture), Menu_Buttons(26).Left, Menu_Buttons(26).Top, 0, Menu_Buttons(26).state * (Menu_Buttons(26).Height / 3), Menu_Buttons(26).Witdh, Menu_Buttons(26).Height / 3, Menu_Buttons(26).Witdh, Menu_Buttons(26).Height / 3
        End If
        
        i = 17
        With Menu_Buttons(i)
            If .Visible = True Then
                RenderTexture Tex_Menu_Buttons(.Picture), .Left, .Top, 0, .state * (.Height / 3), .Witdh, .Height / 3, .Witdh, .Height / 3
            End If
        End With
    End If
    
    ' Deletar Personagem?
    i = MENU_DEL
    If Menu_Win(i).Visible = True Then
        'Janela
        RenderTexture Tex_Menu_Win(i), 289, 186, 0, 0, Tex_Menu_Win(i).Width, Tex_Menu_Win(i).Height, Tex_Menu_Win(i).Width, Tex_Menu_Win(i).Height
     
        RenderText Font_Default, "Deseja Realmente excluir o", 305, 200, White, 0
        RenderText Font_Default, "Personagem: " & Trim$(SelectCharInfo(SelectCharSlot).Name) & ". Não", 305, 215, White, 0
        RenderText Font_Default, "é possivel reverter este processo.", 305, 228, White, 0
        
        ' Botões
        For i = 15 To 16
            With Menu_Buttons(i)
                If .Visible = True Then
                    RenderTexture Tex_Menu_Buttons(.Picture), .Left, .Top, 0, .state * (.Height / 3), .Witdh, .Height / 3, .Witdh, .Height / 3
                End If
            End With
        Next
     End If
     
    ' Select Char
    i = MENU_SELECTSLOT
    If Menu_Win(i).Visible = True Then
        'Janela
        RenderTexture Tex_Menu_Win(i), 0, 294, 0, 0, Tex_Menu_Win(i).Width, Tex_Menu_Win(i).Height, Tex_Menu_Win(i).Width, Tex_Menu_Win(i).Height
     End If
    
    ' Menu Character Visual
    i = MENU_CHARACTER
    If Menu_Win(i).Visible = True Then
    
        ' Janela
        RenderTexture Tex_Menu_Win(i), 204, 106, 0, 0, Tex_Menu_Win(i).Width, Tex_Menu_Win(i).Height, Tex_Menu_Win(i).Width, Tex_Menu_Win(i).Height
                
        ' Select Class
        With Tex_Target
            Select Case NewCharClass
            Case 1
                RenderTexture Tex_Target, 209, 277, 0, 0, 96, .Height, 96, .Height, -1
                
                ' Masculino
                If NewCharSex = 0 Then
                    If NewCharCOR = 0 Then
                        sprite = 1
                    Else
                        sprite = 2
                    End If
                End If
                
                ' Feminino
                If NewCharSex = 1 Then
                    If NewCharCOR = 0 Then
                        sprite = 1
                    Else
                        sprite = 1
                    End If
                End If
                
                NewStatus(1) = "Guerreiro"
                NewStatus(2) = 5
                NewStatus(3) = 8
                NewStatus(4) = 2
                NewStatus(5) = 5
                NewStatus(6) = 5
                
            Case 2
                RenderTexture Tex_Target, 266, 277, 0, 0, 96, .Height, 96, .Height, -1
                
                ' Masculino
                If NewCharSex = 0 Then
                    If NewCharCOR = 0 Then
                        sprite = 1
                    Else
                        sprite = 1
                    End If
                End If
                
                ' Feminino
                If NewCharSex = 1 Then
                    If NewCharCOR = 0 Then
                        sprite = 1
                    Else
                        sprite = 1
                    End If
                End If
                
                NewStatus(1) = "Mago"
                NewStatus(2) = 2 'Str
                NewStatus(3) = 2 'Vit
                NewStatus(4) = 3 'Agi
                NewStatus(5) = 10 'Int
                NewStatus(6) = 8 'Dex
                
            Case 3
                RenderTexture Tex_Target, 323, 277, 0, 0, 96, .Height, 96, .Height, -1

                ' Masculino
                If NewCharSex = 0 Then
                    If NewCharCOR = 0 Then
                        sprite = 1
                    Else
                        sprite = 1
                    End If
                End If
                
                ' Feminino
                If NewCharSex = 1 Then
                    If NewCharCOR = 0 Then
                        sprite = 1
                    Else
                        sprite = 1
                    End If
                End If
                
                NewStatus(1) = "Arqueiro"
                NewStatus(2) = 2 'Str
                NewStatus(3) = 2 'Vit
                NewStatus(4) = 8 'Agi
                NewStatus(5) = 3 'Int
                NewStatus(6) = 10 'Dex
                
            Case 4
                RenderTexture Tex_Target, 380, 277, 0, 0, 96, .Height, 96, .Height, -1
                
                ' Masculino
                If NewCharSex = 0 Then
                    If NewCharCOR = 0 Then
                        sprite = 1
                    Else
                        sprite = 1
                    End If
                End If
                
                ' Feminino
                If NewCharSex = 1 Then
                    If NewCharCOR = 0 Then
                        sprite = 1
                    Else
                        sprite = 1
                    End If
                End If
            
                NewStatus(1) = "Ladino"
                NewStatus(2) = 10 'Str
                NewStatus(3) = 5 'Vit
                NewStatus(4) = 3 'Agi
                NewStatus(5) = 1 'Int
                NewStatus(6) = 6 'Dex
                
            Case 5
                RenderTexture Tex_Target, 437, 277, 0, 0, 96, .Height, 96, .Height, -1
                
                ' Masculino
                If NewCharSex = 0 Then
                    If NewCharCOR = 0 Then
                        sprite = 1
                    Else
                        sprite = 1
                    End If
                End If
                
                ' Feminino
                If NewCharSex = 1 Then
                    If NewCharCOR = 0 Then
                        sprite = 1
                    Else
                        sprite = 1
                    End If
                End If
            
                NewStatus(1) = "Clérigo"
                NewStatus(2) = 2 'Str
                NewStatus(3) = 8 'Vit
                NewStatus(4) = 2 'Agi
                NewStatus(5) = 8 'Int
                NewStatus(6) = 5 'Dex
                
            Case 6
                RenderTexture Tex_Target, 494, 277, 0, 0, 96, .Height, 96, .Height, -1
                
                ' Masculino
                If NewCharSex = 0 Then
                    If NewCharCOR = 0 Then
                        sprite = 1
                    Else
                        sprite = 1
                    End If
                End If
                
                ' Feminino
                If NewCharSex = 1 Then
                    If NewCharCOR = 0 Then
                        sprite = 1
                    Else
                        sprite = 1
                    End If
                End If
            
                NewStatus(1) = "Evocador"
                NewStatus(2) = 1 'Str
                NewStatus(3) = 5 'Vit
                NewStatus(4) = 5 'Agi
                NewStatus(5) = 8 'Int
                NewStatus(6) = 7 'Dex
                
            End Select
        End With
        
        ' Status Iniciais
        RenderText Font_Georgia, Trim$(NewStatus(1)), 518 - (EngineGetTextWidth(Font_Georgia, Trim$(NewStatus(1))) / 2), 133, White, 0
        RenderText Font_Georgia, Trim$(NewStatus(2)), 500, 153, White, 0
        RenderText Font_Georgia, Trim$(NewStatus(3)), 536, 173, White, 0
        RenderText Font_Georgia, Trim$(NewStatus(4)), 524, 194, White, 0
        RenderText Font_Georgia, Trim$(NewStatus(5)), 539, 214, White, 0
        RenderText Font_Georgia, Trim$(NewStatus(6)), 522, 234, White, 0
        
        Width = Tex_Character(sprite).Width / 4
        Height = Tex_Character(sprite).Height / 2
        
        Anim = NewCharStepIdle
        
        sRECT.Top = 0
        sRECT.Bottom = sRECT.Top + Height
        sRECT.Left = (Tex_Character(sprite).Width / 4)
        sRECT.Right = sRECT.Left + Width
        
        drect.Top = 0
        drect.Bottom = Height
        drect.Left = 0
        drect.Right = Width
        
        If NewCharDir = 0 Then
            RenderTexture Tex_Character(sprite), 240, 135, sRECT.Left, sRECT.Top, Width, Height, Width, Height
            RenderTexture Tex_Hair(NewCharHair), 240, 135, sRECT.Left, sRECT.Top, Width, Height, Width, Height, D3DColorRGBA(NewCharHairRGB(1), NewCharHairRGB(2), NewCharHairRGB(3), 255)
        Else
            sRECT.Left = (Anim + 1) * (Tex_Character(sprite).Width / 6)
            sRECT.Right = sRECT.Left + Width
            RenderTexture Tex_Character(sprite), 240, 135, sRECT.Left, sRECT.Top, Width, Height, Width * -1, Height
            RenderTexture Tex_Hair(NewCharHair), 240, 135, sRECT.Left, sRECT.Top, Width, Height, Width * -1, Height, D3DColorRGBA(NewCharHairRGB(1), NewCharHairRGB(2), NewCharHairRGB(3), 255)
        End If
        
        ' Botões
        For i = 18 To 25
            With Menu_Buttons(i)
                If .Visible = True Then
                    RenderTexture Tex_Menu_Buttons(.Picture), .Left, .Top, 0, .state * (.Height / 3), .Witdh, .Height / 3, .Witdh, .Height / 3
                End If
            End With
        Next
        
        ' Textos
        i = 11
        With Menu_Text(i)
            If Text_Focus = i Then
                RenderText Font_Default, .Text & ChatShowLine, .Left - (EngineGetTextWidth(Font_Default, .Text) / 2), .Top, .Color, 0, False
            Else
                RenderText Font_Default, .Text, .Left - (EngineGetTextWidth(Font_Default, .Text) / 2), .Top, .Color, 0, False
            End If
        End With
        
        ' Erro
        i = 12
        With Menu_Text(i)
            RenderText Font_Default, NewCharMsgError, .Left - (EngineGetTextWidth(Font_Default, NewCharMsgError) / 2), .Top, .Color, 0, False
        End With
        
    End If
    
    With srcRect
        .X1 = 0
        .X2 = frmMenu.PicMenuDx8.Width
        .Y1 = 0
        .Y2 = frmMenu.PicMenuDx8.Height
    End With
                    
    With destRect
        .X1 = 0
        .X2 = frmMenu.PicMenuDx8.Width
        .Y1 = 0
        .Y2 = frmMenu.PicMenuDx8.Height
    End With
    
    ' Novidades
    i = MENU_NEWS
    If Menu_Win(i).Visible = True Then
        RenderText Font_Default, "Bem Vindo à Darknessfall Online! Novo Aqui!?", 265, 208, White, 0, False
        RenderText Font_Default, "Para começar clique em [Registrar] e cadastre-se.", 260, 222, White, 0, False
    End If
    
    ' Servidor Online?
    RenderText Font_Default, "Servidor:", 567, 425, White, 0, False
    RenderText Font_Default, "Jogadores:", 670, 425, White, 0, False
    
    If ServerStatus = True Then
        RenderText Font_Default, "Online", 625, 425, Green, 0, False
        RenderText Font_Default, "[" & OnlinePlayers & "/70]", 735, 425, Green, 0, False
    Else
        RenderText Font_Default, "Offline", 625, 425, BrightRed, 0, False
        RenderText Font_Default, "[0/70]", 735, 425, BrightRed, 0, False
    End If

    If FadeAmountMenu > 0 Then RenderTexture Tex_Fade, 0, 0, 0, 0, frmMenu.PicMenuDx8.ScaleWidth, frmMenu.PicMenuDx8.ScaleHeight, 32, 32, D3DColorRGBA(255, 255, 255, FadeAmountMenu)

    ' Cursor
    i = MENU_CURSOR
    If Menu_Win(i).Visible = True Then
        RenderTexture Tex_Menu_Win(i), GlobalX - 6, GlobalY - 4, 0, 0, 32, 32, 32, 32
    End If

    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, destRect, frmMenu.PicMenuDx8.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "RenderDx8Menu", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorItem_DrawProjectile()
Dim ItemNum As Long
Dim sRECT As RECT, destRect As D3DRECT
Dim drect As RECT
Dim Width As Long, Height As Long
Dim srcRect As D3DRECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ItemNum = frmEditor_Item.ScrlPicProject.value

    If ItemNum < 1 Or ItemNum > NumProjectTiles Then
        frmEditor_Item.PicProjétil.Cls
        Exit Sub
    End If
    
    If frmEditor_Item.ScrlPicFrame.value = 0 Then
        frmEditor_Item.PicProjétil.Cls
        Exit Sub
    End If
    
    If frmEditor_Item.ScrlPicFrame.value = 1 Then AnimProjetil = 0
    
    ' rect for source
    sRECT.Top = DirProjetil * (Tex_ProjectTiles(ItemNum).Height / 4)
    sRECT.Bottom = sRECT.Top + Tex_ProjectTiles(ItemNum).Height / 4
    sRECT.Left = AnimProjetil * (Tex_ProjectTiles(ItemNum).Width / frmEditor_Item.ScrlPicFrame.value)
    sRECT.Right = sRECT.Left + Tex_ProjectTiles(ItemNum).Width / frmEditor_Item.ScrlPicFrame.value
    
    ' same for destination as source
    drect.Top = 0
    drect.Bottom = Tex_ProjectTiles(ItemNum).Height / 4
    drect.Left = 0
    drect.Right = Tex_ProjectTiles(ItemNum).Width / frmEditor_Item.ScrlPicFrame.value
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_ProjectTiles(ItemNum), sRECT, drect
                    
    With destRect
        .X1 = 0
        .X2 = Tex_ProjectTiles(ItemNum).Width / frmEditor_Item.ScrlPicFrame.value
        .Y1 = 0
        .Y2 = Tex_ProjectTiles(ItemNum).Height / 4
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Item.PicProjétil.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorItem_DrawItem", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' player Projectiles
Public Sub DrawProjectile(ByVal Index As Long, ByVal PlayerProjectile As Long)
Dim x As Long, y As Long, PicNum As Long, i As Long
Dim sRECT As RECT, drect As RECT
Dim DirectionNew As Long
Dim Width As Long, Height As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' check for subscript error
    If Index < 1 Or PlayerProjectile < 1 Or PlayerProjectile > MAX_PLAYER_PROJECTILES Then Exit Sub
    
    ' check to see if it's time to move the Projectile
    If GetTickCount > Player(Index).Projectile(PlayerProjectile).TravelTime Then
        With Player(Index).Projectile(PlayerProjectile)
            ' set next travel time and the current position and then set the actual direction based on RMXP arrow tiles.
            Select Case .Direction
                ' down
                Case 0
                    .y = .y + 1
                    ' check if they reached maxrange
                    If .y = (GetPlayerY(Index) + .Range) + 1 Then ClearProjectile Index, PlayerProjectile: Exit Sub
                ' up
                Case 1
                    .y = .y - 1
                    ' check if they reached maxrange
                    If .y = (GetPlayerY(Index) - .Range) - 1 Then ClearProjectile Index, PlayerProjectile: Exit Sub
                ' right
                Case 2
                    .x = .x + 1
                    ' check if they reached max range
                    If .x = (GetPlayerX(Index) + .Range) + 1 Then ClearProjectile Index, PlayerProjectile: Exit Sub
                ' left
                Case 3
                    .x = .x - 1
                    ' check if they reached maxrange
                    If .x = (GetPlayerX(Index) - .Range) - 1 Then ClearProjectile Index, PlayerProjectile: Exit Sub
            End Select
            .TravelTime = GetTickCount + .speed
        End With
    End If
    
    ' set the x, y & pic values for future reference
    x = Player(Index).Projectile(PlayerProjectile).x
    y = Player(Index).Projectile(PlayerProjectile).y
    PicNum = Player(Index).Projectile(PlayerProjectile).pic
    
    ' check if left map
    If x > Map.MaxX Or y > Map.MaxY Or x < 0 Or y < 0 Then
        ClearProjectile Index, PlayerProjectile
        Exit Sub
    End If
    
    ' check if we hit a block
    If Map.Tile(x, y).Type = TILE_TYPE_BLOCKED Then
        MeAnimation Player(Index).Projectile(PlayerProjectile).Animation, ConvertMapX(x), ConvertMapY(y)
        ClearProjectile Index, PlayerProjectile
        Exit Sub
    End If
    
    ' check for player hit
    For i = 1 To Player_HighIndex
        If x = GetPlayerX(i) And y = GetPlayerY(i) Then
            ' they're hit, remove it
            If Not x = Player(MyIndex).x Or Not y = GetPlayerY(MyIndex) Then
                ClearProjectile Index, PlayerProjectile
                Exit Sub
            End If
        End If
    Next
    
    ' check for npc hit
    For i = 1 To MAX_MAP_NPCS
        If x = MapNpc(i).x And y = MapNpc(i).y Then
            ' they're hit, remove it
            MeAnimation Player(Index).Projectile(PlayerProjectile).Animation, ConvertMapX(MapNpc(i).x), ConvertMapY(MapNpc(i).y)
            ClearProjectile Index, PlayerProjectile
            Exit Sub
        End If
    Next
    
    ' Evitar Erro
    If Player(Index).Projectile(PlayerProjectile).FrameCount < Player(Index).Projectile(PlayerProjectile).LoopAnim Then Player(Index).Projectile(PlayerProjectile).LoopAnim = 0
        
    'New Direction
    Select Case Player(Index).Projectile(PlayerProjectile).Direction
    Case 0: DirectionNew = 3
    Case 1: DirectionNew = 0
    Case 2: DirectionNew = 1
    Case 3: DirectionNew = 2
    End Select
    
    ' rect for source
    sRECT.Top = DirectionNew * (Tex_ProjectTiles(PicNum).Height / 4)
    sRECT.Bottom = sRECT.Top + Tex_ProjectTiles(PicNum).Height / 4
    sRECT.Left = Player(Index).Projectile(PlayerProjectile).LoopAnim * (Tex_ProjectTiles(PicNum).Width / Player(Index).Projectile(PlayerProjectile).FrameCount)
    sRECT.Right = sRECT.Left + Tex_ProjectTiles(PicNum).Width / Player(Index).Projectile(PlayerProjectile).FrameCount
    
    x = ConvertMapX(x * 32)
    y = ConvertMapY(y * 32) - 12
    
    RenderTexture Tex_ProjectTiles(PicNum), x, y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, -1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawProjectile", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawScreenHudDx8()
Dim i As Long, x As Long, y As Long
Dim sRECT As RECT, drect As RECT
Dim LockState As Long, InvValue As String

    ' Player Hp
    If ScreenshotMapMode = True Then Exit Sub
    
    Call DrawHudDx8

    ' Janela de Chat
    If ChatOpen = True Then
           
        With Tex_Main_Win(10)
            RenderTexture Tex_Main_Win(10), Main_Win(9).Left, Main_Win(9).Top, 0, 0, .Width, .Height, .Width, .Height
            RenderText Font_Georgia, RenderTextChat & ChatCursor, 108, 484, White, 1
        End With
        
    End If
    
    ChatText_Draw
    
    ' Esc Options
    If Main_Win(11).Visible = True Then
        With Tex_Main_Win(14)
            RenderTexture Tex_Main_Win(14), Main_Win(11).Left, Main_Win(11).Top, 0, 0, .Width, .Height, .Width, .Height
        End With
    End If
    
    ' Janelas Dx8 - Personagem
    If Main_Win(2).Visible = True Then DrawCharacterInfoDx8
    
    ' Janela Party
    If Main_Win(15).Visible = True Then DrawPartyDx8
    
    ' Janelas Dx8 - Habilidades
    If Main_Win(3).Visible = True Then DrawSpellsDx8
    
    ' Janelas Dx8 - Shop
    If Main_Win(13).Visible = True Then DrawShopDx8
    
    ' Janelas Dx8 - Inventario
    If Main_Win(1).Visible = True Then
        DrawInventoryDx8
        DrawEquipmentDx8
    End If
    
    ' Janelas Dx8 - Sons Config
    If Main_Win(20).Visible = True Then DrawSoundConfigDx8
    
    ' Janelas Dx8 - Quest Dialog
    If Main_Win(19).Visible = True Then DrawDialogQuestDx8
    
    ' Janelas Dx8 - Quests
    If Main_Win(18).Visible = True Then DrawQuestsDx8
    
    ' Janela Dx8 - Refino
    If Main_Win(21).Visible = True Then
        DrawRefineDx8
    End If
    
    ' Janela Dx8 - Craft Item
    If Main_Win(27).Visible = True Then
        DrawCraftItem
    End If
    
    ' Janelas Dx8 - CurrencyDrop
    If Main_Win(4).Visible = True Then
    
        With Tex_Main_Win(4)
            RenderTexture Tex_Main_Win(4), Main_Win(4).Left, Main_Win(4).Top, 0, 0, .Width, .Height, .Width, .Height
            RenderText Font_Georgia, Trim$(DropCurrencyText), Main_Win(4).Left + 130 - (EngineGetTextWidth(Font_Georgia, Trim$(DropCurrencyText)) / 2), Main_Win(4).Top + 34, White, 0, True
            RenderText Font_Georgia, Trim$(Main_Text(2).Text), Main_Win(4).Left + 115 - (EngineGetTextWidth(Font_Georgia, Trim$(Main_Text(2).Text)) / 2), Main_Win(4).Top + 66, White, 0, True
            RenderTexture Tex_Item(Item(DropCurrencyItem).pic), Main_Win(4).Left + 22, Main_Win(4).Top + 27, 0, 0, 32, 32, 32, 32
            'InvValue = Format$(ConvertCurrency(PlayerInv(tmpCurrencyItem).value), "#,###,###,###")
            
            RenderText Font_Georgia, Trim$(InvValue), Main_Win(4).Left + 34 - (EngineGetTextWidth(Font_Georgia, Trim$(InvValue)) / 2), Main_Win(4).Top + 45, White, 0, True
        End With
        
         ' Botões Ok/Cancel
        For i = 17 To 18
            If MAIN_BUTTON(i).Visible = True Then
                With Tex_Main_Buttons(MAIN_BUTTON(i).Picture)
                    If MAIN_BUTTON(i).LockState = False Then
                        x = ConvertMapX(3 * 32) + Tex_Main_Buttons(MAIN_BUTTON(i).Picture).Width + 25
                        y = ConvertMapY(3 * 32)
                        RenderTexture Tex_Main_Buttons(MAIN_BUTTON(i).Picture), MAIN_BUTTON(i).Left, MAIN_BUTTON(i).Top, 0, MAIN_BUTTON(i).state * (.Height / 3), .Width, .Height / 3, .Width, .Height / 3
                    Else
                        x = ConvertMapX(3 * 32) + Tex_Main_Buttons(MAIN_BUTTON(i).Picture).Width + 25
                        y = ConvertMapY(3 * 32)
                        
                        Select Case MAIN_BUTTON(i).state
                            Case 0: LockState = 2
                            Case 1: LockState = 1
                            Case 2: LockState = 0
                        End Select
                        RenderTexture Tex_Main_Buttons(MAIN_BUTTON(i).Picture), MAIN_BUTTON(i).Left, MAIN_BUTTON(i).Top, 0, LockState * (.Height / 3), .Width, .Height / 3, .Width, .Height / 3
                    End If
                End With
            End If
        Next
        
        RenderText Font_Georgia, "OK", Main_Win(4).Left + 76 - (EngineGetTextWidth(Font_Georgia, "OK") / 2), Main_Win(4).Top + 92, White, 0, True
        RenderText Font_Georgia, "Cancelar", Main_Win(4).Left + 151 - (EngineGetTextWidth(Font_Georgia, "Cancelar") / 2), Main_Win(4).Top + 92, White, 0, True
            
    End If
    
    ' Botões Dx8
    For i = 1 To MAX_MAIN_BUTTONS
        Select Case i
        Case 10, 11, 17, 18, 42, 43, 44, 45, 46, 47, 48, 49
        ' Não Fazer Nada Rç
        Case Else
            If MAIN_BUTTON(i).Visible = True Then
            
                With Tex_Main_Buttons(MAIN_BUTTON(i).Picture)
                    
                    If MAIN_BUTTON(i).LockState = False Then
                        x = ConvertMapX(3 * 32) + Tex_Main_Buttons(MAIN_BUTTON(i).Picture).Width + 25
                        y = ConvertMapY(3 * 32)
                        RenderTexture Tex_Main_Buttons(MAIN_BUTTON(i).Picture), MAIN_BUTTON(i).Left, MAIN_BUTTON(i).Top, 0, MAIN_BUTTON(i).state * (.Height / 3), .Width, .Height / 3, .Width, .Height / 3
                    Else
                        x = ConvertMapX(3 * 32) + Tex_Main_Buttons(MAIN_BUTTON(i).Picture).Width + 25
                        y = ConvertMapY(3 * 32)
                        
                        Select Case MAIN_BUTTON(i).state
                            Case 0: LockState = 2
                            Case 1: LockState = 1
                            Case 2: LockState = 0
                        End Select
                        
                        RenderTexture Tex_Main_Buttons(MAIN_BUTTON(i).Picture), MAIN_BUTTON(i).Left, MAIN_BUTTON(i).Top, 0, LockState * (.Height / 3), .Width, .Height / 3, .Width, .Height / 3
                    End If
                    
                End With
            End If
        End Select
    Next
    
    ' Reviver
    If Main_Win(10).Visible = True Then
        RenderTexture Tex_Main_Win(11), Main_Win(10).Left, Main_Win(10).Top, 0, 0, Tex_Main_Win(11).Width, Tex_Main_Win(11).Height, Tex_Main_Win(11).Width, Tex_Main_Win(11).Height
    
    For i = 10 To 11
                If MAIN_BUTTON(i).Visible = True Then
            
                With Tex_Main_Buttons(MAIN_BUTTON(i).Picture)
                    
                    If MAIN_BUTTON(i).LockState = False Then
                        x = ConvertMapX(3 * 32) + Tex_Main_Buttons(MAIN_BUTTON(i).Picture).Width + 25
                        y = ConvertMapY(3 * 32)
                        RenderTexture Tex_Main_Buttons(MAIN_BUTTON(i).Picture), MAIN_BUTTON(i).Left, MAIN_BUTTON(i).Top, 0, MAIN_BUTTON(i).state * (.Height / 3), .Width, .Height / 3, .Width, .Height / 3
                    Else
                        x = ConvertMapX(3 * 32) + Tex_Main_Buttons(MAIN_BUTTON(i).Picture).Width + 25
                        y = ConvertMapY(3 * 32)
                        
                        Select Case MAIN_BUTTON(i).state
                            Case 0: LockState = 2
                            Case 1: LockState = 1
                            Case 2: LockState = 0
                        End Select
                        
                        RenderTexture Tex_Main_Buttons(MAIN_BUTTON(i).Picture), MAIN_BUTTON(i).Left, MAIN_BUTTON(i).Top, 0, LockState * (.Height / 3), .Width, .Height / 3, .Width, .Height / 3
                    End If
                    
                End With
            End If
    Next
    
    End If
    
    ' Janela Dx8 - Dialogo
    If Main_Win(17).Visible = True Then DrawDialogDx8
    
    ' Janela Dx8 - Banco
    If Main_Win(14).Visible = True Then DrawBankDx8
    
    ' Janela Dx8 - Trade
    If Main_Win(16).Visible = True Then DrawTradeDx8
    
    ' Janela Dx8 - Outfit
    If Main_Win(22).Visible = True Then DrawOutfitDx8
    
    ' Janela Dx8 - Pet Control
    If Main_Win(23).Visible = True Then DrawPetControlDx8
    
    ' Janela Dx8 - Drop Npc Window
    If Main_Win(24).Visible = True Then DrawDropNpc
    If Main_Win(25).Visible = True Then DrawDropList
    
    ' Refine Buttons Desc
    If Main_Win(21).Visible = True Then
        RenderText Font_Georgia, "Refinar", Main_Win(21).Left + 285 - (EngineGetTextWidth(Font_Georgia, "Refinar") / 2), Main_Win(21).Top + 159, White, 0, True
        RenderText Font_Georgia, "Sair", Main_Win(21).Left + 285 - (EngineGetTextWidth(Font_Georgia, "Sair") / 2), Main_Win(21).Top + 186, White, 0, True
    End If
    
        ' Texto Botões Loja
    If MAIN_BUTTON(28).Visible = True Then
        RenderText Font_Georgia, "Comprar", Main_Win(13).Left + 278 - (EngineGetTextWidth(Font_Georgia, "Comprar") / 2), Main_Win(13).Top + 190, White, 0, True
        RenderText Font_Georgia, "Vender", Main_Win(13).Left + 278 - (EngineGetTextWidth(Font_Georgia, "Vender") / 2), Main_Win(13).Top + 217, White, 0, True
        RenderText Font_Georgia, "Sair", Main_Win(13).Left + 278 - (EngineGetTextWidth(Font_Georgia, "Sair") / 2), Main_Win(13).Top + 244, White, 0, True
    End If
    
    ' Descrição de Item Inventario
    If DragInvItemDx8 > 0 Then DrawDragInvItemDx8
    
    If DescItemSystem.Item > 0 Then
        If ShiftDown = False Then
            DrawGlobalDescItemDx8
        Else
            DrawGlobalDescItemDx8P2
        End If
    End If

    ' Descrição de Spell
    If DescSpellPlayerDx8 > 0 Then DrawSpellDescDx8
    If DragPlayerSpellDx8 > 0 Then DrawDragPlayerSpellDx8
    
    ' Cursor
    RenderTexture Tex_Menu_Win(10), GlobalX - 6, GlobalY - 4, 0, 0, 32, 32, 32, 32
   
End Sub

Public Sub DrawPartyDx8()
Dim i As Long, SpriteP As Long
Dim HairP As Long, HairC1 As Long, HairC2 As Long, HairC3 As Long
Dim Member As Long, MemberName As String
Dim WitdhHP As Long, WitdhMP As Long
Dim SpriteParty As Long
Dim Ordem(1 To 3) As Long

    ' Evitar Problemas
    If Party.MemberCount <= 0 Then Exit Sub
    
    ' Ordem de Exibição
    Select Case MyPartySlot
    Case 1: Ordem(1) = 2: Ordem(2) = 3: Ordem(3) = 4
    Case 2: Ordem(1) = 1: Ordem(2) = 3: Ordem(3) = 4
    Case 3: Ordem(1) = 1: Ordem(2) = 2: Ordem(3) = 4
    Case 4: Ordem(1) = 1: Ordem(2) = 2: Ordem(3) = 3
    End Select
    
    ' Party Member (1)
    With Tex_Main_Win(22)
        RenderTexture Tex_Main_Win(22), Main_Win(15).Left, Main_Win(15).Top, 0, 0, .Width, .Height, .Width, .Height
    End With
    
    ' Variaveis
    Member = Ordem(1)
    
    With Party.MemberInfo(Member)
        HairP = .Hair
        HairC1 = .HairColor(1)
        HairC2 = .HairColor(2)
        HairC3 = .HairColor(3)
        MemberName = Trim$(.Name)
        SpriteParty = .sprite
    End With

    ' Nome
    Call RenderText(Font_Georgia, MemberName, Main_Win(15).Left + 50, Main_Win(15).Top + 10, White, 0, False)
    Call RenderText(Font_Georgia, Party.MemberInfo(Member).Level, Main_Win(15).Left + 72, Main_Win(15).Top + 39, White, 0, False)
        
    ' Cabeça e Cabelo
    With Tex_Character(SpriteParty)
        RenderTexture Tex_Character(SpriteParty), Main_Win(15).Left + 8, Main_Win(15).Top + 2, 32, 0, .Width / 3, .Height / 4, .Width / 3, .Height / 4
    End With
    
    With Tex_Hair(HairP)
        RenderTexture Tex_Hair(HairP), Main_Win(15).Left + 8, Main_Win(15).Top + 2, 32, 0, .Width / 3, .Height / 4, .Width / 3, .Height / 4, D3DColorRGBA(HairC1, HairC2, HairC3, 255)
    End With
    
    Member = Party.Member(Member)
    
    ' Barra de Vida/Mana
    If GetPlayerVital(Member, Vitals.HP) > 0 And GetPlayerMaxVital(Member, Vitals.HP) > 0 Then
        WitdhHP = ((GetPlayerVital(Member, Vitals.HP) / Tex_Main_Win(23).Width) / (GetPlayerMaxVital(Member, Vitals.HP) / Tex_Main_Win(23).Width)) * Tex_Main_Win(23).Width
        RenderTexture Tex_Main_Win(23), Main_Win(15).Left + 2, Main_Win(15).Top + 36, 0, 0, WitdhHP, 6, WitdhHP, 6
    End If
        
    If GetPlayerVital(Member, Vitals.MP) > 0 And GetPlayerMaxVital(Member, Vitals.MP) > 0 Then
        WitdhMP = ((GetPlayerVital(Member, Vitals.MP) / Tex_Main_Win(23).Width) / (GetPlayerMaxVital(Member, Vitals.MP) / Tex_Main_Win(23).Width)) * Tex_Main_Win(23).Width
        RenderTexture Tex_Main_Win(23), Main_Win(15).Left + 2, Main_Win(15).Top + 42, 0, 6, WitdhMP, 6, WitdhMP, 6
    End If
    
    ' Party Member (3)
    If Party.MemberCount < 3 Then Exit Sub
    
    ' Variaveis
    Member = Ordem(2)
    
    With Party.MemberInfo(Member)
        HairP = .Hair
        HairC1 = .HairColor(1)
        HairC2 = .HairColor(2)
        HairC3 = .HairColor(3)
        MemberName = Trim$(.Name)
        SpriteParty = .sprite
    End With
    
    With Tex_Main_Win(22)
        RenderTexture Tex_Main_Win(22), Main_Win(15).Left, Main_Win(15).Top + 64, 0, 0, .Width, .Height, .Width, .Height
    End With

    ' Nome
    Call RenderText(Font_Georgia, MemberName, Main_Win(15).Left + 50, Main_Win(15).Top + 74, White, 0, False)
    Call RenderText(Font_Georgia, Party.MemberInfo(Member).Level, Main_Win(15).Left + 72, Main_Win(15).Top + 103, White, 0, False)
        
    ' Cabeça e Cabelo
    With Tex_Character(SpriteParty)
        RenderTexture Tex_Character(SpriteParty), Main_Win(15).Left + 8, Main_Win(15).Top + 66, 32, 0, .Width / 3, .Height / 4, .Width / 3, .Height / 4
    End With
    
    With Tex_Hair(HairP)
        RenderTexture Tex_Hair(HairP), Main_Win(15).Left + 8, Main_Win(15).Top + 66, 32, 0, .Width / 3, .Height / 4, .Width / 3, .Height / 4, D3DColorRGBA(HairC1, HairC2, HairC3, 255)
    End With
    
    Member = Party.Member(Member)
    
    ' Barra de Vida/Mana
    If GetPlayerVital(Member, Vitals.HP) > 0 And GetPlayerMaxVital(Member, Vitals.HP) > 0 Then
        WitdhHP = ((GetPlayerVital(Member, Vitals.HP) / Tex_Main_Win(23).Width) / (GetPlayerMaxVital(Member, Vitals.HP) / Tex_Main_Win(23).Width)) * Tex_Main_Win(23).Width
        RenderTexture Tex_Main_Win(23), Main_Win(15).Left + 2, Main_Win(15).Top + 100, 0, 0, WitdhHP, 6, WitdhHP, 6
    End If
        
    If GetPlayerVital(Member, Vitals.MP) > 0 And GetPlayerMaxVital(Member, Vitals.MP) > 0 Then
        WitdhMP = ((GetPlayerVital(Member, Vitals.MP) / Tex_Main_Win(23).Width) / (GetPlayerMaxVital(Member, Vitals.MP) / Tex_Main_Win(23).Width)) * Tex_Main_Win(23).Width
        RenderTexture Tex_Main_Win(23), Main_Win(15).Left + 2, Main_Win(15).Top + 106, 0, 6, WitdhMP, 6, WitdhMP, 6
    End If
    
    ' Party Member (4)
    If Party.MemberCount < 4 Then Exit Sub
    
    ' Variaveis
    Member = Ordem(3)
    
    With Party.MemberInfo(Member)
        HairP = .Hair
        HairC1 = .HairColor(1)
        HairC2 = .HairColor(2)
        HairC3 = .HairColor(3)
        MemberName = Trim$(.Name)
        SpriteParty = .sprite
    End With
    
    With Tex_Main_Win(22)
        RenderTexture Tex_Main_Win(22), Main_Win(15).Left, Main_Win(15).Top + 64, 0, 0, .Width, .Height, .Width, .Height
    End With

    ' Nome
    Call RenderText(Font_Georgia, MemberName, Main_Win(15).Left + 50, Main_Win(15).Top + 138, White, 0, False)
    Call RenderText(Font_Georgia, Party.MemberInfo(Member).Level, Main_Win(15).Left + 72, Main_Win(15).Top + 167, White, 0, False)
        
    ' Cabeça e Cabelo
    With Tex_Character(SpriteParty)
        RenderTexture Tex_Character(SpriteParty), Main_Win(15).Left + 8, Main_Win(15).Top + 130, 32, 0, .Width / 4, .Height / 2, .Width / 4, .Height / 2
    End With
    
    With Tex_Hair(HairP)
        RenderTexture Tex_Hair(HairP), Main_Win(15).Left + 8, Main_Win(15).Top + 130, 32, 0, .Width / 4, .Height / 2, .Width / 4, .Height / 2, D3DColorRGBA(HairC1, HairC2, HairC3, 255)
    End With
    
    Member = Party.Member(Member)
    
    ' Barra de Vida/Mana
    If GetPlayerVital(Member, Vitals.HP) > 0 And GetPlayerMaxVital(Member, Vitals.HP) > 0 Then
        WitdhHP = ((GetPlayerVital(Member, Vitals.HP) / Tex_Main_Win(23).Width) / (GetPlayerMaxVital(Member, Vitals.HP) / Tex_Main_Win(23).Width)) * Tex_Main_Win(23).Width
        RenderTexture Tex_Main_Win(23), Main_Win(15).Left + 2, Main_Win(15).Top + 164, 0, 0, WitdhHP, 6, WitdhHP, 6
    End If
        
    If GetPlayerVital(Member, Vitals.MP) > 0 And GetPlayerMaxVital(Member, Vitals.MP) > 0 Then
        WitdhMP = ((GetPlayerVital(Member, Vitals.MP) / Tex_Main_Win(23).Width) / (GetPlayerMaxVital(Member, Vitals.MP) / Tex_Main_Win(23).Width)) * Tex_Main_Win(23).Width
        RenderTexture Tex_Main_Win(23), Main_Win(15).Left + 2, Main_Win(15).Top + 170, 0, 6, WitdhMP, 6, WitdhMP, 6
    End If
End Sub

Public Sub DrawInventoryDx8()
    Dim i As Long, ItemNum As Long, PicNum As Long
    Dim sRECT As RECT, Rec As RECT
    Dim Text As String

    If Not InGame Then Exit Sub

    ' Background Inventory
    With Tex_Main_Win(15)
        RenderTexture Tex_Main_Win(15), Main_Win(1).Left, Main_Win(1).Top, 0, 0, .Width, .Height, .Width, .Height
    End With

    ' Peso do jogador
    Text = Player(MyIndex).Peso(1) & "/" & Player(MyIndex).Peso(2)
    Dim PesoColor As Long
    If Player(MyIndex).Peso(1) >= (Player(MyIndex).Peso(2) * 0.85) Then
        PesoColor = BrightRed
    Else
        PesoColor = White
    End If
    RenderText Font_Georgia, "Peso: " & Text, _
        Main_Win(1).Left + 219 - EngineGetTextWidth(Font_Georgia, "Peso: " & Text) / 2, _
        Main_Win(1).Top + 267, PesoColor, 0, False

    ' Ouro
    If HasItem(1) > 0 Then
        Text = HasItem(1)
        RenderText Font_Georgia, "Ouro: " & Text, _
            Main_Win(1).Left + 400 - EngineGetTextWidth(Font_Georgia, "Ouro: " & Text) / 2, _
            Main_Win(1).Top + 267, White, 0, False
    End If

    ' Itens do inventário
    For i = 1 To MAX_INV
        ItemNum = PlayerInv(i).num
        If ItemNum > 0 Then
            PicNum = Item(ItemNum).pic

            ' Calcula posição do slot
            With sRECT
                .Top = Main_Win(1).Top + 38 + 36 * ((i - 1) \ 10)
                .Bottom = .Top + PIC_Y
                .Left = Main_Win(1).Left + 132 + 36 * ((i - 1) Mod 10)
                .Right = .Left + PIC_X
            End With

            ' Seleciona frame correto do item
            If Tex_Item(PicNum).Width > 32 Then
                With Rec
                    .Top = 0
                    .Bottom = 32
                    .Left = InvItemFrame(i) * 32
                    .Right = .Left + 32
                End With
            Else
                With Rec
                    .Top = 0
                    .Bottom = PIC_Y
                    .Left = 0
                    .Right = PIC_X
                End With
            End If

            ' Renderiza o item
            RenderTexture Tex_Item(PicNum), sRECT.Left, sRECT.Top, _
                Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, _
                Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, D3DColorRGBA(255, 255, 255, 255)

            ' Quantidade
            If PlayerInv(i).value > 1 Then
                RenderText Font_Georgia, Format$(ConvertCurrency(CStr(PlayerInv(i).value)), "#,###,###,###"), _
                    sRECT.Left, sRECT.Top + 18, White, 0, False
            End If
        End If
    Next

    ' Slots extras vazios
    For i = 1 + Player(MyIndex).SlotExtra To 30
        With sRECT
            .Top = Main_Win(1).Top + 146 + 36 * ((i - 1) \ 10)
            .Bottom = .Top + PIC_Y
            .Left = Main_Win(1).Left + 132 + 36 * ((i - 1) Mod 10)
            .Right = .Left + PIC_X
        End With
        RenderTexture Tex_Menu_Win(12), sRECT.Left, sRECT.Top, 0, 0, 32, 32, 32, 32
    Next
End Sub


Public Sub DrawAnimatedMapItems()
Dim i As Long
Dim ItemNum As Long, itempic As Long
Dim x As Long, y As Long
Dim MaxFrames As Byte
Dim Amount As Long
Dim Rec As RECT, Rec_Pos As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    'Evitar OverFlow
    If Not InGame Then Exit Sub
    
    ' check for map animation changes#
    For i = 1 To MAX_MAP_ITEMS
        If MapItem(i).num > 0 Then
            itempic = Item(MapItem(i).num).pic

            If itempic < 1 Or itempic > numitems Then Exit Sub 'Evitar OverFlow
            MaxFrames = (Tex_Item(itempic).Width) / 32  ' Work out how many frames there are. /2 because of inventory icons as well as ingame

            If MapItem(i).Frame < MaxFrames - 1 Then
                MapItem(i).Frame = MapItem(i).Frame + 1
            Else
                MapItem(i).Frame = 0
            End If
        End If
    Next
    
    If Main_Win(1).Visible = True Then
    
        For i = 1 To Equipment.Equipment_Count - 1
            ItemNum = GetPlayerEquipment(MyIndex, i)
            If ItemNum > 0 Then itempic = Item(ItemNum).pic
            
            If Tex_Item(itempic).Width > 32 Then
                MaxFrames = (Tex_Item(itempic).Width) / 32
                    
                If EquipFrame(i) < MaxFrames - 1 Then
                        EquipFrame(i) = EquipFrame(i) + 1
                    Else
                        EquipFrame(i) = 0
                End If
            End If
        Next
        
        For i = 1 To EquipmentExtra.Equipment_Count - 1
            ItemNum = GetPlayerEquipmentExtra(MyIndex, i)
            If ItemNum > 0 Then itempic = Item(ItemNum).pic
            
            If Tex_Item(itempic).Width > 32 Then
                MaxFrames = (Tex_Item(itempic).Width) / 32
                    
                If EquipFrame(i + 4) < MaxFrames - 1 Then
                        EquipFrame(i + 4) = EquipFrame(i + 4) + 1
                    Else
                        EquipFrame(i + 4) = 0
                End If
            End If
        Next
    
        For i = 1 To MAX_INV
                ItemNum = GetPlayerInvItemNum(MyIndex, i)
                If ItemNum > 0 Then itempic = Item(ItemNum).pic
                
                If Tex_Item(itempic).Width > 32 Then
                    MaxFrames = (Tex_Item(itempic).Width) / 32  ' Work out how many frames there are. /2 because of inventory icons as well as ingame

                    If InvItemFrame(i) < MaxFrames - 1 Then
                        InvItemFrame(i) = InvItemFrame(i) + 1
                    Else
                        InvItemFrame(i) = 0
                    End If
                    
                End If
        Next
    End If
    
    ' Shop Animated Item
    If Main_Win(13).Visible = True Then
        If InShop > 0 Then
            For i = 1 To MAX_TRADES
                ItemNum = Shop(InShop).TradeItem(i).Item
                If ItemNum > 0 Then itempic = Item(Shop(InShop).TradeItem(i).Item).pic
                
                If Tex_Item(itempic).Width > 32 Then
                    MaxFrames = (Tex_Item(itempic).Width) / 32  ' Work out how many frames there are. /2 because of inventory icons as well as ingame

                    If ShopItemFrame(i) < MaxFrames - 1 Then
                        ShopItemFrame(i) = ShopItemFrame(i) + 1
                    Else
                        ShopItemFrame(i) = 0
                    End If
                    
                End If
            Next
        End If
    End If
                    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawAnimatedInvItems", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawOutfitDx8()
Dim sRECT As RECT, BarWidth As Long
Dim DirHead As Byte, Anim As Byte
Dim SpriteNum As Long, i As Long
Dim x As Long, y As Long, LockState As Long
Dim HairNum As Long, CabeloView As Boolean
Dim XSprite As Long, YSprite As Long

    ' Character
    Anim = Player(MyIndex).Resp
    With sRECT
        .Top = 0
        .Bottom = .Top + (Tex_Character(GetPlayerSprite(MyIndex)).Height / 2)
        .Left = Anim * (Tex_Character(GetPlayerSprite(MyIndex)).Width / 4)
        .Right = .Left + (Tex_Character(GetPlayerSprite(MyIndex)).Width / 4)
    End With
        
    XSprite = 110
    YSprite = 25
        
    ' Evitar OverFlow
    If SelOutfit = 0 Then Exit Sub
    
    If Player(MyIndex).Sex = 0 Then
        SpriteNum = Outfit(POutList(SelOutfit)).MSprite
        HairNum = Outfit(POutList(SelOutfit)).MHair 'PHairList(SelOutHair)
    Else
        SpriteNum = Outfit(POutList(SelOutfit)).FSprite
        HairNum = Outfit(POutList(SelOutfit)).FHair 'PHairList(SelOutHair)
    End If
    
    ' Caso Seja conta com Cor Negra
    If Player(MyIndex).Cor = 1 Then SpriteNum = SpriteNum + 1
    CabeloView = True

    ' Background Outfit
    With Tex_Main_Win(33)
        RenderTexture Tex_Main_Win(33), Main_Win(22).Left, Main_Win(22).Top, 0, 0, .Width, .Height, .Width, .Height
    End With
    
    ' Evitar OverFlow
    If SpriteNum < 0 Or SpriteNum > NumCharacters Then Exit Sub
    If HairNum < 0 Or HairNum > HairNum Then Exit Sub
    
    Call RenderText(Font_Georgia, Outfit(POutList(SelOutfit)).Name, Main_Win(22).Left + 132 - (getWidth(Font_Georgia, Outfit(POutList(SelOutfit)).Name) / 2), Main_Win(22).Top + 89, White, 0, False)
    
    ' Masculino
    If Player(MyIndex).Sex = 0 Then
    
        ' Sprite
        RenderTexture Tex_Character(SpriteNum), Main_Win(22).Left + XSprite, Main_Win(22).Top + YSprite, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top
        
        ' Cabelo
        If SelAddon(1) = True And Outfit(POutList(SelOutfit)).HMAddOff1 = True Then CabeloView = False
        If SelAddon(2) = True And Outfit(POutList(SelOutfit)).HMAddOff2 = True Then CabeloView = False
        If Outfit(POutList(SelOutfit)).MAddon1 = 13 And SelAddon(1) = True Then HairNum = 24
        
        If CabeloView = True Then
            RenderTexture Tex_Hair(HairNum), Main_Win(22).Left + XSprite, Main_Win(22).Top + YSprite, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(SelOutRGB(1), SelOutRGB(2), SelOutRGB(3), 255)
        End If
        
        ' Addon 1/2
        If SelAddon(1) = True And Outfit(POutList(SelOutfit)).MAddon1 > 0 Then RenderTexture Tex_Addon(Outfit(POutList(SelOutfit)).MAddon1), Main_Win(22).Left + XSprite, Main_Win(22).Top + YSprite, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top
        
        Select Case POutList(SelOutfit)
        Case 14, 15, 11, 9
            Select Case Outfit(POutList(SelOutfit)).MAddon2
            Case 32, 34, 35, 23
                XSprite = XSprite - 6
                YSprite = YSprite + 20
            End Select
        End Select
        
        With sRECT
            .Top = 0
            .Bottom = .Top + (Tex_Addon(Outfit(POutList(SelOutfit)).MAddon2).Height / 4)
            .Left = Anim * (Tex_Addon(Outfit(POutList(SelOutfit)).MAddon2).Width / 6)
            .Right = .Left + (Tex_Addon(Outfit(POutList(SelOutfit)).MAddon2).Width / 6)
        End With
        
        If SelAddon(2) = True And Outfit(POutList(SelOutfit)).MAddon2 > 0 Then
            RenderTexture Tex_Addon(Outfit(POutList(SelOutfit)).MAddon2), Main_Win(22).Left + XSprite, Main_Win(22).Top + YSprite, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top
        End If

    Else ' Ferminino
    
        ' Sprite
        RenderTexture Tex_Character(SpriteNum), Main_Win(22).Left + XSprite, Main_Win(22).Top + YSprite, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top
        
        ' Cabelo
        If SelAddon(1) = True And Outfit(POutList(SelOutfit)).HFAddOff1 = True Then CabeloView = False
        If SelAddon(2) = True And Outfit(POutList(SelOutfit)).HFAddOff2 = True Then CabeloView = False
        
        Select Case Player(MyIndex).Outfit
        Case 14, 15, 11
            Select Case SpriteNum
            Case 32, 34, 35
                XSprite = XSprite - 6
                YSprite = YSprite + 20
            End Select
        End Select
        
        If CabeloView = True Then
            RenderTexture Tex_Hair(HairNum), Main_Win(22).Left + XSprite, Main_Win(22).Top + YSprite, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(SelOutRGB(1), SelOutRGB(2), SelOutRGB(3), 255)
        End If
        
        ' Addon 1/2
        If SelAddon(1) = True And Outfit(POutList(SelOutfit)).FAddon1 > 0 Then RenderTexture Tex_Addon(Outfit(POutList(SelOutfit)).FAddon1), Main_Win(22).Left + XSprite, Main_Win(22).Top + YSprite, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top
        If SelAddon(2) = True And Outfit(POutList(SelOutfit)).FAddon2 > 0 Then RenderTexture Tex_Addon(Outfit(POutList(SelOutfit)).FAddon2), Main_Win(22).Left + XSprite, Main_Win(22).Top + YSprite, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top

    End If
    
    ' Button Locked
    For i = 42 To 49
            With Tex_Main_Buttons(MAIN_BUTTON(i).Picture)
                
                If MAIN_BUTTON(i).LockState = False Then
                    x = ConvertMapX(3 * 32) + Tex_Main_Buttons(MAIN_BUTTON(i).Picture).Width + 25
                    y = ConvertMapY(3 * 32)
                    RenderTexture Tex_Main_Buttons(MAIN_BUTTON(i).Picture), MAIN_BUTTON(i).Left, MAIN_BUTTON(i).Top, 0, MAIN_BUTTON(i).state * (.Height / 3), .Width, .Height / 3, .Width, .Height / 3
                Else
                    x = ConvertMapX(3 * 32) + Tex_Main_Buttons(MAIN_BUTTON(i).Picture).Width + 25
                    y = ConvertMapY(3 * 32)
                    
                    Select Case MAIN_BUTTON(i).state
                        Case 0: LockState = 2
                        Case 1: LockState = 1
                        Case 2: LockState = 0
                    End Select
                    
                    RenderTexture Tex_Main_Buttons(MAIN_BUTTON(i).Picture), MAIN_BUTTON(i).Left, MAIN_BUTTON(i).Top, 0, LockState * (.Height / 3), .Width, .Height / 3, .Width, .Height / 3
                End If
                
            End With
    Next
    
    RenderText Font_Georgia, "Confirmar", Main_Win(22).Left + 63, Main_Win(22).Top + 190, White, 0, False
    RenderText Font_Georgia, "Cancelar", Main_Win(22).Left + 143, Main_Win(22).Top + 190, White, 0, False
    
End Sub

Public Sub DrawTradeDx8()
Dim MyName As String, i As Long
Dim Rec As RECT, ItemNum As Long, itempic As Long
Dim ItemValue As Long

    MyName = Trim$(GetPlayerName(MyIndex))

    ' Background Shop
    With Tex_Main_Win(24)
        RenderTexture Tex_Main_Win(24), Main_Win(16).Left, Main_Win(16).Top, 0, 0, .Width, .Height, .Width, .Height
    End With
    
    ' Evitar Desastres
    If InTrade = False Then Exit Sub
    
    For i = 1 To 10
    ' Draw your own offer
        ItemNum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)
        ItemValue = TradeYourOffer(i).value
        
            If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
                itempic = Item(ItemNum).pic
                
            With Rec
                .Top = Main_Win(16).Top + 39 + ((36) * ((i - 1) \ 5))
                .Bottom = .Top + PIC_Y
                .Left = Main_Win(16).Left + 28 + ((36) * (((i - 1) Mod 5)))
                .Right = .Left + PIC_X
            End With
            
            RenderTexture Tex_Item(itempic), Rec.Left, Rec.Top, 0, 0, 32, 32, 32, 32
            Call RenderText(Font_Georgia, Format$(ConvertCurrency(Str(ItemValue)), "#,###,###,###"), Rec.Left, Rec.Top + 18, White, 0, False)
        End If
    Next
    
    ' My Name
    Call RenderText(Font_Georgia, MyName, Main_Win(16).Left + 95 - (getWidth(Font_Georgia, MyName) / 2), Main_Win(16).Top + 5, White, 0, False)
    
    ' Jogador Alvo
    Call RenderText(Font_Georgia, Trim$(TradeName), Main_Win(16).Left + 95 - (getWidth(Font_Georgia, Trim$(TradeName)) / 2), Main_Win(16).Top + 149, White, 0, False)
 
    ' OFERTA DO OUTRO JOGADOR
    For i = 1 To 10
    
    ItemNum = TradeTheirOffer(i).num
    ItemValue = TradeTheirOffer(i).value
    
        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
        
            itempic = Item(ItemNum).pic
        
            With Rec
                .Top = Main_Win(16).Top + 183 + ((36) * ((i - 1) \ 5))
                .Bottom = .Top + PIC_Y
                .Left = Main_Win(16).Left + 28 + ((36) * (((i - 1) Mod 5)))
                .Right = .Left + PIC_X
            End With
            
            RenderTexture Tex_Item(itempic), Rec.Left, Rec.Top, 0, 0, 32, 32, 32, 32
            Call RenderText(Font_Georgia, Format$(ConvertCurrency(Str(ItemValue)), "#,###,###,###"), Rec.Left, Rec.Top + 18, White, 0, False)
        End If
    Next
       
End Sub

Public Sub DrawBankDx8()
Dim i As Long, x As Long, y As Long, ItemNum As Long, itempic As Long, srcRect As D3DRECT, destRect As D3DRECT
Dim Amount As String, ItemValue As Long
Dim Rec As RECT, Rec_Pos As RECT
Dim Colour As Long, PageC(1 To 2) As Byte

    If Not InGame Then Exit Sub
    
    ' Background Shop
    With Tex_Main_Win(21)
        RenderTexture Tex_Main_Win(21), Main_Win(14).Left, Main_Win(14).Top, 0, 0, .Width, .Height, .Width, .Height
    End With
    
    Select Case PageBank
    Case 0: PageC(1) = 1: PageC(2) = 30
    Case 1: PageC(1) = 31: PageC(2) = 60
    Case 2: PageC(1) = 61: PageC(2) = 90
    End Select
    
    For i = PageC(1) To PageC(2)
        ItemNum = Bank.Item(i).num
        ItemValue = Bank.Item(i).value
        
        If ItemNum > 0 Then
            With Rec
                .Top = Main_Win(13).Top + 39 + ((36) * ((i - PageC(1)) \ 5))
                .Bottom = .Top + PIC_Y
                .Left = Main_Win(13).Left + 28 + ((36) * (((i - PageC(1)) Mod 5)))
                .Right = .Left + PIC_X
            End With
            
            RenderTexture Tex_Item(Item(ItemNum).pic), Rec.Left, Rec.Top, 0, 0, 32, 32, 32, 32
            Call RenderText(Font_Georgia, Format$(ConvertCurrency(Str(ItemValue)), "#,###,###,###"), Rec.Left, Rec.Top + 18, White, 0, False)
        End If
    Next
    
End Sub

Public Sub DrawShopDx8()
Dim i As Long, x As Long, y As Long, ItemNum As Long, itempic As Long, srcRect As D3DRECT, destRect As D3DRECT
Dim Amount As String, ItemValue As Long
Dim Rec As RECT, Rec_Pos As RECT, Rec2 As RECT
Dim Colour As Long
Dim RefLvl As Long, HeadHair As Long

    If Not InGame Then Exit Sub
    
    ' Background Shop
    With Tex_Main_Win(20)
        RenderTexture Tex_Main_Win(20), Main_Win(13).Left, Main_Win(13).Top, 0, 0, .Width, .Height, .Width, .Height
    End With
    
    For i = 1 To MAX_TRADES
        ItemNum = Shop(InShop).TradeItem(i).Item
        ItemValue = Shop(InShop).TradeItem(i).ItemValue
        
        If ItemNum > 0 Then
            With Rec
                .Top = Main_Win(13).Top + 39 + ((36) * ((i - 1) \ 5))
                .Bottom = .Top + PIC_Y
                .Left = Main_Win(13).Left + 28 + ((36) * (((i - 1) Mod 5)))
                .Right = .Left + PIC_X
            End With
            
            If Tex_Item(Item(ItemNum).pic).Width > 32 Then ' has more than 1 frame
                With Rec2
                    .Top = 0
                    .Bottom = 32
                    .Left = (ShopItemFrame(i) * 32)
                    .Right = .Left + 32
                End With
            Else
                With Rec2
                    .Top = 0
                    .Bottom = PIC_Y
                    .Left = 0
                    .Right = PIC_X
                End With
            End If
            
            RenderTexture Tex_Item(Item(ItemNum).pic), Rec.Left, Rec.Top, Rec2.Left, Rec2.Top, Rec2.Right - Rec2.Left, Rec2.Bottom - Rec2.Top, Rec2.Right - Rec2.Left, Rec2.Bottom - Rec2.Top, D3DColorRGBA(255, 255, 255, 255)
            'RenderTexture Tex_Item(Item(ItemNum).pic), Rec.Left, Rec.Top, 0, 0, 32, 32, 32, 32
        End If
        
        ' Valores
        If Shop(InShop).TradeItem(i).ItemValue > 1 Then Call RenderText(Font_Georgia, Format$(ConvertCurrency(Str(ItemValue)), "#,###,###,###"), Rec.Left, Rec.Top + 18, White, 0, False)
        
    Next
    
End Sub

Public Sub DrawSoundConfigDx8()
Dim sRECT As RECT
Dim i As Long

    ' Evitar Merda
    If Not InGame Then Exit Sub
    
    ' Background Quests
    With Tex_Main_Win(28)
        RenderTexture Tex_Main_Win(28), Main_Win(20).Left, Main_Win(20).Top, 0, 0, .Width, .Height, .Width, .Height
    End With
    
    With Tex_Main_Buttons(28)
    If Options.Music = 1 Then RenderTexture Tex_Main_Buttons(28), Main_Win(20).Left + 30, Main_Win(20).Top + 30, 0, 0, .Width, .Height, .Width, .Height
    If Options.Sound = 1 Then RenderTexture Tex_Main_Buttons(28), Main_Win(20).Left + 30, Main_Win(20).Top + 49, 0, 0, .Width, .Height, .Width, .Height
    End With
    
    ' Music Level Text
    Call RenderText(Font_Georgia, Options.MusicLevel, Main_Win(20).Left + 137 - (getWidth(Font_Georgia, Options.MusicLevel) / 2), Main_Win(20).Top + 30, White, 0)
    
End Sub

Public Sub DrawDialogQuestDx8()
Dim sRECT As RECT
Dim i As Long

    ' Evitar Merda
    If Not InGame Then Exit Sub
    
    ' Background Quests
    With Tex_Main_Win(27)
        RenderTexture Tex_Main_Win(27), Main_Win(19).Left, Main_Win(19).Top, 0, 0, .Width, .Height, .Width, .Height
    End With
    
    Call RenderText(Font_Georgia, Trim$(QuestDialogName), Main_Win(19).Left + 148 - (getWidth(Font_Georgia, Trim$(QuestDialogName)) / 2), Main_Win(19).Top + 11, White, 0)
    Call RenderText(Font_Georgia, WordWrap(Font_Georgia, Trim$(QuestDialogDesc), 270), Main_Win(19).Left + 18, Main_Win(19).Top + 33, White, 0)
    Call RenderText(Font_Georgia, Trim$(QuestChoice(1)), Main_Win(19).Left + 62 - (getWidth(Font_Georgia, Trim$(QuestChoice(1))) / 2), Main_Win(19).Top + 104, White, 0)
    Call RenderText(Font_Georgia, Trim$("Fechar"), Main_Win(19).Left + 250 - (getWidth(Font_Georgia, Trim$("Fechar")) / 2), Main_Win(19).Top + 104, White, 0)
    
End Sub

Public Sub DrawQuestsDx8()
Dim sRECT As RECT
Dim i As Long
Dim Page(1 To 2) As Byte
Dim Obj(1 To 2) As Long
Dim AtualTask As Byte
Dim NpcName As String

    ' Evitar Merda
    If Not InGame Then Exit Sub
    
    ' Background Quests
    With Tex_Main_Win(25)
        RenderTexture Tex_Main_Win(25), Main_Win(18).Left, Main_Win(18).Top, 0, 0, .Width, .Height, .Width, .Height
    End With
    
    i = QuestLogPage
    
        ' Pagina Inicial
        If i = 0 Then
            Page(1) = 1
            Page(2) = 12
        End If
        
        ' Meio
        If i > 0 And i < 19 Then
            Page(1) = 13 * i
            Page(2) = 12 + (13 * i)
        End If
        
        ' Pagina Final
        If i = 19 Then
            Page(1) = 247
            Page(2) = 255
        End If
    
    Call RenderText(Font_Georgia, "Página: " & QuestLogPage + 1 & "/20", Main_Win(18).Left + 109 - (getWidth(Font_Georgia, "Page: " & QuestLogPage + 1 & "/20") / 2), Main_Win(18).Top + 278, White, 0)
    
    ' Nomes Quests
    For i = Page(1) To Page(2)
        If QuestIdLog(i) > 0 Then
            If i = QLogSelect Then
                Call RenderText(Font_Georgia, Trim$(Quest(QuestIdLog(i)).Name), Main_Win(18).Left + 28, Main_Win(18).Top + 35 + ((i - 1) * 20), Yellow, 0)
            Else
                Call RenderText(Font_Georgia, Trim$(Quest(QuestIdLog(i)).Name), Main_Win(18).Left + 28, Main_Win(18).Top + 35 + ((i - 1) * 20), White, 0)
            End If
        End If
    Next
    
    ' Quest Selecionada Informações
    If QLogSelect > 0 Then
        If QuestIdLog(QLogSelect) > 0 Then
        
        AtualTask = Player(MyIndex).PlayerQuest(QuestIdLog(QLogSelect)).ActualTask
        Obj(1) = Player(MyIndex).PlayerQuest(QuestIdLog(QLogSelect)).CurrentCount
        Obj(2) = Quest(QuestIdLog(QLogSelect)).Task(AtualTask).Amount
        
            Call RenderText(Font_Georgia, WordWrap(Font_Georgia, Trim$(Quest(QuestIdLog(QLogSelect)).Chat(1)), 172), Main_Win(18).Left + 215, Main_Win(18).Top + 75, White, 0)
            Call RenderText(Font_Georgia, WordWrap(Font_Georgia, Trim$(Quest(QuestIdLog(QLogSelect)).Task(AtualTask).TaskLog), 148), Main_Win(18).Left + 393, Main_Win(18).Top + 75, White, 0)
            Call RenderText(Font_Georgia, WordWrap(Font_Georgia, Trim$(Quest(QuestIdLog(QLogSelect)).RewardExp), 148), Main_Win(18).Left + 415, Main_Win(18).Top + 238, White, 0)
            Call RenderText(Font_Georgia, WordWrap(Font_Georgia, Quest(QuestIdLog(QLogSelect)).RewardItemAmount(6), 148), Main_Win(18).Left + 422, Main_Win(18).Top + 253, White, 0)
        
        If Quest(QuestIdLog(QLogSelect)).Task(AtualTask).NPC > 0 Then
            NpcName = Trim$(NPC(Quest(QuestIdLog(QLogSelect)).Task(AtualTask).NPC).Name)
            Call RenderText(Font_Georgia, WordWrap(Font_Georgia, NpcName & " [" & Obj(1) & "/" & Obj(2) & "]", 148), Main_Win(18).Left + 393, Main_Win(18).Top + 190, White, 0)
        End If
        
        If Quest(QuestIdLog(QLogSelect)).Task(AtualTask).Resource > 0 Then
            NpcName = Trim$(Resource(Quest(QuestIdLog(QLogSelect)).Task(AtualTask).Resource).Name)
            Call RenderText(Font_Georgia, WordWrap(Font_Georgia, NpcName & " [" & Obj(1) & "/" & Obj(2) & "]", 148), Main_Win(18).Left + 393, Main_Win(18).Top + 190, White, 0)
        End If
        
        For i = 1 To 5
            If Quest(QuestIdLog(QLogSelect)).RewardItem(i) > 0 Then
                RenderTexture Tex_Item(Item(Quest(QuestIdLog(QLogSelect)).RewardItem(i)).pic), Main_Win(18).Left + 213 + ((i - 1) * 34), Main_Win(18).Top + 239, 0, 0, 32, 32, 32, 32
            End If
        Next
        
        End If
    End If
        
End Sub

Public Sub DrawRefineDx8()
Dim sRECT As RECT
Dim i As Long, ItemNum As Long, ItemLevel As Long
Dim ItemRarity As Long, PriceStr As String
Dim Stone As Long, Scroll As Long

    ' Background Dialog
    With Tex_Main_Win(30)
        RenderTexture Tex_Main_Win(30), Main_Win(21).Left, Main_Win(21).Top, 0, 0, .Width, .Height, .Width, .Height
    End With
    
    ItemNum = Player(MyIndex).Refine.RefineItem(1)
    ItemLevel = Player(MyIndex).Refine.Refinelevel
    Stone = Player(MyIndex).Refine.RefineItem(2)
    Scroll = Player(MyIndex).Refine.RefineItem(3)
    
    ' Pedra
    If Stone > 0 Then
        RenderTexture Tex_Item(Item(Stone).pic), Main_Win(21).Left + 30, Main_Win(21).Top + 36, 0, 0, 32, 32, 32, 32
    End If
    
    ' Pergaminho
    If Scroll > 0 Then
        RenderTexture Tex_Item(Item(Scroll).pic), Main_Win(21).Left + 150, Main_Win(21).Top + 36, 0, 0, 32, 32, 32, 32
    End If
    
    ' Item a Ser Refinado
    If ItemNum = 0 Then
        Call RenderText(Font_Georgia, WordWrap(Font_Georgia, "Adicione o Item.", 250), Main_Win(21).Left + 55, Main_Win(21).Top + 76, White, 0)
        Call RenderText(Font_Georgia, "Ouro: 0", Main_Win(21).Left + 105 - (EngineGetTextWidth(Font_Georgia, "Ouro: 0") / 2), Main_Win(21).Top + 157, White, 0)
    Else
    
        Select Case Item(ItemNum).Rarity
            Case 1: ItemRarity = BrightGreen: PriceStr = "Ouro: 500"
            Case 2: ItemRarity = BrightBlue: PriceStr = "Ouro: 1.000"
            Case 3: ItemRarity = Yellow: PriceStr = "Ouro: 2.000"
            Case 4: ItemRarity = Pink: PriceStr = "Ouro: 3.000"
            Case 5: ItemRarity = Orange: PriceStr = "Ouro: 3.500"
            Case Else: ItemRarity = White: PriceStr = "Ouro: 100"
        End Select
    
        RenderTexture Tex_Item(Item(ItemNum).pic), Main_Win(21).Left + 90, Main_Win(21).Top + 36, 0, 0, 32, 32, 32, 32
        Call RenderText(Font_Georgia, Trim$(Item(ItemNum).Name) & " +" & ItemLevel, Main_Win(21).Left + 105 - (EngineGetTextWidth(Font_Georgia, Trim$(Item(ItemNum).Name) & " +" & ItemLevel) / 2), Main_Win(21).Top + 76, ItemRarity, 0)
        Call RenderText(Font_Georgia, PriceStr, Main_Win(21).Left + 105 - (EngineGetTextWidth(Font_Georgia, PriceStr) / 2), Main_Win(21).Top + 157, White, 0)
        
        ' Chance de Refino
        
        If Player(MyIndex).Refine.Refinelevel <= 11 Then
            Call RenderText(Font_Georgia, "Próx Level: +" & Player(MyIndex).Refine.Refinelevel + 1, Main_Win(21).Left + 105 - (EngineGetTextWidth(Font_Georgia, "Próx Level: +" & Player(MyIndex).Refine.Refinelevel + 1) / 2), Main_Win(21).Top + 96, White, 0)
            Call RenderText(Font_Georgia, "Chance: " & Player(MyIndex).Refine.Chance & "%", Main_Win(21).Left + 105 - (EngineGetTextWidth(Font_Georgia, "Chance: " & Player(MyIndex).Refine.Chance & "%") / 2), Main_Win(21).Top + 116, White, 0)
        Else
            Call RenderText(Font_Georgia, "Level Máximo", Main_Win(21).Left + 105 - (EngineGetTextWidth(Font_Georgia, "Level Máximo") / 2), Main_Win(21).Top + 96, White, 0)
        End If
    End If
    
End Sub

Public Sub DrawDialogDx8()
Dim sRECT As RECT
Dim i As Long
Dim NpcSprite As Long, Anim As Byte, NpcName As String
Dim LocX As Long, LocY As Long

    ' Evitar Merda
    If Not InGame Then Exit Sub
    If Player(MyIndex).InDialogue = False Then Exit Sub
    If DialogueNpcIndex = 0 Then Exit Sub
    
    ' Background Dialog
    With Tex_Main_Win(26)
        RenderTexture Tex_Main_Win(26), Main_Win(17).Left, Main_Win(17).Top, 0, 0, .Width, .Height, .Width, .Height
    End With
    
    ' NPC
    NpcName = Trim$(NPC(Player(MyIndex).InDialogue).Name)
    NpcSprite = NPC(Player(MyIndex).InDialogue).sprite
    Anim = Player(MyIndex).Resp
    
    'Evitar OverFlow
    If NpcSprite = 0 Or NpcSprite > NumCharacters Then Exit Sub
    
    With sRECT
        .Top = 0
        .Bottom = .Top + (Tex_Character(NpcSprite).Height / 2)
        .Left = Anim * (Tex_Character(NpcSprite).Width / 4)
        .Right = .Left + (Tex_Character(NpcSprite).Width / 4)
    End With
    
    LocX = 56 - ((Tex_Character(NpcSprite).Height / 2) / 2)
    LocY = 54 - ((Tex_Character(NpcSprite).Width / 4) / 2)
        
    RenderTexture Tex_Character(NpcSprite), Main_Win(17).Left + LocX, Main_Win(17).Top + LocY, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top
        
    
    ' Textos
    Call RenderText(Font_Georgia, NpcName, Main_Win(17).Left + 245 - (EngineGetTextWidth(Font_Georgia, NpcName) / 2), Main_Win(17).Top + 23, White, 0, False)
    Call RenderText(Font_Georgia, WordWrap(Font_Georgia, Trim$(DialogDesc), 267), Main_Win(17).Left + 110, Main_Win(17).Top + 52, White, 0, False)
    
    For i = 1 To 4
        Call RenderText(Font_Georgia, DialogChoice(i), MAIN_BUTTON(35 + i).Left + 97 - (EngineGetTextWidth(Font_Georgia, DialogChoice(i)) / 2), MAIN_BUTTON(35 + i).Top + 8, White, 0, False)
    Next
    
End Sub

Public Sub DrawEquipmentDx8()
Dim i As Long, ItemNum As Long
Dim sRECT As RECT, x As Long, y As Long
Dim Texto As String, TwoH As Boolean

    If InGame = False Then Exit Sub
    
    ' Equipamento 1 - Espada
    If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
        ItemNum = GetPlayerEquipment(MyIndex, Weapon)
        x = Main_Win(1).Left + 15: y = Main_Win(1).Top + 73
        
        ' Renderizar
        If Tex_Item(Item(ItemNum).pic).Width > 32 Then
            RenderTexture Tex_Item(Item(ItemNum).pic), x, y, EquipFrame(1) * 32, 0, 32, 32, 32, 32
        Else
            RenderTexture Tex_Item(Item(ItemNum).pic), x, y, 0, 0, 32, 32, 32, 32
        End If
        
        If Item(ItemNum).TwoHands = True Then
            TwoH = True
        End If
    Else
        x = Main_Win(1).Left + 15: y = Main_Win(1).Top + 73
        RenderTexture Tex_Item(874), x, y, 0, 0, 32, 32, 32, 32
    End If
    
    ' Equipamento 1 - Elmo
    If GetPlayerEquipment(MyIndex, Helmet) > 0 Then
        ItemNum = GetPlayerEquipment(MyIndex, Helmet)
        x = Main_Win(1).Left + 50: y = Main_Win(1).Top + 38
        
        ' Renderizar
        If Tex_Item(Item(ItemNum).pic).Width > 32 Then
            RenderTexture Tex_Item(Item(ItemNum).pic), x, y, EquipFrame(3) * 32, 0, 32, 32, 32, 32
        Else
            RenderTexture Tex_Item(Item(ItemNum).pic), x, y, 0, 0, 32, 32, 32, 32
        End If
        
    Else
        x = Main_Win(1).Left + 50: y = Main_Win(1).Top + 38
        RenderTexture Tex_Item(876), x, y, 0, 0, 32, 32, 32, 32
    End If
    
    ' Equipamento 1 - Escudo
    If GetPlayerEquipment(MyIndex, Shield) > 0 Then
        ItemNum = GetPlayerEquipment(MyIndex, Shield)
        x = Main_Win(1).Left + 85: y = Main_Win(1).Top + 73
        
        ' Renderizar
        If Tex_Item(Item(ItemNum).pic).Width > 32 Then
            RenderTexture Tex_Item(Item(ItemNum).pic), x, y, EquipFrame(4) * 32, 0, 32, 32, 32, 32
        Else
            RenderTexture Tex_Item(Item(ItemNum).pic), x, y, 0, 0, 32, 32, 32, 32
        End If
        
    Else
        If TwoH = False Then
            x = Main_Win(1).Left + 85: y = Main_Win(1).Top + 73
            RenderTexture Tex_Item(872), x, y, 0, 0, 32, 32, 32, 32
            Else
            x = Main_Win(1).Left + 85: y = Main_Win(1).Top + 73
            RenderTexture Tex_Item(868), x, y, 0, 0, 32, 32, 32, 32
        End If
    End If
    
    ' Equipamento 1 - Armadura
    If GetPlayerEquipment(MyIndex, Armor) > 0 Then
        ItemNum = GetPlayerEquipment(MyIndex, Armor)
        x = Main_Win(1).Left + 50: y = Main_Win(1).Top + 73
        
        ' Renderizar
        If Tex_Item(Item(ItemNum).pic).Width > 32 Then
            RenderTexture Tex_Item(Item(ItemNum).pic), x, y, EquipFrame(2) * 32, 0, 32, 32, 32, 32
        Else
            RenderTexture Tex_Item(Item(ItemNum).pic), x, y, 0, 0, 32, 32, 32, 32
        End If
        
    Else
        x = Main_Win(1).Left + 50: y = Main_Win(1).Top + 73
        RenderTexture Tex_Item(873), x, y, 0, 0, 32, 32, 32, 32
    End If
    
    ' Equipamento 2 - Luva/Anel²
    If GetPlayerEquipmentExtra(MyIndex, Boot) > 0 Then
        ItemNum = GetPlayerEquipmentExtra(MyIndex, Boot)
        x = Main_Win(1).Left + 15: y = Main_Win(1).Top + 108
        
        ' Renderizar
        If Tex_Item(Item(ItemNum).pic).Width > 32 Then
            RenderTexture Tex_Item(Item(ItemNum).pic), x, y, EquipFrame(5) * 32, 0, 32, 32, 32, 32
        Else
            RenderTexture Tex_Item(Item(ItemNum).pic), x, y, 0, 0, 32, 32, 32, 32
        End If
        
    Else
        x = Main_Win(1).Left + 15: y = Main_Win(1).Top + 108
        RenderTexture Tex_Item(869), x, y, 0, 0, 32, 32, 32, 32
    End If
    
    ' Equipamento 2 - Anel
    If GetPlayerEquipmentExtra(MyIndex, Ring) > 0 Then
        ItemNum = GetPlayerEquipmentExtra(MyIndex, Ring)
        x = Main_Win(1).Left + 85: y = Main_Win(1).Top + 108
        
        ' Renderizar
        If Tex_Item(Item(ItemNum).pic).Width > 32 Then
            RenderTexture Tex_Item(Item(ItemNum).pic), x, y, EquipFrame(6) * 32, 0, 32, 32, 32, 32
        Else
            RenderTexture Tex_Item(Item(ItemNum).pic), x, y, 0, 0, 32, 32, 32, 32
        End If
        
    Else
        x = Main_Win(1).Left + 85: y = Main_Win(1).Top + 108
        RenderTexture Tex_Item(869), x, y, 0, 0, 32, 32, 32, 32
    End If
    
    ' Equipamento 2 - Calça
    If GetPlayerEquipmentExtra(MyIndex, Legs) > 0 Then
        ItemNum = GetPlayerEquipmentExtra(MyIndex, Legs)
        x = Main_Win(1).Left + 50: y = Main_Win(1).Top + 108
        
        ' Renderizar
        If Tex_Item(Item(ItemNum).pic).Width > 32 Then
            RenderTexture Tex_Item(Item(ItemNum).pic), x, y, EquipFrame(7) * 32, 0, 32, 32, 32, 32
        Else
            RenderTexture Tex_Item(Item(ItemNum).pic), x, y, 0, 0, 32, 32, 32, 32
        End If
        
    Else
        x = Main_Win(1).Left + 50: y = Main_Win(1).Top + 108
        RenderTexture Tex_Item(871), x, y, 0, 0, 32, 32, 32, 32
    End If
    
    ' Equipamento 2 - Amuleto
    If GetPlayerEquipmentExtra(MyIndex, Amulet) > 0 Then
        ItemNum = GetPlayerEquipmentExtra(MyIndex, Amulet)
        x = Main_Win(1).Left + 15: y = Main_Win(1).Top + 38
        
        ' Renderizar
        If Tex_Item(Item(ItemNum).pic).Width > 32 Then
            RenderTexture Tex_Item(Item(ItemNum).pic), x, y, EquipFrame(8) * 32, 0, 32, 32, 32, 32
        Else
            RenderTexture Tex_Item(Item(ItemNum).pic), x, y, 0, 0, 32, 32, 32, 32
        End If
        
    Else
        x = Main_Win(1).Left + 15: y = Main_Win(1).Top + 38
        RenderTexture Tex_Item(867), x, y, 0, 0, 32, 32, 32, 32
    End If
    
    ' Equipamento 1 - Flecha Equipada
    If Player(MyIndex).AmmoEquip > 0 Then
        ItemNum = Player(MyIndex).AmmoEquip
        x = Main_Win(1).Left + 85: y = Main_Win(1).Top + 38
        RenderTexture Tex_Item(Item(ItemNum).pic), x, y, 0, 0, 32, 32, 32, 32
        
        If Player(MyIndex).AmmoEquip > 0 Then
            Call RenderText(Font_Georgia, Format$(ConvertCurrency(Str(PlayerAmmoEquip)), "#,###,###,###"), x + -4, y + 22, White, 0, False)
        End If
    Else
        x = Main_Win(1).Left + 85: y = Main_Win(1).Top + 38
        RenderTexture Tex_Item(875), x, y, 0, 0, 32, 32, 32, 32
    End If
    
End Sub

    Public Sub DrawHudDx8()
    Dim sRECT As RECT, BarWidth As Long
    Dim HairView As Boolean
    Dim px As Long, py As Long
    Dim SpriteNum As Long, Anim As Byte
    
    ' --- Render Janela Principal (Inventário) ---
    With Tex_Main_Win(5)
        RenderTexture Tex_Main_Win(5), 8, 8, 0, 0, .Width, .Height, .Width, .Height
    End With
    
    HairView = True
    
    ' --- HP e MP Bars ---
    DrawHudBar 67, 28, Player(MyIndex).Vital(1), Player(MyIndex).MaxVital(1) ' HP
    DrawHudBar 67, 62, Player(MyIndex).Vital(2), Player(MyIndex).MaxVital(2) ' MP
    
    ' --- EXP Bar ---
    DrawHudBarExp 16, 97, Player(MyIndex).Exp, GetPlayerNextLevel(MyIndex)
    
    ' --- Sprite do Personagem ---
    Anim = Player(MyIndex).Resp
    SpriteNum = GetPlayerSprite(MyIndex)
    If Player(MyIndex).Cor = 1 Then SpriteNum = SpriteNum + 1
    
    DrawPlayerSprite 8, 18, SpriteNum, Anim, HairView
    
    ' --- Addons ---
    DrawPlayerAddons 8, 18, HairView
    
    ' --- Nível do Personagem ---
    DrawPlayerLevel 30, 5, MyIndex
    
    ' --- Health, Mana e EXP Texto ---
    DrawPlayerStats 144, 5, 144, 41, 119, 75, MyIndex
    
    ' --- VIP Days ---
    If Premium = "Sim" Then RenderText Font_Georgia, RPremium, 10, 112, Yellow, 0, False
    
    ' --- Hotbar e Botões ---
    RenderTexture Tex_Main_Win(8), Main_Win(8).Left, Main_Win(8).Top, 0, 0, Tex_Main_Win(8).Width, Tex_Main_Win(8).Height, Tex_Main_Win(8).Width, Tex_Main_Win(8).Height
    DrawHotbarSlots
    
    If MAIN_BUTTON(8).LockState = False Then
        RenderTexture Tex_Main_Win(9), 764, 327, 0, 0, Tex_Main_Win(9).Width, Tex_Main_Win(9).Height, Tex_Main_Win(9).Width, Tex_Main_Win(9).Height
    End If
    
    ' --- Mensagem de Sem Flechas ---
    DrawNoAmmoMessage
    
    ' --- MiniMap ---
    'DrawMiniMap
End Sub


Public Sub DrawHotbarSlots()
Dim i As Long, n As Long
Dim Temp As Long, Temp2 As Long

    For i = 1 To 12
        If Hotbar(i).Slot > 0 Then
            If Hotbar(i).sType = 1 Then ' Item
                If Item(Hotbar(i).Slot).pic > 0 Then
                    Select Case Item(Hotbar(i).Slot).Type
                    Case ITEM_TYPE_CONSUME, ITEM_TYPE_CURRENCY, ITEM_TYPE_TORCH
                    If Hotbar(i).Amount > 0 Then
                        RenderTexture Tex_Item(Item(Hotbar(i).Slot).pic), Main_Win(8).Left + 4 + ((i - 1) * 38), Main_Win(8).Top + 4, 0, 0, 32, 32, 32, 32
                        RenderText Font_Georgia, Hotbar(i).Amount, Main_Win(8).Left + 18 + ((i - 1) * 38) - (getWidth(Font_Georgia, Hotbar(i).Amount) / 2), Main_Win(8).Top + 12, BrightGreen, 0
                    Else
                        RenderTexture Tex_Item(Item(Hotbar(i).Slot).pic), Main_Win(8).Left + 4 + ((i - 1) * 38), Main_Win(8).Top + 4, 0, 0, 32, 32, 32, 32, D3DColorARGB(100, 255, 0, 0)
                    End If
                    
                    Case Else
                        RenderTexture Tex_Item(Item(Hotbar(i).Slot).pic), Main_Win(8).Left + 4 + ((i - 1) * 38), Main_Win(8).Top + 4, 0, 0, 32, 32, 32, 32
                    End Select
                End If
            
            ElseIf Hotbar(i).sType = 2 Then ' Spell
                    If SpellCD(i) = 0 Then
                        RenderTexture Tex_SpellIcon(Spell(Hotbar(i).Slot).Level(Hotbar(i).LvlSpell).Icon), Main_Win(8).Left + 4 + ((i - 1) * 38), Main_Win(8).Top + 4, 0, 0, 32, 32, 32, 32
                    Else
                        Temp = Int(((SpellCD(i) - GetTickCount) / 1000))
                        Temp2 = Hotbar(i).SpellCD + 1 + Temp
                        RenderTexture Tex_SpellIcon(Spell(Hotbar(i).Slot).Level(Hotbar(i).LvlSpell).Icon), Main_Win(8).Left + 4 + ((i - 1) * 38), Main_Win(8).Top + 4, 32, 0, 32, 32, 32, 32
                        RenderText Font_Georgia, Temp2, Main_Win(8).Left + 18 + ((i - 1) * 38) - (getWidth(Font_Georgia, Temp2) / 2), Main_Win(8).Top + 12, BrightRed, 0
                    End If
            End If
        End If

        Select Case i
        Case 11
            RenderText Font_Georgia, "-", Main_Win(8).Left + 4 + ((i - 1) * 38), Main_Win(8).Top + 20, White, 0
        Case 12
            RenderText Font_Georgia, "=", Main_Win(8).Left + 4 + ((i - 1) * 38), Main_Win(8).Top + 20, White, 0
        Case Else
            RenderText Font_Georgia, i, Main_Win(8).Left + 4 + ((i - 1) * 38), Main_Win(8).Top + 20, White, 0
        End Select
    Next

End Sub

Public Sub DrawSpellDescDx8()
Dim i As Long, SpellNum As Long
Dim sRECT As RECT, x As Long, y As Long
Dim SpellName As String, SpellDesc As String, ManaCost As String

    If InGame = False Then Exit Sub
    If DescSpellPlayerDx8 = 0 Then Exit Sub
    If DescSpellPlayerDx8Slot = 0 Then Exit Sub
    
    With Tex_Main_Win(16)
        RenderTexture Tex_Main_Win(16), GlobalX + 25, GlobalY - 25, 0, 0, .Width, .Height, .Width, .Height, -1
    End With
    
    SpellNum = SkillTree(DescSpellPlayerDx8Slot).Habilidade(1)
    
    If SpellNum > 0 Then
    x = GlobalX + 25
    y = GlobalY - 25
    
    SpellName = Trim$(Spell(SpellNum).Name)
    ManaCost = "Mana: " & Spell(SpellNum).Level(1).MPCost
    
    RenderTexture Tex_SpellIcon(SkillTree(DescSpellPlayerDx8Slot).Icon), x + 13, y + 20, 0, 0, 32, 32, 32, 32 ' SpellIcon
    Call RenderText(Font_Georgia, SpellName, x + 131 - (EngineGetTextWidth(Font_Georgia, Trim$(Spell(SpellNum).Name)) / 2), y + 19, White, False)
    Call RenderText(Font_Georgia, ManaCost, x + 131 - (EngineGetTextWidth(Font_Georgia, ManaCost) / 2), y + 39, White, False)
    
    ' Descrição
    Call RenderText(Font_Georgia, WordWrap(Font_Georgia, Trim$(Spell(SpellNum).Desc), 173), x + 20, y + 63, White, 0)

    End If
    
End Sub

Public Sub DrawShopDescDx8()
Dim i As Long, ItemNum As Long
Dim sRECT As RECT, x As Long, y As Long
Dim ItemInfo(1 To 20) As String
Dim ItemName As String
Dim ItemRarity As Byte
Dim ItemRariStr As String

    If InGame = False Then Exit Sub
    If DragInvItemDx8 > 0 Then Exit Sub
    If DescShopItemDx8 = 0 Then Exit Sub
    If InShop = 0 Then Exit Sub

    ' Setar Item
    ItemNum = DescShopItemDx8
    
    ' Raridade
    Select Case Item(ItemNum).Rarity
        Case 1: ItemRarity = BrightGreen: ItemRariStr = "[Raro]"
        Case 2: ItemRarity = BrightBlue: ItemRariStr = "[Muito Raro]"
        Case 3: ItemRarity = Yellow: ItemRariStr = "[Épico]"
        Case 4: ItemRarity = Pink: ItemRariStr = "[Lendario]"
        Case 5: ItemRarity = Orange: ItemRariStr = "[Único]"
        Case Else: ItemRarity = White: ItemRariStr = "[Comum]"
    End Select
    
    ' Pegar Valores
    ItemName = Trim$(Item(ItemNum).Name)
    
    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Type = ITEM_TYPE_CONSUME Or Item(ItemNum).Type = ITEM_TYPE_TORCH Then
        ItemInfo(1) = "Qtd: " & Shop(InShop).TradeItem(DescShopItemDx8Slot).ItemValue
    Else
        ItemInfo(1) = ItemRariStr
    End If
    
    ' Equipamentos Normais
    Select Case Item(ItemNum).Type
    Case ITEM_TYPE_CURRENCY, ITEM_TYPE_SPELL
    
    x = GlobalX + 25
    y = GlobalY - 40
    
    With Tex_Main_Win(19)
        RenderTexture Tex_Main_Win(19), x, y, 0, 0, .Width, .Height, .Width, .Height, -1
    End With
    
    RenderTexture Tex_Item(Item(ItemNum).pic), x + 13, y + 20, 0, 0, 32, 32, 32, 32 ' SpellIcon
    Call RenderText(Font_Georgia, Trim$(Item(ItemNum).Name), x + 131 - (EngineGetTextWidth(Font_Georgia, Trim$(Item(ItemNum).Name)) / 2), y + 19, ItemRarity, False)
    Call RenderText(Font_Georgia, ItemInfo(1), x + 131 - (EngineGetTextWidth(Font_Georgia, ItemInfo(1)) / 2), y + 39, ItemRarity, False)
    
    ' Descrição
    Call RenderText(Font_Georgia, WordWrap(Font_Georgia, Trim$(Item(ItemNum).Desc), 173), x + 20, y + 63, White, 0)

    Exit Sub
    End Select
    
    ' Stats
    ItemInfo(2) = "Dano: " & Item(ItemNum).Data2
    ItemInfo(3) = "For: +" & Item(ItemNum).Add_Stat(1)
    ItemInfo(4) = "Vit: +" & Item(ItemNum).Add_Stat(2)
    ItemInfo(5) = "Int: +" & Item(ItemNum).Add_Stat(3)
    ItemInfo(6) = "Agi: +" & Item(ItemNum).Add_Stat(4)
    ItemInfo(7) = "Will: +" & Item(ItemNum).Add_Stat(5)
    
    ' Stats Req
    ItemInfo(8) = "Defesa: " & Item(ItemNum).Data4
    ItemInfo(9) = "For: " & Item(ItemNum).Stat_Req(1)
    ItemInfo(10) = "Vit: " & Item(ItemNum).Stat_Req(2)
    ItemInfo(11) = "Int: " & Item(ItemNum).Stat_Req(3)
    ItemInfo(12) = "Agi: " & Item(ItemNum).Stat_Req(4)
    ItemInfo(13) = "Will: " & Item(ItemNum).Stat_Req(5)
    ItemInfo(14) = "Peso: " & Item(ItemNum).Peso
    
    
    'RenderTexture Tex_Main_Win(3), GlobalX + 15, GlobalY + 15, 0, 0, 210, 238, 210, 238, -1
    x = GlobalX + 25
    y = GlobalY - 40
    
    With Tex_Main_Win(3)
        RenderTexture Tex_Main_Win(3), x, y, 0, 0, .Width, .Height, .Width, .Height, -1
    End With
    
    RenderTexture Tex_Item(Item(ItemNum).pic), x + 17, y + 20, 0, 0, 32, 32, 32, 32
    
    Call RenderText(Font_Georgia, Trim$(ItemName), x + 133 - ((EngineGetTextWidth(Font_Georgia, Trim$(ItemName)) / 2)), y + 19, ItemRarity, 0, False)
    Call RenderText(Font_Georgia, Trim$(ItemInfo(1)), x + 133 - ((EngineGetTextWidth(Font_Georgia, Trim$(ItemInfo(1))) / 2)), y + 39, ItemRarity, 0, False)
    Call RenderText(Font_Georgia, Trim$(ItemInfo(2)), x + 64 - ((EngineGetTextWidth(Font_Georgia, Trim$(ItemInfo(2))) / 2)), y + 59, White, 0, False)
    Call RenderText(Font_Georgia, Trim$(ItemInfo(8)), x + 166 - ((EngineGetTextWidth(Font_Georgia, Trim$(ItemInfo(8))) / 2)), y + 59, White, 0, False)
    
    ' Stats
    For i = 1 To Stats.Stat_Count - 1
        Call RenderText(Font_Georgia, Trim$(ItemInfo(3 + i - 1)), x + 64 - ((EngineGetTextWidth(Font_Georgia, Trim$(ItemInfo(3 + i - 1))) / 2)), y + 79 + ((i - 1) * 20), BrightGreen, 0, False)
        Call RenderText(Font_Georgia, Trim$(ItemInfo(9 + i - 1)), x + 166 - ((EngineGetTextWidth(Font_Georgia, Trim$(ItemInfo(9 + i - 1))) / 2)), y + 79 + ((i - 1) * 20), Orange, 0, False)
    Next
    
    ' Peso
    Call RenderText(Font_Georgia, Trim$(ItemInfo(14)), x + 166 - ((EngineGetTextWidth(Font_Georgia, Trim$(ItemInfo(14))) / 2)), y + 219, Orange, 0, False)
    
End Sub

' Descrição do Item
Public Sub DrawInvDescDx8Part1()
    Dim ItemNum As Long, BackNum As Byte
    Dim x As Long, y As Long
    Dim ItemName As String, ItemType As String
    Dim ItemRarity As Long, ItemRariStr As String

    ' Checagens básicas
    If Not InGame Then Exit Sub
    If DragInvItemDx8 > 0 Then Exit Sub
    If DescInvItemDx8 = 0 Then Exit Sub

    ItemNum = DescInvItemDx8
    If ItemNum = 0 Or ItemNum > MAX_ITEMS Then Exit Sub

    ' ----------------------------
    ' Raridade do item
    ' ----------------------------
    Select Case Item(ItemNum).Rarity
        Case 1: ItemRarity = BrightGreen: ItemRariStr = "[Raro]"
        Case 2: ItemRarity = BrightBlue: ItemRariStr = "[Muito Raro]"
        Case 3: ItemRarity = Yellow: ItemRariStr = "[Épico]"
        Case 4: ItemRarity = Pink: ItemRariStr = "[Lendário]"
        Case 5: ItemRarity = Orange: ItemRariStr = "[Único]"
        Case Else: ItemRarity = White: ItemRariStr = "[Comum]"
    End Select

    ' ----------------------------
    ' Nome do item
    ' ----------------------------
    If PlayerInv(DescInvItemDx8Slot).Refine > 0 Then
        Select Case Item(ItemNum).Type
            Case ITEM_TYPE_OUTFIT
                ItemName = Trim$(Item(ItemNum).Name) & " [" & Trim$(Outfit(PlayerInv(DescInvItemDx8Slot).Refine).Name) & "]"
            Case ITEM_TYPE_CONSUME, ITEM_TYPE_CURRENCY, ITEM_TYPE_RESET, ITEM_TYPE_TORCH
                ItemName = Trim$(Item(ItemNum).Name)
            Case ITEM_TYPE_SPELL
                ItemName = Trim$(Item(ItemNum).Name) & " [" & Trim$(Spell(PlayerInv(DescInvItemDx8Slot).Refine).Name) & "]"
            Case Else
                ItemName = Trim$(Item(ItemNum).Name) & " +" & PlayerInv(DescInvItemDx8Slot).Refine
        End Select
    Else
        ItemName = Trim$(Item(ItemNum).Name)
    End If

    ' ----------------------------
    ' Tipo do item e fundo
    ' ----------------------------
    BackNum = 32
    Select Case Item(ItemNum).Type
        Case ITEM_TYPE_WEAPON
            BackNum = 19
            Select Case Item(ItemNum).WeaponCategory
                Case 1: ItemType = "Espada"
                Case 2: ItemType = "Espada Duas Mãos"
                Case 3: ItemType = "Lança"
                Case 4: ItemType = "Arco"
                Case 5: ItemType = "Besta"
                Case 6: ItemType = "Machado"
                Case 7: ItemType = "Machado Duas Mãos"
                Case 8: ItemType = "Cetro"
                Case 9: ItemType = "Cajado"
                Case 10: ItemType = "Orbe"
                Case 11: ItemType = "Adaga"
                Case 12: ItemType = "Adaga Dupla"
                Case 13: ItemType = "Garra"
                Case 14: ItemType = "Punho"
                Case 15: ItemType = "Livro"
                Case 16: ItemType = "Porrete"
                Case 17: ItemType = "Porrete Duas Mãos"
                Case 18: ItemType = "Chicote"
                Case 19: ItemType = "Instrumento"
                Case 20: ItemType = "Foice"
                Case Else: ItemType = "Indefinido"
            End Select
        Case ITEM_TYPE_ARMOR: ItemType = "Armadura": BackNum = 19
        Case ITEM_TYPE_HELMET: ItemType = "Elmo": BackNum = 19
        Case ITEM_TYPE_SHIELD: ItemType = "Escudo": BackNum = 19
        Case ITEM_TYPE_CONSUME: ItemType = "Consumível": BackNum = 32
        Case ITEM_TYPE_KEY: ItemType = "Chave": BackNum = 32
        Case ITEM_TYPE_CURRENCY, ITEM_TYPE_TORCH: ItemType = "Qtd: " & PlayerInv(DescInvItemDx8Slot).value: BackNum = 32
        Case ITEM_TYPE_SPELL: ItemType = "Habilidade": BackNum = 32
        Case ITEM_TYPE_BOOT: ItemType = "Botas": BackNum = 19
        Case ITEM_TYPE_RING: ItemType = "Anél/Pet": BackNum = 19
        Case ITEM_TYPE_AMULET: ItemType = "Amuleto": BackNum = 19
        Case ITEM_TYPE_LEGS: ItemType = "Calça": BackNum = 19
        Case ITEM_TYPE_RESET: ItemType = "Resetador": BackNum = 32
        Case ITEM_TYPE_OUTFIT: ItemType = "Roupa": BackNum = 32
        Case ITEM_TYPE_VIP: ItemType = "VIP": BackNum = 32
        Case Else: ItemType = "Indefinido": BackNum = 32
    End Select

    ' ----------------------------
    ' Posições
    ' ----------------------------
    x = GlobalX + 25
    y = GlobalY - 40

    ' ----------------------------
    ' Fundo do tooltip
    ' ----------------------------
    With Tex_Main_Win(BackNum)
        RenderTexture Tex_Main_Win(BackNum), x, y, 0, 0, .Width, .Height, .Width, .Height, -1
    End With

    ' ----------------------------
    ' Renderiza ícone e textos
    ' ----------------------------
    If Item(ItemNum).pic > 0 Then RenderTexture Tex_Item(Item(ItemNum).pic), x + 13, y + 20, 0, 0, 32, 32, 32, 32
    If ItemName <> vbNullString Then RenderText Font_Georgia, ItemName, x + 131 - EngineGetTextWidth(Font_Georgia, ItemName) / 2, y + 19, ItemRarity, False
    If ItemType <> vbNullString Then RenderText Font_Georgia, ItemType, x + 131 - EngineGetTextWidth(Font_Georgia, ItemType) / 2, y + 39, ItemRarity, False

    ' ----------------------------
    ' Descrição
    ' ----------------------------
    RenderText Font_Georgia, WordWrap(Font_Georgia, Trim$(Item(ItemNum).Desc), 173), x + 20, y + 63, White, 0

    ' ----------------------------
    ' Restrições de classe
    ' ----------------------------
    With Tex_Main_Win(29)
        Dim offset As Long
        For offset = 0 To 5
            If Not Item(ItemNum).ClassRest(offset + 1) Then
                RenderTexture Tex_Main_Win(29), x + 12 + (offset * 18), y + 164, 0, 16 * offset, 16, 16, 16, 16, -1
            End If
        Next
    End With

    ' ----------------------------
    ' Peso
    ' ----------------------------
    RenderText Font_Georgia, "Peso: " & Trim$(Item(ItemNum).Peso), x + 164 - EngineGetTextWidth(Font_Georgia, "Peso: " & Trim$(Item(ItemNum).Peso)) / 2, y + 163, White, 0
End Sub

Public Sub DrawInvDescDx8Part2()
    Dim ItemNum As Long, x As Long, y As Long
    Dim ItemName As String, ItemRarity As Long, ItemRariStr As String
    Dim DmgItem As Long, DefItem As Long
    Dim i As Long
    Dim ElementStr As String
    Dim ItemInfo(1 To 20) As String

    ' Checagens básicas
    If Not InGame Then Exit Sub
    If DragInvItemDx8 > 0 Then Exit Sub
    If DescInvItemDx8 = 0 Then Exit Sub

    ItemNum = DescInvItemDx8
    If ItemNum = 0 Or ItemNum > MAX_ITEMS Then Exit Sub

    ' Itens que não precisam de stats detalhados
    Select Case Item(ItemNum).Type
        Case 5, 6, 7, 8, 13
            DrawInvDescDx8Part1
            Exit Sub
    End Select

    ' Raridade
    Select Case Item(ItemNum).Rarity
        Case 1: ItemRarity = BrightGreen: ItemRariStr = "[Raro]"
        Case 2: ItemRarity = BrightBlue: ItemRariStr = "[Muito Raro]"
        Case 3: ItemRarity = Yellow: ItemRariStr = "[Épico]"
        Case 4: ItemRarity = Pink: ItemRariStr = "[Lendário]"
        Case 5: ItemRarity = Orange: ItemRariStr = "[Único]"
        Case Else: ItemRarity = White: ItemRariStr = "[Comum]"
    End Select

    ' Nome do item com refine
    If PlayerInv(DescInvItemDx8Slot).Refine > 0 Then
        ItemName = Trim$(Item(ItemNum).Name) & " +" & PlayerInv(DescInvItemDx8Slot).Refine
    Else
        ItemName = Trim$(Item(ItemNum).Name)
    End If

    ' Dano e defesa com refine
    DmgItem = Item(ItemNum).Data2
    DefItem = Item(ItemNum).Data4

    If DefItem > 0 Then DefItem = DefItem + ((Item(ItemNum).Data4 * (PlayerInv(DescInvItemDx8Slot).Refine * 8.5)) / 100)
    If DmgItem > 0 Then DmgItem = DmgItem + ((Item(ItemNum).Data2 * (PlayerInv(DescInvItemDx8Slot).Refine * 20.8)) / 100)

    ' Prepara stats e requisitos
    ItemInfo(2) = "Dano: " & DmgItem
    ItemInfo(3) = "For: +" & Item(ItemNum).Add_Stat(1)
    ItemInfo(4) = "Vit: +" & Item(ItemNum).Add_Stat(2)
    ItemInfo(5) = "Int: +" & Item(ItemNum).Add_Stat(3)
    ItemInfo(6) = "Agi: +" & Item(ItemNum).Add_Stat(4)
    ItemInfo(7) = "Des: +" & Item(ItemNum).Add_Stat(5)

    ItemInfo(8) = "Defesa: " & DefItem
    For i = 1 To 5
        ItemInfo(8 + i) = Item(ItemNum).Stat_Req(i)
    Next
    ItemInfo(14) = "Peso: " & Item(ItemNum).Peso

    ' ----------------------------
    ' Posições
    ' ----------------------------
    x = GlobalX + 25
    y = GlobalY - 40

    ' Fundo
    With Tex_Main_Win(3)
        RenderTexture Tex_Main_Win(3), x, y, 0, 0, .Width, .Height, .Width, .Height, -1
    End With

    ' Ícone e nome
    If Item(ItemNum).pic > 0 Then RenderTexture Tex_Item(Item(ItemNum).pic), x + 17, y + 20, 0, 0, 32, 32, 32, 32
    RenderText Font_Georgia, ItemName, x + 133 - EngineGetTextWidth(Font_Georgia, ItemName) / 2, y + 19, ItemRarity, False
    RenderText Font_Georgia, "Dano/Defesa", x + 133 - EngineGetTextWidth(Font_Georgia, "Dano/Defesa") / 2, y + 39, ItemRarity, False

    ' Stats principais
    RenderText Font_Georgia, ItemInfo(2), x + 64 - EngineGetTextWidth(Font_Georgia, ItemInfo(2)) / 2, y + 59, White, False
    RenderText Font_Georgia, ItemInfo(8), x + 166 - EngineGetTextWidth(Font_Georgia, ItemInfo(8)) / 2, y + 59, White, False

    RenderText Font_Georgia, "Atributos:", x + 33, y + 79, BrightGreen, False
    RenderText Font_Georgia, "Requisitos:", x + 133, y + 79, BrightRed, False

    ' Loop de atributos e requisitos
    For i = 1 To 5
        RenderText Font_Georgia, ItemInfo(2 + i), x + 64 - EngineGetTextWidth(Font_Georgia, ItemInfo(2 + i)) / 2, y + 119 + (i - 1) * 20, BrightGreen, False

        If GetPlayerStat(MyIndex, i) < Item(ItemNum).Stat_Req(i) Then
            RenderText Font_Georgia, ItemInfo(8 + i), x + 166 - EngineGetTextWidth(Font_Georgia, ItemInfo(8 + i)) / 2, y + 119 + (i - 1) * 20, BrightRed, False
        Else
            RenderText Font_Georgia, ItemInfo(8 + i), x + 166 - EngineGetTextWidth(Font_Georgia, ItemInfo(8 + i)) / 2, y + 119 + (i - 1) * 20, White, False
        End If
    Next

    ' Elemento
    Select Case Item(ItemNum).Elemento
        Case 0: ElementStr = "Neutro"
        Case 1: ElementStr = "Fogo"
        Case 2: ElementStr = "Água"
        Case 3: ElementStr = "Terra"
        Case 4: ElementStr = "Vento"
        Case 5: ElementStr = "Veneno"
        Case 6: ElementStr = "Sagrado"
        Case 7: ElementStr = "Trevas"
        Case Else: ElementStr = "???"
    End Select
    RenderText Font_Georgia, ElementStr, x + 64 - EngineGetTextWidth(Font_Georgia, ElementStr) / 2, y + 99, BrightGreen, False

    ' Level
    If GetPlayerLevel(MyIndex) < Item(ItemNum).LevelReq Then
        RenderText Font_Georgia, "Level: " & Item(ItemNum).LevelReq, x + 166 - EngineGetTextWidth(Font_Georgia, "Level: " & Item(ItemNum).LevelReq) / 2, y + 99, BrightRed, False
    Else
        RenderText Font_Georgia, "Level: " & Item(ItemNum).LevelReq, x + 166 - EngineGetTextWidth(Font_Georgia, "Level: " & Item(ItemNum).LevelReq) / 2, y + 99, White, False
    End If

    ' Classes
    With Tex_Main_Win(29)
        Dim offset As Long
        For offset = 0 To 5
            If Not Item(ItemNum).ClassRest(offset + 1) Then
                RenderTexture Tex_Main_Win(29), x + 16 + (offset * 18), y + 220, 0, 16 * offset, 16, 16, 16, 16, -1
            End If
        Next
    End With

    ' Peso
    RenderText Font_Georgia, ItemInfo(14), x + 166 - EngineGetTextWidth(Font_Georgia, ItemInfo(14)) / 2, y + 219, White, False
End Sub


Public Sub DrawInvDescDx8()
Dim i As Long, ItemNum As Long
Dim sRECT As RECT, x As Long, y As Long
Dim ItemInfo(1 To 20) As String
Dim ItemName As String
Dim ItemRarity As Byte
Dim ItemRariStr As String
Dim ElementStr As String

    If InGame = False Then Exit Sub
    If DragInvItemDx8 > 0 Then Exit Sub
    If DescInvItemDx8 = 0 Then Exit Sub
    
    ' Setar Item
    ItemNum = DescInvItemDx8
    
    ' Raridade
    Select Case Item(DescInvItemDx8).Rarity
        Case 1: ItemRarity = BrightGreen: ItemRariStr = "[Raro]"
        Case 2: ItemRarity = BrightBlue: ItemRariStr = "[Muito Raro]"
        Case 3: ItemRarity = Yellow: ItemRariStr = "[Épico]"
        Case 4: ItemRarity = Pink: ItemRariStr = "[Lendario]"
        Case 5: ItemRarity = Orange: ItemRariStr = "[Único]"
        Case Else: ItemRarity = White: ItemRariStr = "[Comum]"
    End Select
    
    ' Pegar Valores
    If PlayerInv(DescInvItemDx8Slot).Refine > 0 Then
        ItemName = Trim$(Item(DescInvItemDx8).Name) & " +" & PlayerInv(DescInvItemDx8Slot).Refine
    Else
        ItemName = Trim$(Item(DescInvItemDx8).Name)
    End If
    
    If Item(DescInvItemDx8).Type = ITEM_TYPE_CURRENCY Or Item(DescInvItemDx8).Type = ITEM_TYPE_CONSUME Or Item(DescInvItemDx8).Type = ITEM_TYPE_TORCH Then
        ItemInfo(1) = "Qtd: " & PlayerInv(DescInvItemDx8Slot).value
    Else
        ItemInfo(1) = ItemRariStr
    End If
    
    ' Equipamentos Normais
    Select Case Item(ItemNum).Type
    Case ITEM_TYPE_CURRENCY, ITEM_TYPE_SPELL, ITEM_TYPE_CONSUME
    
    'RenderTexture Tex_Main_Win(3), GlobalX + 15, GlobalY + 15, 0, 0, 210, 238, 210, 238, -1
    x = GlobalX + 25
    y = GlobalY - 40
    
    With Tex_Main_Win(19)
        RenderTexture Tex_Main_Win(19), x, y, 0, 0, .Width, .Height, .Width, .Height, -1
    End With
    
    RenderTexture Tex_Item(Item(ItemNum).pic), x + 13, y + 20, 0, 0, 32, 32, 32, 32 ' SpellIcon
    Call RenderText(Font_Georgia, Trim$(Item(ItemNum).Name), x + 131 - (EngineGetTextWidth(Font_Georgia, Trim$(ItemName)) / 2), y + 19, ItemRarity, False)
    Call RenderText(Font_Georgia, ItemInfo(1), x + 131 - (EngineGetTextWidth(Font_Georgia, ItemInfo(1)) / 2), y + 39, ItemRarity, False)
    ' Descrição
    Call RenderText(Font_Georgia, WordWrap(Font_Georgia, Trim$(Item(ItemNum).Desc), 173), x + 20, y + 63, White, 0)

    With Tex_Main_Win(29)
        If Item(ItemNum).ClassRest(1) = False Then: RenderTexture Tex_Main_Win(29), x + 12, y + 164, 0, 0, 16, 16, 16, 16, -1
        If Item(ItemNum).ClassRest(2) = False Then: RenderTexture Tex_Main_Win(29), x + 30, y + 164, 0, 16, 16, 16, 16, 16, -1
        If Item(ItemNum).ClassRest(3) = False Then: RenderTexture Tex_Main_Win(29), x + 48, y + 164, 0, 32, 16, 16, 16, 16, -1
        If Item(ItemNum).ClassRest(4) = False Then: RenderTexture Tex_Main_Win(29), x + 66, y + 164, 0, 48, 16, 16, 16, 16, -1
        If Item(ItemNum).ClassRest(5) = False Then: RenderTexture Tex_Main_Win(29), x + 84, y + 164, 0, 64, 16, 16, 16, 16, -1
        If Item(ItemNum).ClassRest(6) = False Then: RenderTexture Tex_Main_Win(29), x + 102, y + 164, 0, 80, 16, 16, 16, 16, -1
    End With
    
    ' Peso
    Call RenderText(Font_Georgia, "Peso: " & Trim$(Item(DescInvItemDx8).Peso), x + 164 - (EngineGetTextWidth(Font_Georgia, "Peso: " & Trim$(Item(DescInvItemDx8).Peso)) / 2), y + 163, White, 0)

    Exit Sub
    End Select
    
    ' Stats
    ItemInfo(2) = "Dano: " & Item(DescInvItemDx8).Data2
    ItemInfo(3) = "For: +" & Item(DescInvItemDx8).Add_Stat(1)
    ItemInfo(4) = "Vit: +" & Item(DescInvItemDx8).Add_Stat(2)
    ItemInfo(5) = "Int: +" & Item(DescInvItemDx8).Add_Stat(3)
    ItemInfo(6) = "Agi: +" & Item(DescInvItemDx8).Add_Stat(4)
    ItemInfo(7) = "Des: +" & Item(DescInvItemDx8).Add_Stat(5)
    
    ' Stats Req
    ItemInfo(8) = "Defesa: " & Item(DescInvItemDx8).Data4
    ItemInfo(9) = "For: " & Item(DescInvItemDx8).Stat_Req(1)
    ItemInfo(10) = "Vit: " & Item(DescInvItemDx8).Stat_Req(2)
    ItemInfo(11) = "Int: " & Item(DescInvItemDx8).Stat_Req(3)
    ItemInfo(12) = "Agi: " & Item(DescInvItemDx8).Stat_Req(4)
    ItemInfo(13) = "Des: " & Item(DescInvItemDx8).Stat_Req(5)
    ItemInfo(14) = "Peso: " & Item(DescInvItemDx8).Peso
    
    x = GlobalX + 25
    y = GlobalY - 40
    
    With Tex_Main_Win(3)
        RenderTexture Tex_Main_Win(3), x, y, 0, 0, .Width, .Height, .Width, .Height, -1
    End With
    
    RenderTexture Tex_Item(Item(DescInvItemDx8).pic), x + 17, y + 20, 0, 0, 32, 32, 32, 32
    
    Call RenderText(Font_Georgia, Trim$(ItemName), x + 133 - ((EngineGetTextWidth(Font_Georgia, Trim$(ItemName)) / 2)), y + 19, ItemRarity, 0, False)
    Call RenderText(Font_Georgia, Trim$(ItemInfo(1)), x + 133 - ((EngineGetTextWidth(Font_Georgia, Trim$(ItemInfo(1))) / 2)), y + 39, ItemRarity, 0, False)
    Call RenderText(Font_Georgia, Trim$(ItemInfo(2)), x + 64 - ((EngineGetTextWidth(Font_Georgia, Trim$(ItemInfo(2))) / 2)), y + 59, White, 0, False)
    Call RenderText(Font_Georgia, Trim$(ItemInfo(8)), x + 166 - ((EngineGetTextWidth(Font_Georgia, Trim$(ItemInfo(8))) / 2)), y + 59, White, 0, False)
    
    Call RenderText(Font_Georgia, "Atributos:", x + 33, y + 79, BrightGreen, 0, False)
    Call RenderText(Font_Georgia, "Requisitos:", x + 133, y + 79, BrightRed, 0, False)
    
    For i = 1 To Stats.Stat_Count - 1
        Call RenderText(Font_Georgia, Trim$(ItemInfo(3 + i - 1)), x + 64 - ((EngineGetTextWidth(Font_Georgia, Trim$(ItemInfo(3 + i - 1))) / 2)), y + 119 + ((i - 1) * 20), BrightGreen, 0, False)
    If GetPlayerStat(MyIndex, i) < Item(ItemNum).Stat_Req(i) Then
        Call RenderText(Font_Georgia, Trim$(ItemInfo(9 + i - 1)), x + 166 - ((EngineGetTextWidth(Font_Georgia, Trim$(ItemInfo(9 + i - 1))) / 2)), y + 119 + ((i - 1) * 20), BrightRed, 0, False)
    Else
        Call RenderText(Font_Georgia, Trim$(ItemInfo(9 + i - 1)), x + 166 - ((EngineGetTextWidth(Font_Georgia, Trim$(ItemInfo(9 + i - 1))) / 2)), y + 119 + ((i - 1) * 20), White, 0, False)
    End If
    Next
        
    Select Case Item(ItemNum).Elemento
    Case 0: ElementStr = "Neutro"
    Case 1: ElementStr = "Fogo"
    Case 2: ElementStr = "Água"
    Case 3: ElementStr = "Terra"
    Case 4: ElementStr = "Vento"
    Case 5: ElementStr = "Veneno"
    Case 6: ElementStr = "Sagrado"
    Case 7: ElementStr = "Trevas"
    Case Else: ElementStr = "???"
    End Select
    
    Call RenderText(Font_Georgia, ElementStr, x + 64 - ((EngineGetTextWidth(Font_Georgia, ElementStr) / 2)), y + 99, BrightGreen, 0, False)
    
    If GetPlayerLevel(MyIndex) < Item(ItemNum).LevelReq Then
        Call RenderText(Font_Georgia, "Level: " & Item(ItemNum).LevelReq, x + 166 - ((EngineGetTextWidth(Font_Georgia, "Level: " & Item(ItemNum).LevelReq) / 2)), y + 99, BrightRed, 0, False)
    Else
        Call RenderText(Font_Georgia, "Level: " & Item(ItemNum).LevelReq, x + 166 - ((EngineGetTextWidth(Font_Georgia, "Level: " & Item(ItemNum).LevelReq) / 2)), y + 99, White, 0, False)
    End If
    
    With Tex_Main_Win(29)
        If Item(ItemNum).ClassRest(1) = False Then: RenderTexture Tex_Main_Win(29), x + 16, y + 220, 0, 0, 16, 16, 16, 16, -1
        If Item(ItemNum).ClassRest(2) = False Then: RenderTexture Tex_Main_Win(29), x + 34, y + 220, 0, 16, 16, 16, 16, 16, -1
        If Item(ItemNum).ClassRest(3) = False Then: RenderTexture Tex_Main_Win(29), x + 52, y + 220, 0, 32, 16, 16, 16, 16, -1
        If Item(ItemNum).ClassRest(4) = False Then: RenderTexture Tex_Main_Win(29), x + 70, y + 220, 0, 48, 16, 16, 16, 16, -1
        If Item(ItemNum).ClassRest(5) = False Then: RenderTexture Tex_Main_Win(29), x + 88, y + 220, 0, 64, 16, 16, 16, 16, -1
        If Item(ItemNum).ClassRest(6) = False Then: RenderTexture Tex_Main_Win(29), x + 106, y + 220, 0, 80, 16, 16, 16, 16, -1
    End With
    
    ' Peso
    Call RenderText(Font_Georgia, Trim$(ItemInfo(14)), x + 166 - ((EngineGetTextWidth(Font_Georgia, Trim$(ItemInfo(14))) / 2)), y + 219, White, 0, False)
    
End Sub

Public Sub DrawDragPlayerSpellDx8()
Dim SpellNum As Long, SpellLevel As Long

    SpellNum = SkillTree(DragPlayerSpellDx8).Habilidade(1)
    SpellLevel = Talentos.LvlSkill(DragPlayerSpellDx8)

    If SpellNum = 0 Or SpellNum > MAX_SPELLS Then Exit Sub
    If SpellLevel = 0 Then Exit Sub
    If SkillTree(DragPlayerSpellDx8).Icon = 0 Then Exit Sub
    RenderTexture Tex_SpellIcon(SkillTree(DragPlayerSpellDx8).Icon), GlobalX - 25, GlobalY - 25, 0, 0, 32, 32, 32, 32

End Sub

Public Sub DrawDragInvItemDx8()

    If DragInvItemDx8 = 0 Then Exit Sub
    RenderTexture Tex_Item(Item(PlayerInv(DragInvItemDx8).num).pic), GlobalX - 25, GlobalY - 25, 0, 0, 32, 32, 32, 32

End Sub

Function HasItem(ByVal ItemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If InGame = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If PlayerInv(i).num = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Type = ITEM_TYPE_CONSUME Or Item(ItemNum).Type = ITEM_TYPE_TORCH Then
                HasItem = PlayerInv(i).value
            Else
                HasItem = 1
            End If

            Exit Function
        End If

    Next

End Function

Function HasItemSlot(ByVal ItemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If InGame = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If PlayerInv(i).num = ItemNum Then
            HasItemSlot = i
            Exit Function
        End If

    Next

End Function

Public Sub DrawDeadBackground()
    RenderTexture Tex_Fade, 0, 0, 0, 0, 928, 512, 32, 32, D3DColorARGB(150, 255, 0, 0)
End Sub

Public Sub DrawSkullsPlayer(ByVal Index As Long)
Dim x As Long, y As Long, Anim As Byte
    
    ' Evitar OverFlow
    If Index = 0 Or Index > MAX_PLAYERS Then Exit Sub
    
    ' Loc X, Y
    x = ConvertMapX(Player(Index).x * 32) + Player(Index).xOffset ' + (getWidth(Font_Georgia, (Trim$(GetPlayerName(Index)))) / 2) + 8
    y = ConvertMapY(Player(Index).y * 32) + Player(Index).yOffset - 46
    
    If THover = Index Then
        y = ConvertMapY(Player(THover).y * 32) + Player(THover).yOffset - 64
        If THover > 0 Then THover = 0
    End If

    ' Desenhar Caveiras
    Select Case Player(Index).PlayerKiller(1)
    Case 1: RenderTexture Tex_Skull, x, y, 64, 0, 32, 32, 32, 32 'Pk Amarelo
    Case 2: RenderTexture Tex_Skull, x, y, 0, 0, 32, 32, 32, 32 'Pk Branco
    Case 3: RenderTexture Tex_Skull, x, y, 32, 0, 32, 32, 32, 32 'Pk Vermelho
    Case 4: RenderTexture Tex_Skull, x, y, 96, 0, 32, 32, 32, 32 'Pk Preto
    End Select
    
End Sub

Public Sub DrawGhostPlayer(ByVal Index As Long)
Dim x As Long, y As Long, Anim As Byte
    
    ' Evitar OverFlow
    If Index = 0 Or Index > MAX_PLAYERS Then Exit Sub
    
    x = ConvertMapX(Player(Index).x * 32) + Player(Index).xOffset
    y = ConvertMapY(Player(Index).y * 32) + Player(Index).yOffset

    Select Case GhostAnim(Index)
    Case 0: Anim = 0
    Case 1: Anim = 64
    Case 2: Anim = 128
    End Select
    
    ' Desenhar Gasparzinho no Tumulo
    RenderTexture Tex_Ghost, x - 16, y - 25, Anim, 0, 64, 64, 64, 64
End Sub

Public Sub DrawPlayerBuffs()
Dim sRECT As RECT, destRect As D3DRECT
Dim drect As RECT
Dim Witdh As Long, Height As Long
Dim i As Long, SpellNum As Long, SpellLevel As Long
Dim BuffIcon As Long, BuffTime As Long

If HudDx8Visible = False Then Exit Sub

    For i = 1 To BuffsAtivos
        SpellNum = Player(MyIndex).Buffs(i).Spell
        SpellLevel = Player(MyIndex).Buffs(i).SpellLevel
        
        'Evitar Overflow
        If SpellLevel = 0 Then Exit Sub
        If SpellNum = 0 Then Exit Sub
        
        If i <= 14 Then
        
            If SpellNum > 0 Then
                BuffIcon = Spell(SpellNum).Level(SpellLevel).Buff.Icon
                BuffTime = Player(MyIndex).Buffs(i).Time
                RenderTexture Tex_Buff(BuffIcon), 239 + ((i - 1) * 34), 8, 0, 0, 32, 32, 32, 32, D3DColorRGBA(255, 255, 255, 150)
                RenderText Font_Georgia, BuffTime, 253 + ((i - 1) * 34) - (EngineGetTextWidth(Font_Georgia, BuffTime) / 2), 37, White, 150
            End If
        
        Else
        
            If SpellNum > 0 Then
                BuffIcon = Spell(SpellNum).Level(SpellLevel).Buff.Icon
                BuffTime = Player(MyIndex).Buffs(i).Time
                RenderTexture Tex_Buff(BuffIcon), 239 + ((i - 15) * 34), 52, 0, 0, 32, 32, 32, 32, D3DColorRGBA(255, 255, 255, 150)
                RenderText Font_Georgia, BuffTime, 253 + ((i - 15) * 34) - (EngineGetTextWidth(Font_Georgia, BuffTime) / 2), 81, White, 150
            End If
        
        End If
        
    Next
    
End Sub

'Projétil Target
Public Sub EditorItem_DrawProjectileTarget()
Dim ItemNum As Long
Dim sRECT As RECT, destRect As D3DRECT
Dim drect As RECT
Dim Witdh As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ItemNum = frmEditor_Item.scrlProjectilePic.value

    If ItemNum < 1 Or ItemNum > NumProjectilesTarget Then
        frmEditor_Item.picProjectile.Cls
        Exit Sub
    End If

    Witdh = Tex_ProjectTilesTarget(ItemNum).Width
    Height = Tex_ProjectTilesTarget(ItemNum).Height

    ' rect for source
    sRECT.Top = 0
    sRECT.Bottom = Height
    sRECT.Left = 0
    sRECT.Right = Witdh
    
    ' same for destination as source
    drect = sRECT
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_ProjectTilesTarget(ItemNum), sRECT, drect
    With destRect
        .X1 = 0
        .X2 = Witdh
        .Y1 = 0
        .Y2 = Height
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Item.picProjectile.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorItem_DrawProjectile", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawProjectileTarget()
Dim Angle As Long, x As Long, y As Long, i As Long
Dim Witdh As Long, Height As Long
    If LastProjectile > 0 Then

        ' ****** Create Particle ******
        For i = 1 To LastProjectile
            With ProjectileList(i)
                If .Graphic Then
                
                    ' ****** Update Position ******
                    Angle = DegreeToRadian * Engine_GetAngle(.x, .y, .tx, .ty)
                    .x = .x + (Sin(Angle) * ElapsedTime * 0.35)
                    .y = .y - (Cos(Angle) * ElapsedTime * 0.35)
                    x = .x
                    y = .y
                    
                    ' ****** Update Rotation ******
                    If .RotateSpeed > 0 Then
                        .Rotate = .Rotate + (.RotateSpeed * ElapsedTime * 0.3)
                        Do While .Rotate > 360
                            .Rotate = .Rotate - 360
                        Loop
                    End If
                    
                    Witdh = Tex_ProjectTilesTarget(.Graphic).Width
                    Height = Tex_ProjectTilesTarget(.Graphic).Height
                    
                    ' ****** Render Projectile ******
                    If .Rotate = 360 Then
                        Call RenderTexture(Tex_ProjectTilesTarget(.Graphic), ConvertMapX(x) + 2, ConvertMapY(y), 0, 0, Witdh, Height, Witdh, Height)
                    Else
                        Call RenderTexture(Tex_ProjectTilesTarget(.Graphic), ConvertMapX(x), ConvertMapY(y), 0, 0, Witdh, Height, Witdh, Height, , .Rotate)
                    End If
                    
                End If
            End With
        Next
        
        ' ****** Erase Projectile ******    Seperate Loop For Erasing
        For i = LastProjectile To LastProjectile Step -1
            If ProjectileList(i).Graphic Then
                If Abs(ProjectileList(i).x - ProjectileList(i).tx) < 10 Then
                    If Abs(ProjectileList(i).y - ProjectileList(i).ty) < 10 Then
                        Call ClearProjectileTarget(i)
                    End If
                End If
            End If
        Next
        
    End If
End Sub

Public Sub EditorMap_DrawLight()
Dim Height As Long, Width As Long
Dim sRECT As RECT, destRect As D3DRECT
Dim drect As RECT

Height = 128
Width = 128
    
    sRECT.Top = 0
    sRECT.Bottom = 128
    sRECT.Left = 0
    sRECT.Right = 128
    
    drect = sRECT
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    
    RenderTextureByRects Tex_Light, sRECT, drect, D3DColorARGB(frmMain.scrlA, frmMain.scrlR, frmMain.scrlG, frmMain.scrlB)
    'DirectX8.RenderTexture Tex_Light, 0, 0, 0, 0, Width, Height, Width, Height
    With destRect
        .X1 = 0
        .X2 = 128
        .Y1 = 0
        .Y2 = 128
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmMain.picLight.hWnd, ByVal (0)
    
End Sub

Public Sub DrawLight(ByVal x As Long, ByVal y As Long, ByVal a As Long, ByVal r As Long, ByVal G As Long, ByVal B As Long)
   If Options.Debug = 1 Then On Error GoTo errorhandler

    RenderTexture Tex_Light, ConvertMapX(x) - 48, ConvertMapY(y) - 48, 0, 0, 128, 128, 128, 128, D3DColorARGB(Abs(Int(a) - Rand(0, 25)), Int(r), Int(G), Int(B))

   Exit Sub
errorhandler:
    HandleError "DrawLight", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawCharacterInfoDx8()
Dim i As Long, ItemNum As Long
Dim sRECT As RECT
Dim MaestriaWitdh(1 To 6) As Long
Dim Text As String

    ' Background Character
    With Tex_Main_Win(1)
        RenderTexture Tex_Main_Win(1), Main_Win(2).Left, Main_Win(2).Top, 0, 0, .Width, .Height, .Width, .Height
    End With
    
    ' Informações Personagem
    Call RenderText(Font_Georgia, Trim$(GetPlayerName(MyIndex)), Main_Win(2).Left + 63, Main_Win(2).Top + 38, White, 0, False)
    Call RenderText(Font_Georgia, Trim$(Class(GetPlayerClass(MyIndex)).Name), Main_Win(2).Left + 63, Main_Win(2).Top + 58, White, 0, False)
    Call RenderText(Font_Georgia, Trim$(GetPlayerLevel(MyIndex)), Main_Win(2).Left + 54, Main_Win(2).Top + 78, White, 0, False)
    Call RenderText(Font_Georgia, GetPlayerVital(MyIndex, HP) & "/" & GetPlayerMaxVital(MyIndex, HP), Main_Win(2).Left + 86, Main_Win(2).Top + 98, White, 0, False)
    Call RenderText(Font_Georgia, GetPlayerVital(MyIndex, MP) & "/" & GetPlayerMaxVital(MyIndex, MP), Main_Win(2).Left + 72, Main_Win(2).Top + 118, White, 0, False)
    Call RenderText(Font_Georgia, Trim$(GetPlayerExp(MyIndex)) & "/" & GetPlayerNextLevel(MyIndex), Main_Win(2).Left + 93, Main_Win(2).Top + 138, White, 0, False)
    
    ' Status
    Call RenderText(Font_Georgia, Trim$(GetPlayerPOINTS(MyIndex)), Main_Win(2).Left + 117, Main_Win(2).Top + 162, White, 0, False)
    Call RenderText(Font_Georgia, Trim$(GetPlayerStat(MyIndex, Strength)) & " ¦02+ " & Trim$(GetPlayerStat2(MyIndex, Strength)), Main_Win(2).Left + 62, Main_Win(2).Top + 182, White, 0, False)
    Call RenderText(Font_Georgia, Trim$(GetPlayerStat(MyIndex, Endurance)) & " ¦02+ " & Trim$(GetPlayerStat2(MyIndex, Endurance)), Main_Win(2).Left + 97, Main_Win(2).Top + 202, White, 0, False)
    Call RenderText(Font_Georgia, Trim$(GetPlayerStat(MyIndex, Agility)) & " ¦02+ " & Trim$(GetPlayerStat2(MyIndex, Agility)), Main_Win(2).Left + 85, Main_Win(2).Top + 222, White, 0, False)
    Call RenderText(Font_Georgia, Trim$(GetPlayerStat(MyIndex, Intelligence)) & " ¦02+ " & Trim$(GetPlayerStat2(MyIndex, Intelligence)), Main_Win(2).Left + 100, Main_Win(2).Top + 242, White, 0, False)
    Call RenderText(Font_Georgia, Trim$(GetPlayerStat(MyIndex, Willpower)) & " ¦02+ " & Trim$(GetPlayerStat2(MyIndex, Willpower)), Main_Win(2).Left + 83, Main_Win(2).Top + 262, White, 0, False)

    DrawMaestriaDx8
    
End Sub

Public Sub DrawPetControlDx8()
Dim i As Long, ItemNum As Long
Dim sRECT As RECT

    ' Background Character
    With Tex_Main_Win(34)
        RenderTexture Tex_Main_Win(34), Main_Win(23).Left, Main_Win(23).Top, 0, 0, .Width, .Height, .Width, .Height
    End With

End Sub

Public Sub DrawDropList()
Dim i As Long, ItemNum As Long
Dim sRECT As RECT
Dim Text As String
Dim Text2 As String
Dim Text3 As String
Dim AB As Long, BC As Long

'Evitar Problemas
If LootPage = 0 Then LootPage = 1

    If Main_Win(25).Visible = True Then
        With Tex_Main_Win(37)
            RenderTexture Tex_Main_Win(37), Main_Win(25).Left, Main_Win(25).Top, 0, 0, .Width, .Height, .Width, .Height
        End With
    
        Text2 = LootPage & "/51"
        Call RenderText(Font_Georgia, Text2, Main_Win(25).Left + 60 - (getWidth(Font_Georgia, Text2) / 2), Main_Win(25).Top + 91, White, 0, False)
        
        BC = LootPage * 5
        AB = BC - 4
        
        For i = AB To BC
        
            If ListLoot(i).id > 0 Then
                Text3 = Trim$(MapDrop(ListLoot(i).id).LootName)
                'If Text3 = vbNullString Then Text3 = "Null"
            
                If LootSelect = i Then
                    Call RenderText(Font_Georgia, Trim$(Text3), Main_Win(25).Left + 60 - (getWidth(Font_Georgia, Trim$(Text3)) / 2), Main_Win(25).Top + 10 + ((i - 1) * 15 - ((LootPage - 1) * 75)), Yellow, 0, False)
                Else
                    Call RenderText(Font_Georgia, Trim$(Text3), Main_Win(25).Left + 60 - (getWidth(Font_Georgia, Trim$(Text3)) / 2), Main_Win(25).Top + 10 + ((i - 1) * 15 - ((LootPage - 1) * 75)), White, 0, False)
                End If
            End If
            
        Next
    
    End If
    
End Sub

Public Sub DrawDropNpc()
Dim i As Long, ItemNum As Long
Dim sRECT As RECT
Dim Text As String, Text2 As String, Text3 As String
Dim AB As Byte, BC As Byte

    Text = Trim$(MonsterLoot.LootName)
    If Text = vbNullString Then Text = "Null"

    ' Background Character
    With Tex_Main_Win(36)
        RenderTexture Tex_Main_Win(36), Main_Win(24).Left, Main_Win(24).Top, 0, 0, .Width, .Height, .Width, .Height
    End With
    
    Call RenderText(Font_Georgia, Text, Main_Win(24).Left + 105 - (getWidth(Font_Georgia, Text) / 2), Main_Win(24).Top + 5, White, 0, False)
        
    For i = 1 To 10
        If MonsterLoot.Itens(i) > 0 Then
            With sRECT
                .Top = Main_Win(24).Top + 36 + ((36) * ((i - 1) \ 5))
                .Bottom = .Top + PIC_Y
                .Left = Main_Win(24).Left + 18 + ((36) * (((i - 1) Mod 5)))
                .Right = .Left + PIC_X
            End With
            
            RenderTexture Tex_Item(Item(MonsterLoot.Itens(i)).pic), sRECT.Left, sRECT.Top, 0, 0, 32, 32, 32, 32
            
            If MonsterLoot.value(i) > 1 Then
                Text = Format$(ConvertCurrency(Str(MonsterLoot.value(i))), "#,###,###,###")
                Call RenderText(Font_Georgia, Text, sRECT.Left, sRECT.Top + 18, White, 0, False)
            End If
            
        End If
    Next

End Sub

Public Sub DrawMaestriaDx8()
Dim sRECT As RECT
Dim MaestriaWitdh(1 To 6) As Long
Dim Text As String
Dim WLeft As Long, WTop As Long

WLeft = 54
WTop = 107

    ' Background Character
    With Tex_Main_Win(38)
        RenderTexture Tex_Main_Win(38), 54, 107, 0, 0, .Width, .Height, .Width, .Height
    End With

' Barras Maestria
    With Tex_Main_Win(12) ' Fisico
        MaestriaWitdh(1) = (Player(MyIndex).Maestria(1).Exp / .Width) / (GetPlayerMaestriaExp(1) / .Width) * .Width
        RenderTexture Tex_Main_Win(12), WLeft + 28, WTop + 52, 0, 0, MaestriaWitdh(1), .Height, MaestriaWitdh(1), .Height
    End With
    
    With Tex_Main_Win(12) ' Distancia
        MaestriaWitdh(2) = (Player(MyIndex).Maestria(2).Exp / .Width) / (GetPlayerMaestriaExp(2) / .Width) * .Width
        RenderTexture Tex_Main_Win(12), WLeft + 150, WTop + 52, 0, 0, MaestriaWitdh(2), .Height, MaestriaWitdh(2), .Height
    End With
    
    With Tex_Main_Win(12) ' Magia
        MaestriaWitdh(3) = (Player(MyIndex).Maestria(3).Exp / .Width) / (GetPlayerMaestriaExp(3) / .Width) * .Width
        RenderTexture Tex_Main_Win(12), WLeft + 28, WTop + 88, 0, 0, MaestriaWitdh(3), .Height, MaestriaWitdh(3), .Height
    End With
    
    With Tex_Main_Win(12) ' Bloqueio
        MaestriaWitdh(4) = (Player(MyIndex).Maestria(4).Exp / .Width) / (GetPlayerMaestriaExp(4) / .Width) * .Width
        RenderTexture Tex_Main_Win(12), WLeft + 150, WTop + 88, 0, 0, MaestriaWitdh(4), .Height, MaestriaWitdh(4), .Height
    End With
    
    ' Informação Maestria
    Text = "Level: " & Trim$(Player(MyIndex).Maestria(1).Level)
    Call RenderText(Font_Georgia, Text, (WLeft + 81) - ((EngineGetTextWidth(Font_Georgia, Text)) / 2), WTop + 50, White, 0, False)
    
    Text = "Level: " & Trim$(Player(MyIndex).Maestria(2).Level)
    Call RenderText(Font_Georgia, Text, (WLeft + 203) - ((EngineGetTextWidth(Font_Georgia, Text)) / 2), WTop + 50, White, 0, False)
    
    Text = "Level: " & Trim$(Player(MyIndex).Maestria(3).Level)
    Call RenderText(Font_Georgia, Text, (WLeft + 81) - ((EngineGetTextWidth(Font_Georgia, Text)) / 2), WTop + 87, White, 0, False)
    
    Text = "Level: " & Trim$(Player(MyIndex).Maestria(4).Level)
    Call RenderText(Font_Georgia, Text, (WLeft + 203) - ((EngineGetTextWidth(Font_Georgia, Text)) / 2), WTop + 87, White, 0, False)
    
    Text = "Pontos: " & Player(MyIndex).StPoints
    Call RenderText(Font_Georgia, Text, (Main_Win(3).Left + 65) - ((EngineGetTextWidth(Font_Georgia, Text)) / 2), Main_Win(3).Top + 388, White, 0, False)

End Sub

Public Sub DrawCraftItem()
Dim i As Long
Dim ItemName As String, ItemNum As Long
Dim ItemSName As String, Chance As String, Custo As String, ItemV As String, ItemSValue As String
Dim PA As Integer, PB As Integer

    ' Fundo Janela
    With Tex_Main_Win(42)
        RenderTexture Tex_Main_Win(42), Main_Win(27).Left, Main_Win(27).Top, 0, 0, .Width, .Height, .Width, .Height
    End With
    
    ' Configurar Página
    If CraftPage = 0 Then CraftPage = 1
    If CrafterSelectSlot = 0 Then CrafterSelectSlot = 1
    
    ' Carregar Itens da Página
    PB = CraftPage * 5
    PA = PB - 4
    
    ' Items Disponiveis para Criação
    For i = PA To PB
        If CraftControl(i).Resultado > 0 Then
            ItemNum = CraftControl(i).Resultado
            ItemName = Trim$(Item(ItemNum).Name)
            RenderTexture Tex_Item(Item(ItemNum).pic), Main_Win(27).Left + 23, Main_Win(27).Top + 34 + ((i - PA) * 41), 0, 0, 32, 32, 32, 32
        
            If i = CrafterSelectSlot Then
                RenderText Font_Georgia, ItemName, Main_Win(27).Left + 118 - (getWidth(Font_Georgia, ItemName) / 2), Main_Win(27).Top + 40 + ((i - PA) * 41), Yellow, 0
            Else
                RenderText Font_Georgia, ItemName, Main_Win(27).Left + 118 - (getWidth(Font_Georgia, ItemName) / 2), Main_Win(27).Top + 40 + ((i - PA) * 41), White, 0
            End If
        
        End If
    Next
    
    ' Item Selecionado
    If CrafterSelectSlot > 0 Then
        If CraftControl(CrafterSelectSlot).Resultado > 0 Then
        
            With CraftControl(CrafterSelectSlot)
            ItemSName = Trim$(Item(.Resultado).Name)
            Chance = "Chance: " & .Chance & "%"
            Custo = "Ouro: " & .Cost
            ItemSValue = "Qnt: " & .ResultadoValue
            RenderText Font_Georgia, ItemSName, Main_Win(27).Left + 283 - (getWidth(Font_Georgia, ItemSName) / 2), Main_Win(27).Top + 33, White, 0
            RenderText Font_Georgia, ItemSValue, Main_Win(27).Left + 283 - (getWidth(Font_Georgia, ItemSValue) / 2), Main_Win(27).Top + 48, White, 0
            RenderTexture Tex_Item(Item(CraftControl(CrafterSelectSlot).Resultado).pic), Main_Win(27).Left + 193, Main_Win(27).Top + 34, 0, 0, 32, 32, 32, 32
        
                For i = 1 To 8
                    If .Item(i) > 0 Then
                        
                        If .ItemUnique(i) = False Then
                            ItemV = CCPItemValue(i) & "/" & .ItemValue(i)
                        Else
                            ItemV = "*"
                        End If
                        
                        If i < 5 Then
                            RenderTexture Tex_Item(Item(.Item(i)).pic), Main_Win(27).Left + 192 + ((i - 1) * 39), Main_Win(27).Top + 77, 0, 0, 32, 32, 32, 32
                            RenderText Font_Georgia, ItemV, Main_Win(27).Left + 207 + ((i - 1) * 39) - (getWidth(Font_Georgia, ItemV) / 2), Main_Win(27).Top + 109, White, 0
                        Else
                            RenderTexture Tex_Item(Item(.Item(i)).pic), Main_Win(27).Left + 192 + ((i - 5) * 39), Main_Win(27).Top + 126, 0, 0, 32, 32, 32, 32
                            RenderText Font_Georgia, ItemV, Main_Win(27).Left + 207 + ((i - 5) * 39) - (getWidth(Font_Georgia, ItemV) / 2), Main_Win(27).Top + 158, White, 0
                        End If
                    End If
                Next
            
            
            RenderText Font_Georgia, Chance, Main_Win(27).Left + 266 - (getWidth(Font_Georgia, Chance) / 2), Main_Win(27).Top + 180, White, 0
            RenderText Font_Georgia, Custo, Main_Win(27).Left + 266 - (getWidth(Font_Georgia, Custo) / 2), Main_Win(27).Top + 200, White, 0
                    
            End With
        End If
    End If
    
    'For I = 1 To 4
    '    RenderTexture Tex_Item(15 + I), Main_Win(27).Left + 192 + ((I - 1) * 39), Main_Win(27).Top + 77, 0, 0, 32, 32, 32, 32
    '    RenderText Font_Georgia, ItemV, Main_Win(27).Left + 207 + ((I - 1) * 39) - (getWidth(Font_Georgia, ItemV) / 2), Main_Win(27).Top + 109, White, 0
    'Next
    
    'For I = 5 To 8
    '    RenderTexture Tex_Item(35 + I), Main_Win(27).Left + 192 + ((I - 5) * 39), Main_Win(27).Top + 126, 0, 0, 32, 32, 32, 32
    '    RenderText Font_Georgia, ItemV, Main_Win(27).Left + 207 + ((I - 5) * 39) - (getWidth(Font_Georgia, ItemV) / 2), Main_Win(27).Top + 158, White, 0
    'Next

    RenderText Font_Georgia, CraftPage & "/" & CraftMaxPage, Main_Win(27).Left + 85, Main_Win(27).Top + 237, White, 0
    
End Sub

Public Sub DrawSpellsDx8()
Dim i As Long, SpellNum As Long
Dim SpellIcon As Long
Dim sRECT As RECT, n As Long
Dim MaestriaWitdh(1 To 6) As Long
Dim Text As String
Dim AB(1 To 3) As Long, Page As Byte
Dim PlayerClass As Byte, Color As Byte

    PlayerClass = GetPlayerClass(MyIndex)

    ' Fundo Janela
    With Tex_Main_Win(2)
        RenderTexture Tex_Main_Win(2), Main_Win(3).Left, Main_Win(3).Top, 0, 0, .Width, .Height, .Width, .Height
    End With
    
    ' Fundo Talentos
    With Tex_Talentos(PlayerClass)
        RenderTexture Tex_Talentos(PlayerClass), Main_Win(3).Left + 14, Main_Win(3).Top + 30, 0, 0, .Width, .Height, .Width, .Height
    End With
    
    Select Case PlayerClass
    Case 1 ' Kina
        AB(1) = 1: AB(2) = 42: AB(3) = 1
    Case Else
        AB(1) = (42 * (PlayerClass - 1)) + 1
        AB(2) = AB(1) + 41
        AB(3) = (PlayerClass - 1) * 43
    End Select
    
    ' Guerreiro
    For i = AB(1) To AB(2)
        
    If SkillTree(i).Habilidade(1) > 0 Then
    
        ' Icone
        SpellIcon = SkillTree(i).Icon
        If SpellIcon > 0 And SpellIcon <= NumSpellIcons Then
            With sRECT
                .Top = Main_Win(3).Top + 36 + ((55) * ((i - AB(3)) \ 7))
                .Bottom = .Top + PIC_Y
                .Left = Main_Win(3).Left + 53 + ((68) * (((i - 1) Mod 7)))
                .Right = .Left + PIC_X
            End With
            
        If Player(MyIndex).SkillTree(i).Level = 0 Then
                RenderTexture Tex_SpellIcon(SpellIcon), sRECT.Left, sRECT.Top, 32, 0, 32, 32, 32, 32
            Else
                RenderTexture Tex_SpellIcon(SpellIcon), sRECT.Left, sRECT.Top, 0, 0, 32, 32, 32, 32
            End If
        End If
        
        ' Setas
        If Player(MyIndex).SkillTree(i).Level > 0 Then
            With Tex_Main_Win(39)
                With sRECT
                    .Top = Main_Win(3).Top + 72 + ((55) * ((i - AB(3)) \ 7))
                    .Bottom = .Top + PIC_Y
                    .Left = Main_Win(3).Left + 40 + ((68) * (((i - 1) Mod 7)))
                    .Right = .Left + PIC_X
                End With
            
                RenderTexture Tex_Main_Win(39), sRECT.Left, sRECT.Top, 0, 0, .Width, .Height, .Width, .Height
            End With
        End If
        
        ' Texto
        If Talentos.UsePSkill(i) > 0 Then
            Text = Talentos.LvlSkill(i) & "/" & Player(MyIndex).SkillTree(i).Level + Talentos.UsePSkill(i)
            Color = Yellow
        Else
            Text = Talentos.LvlSkill(i) & "/" & Player(MyIndex).SkillTree(i).Level
            Color = White
        End If
        
        With sRECT
            .Top = Main_Win(3).Top + 69 + ((55) * ((i - AB(3)) \ 7))
            .Bottom = .Top + PIC_Y
            .Left = Main_Win(3).Left + 67 - (getWidth(Font_Georgia, Text) / 2) + ((68) * (((i - 1) Mod 7)))
            .Right = .Left + PIC_X
        End With
        
        RenderText Font_Georgia, Text, sRECT.Left, sRECT.Top, Color, 0
    
    End If
    
    Next
    
    'Meus Pontos
    Text = Player(MyIndex).StPoints - Talentos.UsePoints
    If Talentos.UsePoints > 0 Then Color = Yellow Else Color = White
    RenderText Font_Georgia, Text, Main_Win(3).Left + 100, Main_Win(3).Top + 371, Color, 0
    
End Sub

Public Function IsInTileView(ByVal TileX As Long, ByVal TileY As Long)
Dim x As Long, y As Long

If TileX >= TileView.Left Then
    If TileX <= TileView.Right Then
        If TileY >= TileView.Top Then
            If TileY <= TileView.Bottom Then
                IsInTileView = True
            End If
        End If
    End If
End If

End Function

Public Sub DrawMiniMap()
Dim Rec As RECT, i As Long
Dim TopC As Long, LeftC As Long
Dim x As Long, y As Long

    'Cor
    TopC = 1
    LeftC = 0
    
    With Rec
         .Top = TopC * (Tex_MiniMap.Height / 4)
         .Bottom = Rec.Top + 4
         .Left = LeftC * 4
         .Right = Rec.Left + 4
     End With
    
    ' Defini-lo no minimap
    'RenderTexture Tex_MiniMap, 200, 200, Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top   ', D3DColorARGB(100, 0, 0, 0)

    ' Definir "Chão"
    
End Sub

Public Sub DrawGlobalDescItemDx8()
Dim i As Long, ItemNum As Long, ItemValue As Long, ItemRefine As Long
Dim sRECT As RECT, x As Long, y As Long
Dim ItemInfo(1 To 20) As String
Dim ItemName As String, ItemType As String
Dim ItemRarity As Byte
Dim ItemRariStr As String
Dim ElementStr As String
Dim BackNum As Byte, ShopText As String

    ' Vamo funcionar sem crashar? Brigadu ;3!
    If InGame = False Then Exit Sub
    If DragInvItemDx8 > 0 Then Exit Sub
    
    ' Setar Item
    ItemNum = DescItemSystem.Item
    ItemValue = DescItemSystem.value
    ItemRefine = DescItemSystem.Refine
    
    ' Evitar OverFlow
    If ItemNum = 0 Or ItemNum > MAX_ITEMS Then Exit Sub
    
    ' Raridade
    Select Case Item(ItemNum).Rarity
        Case 1: ItemRarity = BrightGreen: ItemRariStr = "[Raro]"
        Case 2: ItemRarity = BrightBlue: ItemRariStr = "[Muito Raro]"
        Case 3: ItemRarity = Yellow: ItemRariStr = "[Épico]"
        Case 4: ItemRarity = Pink: ItemRariStr = "[Lendario]"
        Case 5: ItemRarity = Orange: ItemRariStr = "[Único]"
        Case Else: ItemRarity = White: ItemRariStr = "[Comum]"
    End Select

    ' Pegar Valores
    If ItemRefine > 0 Then
            
        Select Case Item(ItemNum).Type
        Case ITEM_TYPE_OUTFIT
            ItemName = Trim$(Item(ItemNum).Name) & " [" & Trim$(Outfit(ItemRefine).Name) & "]"
        Case ITEM_TYPE_CONSUME, ITEM_TYPE_CURRENCY, ITEM_TYPE_RESET, ITEM_TYPE_TORCH
            ItemName = Trim$(Item(ItemNum).Name)
        Case ITEM_TYPE_SPELL
            ItemName = Trim$(Item(ItemNum).Name) & " [" & Trim$(Spell(ItemRefine).Name) & "]"
        Case ITEM_TYPE_VIP
            ItemName = Trim$(Item(ItemNum).Name) & " " & ItemRefine & " Dia(s)"
        Case Else
            ItemName = Trim$(Item(ItemNum).Name) & " +" & ItemRefine
        End Select
    
    Else
        ItemName = Trim$(Item(ItemNum).Name)
    End If
    
    Select Case Item(ItemNum).Type
    Case 0: ItemType = "Indefinido"
    Case 1
    BackNum = 19
    Select Case Item(ItemNum).WeaponCategory
        Case 1: ItemType = "Espada"
        Case 2: ItemType = "Espada Duas Mãos"
        Case 3: ItemType = "Lança"
        Case 4: ItemType = "Arco"
        Case 5: ItemType = "Besta"
        Case 6: ItemType = "Machado"
        Case 7: ItemType = "Machado Duas Mãos"
        Case 8: ItemType = "Cetro"
        Case 9: ItemType = "Cajado"
        Case 10: ItemType = "Orbe"
        Case 11: ItemType = "Adaga"
        Case 12: ItemType = "Adaga Dupla"
        Case 13: ItemType = "Garra"
        Case 14: ItemType = "Punho"
        Case 15: ItemType = "Livro"
        Case 16: ItemType = "Porrete"
        Case 17: ItemType = "Porrete Duas Mãos"
        Case 18: ItemType = "Chicote"
        Case 19: ItemType = "Instrumento"
        Case 20: ItemType = "Foice"
    Case Else
        ItemType = "Indefinido"
    End Select
        
    Case 2: ItemType = "Armadura": BackNum = 19
    Case 3: ItemType = "Elmo": BackNum = 19
    Case 4: ItemType = "Escudo": BackNum = 19
    Case 5: ItemType = "Consumivel": BackNum = 32
    Case 6: ItemType = "Chave": BackNum = 32
    Case 7, 15: ItemType = "Qtd: " & ItemValue: BackNum = 32
    Case 8: ItemType = "Habilidade": BackNum = 32
    Case 9: ItemType = "Botas": BackNum = 19
    Case 10: ItemType = "Anél/Pet": BackNum = 19
    Case 11: ItemType = "Amuleto": BackNum = 19
    Case 12: ItemType = "Calça": BackNum = 19
    Case 13: ItemType = "Resetador": BackNum = 32
    Case 14: ItemType = "Roupa": BackNum = 32
    'Case 15: ItemType = "Tocha": BackNum = 32
    Case 16: ItemType = "VIP": BackNum = 32
    Case Else: ItemType = "Indefinido": BackNum = 32
    End Select
    
    x = GlobalX + 25
    y = GlobalY - 40
    
    ' Fundo
    With Tex_Main_Win(BackNum)
        RenderTexture Tex_Main_Win(BackNum), x, y, 0, 0, .Width, .Height, .Width, .Height, -1
    End With
    
    If Item(ItemNum).pic > 0 Then RenderTexture Tex_Item(Item(ItemNum).pic), x + 13, y + 20, 0, 0, 32, 32, 32, 32 ' SpellIcon
    If ItemName <> vbNullString Then Call RenderText(Font_Georgia, Trim$(ItemName), x + 131 - (EngineGetTextWidth(Font_Georgia, Trim$(ItemName)) / 2), y + 19, ItemRarity, False)
    If ItemType <> vbNullString Then Call RenderText(Font_Georgia, Trim$(ItemType), x + 131 - (EngineGetTextWidth(Font_Georgia, Trim$(ItemType)) / 2), y + 39, ItemRarity, False)
    
    ' Descrição
    Call RenderText(Font_Georgia, WordWrap(Font_Georgia, Trim$(Item(ItemNum).Desc), 173), x + 20, y + 63, White, 0)

    With Tex_Main_Win(29)
        If Item(ItemNum).ClassRest(1) = False Then: RenderTexture Tex_Main_Win(29), x + 12, y + 164, 0, 0, 16, 16, 16, 16, -1
        If Item(ItemNum).ClassRest(2) = False Then: RenderTexture Tex_Main_Win(29), x + 30, y + 164, 0, 16, 16, 16, 16, 16, -1
        If Item(ItemNum).ClassRest(3) = False Then: RenderTexture Tex_Main_Win(29), x + 48, y + 164, 0, 32, 16, 16, 16, 16, -1
        If Item(ItemNum).ClassRest(4) = False Then: RenderTexture Tex_Main_Win(29), x + 66, y + 164, 0, 48, 16, 16, 16, 16, -1
        If Item(ItemNum).ClassRest(5) = False Then: RenderTexture Tex_Main_Win(29), x + 84, y + 164, 0, 64, 16, 16, 16, 16, -1
        If Item(ItemNum).ClassRest(6) = False Then: RenderTexture Tex_Main_Win(29), x + 102, y + 164, 0, 80, 16, 16, 16, 16, -1
    End With
    
    ' Peso
    Call RenderText(Font_Georgia, "Peso: " & Trim$(Item(ItemNum).Peso), x + 164 - (EngineGetTextWidth(Font_Georgia, "Peso: " & Trim$(Item(ItemNum).Peso)) / 2), y + 163, White, 0)

    ' Loja
    If InShop > 0 Then
    
        With Tex_Main_Win(41)
            RenderTexture Tex_Main_Win(41), x + 1, y - 20, 0, 0, .Width, .Height, .Width, .Height, -1
        End With
        
        If DescItemSystem.ShopSlot > 0 Then
            If Shop(InShop).TradeItem(DescItemSystem.ShopSlot).CostItem > 0 Then
            
                ShopText = "Preço: " & Shop(InShop).TradeItem(DescItemSystem.ShopSlot).CostValue & " " & Trim$(Item(Shop(InShop).TradeItem(DescItemSystem.ShopSlot).CostItem).Name)
                RenderText Font_Georgia, ShopText, x + 110 - (getWidth(Font_Georgia, ShopText) / 2), y - 20, White, 0
            
            End If
        End If
        
        If DescItemSystem.SInvSlot > 0 Then
            If PlayerInv(DescItemSystem.SInvSlot).num > 0 Then
            
                If Item(PlayerInv(DescItemSystem.SInvSlot).num).Price > 0 Then
                
                    If Item(PlayerInv(DescItemSystem.SInvSlot).num).Price * (Shop(InShop).BuyRate / 100) >= 1 Then
                        ShopText = "Preço: " & Int(Item(PlayerInv(DescItemSystem.SInvSlot).num).Price * (Shop(InShop).BuyRate / 100)) & " Ouro"
                    Else
                        ShopText = "Não pode ser vendido"
                    End If
                Else
                        ShopText = "Não pode ser vendido"
                End If
                    RenderText Font_Georgia, ShopText, x + 110 - (getWidth(Font_Georgia, ShopText) / 2), y - 20, White, 0
                
            End If
        End If
        
    End If

End Sub

Public Sub DrawGlobalDescItemDx8P2()
Dim i As Long, ItemNum As Long, ItemValue As Long, ItemRefine As Long
Dim sRECT As RECT, x As Long, y As Long
Dim ItemInfo(1 To 20) As String
Dim ItemName As String, ItemType As String
Dim ItemRarity As Byte
Dim ItemRariStr As String
Dim ElementStr As String
Dim BackNum As Byte
Dim DmgItem As Long, DefItem As Long
Dim DefAdd As Long

    ' Vamo funcionar sem crashar? Brigadu ;3!
    If InGame = False Then Exit Sub
    If DragInvItemDx8 > 0 Then Exit Sub

    ' Setar Item
    ItemNum = DescItemSystem.Item
    ItemValue = DescItemSystem.value
    ItemRefine = DescItemSystem.Refine
    
    ' Evitar OverFlow
    If ItemNum = 0 Or ItemNum > MAX_ITEMS Then Exit Sub
    
    ' Item que não precisa destas informações
    Select Case Item(ItemNum).Type
    Case 5, 6, 7, 8, 13, 15
        DrawGlobalDescItemDx8
        Exit Sub
    End Select
    
    ' Raridade
    Select Case Item(ItemNum).Rarity
        Case 1: ItemRarity = BrightGreen: ItemRariStr = "[Raro]"
        Case 2: ItemRarity = BrightBlue: ItemRariStr = "[Muito Raro]"
        Case 3: ItemRarity = Yellow: ItemRariStr = "[Épico]"
        Case 4: ItemRarity = Pink: ItemRariStr = "[Lendario]"
        Case 5: ItemRarity = Orange: ItemRariStr = "[Único]"
        Case Else: ItemRarity = White: ItemRariStr = "[Comum]"
    End Select
    
    ' Pegar Valores
    If ItemRefine > 0 Then
        ItemName = Trim$(Item(ItemNum).Name) & " +" & ItemRefine
    Else
        ItemName = Trim$(Item(ItemNum).Name)
    End If
    
    DmgItem = Item(ItemNum).Data2
    DefItem = Item(ItemNum).Data4
    
    If DefItem > 0 Then
        DefItem = DefItem + ((Item(ItemNum).Data4 * (ItemRefine * 8.5)) / 100)
    End If
    
    If DmgItem > 0 Then
        DmgItem = DmgItem + ((Item(ItemNum).Data2 * (ItemRefine * 20.8)) / 100)
    End If
    
    ' Stats
    ItemInfo(2) = "Dano: " & DmgItem
    ItemInfo(3) = "For: +" & Item(ItemNum).Add_Stat(1)
    ItemInfo(4) = "Vit: +" & Item(ItemNum).Add_Stat(2)
    ItemInfo(5) = "Int: +" & Item(ItemNum).Add_Stat(3)
    ItemInfo(6) = "Agi: +" & Item(ItemNum).Add_Stat(4)
    ItemInfo(7) = "Des: +" & Item(ItemNum).Add_Stat(5)
    
    ' Stats Req
    ItemInfo(8) = "Defesa: " & DefItem
    ItemInfo(9) = "For: " & Item(ItemNum).Stat_Req(1)
    ItemInfo(10) = "Vit: " & Item(ItemNum).Stat_Req(2)
    ItemInfo(11) = "Int: " & Item(ItemNum).Stat_Req(3)
    ItemInfo(12) = "Agi: " & Item(ItemNum).Stat_Req(4)
    ItemInfo(13) = "Des: " & Item(ItemNum).Stat_Req(5)
    ItemInfo(14) = "Peso: " & Item(ItemNum).Peso
    
    x = GlobalX + 25
    y = GlobalY - 40
    
    With Tex_Main_Win(3)
        RenderTexture Tex_Main_Win(3), x, y, 0, 0, .Width, .Height, .Width, .Height, -1
    End With
    
    RenderTexture Tex_Item(Item(ItemNum).pic), x + 17, y + 20, 0, 0, 32, 32, 32, 32
    
    Call RenderText(Font_Georgia, Trim$(ItemName), x + 133 - ((EngineGetTextWidth(Font_Georgia, Trim$(ItemName)) / 2)), y + 19, ItemRarity, 0, False)
    Call RenderText(Font_Georgia, Trim$(ItemInfo(1)), x + 133 - ((EngineGetTextWidth(Font_Georgia, Trim$(ItemInfo(1))) / 2)), y + 39, ItemRarity, 0, False)
    Call RenderText(Font_Georgia, Trim$(ItemInfo(2)), x + 64 - ((EngineGetTextWidth(Font_Georgia, Trim$(ItemInfo(2))) / 2)), y + 59, White, 0, False)
    Call RenderText(Font_Georgia, Trim$(ItemInfo(8)), x + 166 - ((EngineGetTextWidth(Font_Georgia, Trim$(ItemInfo(8))) / 2)), y + 59, White, 0, False)
    
    Call RenderText(Font_Georgia, "Atributos:", x + 33, y + 79, BrightGreen, 0, False)
    Call RenderText(Font_Georgia, "Requisitos:", x + 133, y + 79, BrightRed, 0, False)
    
    For i = 1 To Stats.Stat_Count - 1
        Call RenderText(Font_Georgia, Trim$(ItemInfo(3 + i - 1)), x + 64 - ((EngineGetTextWidth(Font_Georgia, Trim$(ItemInfo(3 + i - 1))) / 2)), y + 119 + ((i - 1) * 20), BrightGreen, 0, False)
    If GetPlayerStat(MyIndex, i) < Item(ItemNum).Stat_Req(i) Then
        Call RenderText(Font_Georgia, Trim$(ItemInfo(9 + i - 1)), x + 166 - ((EngineGetTextWidth(Font_Georgia, Trim$(ItemInfo(9 + i - 1))) / 2)), y + 119 + ((i - 1) * 20), BrightRed, 0, False)
    Else
        Call RenderText(Font_Georgia, Trim$(ItemInfo(9 + i - 1)), x + 166 - ((EngineGetTextWidth(Font_Georgia, Trim$(ItemInfo(9 + i - 1))) / 2)), y + 119 + ((i - 1) * 20), White, 0, False)
    End If
    Next
        
    Select Case Item(ItemNum).Elemento
    Case 0: ElementStr = "Neutro"
    Case 1: ElementStr = "Fogo"
    Case 2: ElementStr = "Água"
    Case 3: ElementStr = "Terra"
    Case 4: ElementStr = "Vento"
    Case 5: ElementStr = "Veneno"
    Case 6: ElementStr = "Sagrado"
    Case 7: ElementStr = "Trevas"
    Case Else: ElementStr = "???"
    End Select
    
    Call RenderText(Font_Georgia, ElementStr, x + 64 - ((EngineGetTextWidth(Font_Georgia, ElementStr) / 2)), y + 99, BrightGreen, 0, False)
    
    If GetPlayerLevel(MyIndex) < Item(ItemNum).LevelReq Then
        Call RenderText(Font_Georgia, "Level: " & Item(ItemNum).LevelReq, x + 166 - ((EngineGetTextWidth(Font_Georgia, "Level: " & Item(ItemNum).LevelReq) / 2)), y + 99, BrightRed, 0, False)
    Else
        Call RenderText(Font_Georgia, "Level: " & Item(ItemNum).LevelReq, x + 166 - ((EngineGetTextWidth(Font_Georgia, "Level: " & Item(ItemNum).LevelReq) / 2)), y + 99, White, 0, False)
    End If
    
    With Tex_Main_Win(29)
        If Item(ItemNum).ClassRest(1) = False Then: RenderTexture Tex_Main_Win(29), x + 16, y + 220, 0, 0, 16, 16, 16, 16, -1
        If Item(ItemNum).ClassRest(2) = False Then: RenderTexture Tex_Main_Win(29), x + 34, y + 220, 0, 16, 16, 16, 16, 16, -1
        If Item(ItemNum).ClassRest(3) = False Then: RenderTexture Tex_Main_Win(29), x + 52, y + 220, 0, 32, 16, 16, 16, 16, -1
        If Item(ItemNum).ClassRest(4) = False Then: RenderTexture Tex_Main_Win(29), x + 70, y + 220, 0, 48, 16, 16, 16, 16, -1
        If Item(ItemNum).ClassRest(5) = False Then: RenderTexture Tex_Main_Win(29), x + 88, y + 220, 0, 64, 16, 16, 16, 16, -1
        If Item(ItemNum).ClassRest(6) = False Then: RenderTexture Tex_Main_Win(29), x + 106, y + 220, 0, 80, 16, 16, 16, 16, -1
    End With
    
    ' Peso
    Call RenderText(Font_Georgia, Trim$(ItemInfo(14)), x + 166 - ((EngineGetTextWidth(Font_Georgia, Trim$(ItemInfo(14))) / 2)), y + 219, White, 0, False)
    
End Sub

Public Sub DrawWeaponFrame(ByVal Index As Long, ByVal AttackSpeed As Long)
Dim Rec As RECT
Dim x As Long, y As Long, sprite As Long, Anim As Long
Dim Width As Long, Height As Long

    ' Weapon Frame a ser usada
    If GetPlayerEquipment(Index, Weapon) > 0 Then
        sprite = Item(GetPlayerEquipment(Index, Weapon)).Wframe
    Else
        sprite = 1
    End If
    
    x = ConvertMapX(GetPlayerX(Index) * 32) + Player(Index).xOffset
    y = ConvertMapY(GetPlayerY(Index) * 32) + Player(Index).yOffset - 70

    ' Evitar OverFlow
    If sprite < 1 Or sprite > NumPaperdolls Then Exit Sub
                
    'Controle de animação
    If Player(Index).AttackTimer + (AttackSpeed / 2) - 100 > GetTickCount Then
        Anim = 0
    ElseIf Player(Index).AttackTimer + (AttackSpeed / 2) + 50 > GetTickCount Then
        Anim = 1
    ElseIf Player(Index).AttackTimer + (AttackSpeed / 2) + 100 > GetTickCount Then
        Anim = 2
    ElseIf Player(Index).AttackTimer + (AttackSpeed / 2) + 150 > GetTickCount Then
        Anim = 3
    ElseIf Player(Index).AttackTimer + (AttackSpeed / 2) + 200 > GetTickCount Then
        Anim = 4
    End If
    
    With Rec
        .Top = 0
        .Bottom = .Top + Tex_WFrames(sprite).Height
        .Left = Anim * (Tex_WFrames(sprite).Width / 5)
        .Right = .Left + (Tex_WFrames(sprite).Width / 5)
    End With
    
    'Renderizar na Tela
    RenderTexture Tex_WFrames(sprite), x, y, Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top
        
End Sub

Private Function CanRender() As Boolean
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or _
       Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        HandleDeviceLost
        CanRender = False
        Exit Function
    End If
    
    If frmMain.WindowState = vbMinimized Or GettingMap Then
        CanRender = False
        Exit Function
    End If
    
    CanRender = True
End Function

Private Sub ClearScene()
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorARGB(0, 0, 0, 0), 1#, 0
End Sub

Private Sub PresentScene()
    Dim srcRect As D3DRECT
    With srcRect
        .X1 = 0: .Y1 = 0
        .X2 = frmMain.picScreen.ScaleWidth
        .Y2 = frmMain.picScreen.ScaleHeight
    End With
    
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or _
       Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        HandleDeviceLost
    Else
        Direct3D_Device.Present srcRect, ByVal 0, 0, ByVal 0
        DrawGDI
    End If
End Sub

Private Sub HandleRenderError()
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or _
       Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        HandleDeviceLost
    Else
        If Options.Debug = 1 Then
            HandleError "Render_Graphics", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
            Err.Clear
        End If
        MsgBox "Unrecoverable DX8 error."
        DestroyGame
    End If
End Sub

Private Sub RenderBackground()
    If Map.MapPic > 0 Then
        RenderTexture Tex_BackGround(Map.MapPic), ConvertMapX(0), ConvertMapY(0), _
                      0, 0, Tex_BackGround(Map.MapPic).Width, Tex_BackGround(Map.MapPic).Height, _
                      Tex_BackGround(Map.MapPic).Width, Tex_BackGround(Map.MapPic).Height
    End If
End Sub

Private Sub RenderLowerTiles()
    Dim x As Long, y As Long
    If NumTileSets > 0 Then
        For x = TileView.Left To TileView.Right
            For y = TileView.Top To TileView.Bottom
                If IsValidMapPoint(x, y) Then
                    DrawMapTile x, y
                    If Map.Tile(x, y).Type = TILE_TYPE_CHEST Then DrawChestGraphic x, y
                End If
            Next
        Next
    End If
End Sub

Private Sub RenderDecals()
    Dim i As Long
    For i = 1 To MAX_BYTE
        DrawBlood i
    Next
End Sub

Private Sub RenderCorpses()
    Dim i As Long
    For i = 1 To MAX_MAP_DROPS
        If MapDrop(i).Graphic > 0 Then DrawCorpse i
    Next
End Sub

Private Sub RenderTargetsAndHover()
    Dim i As Long
    ' Target
    If myTarget > 0 Then
        Select Case myTargetType
            Case TARGET_TYPE_PLAYER
                DrawTarget (Player(myTarget).x * 32) + Player(myTarget).xOffset, _
                           (Player(myTarget).y * 32) + Player(myTarget).yOffset
            Case TARGET_TYPE_NPC
                DrawTarget (MapNpc(myTarget).x * 32) + MapNpc(myTarget).xOffset, _
                           (MapNpc(myTarget).y * 32) + MapNpc(myTarget).yOffset
        End Select
    End If
    
    ' Hover
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And Player(i).Map = Player(MyIndex).Map Then
            If CurX = Player(i).x And CurY = Player(i).y Then
                If Not (myTargetType = TARGET_TYPE_PLAYER And myTarget = i) Then
                    DrawHover TARGET_TYPE_PLAYER, i, (Player(i).x * 32) + Player(i).xOffset, (Player(i).y * 32) + Player(i).yOffset
                    DrawPlayerName i
                    THover = i
                End If
            End If
        End If
    Next
    
    For i = 1 To Npc_HighIndex
        If MapNpc(i).num > 0 Then
            If CurX = MapNpc(i).x And CurY = MapNpc(i).y Then
                If Not (myTargetType = TARGET_TYPE_NPC And myTarget = i) Then
                    DrawHover TARGET_TYPE_NPC, i, (MapNpc(i).x * 32) + MapNpc(i).xOffset, (MapNpc(i).y * 32) + MapNpc(i).yOffset
                    DrawNpcName i
                End If
            End If
        End If
    Next
End Sub

Private Sub RenderItems()
    Dim i As Long
    If numitems > 0 Then
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(i).num > 0 Then DrawItem i
        Next
    End If
End Sub

Private Sub RenderMapEvents(Position As Long)
    Dim i As Long
    If Map.CurrentEvents > 0 Then
        For i = 1 To Map.CurrentEvents
            If Map.MapEvents(i).Position = Position Then
                DrawEvent i
            End If
        Next
    End If
End Sub

Private Sub RenderAnimations(Layer As Long)
    Dim i As Long
    If NumAnimations > 0 Then
        For i = 1 To MAX_BYTE
            If AnimInstance(i).Used(Layer) Then DrawAnimation i, Layer
        Next
    End If
End Sub

Private Sub RenderYBasedObjects()
    Dim y As Long, i As Long
    For y = 0 To Map.MaxY
        ' Eventos posicionados por Y
        If Map.CurrentEvents > 0 Then
            For i = 1 To Map.CurrentEvents
                If Map.MapEvents(i).Position = 1 And y = Map.MapEvents(i).y Then
                    DrawEvent i
                End If
            Next
        End If
        
        ' NPCs
        For i = 1 To Npc_HighIndex
            If MapNpc(i).y = y Then
                If MapNpc(i).x >= TileView.Left And MapNpc(i).x <= TileView.Right + 1 And _
                   MapNpc(i).y >= TileView.Top And MapNpc(i).y <= TileView.Bottom + 1 Then
                    If Not ScreenshotMapMode Then DrawNpc i
                End If
            End If
        Next
        
        ' Players
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                If Player(i).y = y Then
                    If Player(i).x + 1 > TileView.Left And Player(i).x < TileView.Right + 1 And _
                       Player(i).y + 1 > TileView.Top And Player(i).y < TileView.Bottom + 1 Then
                        If Not ScreenshotMapMode Then DrawPlayer i
                    End If
                End If
            End If
        Next
        
        ' Recursos
        If NumResources > 0 And Resources_Init And Resource_Index > 0 Then
            For i = 1 To Resource_Index
                If MapResource(i).y = y Then DrawMapResource i
            Next
        End If
    Next
End Sub

Private Sub RenderProjectiles()
    Dim i As Long, x As Long
    
    ' Projéteis de jogadores
    For i = 1 To Player_HighIndex
        For x = 1 To MAX_PLAYER_PROJECTILES
            If Player(i).Projectile(x).pic > 0 Then DrawProjectile i, x
        Next
    Next
    
    ' Projéteis de alvo
    If NumProjectilesTarget > 0 Then DrawProjectileTarget
End Sub

Private Sub RenderUpperTiles()
    Dim x As Long, y As Long
    If NumTileSets > 0 Then
        For x = TileView.Left To TileView.Right
            For y = TileView.Top To TileView.Bottom
                If IsValidMapPoint(x, y) Then DrawMapFringeTile x, y
            Next
        Next
    End If
End Sub

Private Sub RenderBars()
    Dim i As Long
    
    ' Player bars
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
            If IsInTileView(GetPlayerX(i), GetPlayerY(i)) Then DrawPlayerBar i
        End If
    Next
    
    ' NPC bars
    For i = 1 To Npc_HighIndex
        If MapNpc(i).num > 0 And IsInTileView(MapNpc(i).x, MapNpc(i).y) Then
            If MapNpc(i).Vital(1) > 0 Then DrawNpcBar i
        End If
    Next
End Sub

Private Sub RenderLights()
    Dim x As Long, y As Long
    If NumTileSets > 0 Then
        For x = TileView.Left To TileView.Right
            For y = TileView.Top To TileView.Bottom
                If IsValidMapPoint(x, y) And Map.Tile(x, y).Type = TILE_TYPE_LIGHT Then
                    DrawLight x * 32, y * 32, Map.Tile(x, y).Data1, Map.Tile(x, y).Data2, Map.Tile(x, y).Data3, Map.Tile(x, y).Data4
                End If
            Next
        Next
    End If
End Sub

Private Sub RenderWeatherEffects()
    DrawWeather
    DrawFog
    DrawTint
    
    If DrawThunder > 0 Then
        RenderTexture Tex_White, 0, 0, 0, 0, frmMain.picScreen.ScaleWidth, frmMain.picScreen.ScaleHeight, 32, 32, D3DColorRGBA(255, 255, 255, 160)
        DrawThunder = DrawThunder - 1
    End If
    
    If Player(MyIndex).Dead Then DrawDeadBackground
End Sub

Private Sub RenderEditorOverlay()
    Dim x As Long, y As Long
    If InMapEditor Then
        If frmMain.optBlock.value Then
            For x = TileView.Left To TileView.Right
                For y = TileView.Top To TileView.Bottom
                    If IsValidMapPoint(x, y) Then DrawDirection x, y
                Next
            Next
        End If
        DrawTileOutline
        DrawMapAttributes
        If frmMain.optEvent.value Then DrawEvents
    End If
End Sub

Private Sub RenderUI()
    ' Hud e iluminação do jogador
    If HudDx8Visible Then RenderPlayerLight
    
    ' Fade e flash
    If FadeAmount > 0 Then
        RenderTexture Tex_Fade, 0, 0, 0, 0, frmMain.picScreen.ScaleWidth, frmMain.picScreen.ScaleHeight, 32, 32, D3DColorRGBA(255, 255, 255, FadeAmount)
    End If
    
    If FlashTimer > GetTickCount Then
        RenderTexture Tex_White, 0, 0, 0, 0, frmMain.picScreen.ScaleWidth, frmMain.picScreen.ScaleHeight, 32, 32, -1
    End If
    
    ' FPS e Ping
    If BFPS Then
        RenderText Font_Georgia, "FPS: " & CStr(GameFPS), 788, 117, Yellow, 0
        RenderText Font_Georgia, "Ping: " & Ping, 850, 117, Yellow, 0
    End If
    
    ' Coordenadas do jogador
    If BLoc Then
        RenderText Font_Default, "cur x: " & CurX & " y: " & CurY, 350, 1, Yellow, 0
        RenderText Font_Default, "loc x: " & GetPlayerX(MyIndex) & " y: " & GetPlayerY(MyIndex), 350, 15, Yellow, 0
        RenderText Font_Default, " (map " & GetPlayerMap(MyIndex) & ")", 350, 27, Yellow, 0
    End If
    
    ' Mensagens e chat bubbles
    Dim i As Long
    For i = 1 To MAX_BYTE
        If chatBubble(i).active Then DrawChatBubble i
    Next
    
    For i = 1 To Action_HighIndex
        DrawActionMsg i
    Next
    
    ' Nome do mapa
    If Not ScreenshotMapMode Then
        Dim status As String
        status = IIf(Map.Moral = 0, " - [Área de Risco]", " - [Área Segura]")
        RenderText Font_Georgia, Trim$(Map.Name) & status, DrawMapNameX, DrawMapNameY, DrawMapNameColor, 0
    End If
    
    ' Eventos visíveis com nome
    For i = 1 To Map.CurrentEvents
        If Map.MapEvents(i).Visible = 1 And Map.MapEvents(i).ShowName = 1 Then DrawEventName i
    Next
End Sub

Private Sub RenderPlayerBuffs()
    If BuffsAtivos > 0 Then DrawPlayerBuffs
End Sub

Private Sub RenderPlayerLight()
    Dim px As Long, py As Long
    px = ConvertMapX(GetPlayerX(MyIndex) * 32) + Player(MyIndex).xOffset
    py = ConvertMapY(GetPlayerY(MyIndex) * 32) + Player(MyIndex).yOffset

    Select Case Map.LightOff
        Case 0
            ' Sem luz desligada
            RenderTexture Tex_BackGround(1), px - (Tex_BackGround(1).Width / 2) + 20, _
                                         py - (Tex_BackGround(1).Height / 2) - 20, _
                                         0, 0, Tex_BackGround(1).Width, Tex_BackGround(1).Height, _
                                         Tex_BackGround(1).Width, Tex_BackGround(1).Height, _
                                         D3DColorARGB(255, 0, 0, 0)
        Case 1
            If PlayerLight Then
                RenderTexture Tex_BackGround(2), px - (Tex_BackGround(2).Width / 2) + 20, _
                                             py - (Tex_BackGround(2).Height / 2) - 20, _
                                             0, 0, Tex_BackGround(2).Width, Tex_BackGround(2).Height, _
                                             Tex_BackGround(2).Width, Tex_BackGround(2).Height, _
                                             D3DColorARGB(255, 0, 0, 0)
            Else
                RenderTexture Tex_BackGround(3), px - (Tex_BackGround(3).Width / 2) + 20, _
                                             py - (Tex_BackGround(3).Height / 2) - 20, _
                                             0, 0, Tex_BackGround(3).Width, Tex_BackGround(3).Height, _
                                             Tex_BackGround(3).Width, Tex_BackGround(3).Height, _
                                             D3DColorARGB(255, 0, 0, 0)
            End If
    End Select
End Sub

Private Sub RenderHUD()
    If InGame = True And HudDx8Visible = True Then
        DrawScreenHudDx8
    End If
End Sub

' HP / MP
Private Sub DrawHudBar(PosX As Long, PosY As Long, current As Long, max As Long)
    Dim sRECT As RECT, BarWidth As Long
    If current <= 0 Then Exit Sub
    BarWidth = (current / max) * 154
    
    With sRECT
        .Top = IIf(PosY = 28, 0, 6)
        .Left = 0
        .Right = .Left + BarWidth
        .Bottom = .Top + 6
    End With
    
    RenderTexture Tex_Main_Win(7), PosX, PosY, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
End Sub

' EXP
Private Sub DrawHudBarExp(PosX As Long, PosY As Long, Exp As Long, NextLevel As Long)
    Dim sRECT As RECT, BarWidth As Long
    BarWidth = (Exp / NextLevel) * 205
    
    With sRECT
        .Top = 0
        .Left = 0
        .Right = .Left + BarWidth
        .Bottom = .Top + 6
    End With
    
    RenderTexture Tex_Main_Win(6), PosX, PosY, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
End Sub

' Sprite + Hair
Private Sub DrawPlayerSprite(PosX As Long, PosY As Long, SpriteNum As Long, Anim As Byte, HairView As Boolean)
    Dim sRECT As RECT
    
    With sRECT
        .Top = 0
        .Bottom = Tex_Character(SpriteNum).Height / 2
        .Left = Anim * (Tex_Character(SpriteNum).Width / 4)
        .Right = .Left + (Tex_Character(SpriteNum).Width / 4)
    End With
    
    RenderTexture Tex_Character(SpriteNum), PosX, PosY, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top
    
    ' Hair
    If HairView And Player(MyIndex).Hair > 0 And Player(MyIndex).Hair <= NumHairs Then
        RenderTexture Tex_Hair(Player(MyIndex).Hair), PosX, PosY, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, _
                      D3DColorRGBA(Player(MyIndex).HairRGB(1), Player(MyIndex).HairRGB(2), Player(MyIndex).HairRGB(3), 255)
    End If
End Sub

' Addons
Private Sub DrawPlayerAddons(PosX As Long, PosY As Long, ByRef HairView As Boolean)
    Dim OutfitID As Long
    OutfitID = Player(MyIndex).Outfit
    
    ' Checa Addon1
    If Player(MyIndex).Addon(1) Then
        If (Player(MyIndex).Sex = 0 And Outfit(OutfitID).HMAddOff1) Or (Player(MyIndex).Sex = 1 And Outfit(OutfitID).HFAddOff1) Then HairView = False
        RenderTexture Tex_Addon(IIf(Player(MyIndex).Sex = 0, Outfit(OutfitID).MAddon1, Outfit(OutfitID).FAddon1)), PosX + 1, PosY - 2, 0, 0, 32, 32, 32, 32
    End If
    
    ' Checa Addon2
    If Player(MyIndex).Addon(2) Then
        If (Player(MyIndex).Sex = 0 And Outfit(OutfitID).HMAddOff2) Or (Player(MyIndex).Sex = 1 And Outfit(OutfitID).HFAddOff2) Then HairView = False
    End If
End Sub

' Level
Private Sub DrawPlayerLevel(PosX As Long, PosY As Long, Index As Long)
    Dim LevelText As String
    LevelText = "Lvl: " & Format$(GetPlayerLevel(Index), "00")
    RenderText Font_Georgia, LevelText, PosX - EngineGetTextWidth(Font_Georgia, LevelText) / 2, PosY, White, 0, False
End Sub

' HP, MP, EXP texto
Private Sub DrawPlayerStats(HPX As Long, HPY As Long, MPX As Long, MPY As Long, EXPX As Long, EXPY As Long, Index As Long)
    RenderText Font_Georgia, "HP: " & GetPlayerVital(Index, HP) & "/" & GetPlayerMaxVital(Index, HP), HPX - EngineGetTextWidth(Font_Georgia, "HP: " & GetPlayerVital(Index, HP) & "/" & GetPlayerMaxVital(Index, HP)) / 2, HPY, White, 0, False
    RenderText Font_Georgia, "MP: " & GetPlayerVital(Index, MP) & "/" & GetPlayerMaxVital(Index, MP), MPX - EngineGetTextWidth(Font_Georgia, "MP: " & GetPlayerVital(Index, MP) & "/" & GetPlayerMaxVital(Index, MP)) / 2, MPY, White, 0, False
    RenderText Font_Georgia, "EXP: " & GetPlayerExp(Index) & "/" & GetPlayerNextLevel(Index), EXPX - EngineGetTextWidth(Font_Georgia, "EXP: " & GetPlayerExp(Index) & "/" & GetPlayerNextLevel(Index)) / 2, EXPY, White, 0, False
End Sub

' Sem Flechas
Private Sub DrawNoAmmoMessage()
    If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
        If Item(GetPlayerEquipment(MyIndex, Weapon)).Ammo > 0 Then
            If Player(MyIndex).AmmoEquip > 0 And PlayerAmmoEquip = 0 Then
                RenderText Font_Georgia, "Você está sem Flechas!", 464 - EngineGetTextWidth(Font_Georgia, "Você está sem Flechas!") / 2, 240, White, 0, False
            End If
        End If
    End If
End Sub


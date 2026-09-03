Attribute VB_Name = "PLighting"
'/**
' * PLighting - Psu3D Global Light And Fog
' * @description Holds the single directional light and the distance fog every material is shaded against, and exposes the flat shading maths the renderer runs per face. A revision counter is bumped on every change so cached shading tables can detect that they went stale.
' * @author UesleiDev
' * @version 1.0
' * @remarks Fog is quantised into bands so materials can precompute one colour per band instead of blending on every face.
' */

Option Explicit
Option Private Module

'/** @section Module state */

Private m_lightX As Single
Private m_lightY As Single
Private m_lightZ As Single
Private m_ambient As Single
Private m_diffuse As Single

Private m_fogOn As Boolean
Private m_fogStart As Single
Private m_fogEnd As Single
Private m_fogColor As Long
Private m_fogSteps As Long
Private m_fogScale As Single
Private m_fogInvSpan As Single

Private m_revision As Long
Private m_ready As Boolean

'/** @section Lifetime */

'/**
' * @brief Restores the stock lighting and fog setup.
' * @description Called automatically the first time any accessor is touched, so a fresh project renders sensibly without configuration.
' */
Public Sub ResetDefaults()
    m_ready = True

    SetLight 0.38, 0.42, 0.82
    SetLightIntensity 0.46, 0.82
    SetFog 6!, 25!, PCore.ColorPack(160, 185, 210)
    SetFogSteps 32
    m_fogOn = True

    Touch
End Sub

'/** @section Directional light */

'/**
' * @brief Points the global light in a new direction.
' * @param dx The X component of the direction the light travels towards.
' * @param dy The Y component of the direction.
' * @param dz The Z component of the direction.
' * @remarks The vector is normalised internally; a zero length vector is ignored.
' */
Public Sub SetLight(ByVal dx As Single, ByVal dy As Single, ByVal dz As Single)
    Dim mgn As Single

    EnsureReady
    mgn = Sqr(dx * dx + dy * dy + dz * dz)
    If mgn < P_EPSILON Then Exit Sub

    m_lightX = dx / mgn
    m_lightY = dy / mgn
    m_lightZ = dz / mgn

    Touch
End Sub

'/**
' * @brief Sets how much light reaches unlit and lit surfaces.
' * @param ambient The floor brightness applied to every face.
' * @param diffuse How much the light direction adds on top of the ambient floor.
' */
Public Sub SetLightIntensity(ByVal ambient As Single, ByVal diffuse As Single)
    EnsureReady
    m_ambient = ambient
    m_diffuse = diffuse
    Touch
End Sub

'/**
' * @brief Reads the X component of the normalised light direction.
' * @return The direction component.
' */
Public Property Get LightX() As Single
    EnsureReady
    LightX = m_lightX
End Property

'/**
' * @brief Reads the Y component of the normalised light direction.
' * @return The direction component.
' */
Public Property Get LightY() As Single
    EnsureReady
    LightY = m_lightY
End Property

'/**
' * @brief Reads the Z component of the normalised light direction.
' * @return The direction component.
' */
Public Property Get LightZ() As Single
    EnsureReady
    LightZ = m_lightZ
End Property

'/**
' * @brief Reads the ambient floor brightness.
' * @return The ambient term.
' */
Public Property Get Ambient() As Single
    EnsureReady
    Ambient = m_ambient
End Property

'/**
' * @brief Reads the diffuse contribution of the light.
' * @return The diffuse term.
' */
Public Property Get Diffuse() As Single
    EnsureReady
    Diffuse = m_diffuse
End Property

'/** @section Shading */

'/**
' * @brief Computes the brightness multiplier for a surface normal.
' * @param nx The X component of the unit normal.
' * @param ny The Y component of the unit normal.
' * @param nz The Z component of the unit normal.
' * @return The multiplier to apply to the material colour.
' */
Public Function ShadeFactor(ByVal nx As Single, ByVal ny As Single, ByVal nz As Single) As Single
    Dim d As Single

    EnsureReady
    d = nx * m_lightX + ny * m_lightY + nz * m_lightZ
    If d < 0! Then d = 0!

    ShadeFactor = m_ambient + d * m_diffuse
End Function

'/**
' * @brief Applies flat directional shading to a colour.
' * @param baseCol The unlit material colour.
' * @param nx The X component of the unit normal.
' * @param ny The Y component of the unit normal.
' * @param nz The Z component of the unit normal.
' * @return The shaded colour, before fog.
' */
Public Function ShadeColor(ByVal baseCol As Long, ByVal nx As Single, ByVal ny As Single, ByVal nz As Single) As Long
    ShadeColor = PCore.ColorScale(baseCol, ShadeFactor(nx, ny, nz))
End Function

'/**
' * @brief Applies flat shading for one of the six axis-aligned normals.
' * @param baseCol The unlit material colour.
' * @param dirKey The PDirection entry naming the face orientation.
' * @return The shaded colour, before fog.
' */
Public Function ShadeColorByDir(ByVal baseCol As Long, ByVal dirKey As PDirection) As Long
    Dim nx As Single, ny As Single, nz As Single

    DirectionNormal dirKey, nx, ny, nz
    ShadeColorByDir = ShadeColor(baseCol, nx, ny, nz)
End Function

'/**
' * @brief Expands a PDirection entry into its unit normal.
' * @param dirKey The direction key to expand.
' * @param outX Receives the X component.
' * @param outY Receives the Y component.
' * @param outZ Receives the Z component.
' */
Public Sub DirectionNormal(ByVal dirKey As PDirection, ByRef outX As Single, ByRef outY As Single, ByRef outZ As Single)
    outX = 0!
    outY = 0!
    outZ = 0!

    Select Case dirKey
        Case pdPosX: outX = 1!
        Case pdNegX: outX = -1!
        Case pdPosY: outY = 1!
        Case pdNegY: outY = -1!
        Case pdPosZ: outZ = 1!
        Case Else:   outZ = -1!
    End Select
End Sub

'/** @section Distance fog */

'/**
' * @brief Configures the distance fog band.
' * @param startDist The depth at which fog begins to tint geometry.
' * @param endDist The depth at which geometry is fully replaced by the fog colour.
' * @param col The fog colour, which should match the slide background.
' */
Public Sub SetFog(ByVal startDist As Single, ByVal endDist As Single, ByVal col As Long)
    EnsureReady

    If endDist <= startDist Then endDist = startDist + 0.001

    m_fogStart = startDist
    m_fogEnd = endDist
    m_fogColor = col
    m_fogInvSpan = 1! / (m_fogEnd - m_fogStart)
    m_fogScale = m_fogSteps * m_fogInvSpan

    Touch
End Sub

'/**
' * @brief Sets how many discrete bands the fog is quantised into.
' * @param steps The band count, clamped to 1..P_MAX_FOG_STEPS.
' * @remarks More bands look smoother but grow every cached shading table.
' */
Public Sub SetFogSteps(ByVal steps As Long)
    EnsureReady

    m_fogSteps = PCore.ClampLong(steps, 1, P_MAX_FOG_STEPS)
    m_fogScale = m_fogSteps * m_fogInvSpan

    Touch
End Sub

'/**
' * @brief Turns distance fog on or off without losing its settings.
' * @param enabled True to fog geometry by depth, False to render it flat.
' */
Public Sub EnableFog(ByVal enabled As Boolean)
    EnsureReady
    m_fogOn = enabled
    Touch
End Sub

'/**
' * @brief Reports whether distance fog is currently applied.
' * @return True when fog is enabled.
' */
Public Property Get FogEnabled() As Boolean
    EnsureReady
    FogEnabled = m_fogOn
End Property

'/**
' * @brief Reads the depth at which fog starts.
' * @return The near fog distance.
' */
Public Property Get FogStart() As Single
    EnsureReady
    FogStart = m_fogStart
End Property

'/**
' * @brief Reads the depth at which fog becomes opaque.
' * @return The far fog distance, which doubles as the render far plane.
' */
Public Property Get FogEnd() As Single
    EnsureReady
    FogEnd = m_fogEnd
End Property

'/**
' * @brief Reads the fog colour.
' * @return The packed fog colour.
' */
Public Property Get FogColor() As Long
    EnsureReady
    FogColor = m_fogColor
End Property

'/**
' * @brief Reads how many bands the fog is quantised into.
' * @return The band count.
' */
Public Property Get FogSteps() As Long
    EnsureReady
    FogSteps = m_fogSteps
End Property

'/**
' * @brief Maps a view depth onto its fog band.
' * @param depth The distance from the camera in world units.
' * @return The band index, from 0 to FogSteps.
' */
Public Function FogStepOf(ByVal depth As Single) As Long
    Dim s As Long

    EnsureReady
    If Not m_fogOn Then Exit Function
    If depth <= m_fogStart Then Exit Function

    If depth >= m_fogEnd Then
        FogStepOf = m_fogSteps
        Exit Function
    End If

    s = CLng((depth - m_fogStart) * m_fogScale)
    FogStepOf = PCore.ClampLong(s, 0, m_fogSteps)
End Function

'/**
' * @brief Blends a colour towards the fog colour by depth.
' * @param col The shaded surface colour.
' * @param depth The distance from the camera in world units.
' * @return The fogged colour.
' */
Public Function ApplyFog(ByVal col As Long, ByVal depth As Single) As Long
    EnsureReady

    If Not m_fogOn Or depth <= m_fogStart Then
        ApplyFog = col
        Exit Function
    End If

    If depth >= m_fogEnd Then
        ApplyFog = m_fogColor
        Exit Function
    End If

    ApplyFog = PCore.ColorMix(col, m_fogColor, (depth - m_fogStart) * m_fogInvSpan)
End Function

'/**
' * @brief Blends a colour towards the fog colour by band index.
' * @param col The shaded surface colour.
' * @param bandIdx The fog band, from 0 to FogSteps.
' * @return The fogged colour.
' * @remarks This is the form used when baking shading tables, where the depth is already quantised.
' */
Public Function ApplyFogStep(ByVal col As Long, ByVal bandIdx As Long) As Long
    EnsureReady

    If Not m_fogOn Or bandIdx <= 0 Then
        ApplyFogStep = col
        Exit Function
    End If

    If bandIdx >= m_fogSteps Then
        ApplyFogStep = m_fogColor
        Exit Function
    End If

    ApplyFogStep = PCore.ColorMix(col, m_fogColor, bandIdx / CSng(m_fogSteps))
End Function

'/** @section Change tracking */

'/**
' * @brief Reads the counter bumped whenever the light or fog changes.
' * @return The current revision.
' * @remarks Cached shading tables compare this against the revision they were baked at to decide whether to rebuild.
' */
Public Property Get Revision() As Long
    EnsureReady
    Revision = m_revision
End Property

'/** @section Private helpers */

'/**
' * @brief Marks the lighting environment as changed.
' * @remarks Also pokes the material registry, so its baked table can carry a plain dirty flag instead of asking this module for a revision number on every face it shades.
' */
Private Sub Touch()
    m_revision = m_revision + 1
    PMaterials.NotifyChanged
End Sub

'/**
' * @brief Applies the stock setup the first time the module is used.
' */
Private Sub EnsureReady()
    If m_ready Then Exit Sub
    ResetDefaults
End Sub

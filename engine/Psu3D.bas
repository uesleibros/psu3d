Attribute VB_Name = "Psu3D"
'/**
' * Psu3D - Engine Facade
' * @description The one entry point a slide needs: it boots a canvas, a camera, a renderer and a scene, wires them together, and keeps them reachable as engine wide defaults. Everything it hands back is an ordinary object, so a project that outgrows the defaults can build its own set and ignore this module entirely.
' * @author UesleiDev
' * @version 1.0
' * @remarks Import order inside the VBE does not matter, but the module list is PCore, PLighting, PMaterial, PMaterials, PCanvas, PCamera, PRenderer, PScene, PBody, PLevel and finally this facade.
' */

Option Explicit
Option Private Module

'/** @section Version */

'/** @description Library name reported by Version. */
Private Const LIB_NAME As String = "Psu3D"

'/** @description Library version reported by Version. */
Private Const LIB_VERSION As String = "1.0"

'/** @section Module state */

Private m_canvas As PCanvas
Private m_camera As PCamera
Private m_renderer As PRenderer
Private m_scene As PScene
Private m_booted As Boolean

'/** @section Identity */

'/**
' * @brief Reads the library name and version.
' * @return A short identification string.
' */
Public Function Version() As String
    Version = LIB_NAME & " " & LIB_VERSION
End Function

'/** @section Boot */

'/**
' * @brief Builds the default canvas, camera, renderer and scene and binds them to a slide.
' * @param target The Shapes collection of the slide that will be drawn into.
' * @param X The left edge of the view, in slide points; the whole slide is used when width and height are left at zero.
' * @param Y The top edge of the view, in slide points.
' * @param Width The view width, in slide points.
' * @param Height The view height, in slide points.
' * @remarks Also registers the stock material palette and sweeps away polygons left behind by an interrupted run.
' */
Public Sub Boot(ByVal target As Shapes, _
                Optional ByVal X As Single = 0!, Optional ByVal Y As Single = 0!, _
                Optional ByVal Width As Single = 0!, Optional ByVal Height As Single = 0!)
    Set m_canvas = New PCanvas
    Set m_camera = New PCamera
    Set m_renderer = New PRenderer
    Set m_scene = New PScene

    If Width <= 0! Or Height <= 0! Then
        m_canvas.FillSlide
    Else
        m_canvas.Init X, Y, Width, Height
    End If

    PMaterials.CreateDefaults

    m_renderer.Attach target, m_canvas, m_camera
    m_renderer.Purge


    m_booted = True
End Sub

'/**
' * @brief Wipes every polygon this engine drew and releases the default objects.
' * @remarks Call it when leaving the slide, so nothing is left on it once the show moves on.
' */
Public Sub Shutdown()
    If Not m_renderer Is Nothing Then
        m_renderer.ClearShapes
        m_renderer.Purge
    End If

    If Not m_canvas Is Nothing And Not m_renderer Is Nothing Then
        m_canvas.RemoveBackdrop m_renderer.Target
    End If

    Set m_renderer = Nothing
    Set m_scene = Nothing
    Set m_camera = Nothing
    Set m_canvas = Nothing
    m_booted = False
End Sub

'/**
' * @brief Reports whether Boot has run.
' * @return True when the default objects exist.
' */
Public Property Get IsReady() As Boolean
    IsReady = m_booted
End Property

'/** @section Defaults */

'/**
' * @brief Reads the default canvas.
' * @return The canvas built by Boot, or Nothing before it runs.
' */
Public Property Get Canvas() As PCanvas
    Set Canvas = m_canvas
End Property

'/**
' * @brief Reads the default camera.
' * @return The camera built by Boot.
' */
Public Property Get Camera() As PCamera
    Set Camera = m_camera
End Property

'/**
' * @brief Reads the default renderer.
' * @return The renderer built by Boot.
' */
Public Property Get Renderer() As PRenderer
    Set Renderer = m_renderer
End Property

'/**
' * @brief Reads the default scene.
' * @return The scene built by Boot.
' */
Public Property Get Scene() As PScene
    Set Scene = m_scene
End Property

'/** @section Factories */

'/**
' * @brief Creates a canvas, which is how a second view is added to a slide.
' * @param X The left edge in slide points.
' * @param Y The top edge in slide points.
' * @param Width The view width in slide points.
' * @param Height The view height in slide points.
' * @param cvName The canvas name, which also names its backdrop shape.
' * @return The new canvas.
' */
Public Function CreateCanvas(ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, _
                             Optional ByVal cvName As String = vbNullString) As PCanvas
    Dim cv As PCanvas

    Set cv = New PCanvas
    cv.Init X, Y, Width, Height
    If Len(cvName) > 0 Then cv.Name = cvName

    Set CreateCanvas = cv
End Function

'/**
' * @brief Creates a camera.
' * @param X The eye X coordinate.
' * @param Y The eye Y coordinate.
' * @param Z The eye Z coordinate.
' * @param yawRad The heading in radians.
' * @param pitchRad The elevation in radians.
' * @return The new camera.
' */
Public Function CreateCamera(Optional ByVal X As Single = 0!, Optional ByVal Y As Single = 0!, Optional ByVal Z As Single = 0!, _
                             Optional ByVal yawRad As Single = 0!, Optional ByVal pitchRad As Single = 0!) As PCamera
    Dim cam As PCamera

    Set cam = New PCamera
    cam.Init X, Y, Z, yawRad, pitchRad

    Set CreateCamera = cam
End Function

'/**
' * @brief Creates a renderer already bound to a slide, a canvas and a camera.
' * @param target The Shapes collection to draw into.
' * @param cv The canvas to project onto.
' * @param cam The camera to render from.
' * @param prefix The shape name prefix; give every renderer on a slide its own.
' * @return The new renderer.
' */
Public Function CreateRenderer(ByVal target As Shapes, ByVal cv As PCanvas, ByVal cam As PCamera, _
                               Optional ByVal prefix As String = vbNullString) As PRenderer
    Dim rd As PRenderer

    Set rd = New PRenderer
    If Len(prefix) > 0 Then rd.Prefix = prefix
    rd.Attach target, cv, cam

    Set CreateRenderer = rd
End Function

'/**
' * @brief Creates a scene.
' * @param slots How many object slots to reserve up front.
' * @return The new scene.
' */
Public Function CreateScene(Optional ByVal slots As Long = 0) As PScene
    Dim sc As PScene

    Set sc = New PScene
    If slots > 0 Then sc.Reserve slots

    Set CreateScene = sc
End Function

'/**
' * @brief Creates a body bound to a scene.
' * @param sc The scene it moves through; the booted one when none is named.
' * @return The new body, ready to be tuned and advanced.
' */
Public Function CreateBody(Optional ByVal sc As PScene) As PBody
    Dim b As PBody

    Set b = New PBody

    If sc Is Nothing Then
        b.Attach m_scene
    Else
        b.Attach sc
    End If

    Set CreateBody = b
End Function

'/** @section Environment shortcuts */

'/**
' * @brief Points the global light in a new direction.
' * @param dx The X component of the direction.
' * @param dy The Y component of the direction.
' * @param dz The Z component of the direction.
' */
Public Sub SetLight(ByVal dx As Single, ByVal dy As Single, ByVal dz As Single)
    PLighting.SetLight dx, dy, dz
End Sub

'/**
' * @brief Configures the distance fog, and repaints the canvas backdrop to match.
' * @param startDist The depth at which fog begins.
' * @param endDist The depth at which geometry disappears into the fog.
' * @param col The fog colour.
' */
Public Sub SetFog(ByVal startDist As Single, ByVal endDist As Single, ByVal col As Long)
    PLighting.SetFog startDist, endDist, col
    If Not m_canvas Is Nothing Then m_canvas.BackColor = col
End Sub

'/** @section Frame */

'/**
' * @brief Opens a frame on the default renderer and lets the polygon budget follow the frame time.
' * @param dt The duration of the previous frame, in seconds.
' */
Public Sub BeginFrame(ByVal dt As Single)
    If m_renderer Is Nothing Then Exit Sub

    m_renderer.AdaptBudget dt
    m_renderer.BeginFrame
End Sub

'/**
' * @brief Closes the frame, retiring the shapes the previous one left on the slide.
' * @remarks Required when the renderer is double buffered, which it is by default. Call it after every draw call of the frame.
' */
Public Sub EndFrame()
    If m_renderer Is Nothing Then Exit Sub
    m_renderer.EndFrame
End Sub

'/**
' * @brief Draws the default scene through the default renderer.
' * @return How many objects were submitted.
' */
Public Function RenderScene() As Long
    If m_renderer Is Nothing Or m_scene Is Nothing Then Exit Function
    RenderScene = m_scene.Render(m_renderer)
End Function

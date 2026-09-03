VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/**
' * PCanvas - Psu3D Render Target Rectangle
' * @description Defines where on the slide a 3D view is drawn and how the world is projected into it. A canvas owns its position, its size, its field of view and its clipping bounds, so several independent views can share one slide: a full screen game, a rear view mirror, a minimap or a small preview window.
' * @author UesleiDev
' * @version 1.0
' * @scope Private to this VBA project: class modules carry VB_Exposed = False, which is the class level equivalent of Option Private Module.
' * @remarks Every coordinate is expressed in slide points, the same unit PowerPoint uses for shapes, so a canvas can be placed exactly where a placeholder sits on the slide.
' */

Option Explicit

'/** @section Instance state */

Private m_name As String

Private m_x As Single
Private m_y As Single
Private m_w As Single
Private m_h As Single

Private m_halfW As Single
Private m_halfH As Single
Private m_centerX As Single
Private m_centerY As Single

Private m_fov As Single
Private m_focal As Single
Private m_tanH As Single
Private m_tanV As Single
Private m_normH As Single
Private m_normV As Single

Private m_near As Single
Private m_far As Single

Private m_backColor As Long
Private m_backVisible As Boolean
Private m_backShape As Shape

'/** @section Lifetime */

'/**
' * @brief Starts the canvas at a sane default before Init is called.
' */
Private Sub Class_Initialize()
    m_name = "canvas"
    m_near = 0.06
    m_far = 0!
    m_fov = 0.82
    m_backColor = PCore.ColorPack(160, 185, 210)
    m_backVisible = False

    SetRect 0!, 0!, 960!, 540!
End Sub

'/**
' * @brief Places and sizes the canvas in one call.
' * @param X The left edge on the slide, in points.
' * @param Y The top edge on the slide, in points.
' * @param Width The drawable width, in points.
' * @param Height The drawable height, in points.
' * @return The canvas itself, so a canvas can be created and configured in one expression.
' */
Public Function Init(ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As PCanvas
    SetRect X, Y, Width, Height
    Set Init = Me
End Function

'/**
' * @brief Sizes the canvas to cover the whole slide.
' * @param margin An optional inset applied to all four edges, in points.
' * @return The canvas itself.
' */
Public Function FillSlide(Optional ByVal margin As Single = 0!) As PCanvas
    Dim sw As Single
    Dim sh As Single

    With ActivePresentation.PageSetup
        sw = .SlideWidth
        sh = .SlideHeight
    End With

    SetRect margin, margin, sw - margin * 2!, sh - margin * 2!
    Set FillSlide = Me
End Function

'/**
' * @brief Centres a canvas of the given size on the slide.
' * @param Width The drawable width, in points.
' * @param Height The drawable height, in points.
' * @return The canvas itself.
' */
Public Function CenterOnSlide(ByVal Width As Single, ByVal Height As Single) As PCanvas
    Dim sw As Single
    Dim sh As Single

    With ActivePresentation.PageSetup
        sw = .SlideWidth
        sh = .SlideHeight
    End With

    SetRect (sw - Width) * 0.5!, (sh - Height) * 0.5!, Width, Height
    Set CenterOnSlide = Me
End Function

'/**
' * @brief Fits the canvas to the slide while respecting an aspect ratio.
' * @param aspect The width divided by the height the view should keep.
' * @param fit pfContain to letterbox inside the slide, pfCover to fill it and overflow, pfStretch to ignore the ratio.
' * @return The canvas itself.
' */
Public Function FitToSlide(ByVal aspect As Single, Optional ByVal fit As PCanvasFit = pfContain) As PCanvas
    Dim sw As Single
    Dim sh As Single
    Dim w As Single
    Dim h As Single

    With ActivePresentation.PageSetup
        sw = .SlideWidth
        sh = .SlideHeight
    End With

    If aspect < P_EPSILON Then aspect = 1.7778!

    Select Case fit
        Case pfStretch
            w = sw
            h = sh
        Case pfCover
            If sw / sh < aspect Then
                h = sh
                w = sh * aspect
            Else
                w = sw
                h = sw / aspect
            End If
        Case Else
            If sw / sh > aspect Then
                h = sh
                w = sh * aspect
            Else
                w = sw
                h = sw / aspect
            End If
    End Select

    SetRect (sw - w) * 0.5!, (sh - h) * 0.5!, w, h
    Set FitToSlide = Me
End Function

'/** @section Geometry */

'/**
' * @brief Places and sizes the canvas.
' * @param X The left edge on the slide, in points.
' * @param Y The top edge on the slide, in points.
' * @param Width The drawable width, in points.
' * @param Height The drawable height, in points.
' */
Public Sub SetRect(ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single)
    If Width < 1! Then Width = 1!
    If Height < 1! Then Height = 1!

    m_x = X
    m_y = Y
    m_w = Width
    m_h = Height

    Refresh
End Sub

'/**
' * @brief Moves the canvas without resizing it.
' * @param X The new left edge, in points.
' * @param Y The new top edge, in points.
' */
Public Sub SetPosition(ByVal X As Single, ByVal Y As Single)
    SetRect X, Y, m_w, m_h
End Sub

'/**
' * @brief Resizes the canvas without moving it.
' * @param Width The new drawable width, in points.
' * @param Height The new drawable height, in points.
' */
Public Sub SetSize(ByVal Width As Single, ByVal Height As Single)
    SetRect m_x, m_y, Width, Height
End Sub

'/**
' * @brief Copies the position and size of an existing shape, so a placeholder drawn on the slide can define the viewport.
' * @param src The shape whose bounds are adopted.
' */
Public Sub SetFromShape(ByVal src As Shape)
    If src Is Nothing Then Exit Sub
    SetRect src.Left, src.Top, src.Width, src.Height
End Sub

'/** @section Identity */

'/**
' * @brief Reads the canvas name, which is also used to name its backdrop shape.
' * @return The canvas name.
' */
Public Property Get Name() As String
    Name = m_name
End Property

'/**
' * @brief Sets the canvas name.
' * @param value The new name; keep it unique when several canvases share a slide.
' */
Public Property Let Name(ByVal value As String)
    If Len(value) > 0 Then m_name = value
End Property

'/** @section Bounds */

'/**
' * @brief Reads the left edge of the canvas on the slide.
' * @return The X coordinate in points.
' */
Public Property Get X() As Single
    X = m_x
End Property

'/**
' * @brief Moves the left edge of the canvas.
' * @param value The X coordinate in points.
' */
Public Property Let X(ByVal value As Single)
    SetRect value, m_y, m_w, m_h
End Property

'/**
' * @brief Reads the top edge of the canvas on the slide.
' * @return The Y coordinate in points.
' */
Public Property Get Y() As Single
    Y = m_y
End Property

'/**
' * @brief Moves the top edge of the canvas.
' * @param value The Y coordinate in points.
' */
Public Property Let Y(ByVal value As Single)
    SetRect m_x, value, m_w, m_h
End Property

'/**
' * @brief Reads the drawable width.
' * @return The width in points.
' */
Public Property Get Width() As Single
    Width = m_w
End Property

'/**
' * @brief Sets the drawable width.
' * @param value The width in points.
' */
Public Property Let Width(ByVal value As Single)
    SetRect m_x, m_y, value, m_h
End Property

'/**
' * @brief Reads the drawable height.
' * @return The height in points.
' */
Public Property Get Height() As Single
    Height = m_h
End Property

'/**
' * @brief Sets the drawable height.
' * @param value The height in points.
' */
Public Property Let Height(ByVal value As Single)
    SetRect m_x, m_y, m_w, value
End Property

'/**
' * @brief Reads the right edge of the canvas.
' * @return The X coordinate of the right border, in points.
' */
Public Property Get Right() As Single
    Right = m_x + m_w
End Property

'/**
' * @brief Reads the bottom edge of the canvas.
' * @return The Y coordinate of the bottom border, in points.
' */
Public Property Get Bottom() As Single
    Bottom = m_y + m_h
End Property

'/**
' * @brief Reads the slide X coordinate of the canvas centre, which is where the view axis lands.
' * @return The centre X in points.
' */
Public Property Get CenterX() As Single
    CenterX = m_centerX
End Property

'/**
' * @brief Reads the slide Y coordinate of the canvas centre.
' * @return The centre Y in points.
' */
Public Property Get CenterY() As Single
    CenterY = m_centerY
End Property

'/**
' * @brief Reads half the canvas width.
' * @return The half width in points.
' */
Public Property Get HalfWidth() As Single
    HalfWidth = m_halfW
End Property

'/**
' * @brief Reads half the canvas height.
' * @return The half height in points.
' */
Public Property Get HalfHeight() As Single
    HalfHeight = m_halfH
End Property

'/**
' * @brief Reads the width to height ratio of the canvas.
' * @return The aspect ratio.
' */
Public Property Get Aspect() As Single
    Aspect = m_w / m_h
End Property

'/** @section Projection */

'/**
' * @brief Reads the horizontal field of view.
' * @return The half angle in radians.
' */
Public Property Get FieldOfView() As Single
    FieldOfView = m_fov
End Property

'/**
' * @brief Sets the horizontal field of view.
' * @param value The half angle in radians, clamped to a usable range.
' */
Public Property Let FieldOfView(ByVal value As Single)
    m_fov = PCore.Clamp(value, 0.15!, 1.45!)
    Refresh
End Property

'/**
' * @brief Reads the horizontal field of view in degrees.
' * @return The half angle in degrees.
' */
Public Property Get FieldOfViewDeg() As Single
    FieldOfViewDeg = m_fov * P_RAD2DEG
End Property

'/**
' * @brief Sets the horizontal field of view in degrees.
' * @param value The half angle in degrees.
' */
Public Property Let FieldOfViewDeg(ByVal value As Single)
    FieldOfView = value * P_DEG2RAD
End Property

'/**
' * @brief Reads the near clipping distance.
' * @return The distance in world units below which geometry is cut away.
' */
Public Property Get NearPlane() As Single
    NearPlane = m_near
End Property

'/**
' * @brief Sets the near clipping distance.
' * @param value The distance in world units; very small values make nearby faces flicker.
' */
Public Property Let NearPlane(ByVal value As Single)
    m_near = PCore.Clamp(value, 0.01!, 4!)
End Property

'/**
' * @brief Reads the far clipping distance.
' * @return The distance in world units beyond which geometry is skipped, falling back to the fog end when unset.
' */
Public Property Get FarPlane() As Single
    If m_far > 0! Then
        FarPlane = m_far
    Else
        FarPlane = PLighting.FogEnd
    End If
End Property

'/**
' * @brief Sets the far clipping distance.
' * @param value The distance in world units, or zero to follow the fog end automatically.
' */
Public Property Let FarPlane(ByVal value As Single)
    If value < 0! Then value = 0!
    m_far = value
End Property

'/**
' * @brief Reads the focal length derived from the field of view and the canvas width.
' * @return The distance from the eye to the projection plane, in points.
' */
Public Property Get FocalLength() As Single
    FocalLength = m_focal
End Property

'/**
' * @brief Reads the tangent of the horizontal half field of view.
' * @return The horizontal frustum slope.
' */
Public Property Get TanFovH() As Single
    TanFovH = m_tanH
End Property

'/**
' * @brief Reads the tangent of the vertical half field of view, which follows from the canvas shape.
' * @return The vertical frustum slope.
' */
Public Property Get TanFovV() As Single
    TanFovV = m_tanV
End Property

'/**
' * @brief Reads the normalisation factor of the left and right frustum planes.
' * @return The factor used when testing a sphere against the side planes.
' */
Public Property Get NormH() As Single
    NormH = m_normH
End Property

'/**
' * @brief Reads the normalisation factor of the top and bottom frustum planes.
' * @return The factor used when testing a sphere against the vertical planes.
' */
Public Property Get NormV() As Single
    NormV = m_normV
End Property

'/**
' * @brief Projects a point already expressed in view space onto the slide.
' * @param fwd The distance in front of the eye.
' * @param side The distance to the right of the eye.
' * @param up The distance above the eye.
' * @param outX Receives the slide X coordinate in points.
' * @param outY Receives the slide Y coordinate in points.
' * @param shiftY An extra vertical offset in points, used for effects such as head bob.
' * @return False when the point sits behind the near plane and cannot be projected.
' */
Public Function Project(ByVal fwd As Single, ByVal side As Single, ByVal up As Single, _
                        ByRef outX As Single, ByRef outY As Single, _
                        Optional ByVal shiftY As Single = 0!) As Boolean
    Dim invD As Single

    If fwd < m_near Then Exit Function

    invD = m_focal / fwd
    outX = m_centerX + side * invD
    outY = m_centerY - up * invD + shiftY

    Project = True
End Function

'/** @section Point queries */

'/**
' * @brief Tests whether a slide point falls inside the canvas.
' * @param slideX The X coordinate in points.
' * @param slideY The Y coordinate in points.
' * @return True when the point is inside the drawable area.
' * @remarks This is how a caller asks about a pointer. The canvas maps coordinates and never
' * reads hardware: subtract CenterX and CenterY from whatever gives you a pointer position and you
' * have the offset from the middle, which is what mouse look wants. There were CursorX, CursorY,
' * CursorInside and CenterCursor here once. They were convenient, and they made every project that
' * imported the core have to import a cursor module too, wanting a mouse or not.
' */
Public Function Contains(ByVal slideX As Single, ByVal slideY As Single) As Boolean
    Contains = PCore.PointInBox(slideX, slideY, m_x, m_y, m_x + m_w, m_y + m_h)
End Function

'/**
' * @brief Converts a slide point into canvas local coordinates, where the origin is the top left of the canvas.
' * @param slideX The X coordinate in points.
' * @param slideY The Y coordinate in points.
' * @param outX Receives the local X coordinate.
' * @param outY Receives the local Y coordinate.
' */
Public Sub ToLocal(ByVal slideX As Single, ByVal slideY As Single, ByRef outX As Single, ByRef outY As Single)
    outX = slideX - m_x
    outY = slideY - m_y
End Sub

'/**
' * @brief Converts canvas local coordinates back into slide coordinates.
' * @param localX The local X coordinate.
' * @param localY The local Y coordinate.
' * @param outX Receives the slide X coordinate.
' * @param outY Receives the slide Y coordinate.
' */
Public Sub ToSlide(ByVal localX As Single, ByVal localY As Single, ByRef outX As Single, ByRef outY As Single)
    outX = localX + m_x
    outY = localY + m_y
End Sub

'/** @section Backdrop */

'/**
' * @brief Reads the colour painted behind the geometry.
' * @return The packed backdrop colour.
' */
Public Property Get BackColor() As Long
    BackColor = m_backColor
End Property

'/**
' * @brief Sets the colour painted behind the geometry.
' * @param value The packed colour; matching the fog colour hides the horizon seam.
' */
Public Property Let BackColor(ByVal value As Long)
    m_backColor = value
End Property

'/**
' * @brief Reports whether the backdrop rectangle is in use.
' * @return True when EnsureBackdrop will keep a shape alive.
' */
Public Property Get BackVisible() As Boolean
    BackVisible = m_backVisible
End Property

'/**
' * @brief Chooses whether the canvas paints a backdrop rectangle.
' * @param value True to draw a solid rectangle behind the geometry.
' */
Public Property Let BackVisible(ByVal value As Boolean)
    m_backVisible = value
End Property

'/**
' * @brief Creates or updates the persistent backdrop rectangle for this canvas.
' * @param target The shape collection of the slide being rendered.
' * @remarks The backdrop is created once and reused across frames, so it costs nothing per frame; it is pushed to the back so geometry always draws over it.
' */
Public Sub EnsureBackdrop(ByVal target As Shapes)
    If target Is Nothing Then Exit Sub

    If Not m_backVisible Then
        RemoveBackdrop target
        Exit Sub
    End If

    If Not m_backShape Is Nothing Then Exit Sub

    Set m_backShape = FindShape(target, BackdropName)

    If m_backShape Is Nothing Then
        Set m_backShape = target.AddShape(msoShapeRectangle, m_x, m_y, m_w, m_h)
        m_backShape.Name = BackdropName
        m_backShape.Line.Visible = msoFalse
    End If

    With m_backShape
        .Left = m_x
        .Top = m_y
        .Width = m_w
        .Height = m_h
        .Fill.ForeColor.RGB = m_backColor
        .Fill.Transparency = 0!
        .ZOrder msoSendToBack
    End With
End Sub

'/**
' * @brief Raises the backdrop above everything drawn so far.
' * @remarks What a second canvas needs before it draws: the frame of the first one was created after
' * this backdrop and would otherwise show through the smaller view. Lifting it here reopens a clean
' * rectangle for this canvas to draw into, and its own polygons land on top a moment later.
' */
Public Sub LiftBackdrop()
    If m_backShape Is Nothing Then Exit Sub
    m_backShape.ZOrder msoBringToFront
End Sub

'/**
' * @brief Deletes the backdrop rectangle if one exists.
' * @param target The shape collection of the slide being rendered.
' */
Public Sub RemoveBackdrop(ByVal target As Shapes)
    Dim s As Shape

    If target Is Nothing Then Exit Sub

    '/* Resolved by name rather than through the cache, so a reference left dangling by an outside delete cannot be touched. */
    Set s = FindShape(target, BackdropName)
    Set m_backShape = Nothing

    If s Is Nothing Then Exit Sub

    s.Delete
End Sub

'/**
' * @brief Reads the shape name reserved for this canvas backdrop.
' * @return The generated shape name.
' * @remarks Deliberately outside the "p3d_" polygon prefix, so a renderer purge never sweeps the backdrop away.
' */
Public Property Get BackdropName() As String
    BackdropName = "p3dbg_" & m_name
End Property

'/** @section Private helpers */

'/**
' * @brief Finds a shape by name without relying on error trapping.
' * @param target The shape collection to search.
' * @param nm The shape name to look for.
' * @return The shape, or Nothing when the slide does not hold it.
' */
Private Function FindShape(ByVal target As Shapes, ByVal nm As String) As Shape
    Dim s As Shape

    For Each s In target
        If s.Name = nm Then
            Set FindShape = s
            Exit Function
        End If
    Next s
End Function

'/**
' * @brief Recomputes every derived projection metric after the rectangle or the field of view changes.
' */
Private Sub Refresh()
    m_halfW = m_w * 0.5!
    m_halfH = m_h * 0.5!
    m_centerX = m_x + m_halfW
    m_centerY = m_y + m_halfH

    m_tanH = Tan(m_fov)
    m_focal = m_halfW / m_tanH
    m_tanV = m_halfH / m_focal

    m_normH = 1! / Sqr(1! + m_tanH * m_tanH)
    m_normV = 1! / Sqr(1! + m_tanV * m_tanV)
End Sub

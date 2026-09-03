# PCanvas

**Where and how to project**

Defines where on the slide a 3D view is drawn and how the world is projected into it. A canvas owns its position, its size, its field of view and its clipping bounds, so several independent views can share one slide: a full screen game, a rear view mirror, a minimap or a small preview window.

> Every coordinate is expressed in slide points, the same unit PowerPoint uses for shapes, so a canvas can be placed exactly where a placeholder sits on the slide.

> **Scope.** Private to this VBA project: class modules carry VB_Exposed = False, which is the class level equivalent of Option Private Module.

## Index

**Lifetime.** [`Init`](#init), [`FillSlide`](#fillslide), [`CenterOnSlide`](#centeronslide), [`FitToSlide`](#fittoslide)

**Geometry.** [`SetRect`](#setrect), [`SetPosition`](#setposition), [`SetSize`](#setsize), [`SetFromShape`](#setfromshape)

**Identity.** [`Name`](#name)

**Bounds.** [`X`](#x), [`Y`](#y), [`Width`](#width), [`Height`](#height), [`Right`](#right), [`Bottom`](#bottom), [`CenterX`](#centerx), [`CenterY`](#centery), [`HalfWidth`](#halfwidth), [`HalfHeight`](#halfheight), [`Aspect`](#aspect)

**Projection.** [`FieldOfView`](#fieldofview), [`FieldOfViewDeg`](#fieldofviewdeg), [`NearPlane`](#nearplane), [`FarPlane`](#farplane), [`FocalLength`](#focallength), [`TanFovH`](#tanfovh), [`TanFovV`](#tanfovv), [`NormH`](#normh), [`NormV`](#normv), [`Project`](#project)

**Point queries.** [`Contains`](#contains), [`ToLocal`](#tolocal), [`ToSlide`](#toslide)

**Backdrop.** [`BackColor`](#backcolor), [`BackVisible`](#backvisible), [`EnsureBackdrop`](#ensurebackdrop), [`LiftBackdrop`](#liftbackdrop), [`RemoveBackdrop`](#removebackdrop), [`BackdropName`](#backdropname)

## Members

### Init

```vba
Public Function Init(ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As PCanvas
```

Places and sizes the canvas in one call.

| parameter | what it is |
|---|---|
| `X` | The left edge on the slide, in points. |
| `Y` | The top edge on the slide, in points. |
| `Width` | The drawable width, in points. |
| `Height` | The drawable height, in points. |

**Returns.** The canvas itself, so a canvas can be created and configured in one expression.

### FillSlide

```vba
Public Function FillSlide(Optional ByVal margin As Single = 0!) As PCanvas
```

Sizes the canvas to cover the whole slide.

| parameter | what it is |
|---|---|
| `margin` | An optional inset applied to all four edges, in points. |

**Returns.** The canvas itself.

### CenterOnSlide

```vba
Public Function CenterOnSlide(ByVal Width As Single, ByVal Height As Single) As PCanvas
```

Centres a canvas of the given size on the slide.

| parameter | what it is |
|---|---|
| `Width` | The drawable width, in points. |
| `Height` | The drawable height, in points. |

**Returns.** The canvas itself.

### FitToSlide

```vba
Public Function FitToSlide(ByVal aspect As Single, Optional ByVal fit As PCanvasFit = pfContain) As PCanvas
```

Fits the canvas to the slide while respecting an aspect ratio.

| parameter | what it is |
|---|---|
| `aspect` | The width divided by the height the view should keep. |
| `fit` | pfContain to letterbox inside the slide, pfCover to fill it and overflow, pfStretch to ignore the ratio. |

**Returns.** The canvas itself.

### SetRect

```vba
Public Sub SetRect(ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single)
```

Places and sizes the canvas.

| parameter | what it is |
|---|---|
| `X` | The left edge on the slide, in points. |
| `Y` | The top edge on the slide, in points. |
| `Width` | The drawable width, in points. |
| `Height` | The drawable height, in points. |

### SetPosition

```vba
Public Sub SetPosition(ByVal X As Single, ByVal Y As Single)
```

Moves the canvas without resizing it.

| parameter | what it is |
|---|---|
| `X` | The new left edge, in points. |
| `Y` | The new top edge, in points. |

### SetSize

```vba
Public Sub SetSize(ByVal Width As Single, ByVal Height As Single)
```

Resizes the canvas without moving it.

| parameter | what it is |
|---|---|
| `Width` | The new drawable width, in points. |
| `Height` | The new drawable height, in points. |

### SetFromShape

```vba
Public Sub SetFromShape(ByVal src As Shape)
```

Copies the position and size of an existing shape, so a placeholder drawn on the slide can define the viewport.

| parameter | what it is |
|---|---|
| `src` | The shape whose bounds are adopted. |

### Name

```vba
Public Property Get Name() As String
Public Property Let Name(ByVal value As String)
```

Reads the canvas name, which is also used to name its backdrop shape.

**Write.** Sets the canvas name.

| parameter | what it is |
|---|---|
| `value` | The new name; keep it unique when several canvases share a slide. |

**Returns.** The canvas name.

### X

```vba
Public Property Get X() As Single
Public Property Let X(ByVal value As Single)
```

Reads the left edge of the canvas on the slide.

**Write.** Moves the left edge of the canvas.

| parameter | what it is |
|---|---|
| `value` | The X coordinate in points. |

**Returns.** The X coordinate in points.

### Y

```vba
Public Property Get Y() As Single
Public Property Let Y(ByVal value As Single)
```

Reads the top edge of the canvas on the slide.

**Write.** Moves the top edge of the canvas.

| parameter | what it is |
|---|---|
| `value` | The Y coordinate in points. |

**Returns.** The Y coordinate in points.

### Width

```vba
Public Property Get Width() As Single
Public Property Let Width(ByVal value As Single)
```

Reads the drawable width.

**Write.** Sets the drawable width.

| parameter | what it is |
|---|---|
| `value` | The width in points. |

**Returns.** The width in points.

### Height

```vba
Public Property Get Height() As Single
Public Property Let Height(ByVal value As Single)
```

Reads the drawable height.

**Write.** Sets the drawable height.

| parameter | what it is |
|---|---|
| `value` | The height in points. |

**Returns.** The height in points.

### Right

```vba
Public Property Get Right() As Single
```

Reads the right edge of the canvas.

**Returns.** The X coordinate of the right border, in points.

### Bottom

```vba
Public Property Get Bottom() As Single
```

Reads the bottom edge of the canvas.

**Returns.** The Y coordinate of the bottom border, in points.

### CenterX

```vba
Public Property Get CenterX() As Single
```

Reads the slide X coordinate of the canvas centre, which is where the view axis lands.

**Returns.** The centre X in points.

### CenterY

```vba
Public Property Get CenterY() As Single
```

Reads the slide Y coordinate of the canvas centre.

**Returns.** The centre Y in points.

### HalfWidth

```vba
Public Property Get HalfWidth() As Single
```

Reads half the canvas width.

**Returns.** The half width in points.

### HalfHeight

```vba
Public Property Get HalfHeight() As Single
```

Reads half the canvas height.

**Returns.** The half height in points.

### Aspect

```vba
Public Property Get Aspect() As Single
```

Reads the width to height ratio of the canvas.

**Returns.** The aspect ratio.

### FieldOfView

```vba
Public Property Get FieldOfView() As Single
Public Property Let FieldOfView(ByVal value As Single)
```

Reads the horizontal field of view.

**Write.** Sets the horizontal field of view.

| parameter | what it is |
|---|---|
| `value` | The half angle in radians, clamped to a usable range. |

**Returns.** The half angle in radians.

### FieldOfViewDeg

```vba
Public Property Get FieldOfViewDeg() As Single
Public Property Let FieldOfViewDeg(ByVal value As Single)
```

Reads the horizontal field of view in degrees.

**Write.** Sets the horizontal field of view in degrees.

| parameter | what it is |
|---|---|
| `value` | The half angle in degrees. |

**Returns.** The half angle in degrees.

### NearPlane

```vba
Public Property Get NearPlane() As Single
Public Property Let NearPlane(ByVal value As Single)
```

Reads the near clipping distance.

**Write.** Sets the near clipping distance.

| parameter | what it is |
|---|---|
| `value` | The distance in world units; very small values make nearby faces flicker. |

**Returns.** The distance in world units below which geometry is cut away.

### FarPlane

```vba
Public Property Get FarPlane() As Single
Public Property Let FarPlane(ByVal value As Single)
```

Reads the far clipping distance.

**Write.** Sets the far clipping distance.

| parameter | what it is |
|---|---|
| `value` | The distance in world units, or zero to follow the fog end automatically. |

**Returns.** The distance in world units beyond which geometry is skipped, falling back to the fog end when unset.

### FocalLength

```vba
Public Property Get FocalLength() As Single
```

Reads the focal length derived from the field of view and the canvas width.

**Returns.** The distance from the eye to the projection plane, in points.

### TanFovH

```vba
Public Property Get TanFovH() As Single
```

Reads the tangent of the horizontal half field of view.

**Returns.** The horizontal frustum slope.

### TanFovV

```vba
Public Property Get TanFovV() As Single
```

Reads the tangent of the vertical half field of view, which follows from the canvas shape.

**Returns.** The vertical frustum slope.

### NormH

```vba
Public Property Get NormH() As Single
```

Reads the normalisation factor of the left and right frustum planes.

**Returns.** The factor used when testing a sphere against the side planes.

### NormV

```vba
Public Property Get NormV() As Single
```

Reads the normalisation factor of the top and bottom frustum planes.

**Returns.** The factor used when testing a sphere against the vertical planes.

### Project

```vba
Public Function Project(ByVal fwd As Single, ByVal side As Single, ByVal up As Single, ByRef outX As Single, ByRef outY As Single, Optional ByVal shiftY As Single = 0!) As Boolean
```

Projects a point already expressed in view space onto the slide.

| parameter | what it is |
|---|---|
| `fwd` | The distance in front of the eye. |
| `side` | The distance to the right of the eye. |
| `up` | The distance above the eye. |
| `outX` | Receives the slide X coordinate in points. |
| `outY` | Receives the slide Y coordinate in points. |
| `shiftY` | An extra vertical offset in points, used for effects such as head bob. |

**Returns.** False when the point sits behind the near plane and cannot be projected.

### Contains

```vba
Public Function Contains(ByVal slideX As Single, ByVal slideY As Single) As Boolean
```

Tests whether a slide point falls inside the canvas.

| parameter | what it is |
|---|---|
| `slideX` | The X coordinate in points. |
| `slideY` | The Y coordinate in points. |

**Returns.** True when the point is inside the drawable area.

> This is how a caller asks about a pointer. The canvas maps coordinates and never reads hardware: subtract CenterX and CenterY from whatever gives you a pointer position and you have the offset from the middle, which is what mouse look wants. There were CursorX, CursorY, CursorInside and CenterCursor here once. They were convenient, and they made every project that imported the core have to import a cursor module too, wanting a mouse or not.

### ToLocal

```vba
Public Sub ToLocal(ByVal slideX As Single, ByVal slideY As Single, ByRef outX As Single, ByRef outY As Single)
```

Converts a slide point into canvas local coordinates, where the origin is the top left of the canvas.

| parameter | what it is |
|---|---|
| `slideX` | The X coordinate in points. |
| `slideY` | The Y coordinate in points. |
| `outX` | Receives the local X coordinate. |
| `outY` | Receives the local Y coordinate. |

### ToSlide

```vba
Public Sub ToSlide(ByVal localX As Single, ByVal localY As Single, ByRef outX As Single, ByRef outY As Single)
```

Converts canvas local coordinates back into slide coordinates.

| parameter | what it is |
|---|---|
| `localX` | The local X coordinate. |
| `localY` | The local Y coordinate. |
| `outX` | Receives the slide X coordinate. |
| `outY` | Receives the slide Y coordinate. |

### BackColor

```vba
Public Property Get BackColor() As Long
Public Property Let BackColor(ByVal value As Long)
```

Reads the colour painted behind the geometry.

**Write.** Sets the colour painted behind the geometry.

| parameter | what it is |
|---|---|
| `value` | The packed colour; matching the fog colour hides the horizon seam. |

**Returns.** The packed backdrop colour.

### BackVisible

```vba
Public Property Get BackVisible() As Boolean
Public Property Let BackVisible(ByVal value As Boolean)
```

Reports whether the backdrop rectangle is in use.

**Write.** Chooses whether the canvas paints a backdrop rectangle.

| parameter | what it is |
|---|---|
| `value` | True to draw a solid rectangle behind the geometry. |

**Returns.** True when EnsureBackdrop will keep a shape alive.

### EnsureBackdrop

```vba
Public Sub EnsureBackdrop(ByVal target As Shapes)
```

Creates or updates the persistent backdrop rectangle for this canvas.

| parameter | what it is |
|---|---|
| `target` | The shape collection of the slide being rendered. |

> The backdrop is created once and reused across frames, so it costs nothing per frame; it is pushed to the back so geometry always draws over it.

### LiftBackdrop

```vba
Public Sub LiftBackdrop()
```

Raises the backdrop above everything drawn so far.

> What a second canvas needs before it draws: the frame of the first one was created after this backdrop and would otherwise show through the smaller view. Lifting it here reopens a clean rectangle for this canvas to draw into, and its own polygons land on top a moment later.

### RemoveBackdrop

```vba
Public Sub RemoveBackdrop(ByVal target As Shapes)
```

Deletes the backdrop rectangle if one exists.

| parameter | what it is |
|---|---|
| `target` | The shape collection of the slide being rendered. |

### BackdropName

```vba
Public Property Get BackdropName() As String
```

Reads the shape name reserved for this canvas backdrop.

**Returns.** The generated shape name.

> Deliberately outside the "p3d_" polygon prefix, so a renderer purge never sweeps the backdrop away.

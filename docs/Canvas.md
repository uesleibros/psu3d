# Canvas

The canvas is the rectangle of the slide the scene is drawn into, plus the lens it is seen through.

```vba
Dim cv As PCanvas
Set cv = New PCanvas

cv.Init 40, 40, 640, 400          ' x, y, width, height in slide points
cv.FieldOfViewDeg = 46
cv.NearPlane = 0.05
cv.FarPlane = 60
```

Four shortcuts cover the common cases:

```vba
cv.FillSlide                      ' the whole slide
cv.CenterOnSlide 800, 450         ' that size, centred
cv.FitToSlide 1.7778              ' that aspect ratio, fitted to the slide
cv.SetFromShape Shapes("viewport")  ' use a placeholder already in the design
```

`SetFromShape` is the most practical one when somebody drew the frame in PowerPoint: you position the rectangle with the mouse and the library obeys.

## Why the canvas has a position

In an ordinary engine the screen is the screen. Here you almost never want the whole slide. You want a viewport beside some text, inside a frame, in a corner.

And because the canvas has a position, two canvases fit on the same slide. Two cameras, two renderers, one scene:

```vba
Set cvMini = New PCanvas
cvMini.Init 20, 20, 250, 150

Set rdMini = New PRenderer
rdMini.Prefix = "mini_"           ' its own prefix, or one erases the other's frame
rdMini.Attach Shapes, cvMini, camMini
rdMini.PolyBudget = 48            ' a second view is a second pass over the same geometry
```

## The lens

`FieldOfViewDeg` is the horizontal field of view. The vertical one is derived from the shape of the canvas, so changing the height of the rectangle does not distort the image, it only crops or reveals.

`FocalLength`, `TanFovH`, `TanFovV`, `NormH` and `NormV` are derived values the renderer reads once per frame. You do not need them unless you are writing your own pipeline.

`NearPlane` is the closest distance at which anything is still drawn. Too close to zero and the projection explodes into enormous numbers; the default is safe.

## Backdrop

The canvas can keep a background rectangle behind the geometry:

```vba
cv.BackVisible = True
cv.BackColor = PCore.ColorPack(160, 185, 210)
cv.EnsureBackdrop Shapes
```

That rectangle is a shape named `p3dbg_<name>`. The prefix differs from the `p3d_` of the geometry shapes on purpose: the renderer's cleanup deletes everything starting with `p3d_`, and the backdrop has to survive it.

## Coordinates

`ToLocal` and `ToSlide` convert between canvas space and slide space. `Contains` answers whether a slide point falls inside the canvas.

**The canvas does not read hardware.** It maps coordinates, and nothing else. Whoever has a pointer does the arithmetic:

```vba
dx = MyCursorX - cv.CenterX
dy = MyCursorY - cv.CenterY
cam.AddAngles dx * 0.0022, -dy * 0.0022
WarpMyCursorTo cv.CenterX, cv.CenterY
```

`CenterX` and `CenterY` are the centre of the canvas in slide points, so the subtraction gives the offset of the pointer from the middle, which is exactly what mouse look wants. Warping it back afterwards is how relative mouse is done without capturing the device.

There used to be a `CursorX`, a `CursorY`, a `CursorInside` and a `CenterCursor` here. They were convenient, and they forced every project that imported the core to import a cursor module as well, whether it wanted a mouse or not. They were removed, and the core stopped depending on anything.

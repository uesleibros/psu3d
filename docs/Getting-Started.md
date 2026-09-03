# Getting started

## A static figure, with no game in it

No loop, no physics, no keyboard. The shapes stay on the slide when the Sub returns, and you can close the VBE and go on editing the presentation.

Paste this into the slide's own module, `Slide1`, rather than a standard module, so that `Shapes` resolves:

```vba
Public Sub Terrain()
    Dim x As Long, y As Long, h As Single, ground As Long

    Psu3D.Boot Shapes, 40, 40, 640, 400
    ground = PMaterials.Create("ground", PCore.ColorPack(96, 140, 90)).Id

    Psu3D.Camera.SetPosition -16, -16, 13
    Psu3D.Camera.LookAt 0, 0, 0
    Psu3D.Renderer.PolyBudget = 3000

    For y = 0 To 23
        For x = 0 To 23
            h = 2 + Sin(x * 0.4) * Cos(y * 0.35) * 1.6
            Psu3D.Scene.AddBox ground, x - 12, y - 12, x - 11, y - 11, h, h
        Next x
    Next y

    Psu3D.BeginFrame 0
    Psu3D.RenderScene
    Psu3D.EndFrame
End Sub
```

That is 576 boxes inside a 640 by 400 point rectangle of the slide. Replace `h` with a value from a spreadsheet and it becomes a 3D bar chart. Replace it with noise and it becomes terrain. The engine cannot tell the difference.

## What each line does

`Psu3D.Boot Shapes, 40, 40, 640, 400` builds the canvas, camera, renderer and scene in one call, creates the stock material palette, and clears any shapes left over from a previous run. The four numbers are the x, y, width and height of the slide rectangle to draw into. Without them the canvas fills the slide.

`PMaterials.Create` registers a surface and hands back the object. The `.Id` is the number the scene uses from then on, because a `Long` in an array is cheaper than an object pointer.

`Psu3D.Camera.LookAt 0, 0, 0` aims the camera at a world point, working out yaw and pitch.

`Psu3D.Renderer.PolyBudget = 3000` says how many polygons the frame may spend. 576 boxes cost three each, so 1,728, and a ceiling of 3,000 guarantees nothing is cut. Without it the default budget would drop the back of the terrain.

`BeginFrame`, `RenderScene`, `EndFrame` is the frame: swap the shape bank, cull and sort and draw, then retire the previous frame.

## A real loop

Once you want movement the drawing goes inside a loop, and two things appear that exist only in PowerPoint.

```vba
Public Sub Main()
    Dim dt As Single, t0 As Single
    Dim timerShp As Shape

    Set timerShp = Slide1.Shapes("timer")

    Psu3D.Boot Shapes
    Psu3D.Camera.SetPosition 0, 0, 1.7
    Psu3D.Scene.AddBox PMaterials.IdOf("grass"), -20, -20, 20, 20, 0, 1

    t0 = Timer

    Do While ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 1
        dt = Timer - t0
        If dt < 0.001 Then dt = 0.001
        If dt > 0.05 Then dt = 0.05
        t0 = Timer

        Psu3D.BeginFrame dt
        Psu3D.RenderScene
        Psu3D.EndFrame

        timerShp.TextEffect.Text = Format$(Timer, "0.00")
        DoEvents
    Loop

    Psu3D.Shutdown
End Sub
```

The first strange thing is `timerShp`. It is not a clock. Writing to a WordArt is what forces PowerPoint to repaint the slide during a show, and the repaint is how the frame becomes visible at all. Read [the refresh pump](Performance.md#the-refresh-pump) before touching it.

The second is `dt` being clamped between one millisecond and fifty. A frame that took half a second, because Windows decided to do something else, would move everything half a second at once and tunnel through a wall. The upper clamp turns that into slow motion, which is always the better failure.

## Without the facade

`Psu3D` is a convenience layer. Nothing requires it:

```vba
Dim cv As PCanvas, cam As PCamera, rd As PRenderer, sc As PScene

Set cv = New PCanvas
cv.Init 40, 40, 640, 400

Set cam = New PCamera
cam.Init 0, 0, 2, P_HALF_PI, 0

Set rd = New PRenderer
rd.Attach Shapes, cv, cam

Set sc = New PScene
```

And nothing requires the scene either. The renderer draws primitives directly, with nothing stored. See [Renderer and primitives](Renderer-And-Primitives.md).

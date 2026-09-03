# Recipes

## A spinning cube on a slide

```vba
Public Sub Cube()
    Dim i As Long, t As Single, m As Long

    Psu3D.Boot Shapes, 100, 60, 500, 320
    m = PMaterials.Create("cube", PCore.ColorPack(220, 120, 60)).Id
    Psu3D.Camera.SetPosition -6, -6, 4
    Psu3D.Camera.LookAt 0, 0, 0.5

    i = Psu3D.Scene.AddRotatedBox(m, 0, 0, 1, 1, 1, 2)

    For t = 0 To 6.2 Step 0.05
        Psu3D.Scene.SetAngle i, t
        Psu3D.BeginFrame 0.05
        Psu3D.RenderScene
        Psu3D.EndFrame
        Slide1.Shapes("timer").TextEffect.Text = Format$(Timer, "0.000")
        DoEvents
    Next t

    Psu3D.Shutdown
End Sub
```

## A 3D bar chart from cell values

```vba
Public Sub Chart(ByVal data As Variant)
    Dim i As Long, h As Single, m As Long

    Psu3D.Boot Shapes, 40, 40, 640, 380
    m = PMaterials.Create("bar", PCore.ColorPack(70, 130, 200)).Id
    Psu3D.Camera.SetPosition -8, -14, 9
    Psu3D.Camera.LookAt 0, 0, 2
    Psu3D.Renderer.PolyBudget = 400

    For i = LBound(data) To UBound(data)
        h = CSng(data(i))
        Psu3D.Scene.AddBox m, i * 1.4, -0.5, i * 1.4 + 1, 0.5, h, h
    Next i

    Psu3D.BeginFrame 0
    Psu3D.RenderScene
    Psu3D.EndFrame
End Sub
```

Passing the height as the thickness as well makes the bar reach down to zero, which is what a bar should do.

## Minimal first person

```vba
Dim body As PBody
Set body = Psu3D.CreateBody()
body.SetPosition 0, 0, 2

anchorX = GetSystemMetrics(SM_CXSCREEN) \ 2
anchorY = GetSystemMetrics(SM_CYSCREEN) \ 2
ShowCursor 0
SetCursorPos anchorX, anchorY

Do While ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 1
    dt = Timer - t0
    If dt < 0.001 Then dt = 0.001
    If dt > 0.05 Then dt = 0.05
    t0 = Timer

    GetCursorPos p
    Psu3D.Camera.AddAngles (p.X - anchorX) * 0.0022, -(p.Y - anchorY) * 0.0022
    SetCursorPos anchorX, anchorY

    fwd = 0: strafe = 0
    If KeyDown(vbKeyW) Then fwd = fwd + 1
    If KeyDown(vbKeyS) Then fwd = fwd - 1
    If KeyDown(vbKeyD) Then strafe = strafe + 1
    If KeyDown(vbKeyA) Then strafe = strafe - 1

    wishX = Psu3D.Camera.ForwardX * fwd + Psu3D.Camera.RightX * strafe
    wishY = Psu3D.Camera.ForwardY * fwd + Psu3D.Camera.RightY * strafe

    Psu3D.Scene.UpdateMotion dt
    body.Advance dt, wishX, wishY, KeyDown(vbKeySpace), KeyDown(vbKeyShift)
    If body.CarryYaw <> 0 Then Psu3D.Camera.AddAngles body.CarryYaw, 0

    Psu3D.Camera.SetPosition body.X, body.Y, body.Z + 1.62

    Psu3D.BeginFrame dt
    Psu3D.RenderScene
    Psu3D.EndFrame

    timerShp.TextEffect.Text = Format$(Timer, "0.00")
    DoEvents
Loop

ShowCursor 1
```

The pointer is read and warped back to a fixed anchor, which is relative mouse without capturing the device. Because the same anchor is used for both, its exact position never matters. `PDemo.bas` has the API declarations and a ready made `KeyDown`.

## A platform that goes back and forth, with you on it

```vba
p = sc.AddBox(metal, -1.4, -1, 1.4, 1, 5.4, 0.4)
sc.SetMotion p, 1, 0, 0, 5.5, 0.85, 0
```

And in the loop, before the body:

```vba
sc.UpdateMotion dt
body.Advance dt, wishX, wishY, up, down
```

The body is carried on its own. The order is what matters: the scene moves first.

## A disc that turns and takes you with it

```vba
d = sc.AddRotatedBox(metal, 0, 5, 3.6, 0.85, 6.4, 0.4, 0)
sc.SetSpin d, 0.95
```

And add `body.CarryYaw` to the camera yaw, or you turn around the centre of the disc while facing the same way the whole time.

## A lift driven by hand

```vba
zBefore = zNow
zNow = 3.6 + Sin(t * 0.7) * 3.2
sc.SetTopZ lift, zNow
If body.GroundObject = lift Then body.Nudge 0, 0, zNow - zBefore
```

`Nudge` shifts without touching velocity, because being carried is not being accelerated.

## An invisible collider

```vba
Set m = PMaterials.Create("clip", 0, pcSolid)
m.Visible = False
sc.AddBox m.Id, -10, 8, 10, 8.4, 4, 4
```

An invisible object is skipped during collection: never culled, never sorted, never drawn, and it spends no budget.

## Two views on one slide

```vba
Set cvMini = New PCanvas
cvMini.Init 20, 20, 250, 150

Set camMini = New PCamera

Set rdMini = New PRenderer
rdMini.Prefix = "mini_"
rdMini.Attach Shapes, cvMini, camMini
rdMini.PolyBudget = 48

' per frame, after the main view:
camMini.SetPosition body.X - cam.ForwardX * 9, body.Y - cam.ForwardY * 9, body.Z + 10
camMini.LookAt body.X, body.Y, body.Z + 1
rdMini.BeginFrame
sc.Render rdMini
rdMini.EndFrame
```

The different prefix is mandatory. Without it one renderer deletes the other's shapes.

## Loading a level from a file

```vba
If Not PLevel.ParseFile("C:\levels\obby.json", sc) Then
    MsgBox PLevel.Error
    Exit Sub
End If

body.SetPosition PLevel.SpawnX, PLevel.SpawnY, PLevel.SpawnZ
cam.SetAngles PLevel.SpawnYaw, 0
If PLevel.HasKillZ Then body.KillZ = PLevel.KillZ
If PLevel.Gravity > 0 Then body.Gravity = PLevel.Gravity
```

## Running the demo

```vba
PDemo.RunDemo                          ' obby.json, the hard level
PDemo.RunPool                          ' a pool, for testing water
PDemo.RunArena                         ' an arena, with a second view and movers
PDemo.RunFile "C:\levels\mine.json"    ' your level
```

The slide needs a WordArt shape named `timer`. It is the refresh pump, not a clock.

# Renderer and primitives

The renderer is where a 3D face becomes a PowerPoint shape. It is also where the primitives live, because there is no mesh object here: a box is not a thing that is stored, it is a call that happens, in the same way `fillRect` is on a 2D canvas.

```vba
Dim rd As PRenderer
Set rd = New PRenderer

rd.Attach Shapes, cv, cam
rd.PolyBudget = 140
rd.SeamFill = True
```

## Primitives

All of them draw immediately, with nothing stored and nothing allocated:

```vba
rd.DrawBox         matId, x1, y1, x2, y2, top, thickness
rd.DrawBoxLod      matId, x1, y1, x2, y2, top, thickness
rd.DrawBoxRotated  matId, cx, cy, halfW, halfH, top, thickness, angle
rd.DrawRamp        matId, x1, y1, x2, y2, zLow, zHigh, thickness, paX
rd.DrawFloor       matId, x1, y1, x2, y2, z
rd.DrawWall        matId, x1, y1, x2, y2, z, height
rd.DrawBillboard   matId, x, y, z, width, height
rd.DrawSpinner     matId, x, y, z, radius, phase
```

`DrawBox` submits only the faces the eye can see, at most three of the six. `DrawBoxLod` submits only the most visible face, and is what the scene uses once a box has become too small on screen for the other two to be worth a shape each.

`DrawBillboard` is a quad that always turns to face the camera. `DrawSpinner` is two double sided plates crossed at right angles, spinning on their own axis, which is the shape a pickup usually takes. The cross is deliberate: a single plate would vanish twice per rotation when seen edge on.

## Lower still

If the primitives do not fit, hand over the geometry:

```vba
rd.DrawQuad     x0,y0,z0, x1,y1,z1, x2,y2,z2, x3,y3,z3, matId
rd.DrawTriangle x0,y0,z0, x1,y1,z1, x2,y2,z2, matId
rd.DrawPolygon2D ...     ' in canvas space, for a HUD
```

`DrawPolygon2D` skips the whole 3D stage and draws straight in canvas points. It is what the demo's HUD uses, and that is deliberate: a text shape forces a z-order change every frame, and in PowerPoint that repaints the entire slide.

## The frame

```vba
rd.BeginFrame         ' swap the name bank
' ... draw ...
rd.EndFrame           ' delete the previous bank in one call
```

Between `BeginFrame` and `EndFrame` the new shapes are created with names from one bank while the previous frame is still on screen with names from the other. At `EndFrame` the old bank is deleted in a single `Range(...).Delete`.

That is double buffering. PowerPoint has no screen updating switch, so the screen cannot be frozen while drawing; what can be done is never leaving it empty.

## Budget

`PolyBudget` is the ceiling on the polygons a frame may spend. `AdaptBudget dt` raises the ceiling slowly when frames come in fast and drops it quickly when they run late, between the limits set by `SetBudgetRange`. `AutoBudget` turns that on and off.

For a scene of known size, pinning the budget above the worst case is better: a cut that never happens is a cut that can never flicker.

## Seams

Two neighbouring polygons leave a thread of background showing between them, because PowerPoint's antialiasing does not close the join. `SeamFill = True` outlines each polygon in its own colour, which closes the join at no extra COM cost, since the outline colour goes in the same call.

`EdgeInflate` does the same thing by pushing vertices outwards. It works, but it distorts small faces, so `SeamFill` is the better choice.

## Measuring

```vba
rd.Profiling = True
Debug.Print rd.ComSeconds, rd.ComOps, rd.PolyCount
```

`DryRun = True` runs the entire pipeline without touching a single shape. It is what the self test uses, and it is how you separate the cost of your own code from the cost of PowerPoint.

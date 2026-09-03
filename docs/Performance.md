# Performance

## The refresh pump

This is not an optimisation, it is the condition for anything appearing at all. Read it before touching the loop.

```vba
timerShp.TextEffect.Text = m_statText & "  " & Format$(Timer, "0.00")
DoEvents
```

Writing to a WordArt is what forces PowerPoint to repaint the slide during a show, and the repaint is how the frame becomes visible. Skip it, or write the same string twice in a row, and the picture stops updating even though the engine keeps running.

Three consequences:

1. **It has to fire every frame.** It cannot be throttled to every 100 ms.
2. **The value has to actually change.** That is why the raw clock is appended to the text: it guarantees the string differs.
3. **`DoEvents` is where the repaint happens.** Without it the request sits in the queue.

The pump and the `DoEvents` are timed together as one bucket, because they are one event: the write asks PowerPoint to repaint and the `DoEvents` is where it gets to.

## Double buffering

PowerPoint has no screen updating switch, so the screen cannot be frozen while drawing. What can be done is never leaving it empty.

The renderer keeps two banks of shape names. The new frame's shapes go in with names from one bank while the previous frame is still on screen with names from the other. At `EndFrame` the old bank is deleted in a single call:

```vba
target.Range(oldBankNames).Delete
```

One COM call to delete a hundred and forty shapes, instead of a hundred and forty.

## Polygon budget

```vba
rd.SetBudgetRange 120, 150
rd.PolyBudget = 140
rd.AutoBudget = False
```

`AdaptBudget dt` raises the ceiling slowly, by 2, when frames come in fast, and drops it quickly, by 8, when they run late. The hysteresis is asymmetric on purpose: three slow frames in a row are needed to shrink and ten fast ones to grow, otherwise the budget oscillates and background geometry flickers.

**For a scene of known size, pin it.** If the worst case, with everything on screen at once, costs 100 polygons, pinning at 140 means the draw list is never cut, and a cut that never happens is a cut that can never flicker.

`AutoBudget` is for a scene too large for any fixed number, where trading distant geometry for frame rate is the right deal.

The selection has hysteresis of its own: an object already on screen may overspend by a small margin rather than vanish. Without it, an object sitting on the budget line is cut on one frame and drawn on the next. Measured with a deliberately tight budget, fourteen disappearances over four hundred frames became none.

## Nothing is read back from the shapes

A rule of the library: the per frame path never queries a shape property. A COM object is expensive to interrogate, and everything we would ask about is already ours: position, size, colour, angle, all of it lives in an array.

Audited: the only shape reads in the whole engine are `PRenderer.Purge` and `PCanvas.FindShape`, and both run only at boot. Neither `PScene`, nor `PMaterials`, nor the renderer's primitives ever touch a shape.

## A frame that did not change

A body that ended the frame where it started changes nothing on screen. The demo keeps a flag:

```vba
If m_dirty Then
    Psu3D.BeginFrame dt
    Psu3D.RenderScene
    Psu3D.EndFrame
    m_dirty = False
End If
```

Skipping the frame skips every COM call in it: no delete, no `AddPolyline`, no fill. A player standing still costs nothing, which is the cheapest optimisation available here.

Careful: the refresh pump does **not** go inside that `If`. It runs always.

## Precomputed shading

The colour of a face is material times direction times fog band, and that is a table, not a calculation. `PMaterials.ShadeBand` is two bounds checks and one array read.

Only axis aligned faces use the table. A ramp falls to the slow path, which is the price of being able to point anywhere.

## Measuring properly

```vba
rd.Profiling = True
rd.DryRun = True        ' runs everything without touching a shape
```

With `DryRun` on you are measuring only your own code. The difference between the time with and without it is what PowerPoint charges, and that is the number that decides where optimising is worth anything.

The demo prints three of them on the status bar: `render`, `com` and `doevents`. If `doevents` dominates, shrinking the pipeline moves nothing and the lever is how much the repaint has to draw.

## Level of detail

`sc.LodSize = 44` makes a box whose projected diameter has fallen below 44 points draw only its most visible face. A box shows at most three faces, and the two minor ones become slivers long before the dominant one does.

Raise the number to buy frame rate, lower it to buy silhouette. The threshold has hysteresis, so an object drifting across it does not flip on alternate frames.

## Spatial index

Physics queries go through a uniform grid. See [Spatial index](Spatial-Index.md).

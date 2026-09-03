# Concepts

Five objects, each answering one question. Knowing which question belongs to which is knowing where to make a change.

| object | question it answers |
|---|---|
| `PCanvas` | where on the slide, and through what lens |
| `PCamera` | where am I looking from, and at what |
| `PRenderer` | how does a face become a PowerPoint shape |
| `PScene` | what exists in the world, and in what order to paint it |
| `PBody` | what moves, and is stopped by what exists |

## The canvas is not the screen

In an ordinary 3D engine the screen is the screen. Here it is not: the canvas is any rectangle of the slide, with an x, a y, a width and a height in points. That exists because on a slide you almost never want the whole thing. You want a viewport in a corner, beside some text, inside a frame that was already part of the design.

The direct consequence is that two canvases, two cameras and two renderers can share one scene and one slide. Each renderer names its shapes from its own prefix, so neither deletes the other's frame. That is how the demo's arena has a small second view in the corner.

The vertical field of view is derived from the shape of the canvas, so changing the height does not distort the image.

## The scene stores data, not objects

`PScene` does not hold a list of objects. It holds parallel arrays: one of kinds, one of materials, one of x1, one of y1, and so on. Adding a thousand boxes costs a thousand array slots and no allocations per frame.

The price is that you never hold *the box*, you hold a `Long` that indexes it. The whole API works that way:

```vba
Dim box As Long
box = sc.AddBox(stone, -2, -2, 2, 2, 1, 0.5)
sc.SetMaterial box, ice
sc.MoveBy box, 0, 0, 3
```

An id is good for the life of the object, because nothing ever reorders the arrays. The exception is `Remove`, which returns the slot to the store: after removing it, the id names whatever took its place. To hide something you intend to bring back, use `SetActive`, which keeps the slot reserved.

## There is no mesh

A box is not a thing that is stored. It is a call that happens, in the same way a rectangle is on a 2D canvas. There is no stored vertex, no buffer, no model transform.

That is why the primitives are commands on the renderer:

```vba
rd.DrawBox matId, x1, y1, x2, y2, top, thickness
```

rather than methods on a mesh object. And it is why drawing a thousand boxes allocates nothing.

## The material decides almost everything

Colour, whether it is solid, whether you pass through, friction, bounce, buoyancy, whether it can be climbed: all of it lives on the material, not on the object. Two objects with the same material behave the same, and swapping an object's material changes how it is drawn and how it collides in the same line.

```vba
sc.SetMaterial bridge, brittle
```

See [Materials](Materials.md).

## The body reports, it does not decide

`PBody` knows how to walk, jump, step up, climb, swim and be carried by a platform. What it does not know is what any of it means. It hands back the list of triggers it touched and refuses to guess whether they were coins, checkpoints or mines.

```vba
For k = 0 To body.TouchCount - 1
    idx = body.TouchAt(k)
    If sc.TagOf(idx) = MY_CHECKPOINT Then
        ' your rule
    End If
Next k
```

For the same reason the scene carries a `tag` per object, a number of yours, rather than an `IsCheckpoint`. The rules of your domain stay in your code.

## The order of a frame

```
UpdateMotion       whatever moves on its own moves
body.Advance       the body is carried, accelerated, pushed out, and falls
camera.SetPosition the camera follows the body
BeginFrame         swap the shape bank and adjust the budget
RenderScene        cull, budget, order, draw
EndFrame           retire the previous frame
refresh pump       PowerPoint is asked to repaint
DoEvents           and the repaint happens
```

The order of the first two matters. A platform that has already moved can carry the body; one that moves afterwards is one the body spends a frame standing beside.

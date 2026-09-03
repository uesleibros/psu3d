# Scene

`PScene` holds the world and answers two questions per frame: what the camera can see, and what sits inside this box of space.

## Adding

```vba
Dim box As Long, ramp As Long, disc As Long

box = sc.AddBox(stone, -2, -2, 2, 2, 1, 0.5)
'                     x1  y1  x2 y2 top thickness

ramp = sc.AddRamp(stone, 0, 0, 6, 3, 0, 2, 0.4, paX)
'                       x1 y1 x2 y2 low high thickness axis

disc = sc.AddRotatedBox(metal, 0, 5, 3.6, 0.85, 1, 0.4, 0)
'                             cx cy halfW halfH top thickness angle

panel = sc.AddBillboard(decor, 0, 8, 2, 3, 2)
'                             x  y  z  w  h

coin = sc.AddSpinner(gold, 0, 3, 1.2, 0.32)
'                         x  y   z   radius
```

Each returns a `Long`, which is the object's id. An id is good for the life of the object, because nothing ever reorders the arrays.

## Editing

```vba
sc.SetMaterial box, ice
sc.MoveBy box, 0, 0, 3
sc.SetTopZ box, 4.2
sc.SetAngle disc, 1.2
sc.SetActive box, False        ' gone from the screen and from collision, slot kept
sc.Remove box                  ' gone for good, and the slot returns to the store
```

The difference between `SetActive False` and `Remove` matters: after `Remove`, the id names whoever took the slot. Use `SetActive` for something you intend to bring back.

## Motion

Two kinds, and the scene drives both:

```vba
sc.SetMotion platform, 1, 0, 0, 5.5, 0.85, 0
'                     axis x,y,z  amp  speed phase

sc.SetSpin disc, 1.15
'                radians per second
```

`SetMotion` makes a platform oscillate on a sine, which slows it at the ends, which is what makes it possible to step onto one on purpose. Call `sc.UpdateMotion dt` once per frame; it returns how many objects moved, so a caller that only redraws when something changed knows when to redraw.

Whoever was standing on it goes along. See [Body and physics](Body-And-Physics.md#being-carried).

## Queries

```vba
n = sc.QueryBox(x1, y1, x2, y2)
n = sc.QueryRadius(x, y, radius)

For k = 0 To n - 1
    idx = sc.ResultAt(k)
Next k
```

Queries go through the [spatial index](Spatial-Index.md), so they stay cheap with thousands of objects.

To ask about one object:

```vba
z = sc.TopZAt(idx, x, y)              ' surface height there, interpolating a ramp
b = sc.Blocks(idx, True)              ' does the material block, arriving from above?
ok = sc.ContainsXY(idx, x, y)         ' is the point inside, respecting rotation
sc.GetSpanZ idx, bottom, top
sc.GetBoundsXY idx, x1, y1, x2, y2
sc.GetOrientedBox idx, cx, cy, hw, hh, ang
```

`GetOrientedBox` returns the box in its own frame, which is what a solver needs to treat a turned slab as the slab it is rather than as the square that encloses it.

## Tag

One number of yours per object:

```vba
sc.SetTag gate, 7
If sc.TagOf(idx) = 7 Then ...
```

The scene deliberately does not know what a checkpoint, a door or a waypoint is. It carries a number and lets the caller decide what it stands for, which is the difference between a library you can build a game on and a library you can only build this game on.

## Drawing

```vba
drawn = sc.Render(rd)
```

`Render` does everything: frustum culling, spending the budget, working out the paint order, and drawing. It returns how many objects were painted.

Afterwards, `DrawnCount` and `DrawnAt(i)` give back the order the frame used. It is the one thing about a frame that cannot be checked by looking at the shapes afterwards.

## Level of detail

`LodSize` is the projected size, in slide points, below which a box drops to drawing only its most visible face. A box shows at most three faces, and the two minor ones become slivers long before the dominant one does. Past that point they cost a shape each to contribute a few points of colour.

The threshold has hysteresis: an object has to fall well under the line to lose its faces and climb well over it to get them back. With a single threshold, an object drifting across it flips on alternate frames, which reads as a shape blinking.

Zero turns the whole thing off.

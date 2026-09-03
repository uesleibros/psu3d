# Changelog

## 1.2.0

### Three more example levels

`minimal.json` is the smallest document that is still a level, to copy when starting your own. `pool.json` is a swimming pool, for buoyancy, drag, swimming, diving and ladders out of the water. `showcase.json` has one of every material behaviour and every primitive, from the kerb and the wall that differ only in their `step` to the turntable that turns your view with it.

### A level checker

`tools/check-level.py` validates a level against the schema, and then, using the physics the level itself declares, floods outwards from the spawn to find any surface that cannot be reached. A level that parses perfectly and cannot be finished is the mistake nothing else reports.

It caught real problems in the new levels before they shipped: a diving board only reachable by a 2.4 metre jump, a rail floating above the deck it was meant to sit on, and a staircase whose treads stacked in place instead of marching along.

### Documentation

Three new pages: [Examples](https://github.com/uesleibros/psu3d/wiki/Examples), [Troubleshooting](https://github.com/uesleibros/psu3d/wiki/Troubleshooting) and [FAQ](https://github.com/uesleibros/psu3d/wiki/FAQ). A `CONTRIBUTING.md` stating the rules that are not negotiable and the standard of evidence expected for anything that can be wrong rather than merely slow.

### Everything in English

The changelog, the level checker, the wiki sync tool, the repository description and the material names in `obby.json` were all still Portuguese. They are not now.

## 1.1.0

### Flicker fixes

Two things could make an object blink or vanish between two frames that were otherwise identical. Both were found by measuring temporal stability along a smooth camera path, not by looking at a single frame.

**Level of detail had a single threshold.** An object drifting across it flipped between three faces and one on alternate frames, which reads as a shape blinking. It now has two thresholds: it has to fall well under the line to lose its faces and climb well over it to get them back. Measured over 260 frames, four pops became none.

**The budget cut had no hysteresis.** An object sitting exactly on the polygon budget was cut on one frame and drawn on the next, disappearing and returning while nothing about it changed. An object already on screen may now overspend by a small, bounded margin rather than vanish. Measured with a deliberately tight budget over 400 frames, fourteen disappearances became none.

The residual five order flips per 260 frames come from genuine ordering cycles. Five alternative tiebreak rules were measured against them, including breaking by strongly connected component and by least visible damage, and none improved on the current one.

`PScene.WasReduced(slot)` was added alongside `DrawnAt`, so the level of detail decision can be observed. It is what the new `CheckStability` block tests.

### The demo declares its own input

`PDemo` now declares `GetCursorPos`, `SetCursorPos`, `ShowCursor` and `GetSystemMetrics` at the top of the file, next to the keyboard declaration it already had. Mouse look reads the pointer, measures its travel from a fixed anchor, and warps it back, which is relative mouse without capturing the device. Because the same anchor is used to read and to warp, its exact position never matters, so no conversion between screen pixels and slide points is needed.

Psu3D now has exactly one external requirement, and only for one module: `PLevel` needs a JSON parser.

### Documentation

Rewritten in English and reorganised into hand written guides plus twelve generated reference pages, one per module, produced from the docstrings in the source so no signature in the documentation can drift from the code.

## 1.0.1

### The core stopped depending on an external cursor module

`PCanvas` had four members that called an external cursor module: `CursorX`, `CursorY`, `CursorInside` and `CenterCursor`. Because VBA compiles the whole project at once, that forced **every** project importing the core to import a cursor module as well, whether it wanted a mouse or not. It was a real dependency, hidden behind four conveniences.

All four were removed. The canvas maps coordinates and does not read hardware: whoever has a pointer subtracts `CenterX` and `CenterY`, which is the same arithmetic and works with any source of pointer position.

**Migration.** If you used any of the four:

| before | now |
|---|---|
| `cv.CursorX` | `MyCursorX - cv.X` |
| `cv.CursorY` | `MyCursorY - cv.Y` |
| `cv.CursorInside` | `cv.Contains(MyCursorX, MyCursorY)` |
| `cv.CenterCursor` | `WarpCursorTo cv.CenterX, cv.CenterY` |

And for mouse look, directly:

```vba
dx = MyCursorX - cv.CenterX
dy = MyCursorY - cv.CenterY
```

## 1.0.0

First public release.

### Core

Canvas with a position and a size on the slide, camera with real yaw and pitch, renderer with five plane frustum clipping and backface culling, scene in parallel arrays, materials with precomputed shading, directional light and banded fog.

### Physics

`PBody` as a class of the engine: walking, jumping, stepping up by material, climbing, swimming with buoyancy and drag, bouncing, and being carried by a platform that travels **and** by one that turns. The body reports the triggers it touched and does not decide what they mean.

### Levels

`PLevel` reads and writes a scene as JSON. The round trip gives back the same level plus the edit. Numbers are written with `Str` so the output does not depend on the machine's regional settings.

### Depth sorting

Painter's algorithm with separating planes and a topological sort. A screen space filter that discards 60% of pairs before the comparator, a turned box AABB that follows its angle, and a cycle tiebreak by least debt.

Measured against an exact oracle: 0.00% visible artefacts on `obby.json` against 1.85% before, and 1.01% on an adversarial scene against 6.15% before.

### Spatial index

A uniform grid in XY for physics queries, with whatever moves and whatever is too large staying out and being swept. Between 19 and 79 times fewer objects touched per query in large scenes. Proved equivalent to the full sweep across 120,000 queries.

### Robustness

Four paths that ended in a runtime error were closed: a `Long` overflow in the grid with an absurd coordinate, an enormous `repeat` in a level file, a material name with a quote breaking the saved document, and `budget` and `tag` values outside the range of a `Long`.

### Self test

Assertions covering maths, colour, lighting, materials, canvas, camera, scene, index, motion, body, draw order, renderer, levels, saving and robustness.

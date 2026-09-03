# Spatial index

`QueryBox` used to be a sweep over every object, and `PBody` calls it once per collision pass, about five times a frame. With 2,500 objects that is 12,500 iterations per frame just to find out what is underfoot.

There is now a uniform grid in XY, built by counting and prefix summing into a flat array: no list per cell, no allocation while a frame is running.

| objects | sweep touches | grid touches |
|---|---|---|
| 40 | 40 | 24 |
| 576 | 576 | 30 |
| 2500 | 2500 | 31 |

## The cell is sized from the average object

A terrain of one metre tiles indexed in four metre cells would put sixteen tiles in every bucket and index nothing.

## Two things stay out and are swept

**Whatever moves.** Anything that has called `MoveBy`, `SetTopZ`, `SetAngle`, `SetMotion` or `SetSpin` leaves the grid and does not return. Something that moved will very likely move again, and rebuilding the grid every frame for three platforms costs more than the grid saves.

**Whatever is too large.** A ground plane appears in every cell. Indexing it fills each bucket with the one object every query was going to reach anyway, so past 32 cells it is swept alongside the platforms. In a scene with a giant floor that drops the index from 5,879 entries to 2,158.

`DynamicCount` and `CellSize` report what stayed out and how fine the grid ended up.

## An index can be wrong, not merely slow

A query that misses an object is a body falling through the floor. So it was not sampled, it was compared: 120,000 random queries against the full sweep, on scenes built deliberately out of what breaks grids, meaning objects far larger than a cell, objects outside the indexed area, objects switched off, and objects that have moved. Zero divergences.

The same test runs in VBA, in `CheckIndex`.

## What stays linear

The frustum cull in `Render`, one pass per frame. The frustum is a cone, and the rectangle covering it touches so many cells that the lookup would cost more than the sweep. That is a decision, not an oversight.

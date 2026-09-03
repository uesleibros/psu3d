# Known limits

What the library does not do, and why. None of this is a bug; it is the boundary.

## No Z buffer

There is no per pixel depth. Whole polygons are sorted and painted back to front.

**Geometry that interpenetrates has no correct order.** Two boxes passing through each other have no "which is in front", and the best the library can do is pick the least visible mistake.

In practice, measured against an exact oracle: 0.00% visible artefacts on a real level, 1.01% on a scene built deliberately to be hostile, and most of that is genuine cycles. Detail in [Depth sorting](Depth-Sorting.md).

Practical workaround: avoid geometry that passes through other geometry. Touching is easy to order, overlapping is not.

## No rasteriser of our own

The polygon, already transformed, clipped and coloured, is handed to PowerPoint. The fill is theirs. That is what makes a Z buffer impossible, and it is also what makes the library fit in VBA.

## No textures

Flat colour per face. No UVs, no perspective correct interpolation. A per face outline is what exists for detail.

## Flat shading

One colour per face, from the precomputed table. No Gouraud, no Phong, no per vertex normals.

## Rotation only in yaw

`AddRotatedBox` turns on the vertical axis. There is no object pitch or roll, and the camera has no roll.

`DrawQuad` and `DrawTriangle` take three or four arbitrary 3D vertices, so any orientation can be drawn by hand. What is missing is a stored mesh with a model matrix.

## No meshes

There is no OBJ loading and no vertex buffer. Primitives are generated per call. That is a design decision: with no stored mesh, drawing a thousand boxes allocates nothing.

## The frustum cull is linear

One pass over every object per frame. The spatial index covers physics queries, not the cull. The frustum is a cone, and the rectangle covering it touches so many cells that the lookup would cost more than the sweep.

With 2,500 objects that pass is real. If it becomes a problem, the answer is a bounding volume hierarchy, not a grid.

## The grid is capped at 4,096 cells per axis

An absurdly large world degrades into a slow lookup rather than failing. That clamp is deliberate.

## HUD text is not text

The engine draws no text. A text shape forces a z-order change every frame, and in PowerPoint that repaints the entire slide. The demo's HUD is made of rectangles through `DrawPolygon2D`.

The only text shape in the loop is the refresh pump, and it exists precisely because it repaints the slide.

## Single, not Double

All geometry is `Single`. That is about seven significant digits, which is plenty for a world a few hundred units across and not enough for coordinates in the millions. If your scene is enormous, move the origin rather than growing the numbers.

## Windows only

`QueryPerformanceCounter` for the clock, and `GetAsyncKeyState`, `GetCursorPos`, `SetCursorPos` and `ShowCursor` for the demo's input, are Windows API calls. The rendering core itself uses none of them and would run anywhere VBA runs, but the clock and the demo would need replacing.

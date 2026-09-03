# FAQ

## Is this really 3D, or a trick?

Really 3D. Real view transform with yaw and pitch, perspective projection, five plane frustum clipping with polygons actually clipped rather than dropped, and backface culling.

What it does not have is a Z buffer or a rasteriser of its own. The full answer, with the code, is in [How the 3D works](How-The-3D-Works.md).

## How many objects can it handle?

There is no fixed ceiling. Every store grows by doubling.

The practical limit is what PowerPoint will repaint per frame, not what the library will hold. The polygon budget exists exactly to keep that number under control while the object count grows past it.

Physics queries go through a [spatial index](Spatial-Index.md), so they stay cheap into the thousands. The frustum cull is still one linear pass per frame, so that is the number that grows with the scene.

## Does it need any references or DLLs?

No. Import the files and it compiles. The Windows API functions it uses are declared inside the modules that use them.

## Can I use it without the game parts?

Yes, and that is a first class case. The eight core modules know nothing about players, physics or levels. A terrain, a bar chart or a 3D figure on a slide uses the same core and never imports `PBody` or `PLevel`.

## Can I have two views on one slide?

Yes. Two canvases, two cameras, two renderers, one scene. Each renderer needs its own `Prefix`, or one will delete the other's shapes. There is a recipe in [Recipes](Recipes.md#two-views-on-one-slide).

## Why is my HUD text not drawing?

The engine draws no text. A text shape forces a z-order change every frame, and in PowerPoint that repaints the entire slide. Build the HUD out of rectangles through `DrawPolygon2D`, which is what the demo does.

The one exception is the refresh pump, and that exists precisely because it repaints the slide.

## Why is the timer shape mandatory?

It is not a timer. Writing to a WordArt is the only reliable way to make PowerPoint repaint during a slide show. Without it the engine runs and the picture never updates.

## Can I load a model from an OBJ file?

Not out of the box. There is no mesh storage: primitives are generated per call, which is why drawing a thousand boxes allocates nothing.

`DrawQuad` and `DrawTriangle` take arbitrary 3D vertices, so nothing stops you from writing a loader that pushes triangles at them each frame. What you would be giving up is the depth sorting, which works on objects and would not know about your triangles.

## Why is everything `Single` and not `Double`?

Seven significant digits is plenty for a world a few hundred units across, and half the memory across thirty parallel arrays. If your scene is enormous, move the origin rather than growing the numbers.

## Can I change gravity, or make a low gravity level?

Yes, on the body, or in the `level` block of a level file:

```vba
body.Gravity = 6
body.JumpSpeed = 4.5
```

`tools/check-level.py` reads whatever the level declares, so a low gravity level is checked against low gravity reach.

## Does it work on Mac PowerPoint?

The rendering core uses nothing platform specific. The clock and the demo's input use Windows API calls, so those would need replacing. Nobody has tried it.

## Why not use shapes with 3D effects instead?

PowerPoint's own 3D rotation is per shape and has no camera, no depth sorting and no perspective shared between shapes. Two shapes rotated the same way do not agree about where the viewer is, so they cannot occlude each other correctly. That is the whole problem this library solves.

## How is the depth order decided?

An axis aligned separating plane between each pair, the side the eye is on deciding which is behind, and a topological sort over the resulting constraints. When no order satisfies everything, which happens with cycles, the object owing the fewest constraints is forced into place.

Measured against an exact oracle, it is 0.00% wrong on the bundled obstacle course. The detail, including the bug that led to the current rule, is in [Depth sorting](Depth-Sorting.md).

## Can I contribute?

Yes. See [CONTRIBUTING](https://github.com/uesleibros/psu3d/blob/main/CONTRIBUTING.md). The short version: no `On Error`, no COM on the hot path, and anything that can be wrong rather than merely slow has to come with a measurement.

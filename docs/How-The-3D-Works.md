# How the 3D works

A fair question about a 3D engine inside PowerPoint is whether it is really 3D or a fake built on scaling things by distance. It is really 3D, and it is possible to point at exactly where.

## The four stages that define 3D

### 1. View transform

`PCamera.WorldToView` is an orthonormal change of basis with yaw and pitch:

```
fwd  = (dx*cosYaw + dy*sinYaw) * cosPitch + dz * sinPitch
side = -dx*sinYaw + dy*cosYaw
up   = dz*cosPitch - (dx*cosYaw + dy*sinYaw) * sinPitch
```

That is a view matrix written by hand. The important part: **pitch is a real rotation**, not a vertical shift of the image. Shifting the image is the cheap trick a fake engine uses, and it gives itself away the moment you look far up or far down, because the perspective does not follow.

There is a separate `ScreenShiftY` on the camera, but it is there for head bob and weapon recoil, not for looking up.

### 2. Perspective projection

`PCanvas.Project`:

```
invD = focal / fwd
x = centreX + side * invD
y = centreY - up   * invD
```

A textbook pinhole camera. It is why parallels converge correctly and why the distortion towards the edges is the right distortion.

### 3. Frustum clipping

Five planes in view space, near plus the four sides, clipped in the Sutherland and Hodgman style. A polygon crossing the near plane is **clipped**, generating new vertices, not discarded.

This is the dividing line. Anything that merely scales sprites by distance does not have it, and gives itself away when a wall passes through the eye.

### 4. Backface culling

The sign of the dot product between the face normal and the vector from the face to the eye, with the normal coming from a cross product. The normal is deliberately left unnormalised: dividing by its own length costs a square root and three divides per face, and the test only needs the sign, which scaling by a positive number cannot change.

On top of that, directional lighting per normal and fog by depth.

## Where the boundary is

Two things a complete 3D engine has and Psu3D does not.

### There is no Z buffer

There is no per pixel depth. Whole polygons are sorted and painted back to front, which is the painter's algorithm.

The concrete consequence: **geometry that interpenetrates has no correct order**, because there is no answer to which of two boxes that pass through each other is in front. That is why the sorting became the hardest part of the library, and why it breaks cycles rather than resolving them. See [Depth sorting](Depth-Sorting.md).

A Z buffer would end that in one line, and is impossible here for the following reason.

### We do not rasterise

The polygon, already transformed, clipped and coloured, is handed to PowerPoint, and PowerPoint fills it. Psu3D does geometry and lighting; the fill is theirs. It is the same contract an SVG renderer has.

### The rest

Flat colour per face, no textures and no UVs. Flat shading, no Gouraud and no Phong. Object rotation only in yaw, although `DrawQuad` and `DrawTriangle` accept three or four arbitrary 3D vertices, so any orientation can be drawn by hand; what is missing is a stored mesh with a model matrix.

## The honest comparison

This is not three.js. three.js has a GPU, a Z buffer, textures and meshes.

Psu3D is a software 3D renderer of the generation before hardware acceleration: correct transformation and clipping, painter's algorithm, flat shading, primitives generated rather than meshes loaded. Which is literally how 3D worked before 3D cards existed.

So: real 3D, yes. From 1996, running inside PowerPoint.

# Lighting and fog

One global directional light and one global fog. Both apply to the whole scene.

```vba
PLighting.SetLight 0.38, 0.42, 0.82
PLighting.SetLightIntensity 0.55, 0.45      ' ambient, diffuse
PLighting.SetFog 8, 34, PCore.ColorPack(160, 185, 210)
PLighting.SetFogSteps 24
PLighting.EnableFog True
```

The light direction is normalised on the way in, so any vector will do.

`SetLightIntensity` splits the lighting in two: ambient is what a face receives with its back to the light, diffuse is what it gains by facing it. The two adding to somewhere near 1 gives natural contrast; too much ambient flattens the scene.

## Fog in bands

`SetFog` takes the distance at which fog begins, the distance at which it is total, and the colour. Outside that range there is no blending.

`SetFogSteps` is the detail that matters for performance. Fog is not continuous: it is quantised into a number of bands, and the colour of each band is precomputed per material and per face direction. With 24 bands, the colour of any face is an array read rather than an interpolation.

Too few bands and visible banding appears on the ground. Too many and the table grows. The ceiling is `P_MAX_FOG_STEPS`, which is 64.

Match the fog colour to the canvas backdrop colour. If they differ, the horizon becomes a line.

```vba
cv.BackColor = PCore.ColorPack(160, 185, 210)
PLighting.SetFog 8, 34, PCore.ColorPack(160, 185, 210)
```

## Revision

`PLighting` keeps a counter. Every change increments it and notifies `PMaterials`, which throws the shading table away. Nothing ever needs invalidating by hand: changing the light mid game simply works, and costs one table rebuild on the next frame.

## Direct

If you want the colour of a face yourself:

```vba
f = PLighting.ShadeFactor(nx, ny, nz)              ' 0 to 1
col = PLighting.ShadeColor(baseCol, nx, ny, nz)
col = PLighting.ShadeColorByDir(baseCol, pdPosZ)   ' axis aligned face
col = PLighting.ApplyFog(col, distance)
```

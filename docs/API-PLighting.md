# PLighting

**Directional light and global fog**

Holds the single directional light and the distance fog every material is shaded against, and exposes the flat shading maths the renderer runs per face. A revision counter is bumped on every change so cached shading tables can detect that they went stale.

> Fog is quantised into bands so materials can precompute one colour per band instead of blending on every face.

## Index

**Lifetime.** [`ResetDefaults`](#resetdefaults)

**Directional light.** [`SetLight`](#setlight), [`SetLightIntensity`](#setlightintensity), [`LightX`](#lightx), [`LightY`](#lighty), [`LightZ`](#lightz), [`Ambient`](#ambient), [`Diffuse`](#diffuse)

**Shading.** [`ShadeFactor`](#shadefactor), [`ShadeColor`](#shadecolor), [`ShadeColorByDir`](#shadecolorbydir), [`DirectionNormal`](#directionnormal)

**Distance fog.** [`SetFog`](#setfog), [`SetFogSteps`](#setfogsteps), [`EnableFog`](#enablefog), [`FogEnabled`](#fogenabled), [`FogStart`](#fogstart), [`FogEnd`](#fogend), [`FogColor`](#fogcolor), [`FogSteps`](#fogsteps), [`FogStepOf`](#fogstepof), [`ApplyFog`](#applyfog), [`ApplyFogStep`](#applyfogstep)

**Change tracking.** [`Revision`](#revision)

## Members

### ResetDefaults

```vba
Public Sub ResetDefaults()
```

Restores the stock lighting and fog setup.

Called automatically the first time any accessor is touched, so a fresh project renders sensibly without configuration.

### SetLight

```vba
Public Sub SetLight(ByVal dx As Single, ByVal dy As Single, ByVal dz As Single)
```

Points the global light in a new direction.

| parameter | what it is |
|---|---|
| `dx` | The X component of the direction the light travels towards. |
| `dy` | The Y component of the direction. |
| `dz` | The Z component of the direction. |

> The vector is normalised internally; a zero length vector is ignored.

### SetLightIntensity

```vba
Public Sub SetLightIntensity(ByVal ambient As Single, ByVal diffuse As Single)
```

Sets how much light reaches unlit and lit surfaces.

| parameter | what it is |
|---|---|
| `ambient` | The floor brightness applied to every face. |
| `diffuse` | How much the light direction adds on top of the ambient floor. |

### LightX

```vba
Public Property Get LightX() As Single
```

Reads the X component of the normalised light direction.

**Returns.** The direction component.

### LightY

```vba
Public Property Get LightY() As Single
```

Reads the Y component of the normalised light direction.

**Returns.** The direction component.

### LightZ

```vba
Public Property Get LightZ() As Single
```

Reads the Z component of the normalised light direction.

**Returns.** The direction component.

### Ambient

```vba
Public Property Get Ambient() As Single
```

Reads the ambient floor brightness.

**Returns.** The ambient term.

### Diffuse

```vba
Public Property Get Diffuse() As Single
```

Reads the diffuse contribution of the light.

**Returns.** The diffuse term.

### ShadeFactor

```vba
Public Function ShadeFactor(ByVal nx As Single, ByVal ny As Single, ByVal nz As Single) As Single
```

Computes the brightness multiplier for a surface normal.

| parameter | what it is |
|---|---|
| `nx` | The X component of the unit normal. |
| `ny` | The Y component of the unit normal. |
| `nz` | The Z component of the unit normal. |

**Returns.** The multiplier to apply to the material colour.

### ShadeColor

```vba
Public Function ShadeColor(ByVal baseCol As Long, ByVal nx As Single, ByVal ny As Single, ByVal nz As Single) As Long
```

Applies flat directional shading to a colour.

| parameter | what it is |
|---|---|
| `baseCol` | The unlit material colour. |
| `nx` | The X component of the unit normal. |
| `ny` | The Y component of the unit normal. |
| `nz` | The Z component of the unit normal. |

**Returns.** The shaded colour, before fog.

### ShadeColorByDir

```vba
Public Function ShadeColorByDir(ByVal baseCol As Long, ByVal dirKey As PDirection) As Long
```

Applies flat shading for one of the six axis-aligned normals.

| parameter | what it is |
|---|---|
| `baseCol` | The unlit material colour. |
| `dirKey` | The PDirection entry naming the face orientation. |

**Returns.** The shaded colour, before fog.

### DirectionNormal

```vba
Public Sub DirectionNormal(ByVal dirKey As PDirection, ByRef outX As Single, ByRef outY As Single, ByRef outZ As Single)
```

Expands a PDirection entry into its unit normal.

| parameter | what it is |
|---|---|
| `dirKey` | The direction key to expand. |
| `outX` | Receives the X component. |
| `outY` | Receives the Y component. |
| `outZ` | Receives the Z component. |

### SetFog

```vba
Public Sub SetFog(ByVal startDist As Single, ByVal endDist As Single, ByVal col As Long)
```

Configures the distance fog band.

| parameter | what it is |
|---|---|
| `startDist` | The depth at which fog begins to tint geometry. |
| `endDist` | The depth at which geometry is fully replaced by the fog colour. |
| `col` | The fog colour, which should match the slide background. |

### SetFogSteps

```vba
Public Sub SetFogSteps(ByVal steps As Long)
```

Sets how many discrete bands the fog is quantised into.

| parameter | what it is |
|---|---|
| `steps` | The band count, clamped to 1..P_MAX_FOG_STEPS. |

> More bands look smoother but grow every cached shading table.

### EnableFog

```vba
Public Sub EnableFog(ByVal enabled As Boolean)
```

Turns distance fog on or off without losing its settings.

| parameter | what it is |
|---|---|
| `enabled` | True to fog geometry by depth, False to render it flat. |

### FogEnabled

```vba
Public Property Get FogEnabled() As Boolean
```

Reports whether distance fog is currently applied.

**Returns.** True when fog is enabled.

### FogStart

```vba
Public Property Get FogStart() As Single
```

Reads the depth at which fog starts.

**Returns.** The near fog distance.

### FogEnd

```vba
Public Property Get FogEnd() As Single
```

Reads the depth at which fog becomes opaque.

**Returns.** The far fog distance, which doubles as the render far plane.

### FogColor

```vba
Public Property Get FogColor() As Long
```

Reads the fog colour.

**Returns.** The packed fog colour.

### FogSteps

```vba
Public Property Get FogSteps() As Long
```

Reads how many bands the fog is quantised into.

**Returns.** The band count.

### FogStepOf

```vba
Public Function FogStepOf(ByVal depth As Single) As Long
```

Maps a view depth onto its fog band.

| parameter | what it is |
|---|---|
| `depth` | The distance from the camera in world units. |

**Returns.** The band index, from 0 to FogSteps.

### ApplyFog

```vba
Public Function ApplyFog(ByVal col As Long, ByVal depth As Single) As Long
```

Blends a colour towards the fog colour by depth.

| parameter | what it is |
|---|---|
| `col` | The shaded surface colour. |
| `depth` | The distance from the camera in world units. |

**Returns.** The fogged colour.

### ApplyFogStep

```vba
Public Function ApplyFogStep(ByVal col As Long, ByVal bandIdx As Long) As Long
```

Blends a colour towards the fog colour by band index.

| parameter | what it is |
|---|---|
| `col` | The shaded surface colour. |
| `bandIdx` | The fog band, from 0 to FogSteps. |

**Returns.** The fogged colour.

> This is the form used when baking shading tables, where the depth is already quantised.

### Revision

```vba
Public Property Get Revision() As Long
```

Reads the counter bumped whenever the light or fog changes.

**Returns.** The current revision.

> Cached shading tables compare this against the revision they were baked at to decide whether to rebuild.

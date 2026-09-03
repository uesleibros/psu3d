# PCore

**Types, maths, colour and clock**

The layer everything else is written on: the shared types and enumerations, scalar and angle maths, a deterministic random source, colour arithmetic and a clock fine enough to time one frame. Nothing here knows about a slide, a shape or a scene, so any of it can be used on its own.

> Types live in a standard module rather than a class because VBA forbids a class from exposing, in a public signature, a type it declared itself. Four small modules were merged into this one: they had no dependencies of their own, were imported together every time, and separating them bought nothing but four more names in the project tree.

## Enumerations

### `PDirection`

Axis-aligned normal keys used to index precomputed shading tables.

| value | number |
|---|---|
| `pdNone` | -1 |
| `pdPosX` | 0 |
| `pdNegX` | 1 |
| `pdPosY` | 2 |
| `pdNegY` | 3 |
| `pdPosZ` | 4 |
| `pdNegZ` | 5 |

pdNone marks a face whose normal is arbitrary, forcing the renderer onto the slower per-face lighting path.

### `PCollision`

How a material answers the physics query "can a body pass through me?".

| value | number |
|---|---|
| `pcGhost` | 0 |
| `pcSolid` | 1 |
| `pcOneWay` | 2 |
| `pcTrigger` | 3 |

pcGhost is purely decorative, pcSolid blocks from every side, pcOneWay blocks only a body landing on its top face, and pcTrigger reports the overlap without ever stopping the body.

### `PAxis`

Which horizontal axis a ramp climbs along.

| value | number |
|---|---|
| `paX` | 0 |
| `paY` | 1 |

### `PObjectKind`

The primitive a scene object is drawn as.

| value | number |
|---|---|
| `pkBox` | 0 |
| `pkRotatedBox` | 1 |
| `pkRamp` | 2 |
| `pkBillboard` | 3 |
| `pkSpinner` | 4 |

### `PCanvasFit`

Strategy used by PCanvas.FitToSlide when reshaping a canvas against the slide bounds.

| value | number |
|---|---|
| `pfStretch` | 0 |
| `pfContain` | 1 |
| `pfCover` | 2 |

## Constants

| name | type | value | what it is |
|---|---|---|---|
| `P_MAX_FOG_STEPS` | Long | `64` | Highest number of fog bands a shading table may hold. |
| `P_INVALID_ID` | Long | `-1` | Identifier returned whenever a material lookup fails. |
| `P_PI` | Single | `3.14159265` | Ratio of a circle's circumference to its diameter. |
| `P_TWO_PI` | Single | `6.28318531` | A full turn expressed in radians. |
| `P_HALF_PI` | Single | `1.57079633` | A quarter turn expressed in radians. |
| `P_DEG2RAD` | Single | `0.0174532925` | Multiplier converting degrees into radians. |
| `P_RAD2DEG` | Single | `57.2957795` | Multiplier converting radians into degrees. |
| `P_EPSILON` | Single | `0.000001` | Tolerance below which a Single is treated as zero. |
| `P_WHITE` | Long | `16777215` | Opaque white. |
| `P_BLACK` | Long | `0` | Opaque black. |

## Index

**Clamping and interpolation.** [`Clamp`](#clamp), [`ClampLong`](#clamplong), [`Clamp01`](#clamp01), [`Lerp`](#lerp), [`InverseLerp`](#inverselerp), [`MoveTowards`](#movetowards), [`Approach`](#approach)

**Comparison helpers.** [`SignOf`](#signof), [`MinOf`](#minof), [`MaxOf`](#maxof), [`NearlyEqual`](#nearlyequal)

**Angles.** [`WrapAngle`](#wrapangle), [`AngleDelta`](#angledelta), [`ApproachAngle`](#approachangle)

**Planar queries.** [`SpansOverlap`](#spansoverlap), [`BoxesOverlap`](#boxesoverlap), [`PointInBox`](#pointinbox), [`DistSq2D`](#distsq2d)

**Deterministic random.** [`SeedRandom`](#seedrandom), [`RandomSeed`](#randomseed), [`RandomNext`](#randomnext), [`RandomRange`](#randomrange), [`RandomInt`](#randomint), [`RandomChance`](#randomchance)

**Colour packing.** [`ColorPack`](#colorpack), [`ColorRed`](#colorred), [`ColorGreen`](#colorgreen), [`ColorBlue`](#colorblue), [`ColorUnpack`](#colorunpack)

**Colour blending.** [`ColorScale`](#colorscale), [`ColorMix`](#colormix), [`ColorLighten`](#colorlighten), [`ColorDarken`](#colordarken), [`ColorLuminance`](#colorluminance), [`ColorGrayscale`](#colorgrayscale)

**Colour as text.** [`ColorFromHex`](#colorfromhex), [`ColorToHex`](#colortohex)

**Clock.** [`Seconds`](#seconds), [`Available`](#available)

## Members

### Clamp

```vba
Public Function Clamp(ByVal value As Single, ByVal minVal As Single, ByVal maxVal As Single) As Single
```

Constrains a value to an inclusive range.

| parameter | what it is |
|---|---|
| `value` | The value to constrain. |
| `minVal` | The lowest accepted result. |
| `maxVal` | The highest accepted result. |

**Returns.** The value pulled inside the range.

### ClampLong

```vba
Public Function ClampLong(ByVal value As Long, ByVal minVal As Long, ByVal maxVal As Long) As Long
```

Constrains an integral value to an inclusive range.

| parameter | what it is |
|---|---|
| `value` | The value to constrain. |
| `minVal` | The lowest accepted result. |
| `maxVal` | The highest accepted result. |

**Returns.** The value pulled inside the range.

### Clamp01

```vba
Public Function Clamp01(ByVal value As Single) As Single
```

Constrains a value to the unit range.

| parameter | what it is |
|---|---|
| `value` | The value to constrain. |

**Returns.** The value pulled inside 0..1.

### Lerp

```vba
Public Function Lerp(ByVal a As Single, ByVal b As Single, ByVal t As Single) As Single
```

Linearly blends between two values.

| parameter | what it is |
|---|---|
| `a` | The value returned at t = 0. |
| `b` | The value returned at t = 1. |
| `t` | The unclamped blend factor. |

**Returns.** The interpolated value.

### InverseLerp

```vba
Public Function InverseLerp(ByVal a As Single, ByVal b As Single, ByVal value As Single) As Single
```

Recovers the blend factor that maps a value back onto a range.

| parameter | what it is |
|---|---|
| `a` | The value standing for 0. |
| `b` | The value standing for 1. |
| `value` | The value to locate inside the range. |

**Returns.** The blend factor, or zero when the range is degenerate.

### MoveTowards

```vba
Public Function MoveTowards(ByVal current As Single, ByVal target As Single, ByVal maxDelta As Single) As Single
```

Steps a value towards a target without ever overshooting it.

| parameter | what it is |
|---|---|
| `current` | The current value. |
| `target` | The value being approached. |
| `maxDelta` | The largest change allowed on this call. |

**Returns.** The advanced value.

### Approach

```vba
Public Function Approach(ByVal current As Single, ByVal target As Single, ByVal rate As Single, ByVal dt As Single) As Single
```

Smooths a value towards a target at a frame-rate independent rate.

| parameter | what it is |
|---|---|
| `current` | The current value. |
| `target` | The value being approached. |
| `rate` | How aggressively the gap is closed, in units per second. |
| `dt` | The elapsed frame time in seconds. |

**Returns.** The smoothed value.

### SignOf

```vba
Public Function SignOf(ByVal value As Single) As Single
```

Reports the sign of a value.

| parameter | what it is |
|---|---|
| `value` | The value to inspect. |

**Returns.** -1, 0 or 1.

### MinOf

```vba
Public Function MinOf(ByVal a As Single, ByVal b As Single) As Single
```

Returns the smaller of two values.

| parameter | what it is |
|---|---|
| `a` | The first value. |
| `b` | The second value. |

**Returns.** The minimum of the pair.

### MaxOf

```vba
Public Function MaxOf(ByVal a As Single, ByVal b As Single) As Single
```

Returns the larger of two values.

| parameter | what it is |
|---|---|
| `a` | The first value. |
| `b` | The second value. |

**Returns.** The maximum of the pair.

### NearlyEqual

```vba
Public Function NearlyEqual(ByVal a As Single, ByVal b As Single, Optional ByVal tol As Single = P_EPSILON) As Boolean
```

Compares two values within a tolerance.

| parameter | what it is |
|---|---|
| `a` | The first value. |
| `b` | The second value. |
| `tol` | The largest difference still treated as equality. |

**Returns.** True when the values match closely enough.

### WrapAngle

```vba
Public Function WrapAngle(ByVal angleRad As Single) As Single
```

Wraps an angle into the -PI..PI range.

| parameter | what it is |
|---|---|
| `angleRad` | The angle in radians. |

**Returns.** The equivalent angle inside a single turn.

### AngleDelta

```vba
Public Function AngleDelta(ByVal fromRad As Single, ByVal toRad As Single) As Single
```

Measures the shortest signed rotation between two angles.

| parameter | what it is |
|---|---|
| `fromRad` | The starting angle in radians. |
| `toRad` | The destination angle in radians. |

**Returns.** The signed delta inside -PI..PI.

### ApproachAngle

```vba
Public Function ApproachAngle(ByVal current As Single, ByVal target As Single, ByVal rate As Single, ByVal dt As Single) As Single
```

Smooths an angle towards a target across the shortest arc.

| parameter | what it is |
|---|---|
| `current` | The current angle in radians. |
| `target` | The destination angle in radians. |
| `rate` | How aggressively the gap is closed, in units per second. |
| `dt` | The elapsed frame time in seconds. |

**Returns.** The smoothed angle.

### SpansOverlap

```vba
Public Function SpansOverlap(ByVal a1 As Single, ByVal a2 As Single, ByVal b1 As Single, ByVal b2 As Single) As Boolean
```

Tests whether two one dimensional spans overlap.

| parameter | what it is |
|---|---|
| `a1` | The start of the first span. |
| `a2` | The end of the first span. |
| `b1` | The start of the second span. |
| `b2` | The end of the second span. |

**Returns.** True when the spans share any ground.

### BoxesOverlap

```vba
Public Function BoxesOverlap(ByVal ax1 As Single, ByVal ay1 As Single, ByVal ax2 As Single, ByVal ay2 As Single, ByVal bx1 As Single, ByVal by1 As Single, ByVal bx2 As Single, ByVal by2 As Single) As Boolean
```

Tests whether two axis-aligned rectangles overlap.

| parameter | what it is |
|---|---|
| `ax1` | The left edge of the first rectangle. |
| `ay1` | The top edge of the first rectangle. |
| `ax2` | The right edge of the first rectangle. |
| `ay2` | The bottom edge of the first rectangle. |
| `bx1` | The left edge of the second rectangle. |
| `by1` | The top edge of the second rectangle. |
| `bx2` | The right edge of the second rectangle. |
| `by2` | The bottom edge of the second rectangle. |

**Returns.** True when the rectangles intersect.

### PointInBox

```vba
Public Function PointInBox(ByVal px As Single, ByVal py As Single, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Boolean
```

Tests whether a point falls inside an axis-aligned rectangle.

| parameter | what it is |
|---|---|
| `px` | The point X coordinate. |
| `py` | The point Y coordinate. |
| `x1` | The left edge. |
| `y1` | The top edge. |
| `x2` | The right edge. |
| `y2` | The bottom edge. |

**Returns.** True when the point is inside the rectangle or on its border.

### DistSq2D

```vba
Public Function DistSq2D(ByVal ax As Single, ByVal ay As Single, ByVal bx As Single, ByVal by As Single) As Single
```

Squared planar distance between two points.

| parameter | what it is |
|---|---|
| `ax` | The first point X coordinate. |
| `ay` | The first point Y coordinate. |
| `bx` | The second point X coordinate. |
| `by` | The second point Y coordinate. |

**Returns.** The squared distance, avoiding a square root.

### SeedRandom

```vba
Public Sub SeedRandom(ByVal seed As Long)
```

Reseeds the engine random stream so a generated world can be reproduced exactly.

| parameter | what it is |
|---|---|
| `seed` | The seed value; any Long is accepted. |

### RandomSeed

```vba
Public Property Get RandomSeed() As Long
```

Reads the seed currently driving the random stream.

**Returns.** The last seed handed to SeedRandom.

### RandomNext

```vba
Public Function RandomNext() As Double
```

Draws the next value from the Lehmer generator backing the engine.

**Returns.** A double inside the 0..1 range.

### RandomRange

```vba
Public Function RandomRange(ByVal lo As Single, ByVal hi As Single) As Single
```

Draws a value inside an arbitrary range.

| parameter | what it is |
|---|---|
| `lo` | The lowest value that can be produced. |
| `hi` | The highest value that can be produced. |

**Returns.** The random value.

### RandomInt

```vba
Public Function RandomInt(ByVal lo As Long, ByVal hi As Long) As Long
```

Draws an integer inside an inclusive range.

| parameter | what it is |
|---|---|
| `lo` | The lowest value that can be produced. |
| `hi` | The highest value that can be produced. |

**Returns.** The random integer.

### RandomChance

```vba
Public Function RandomChance(ByVal chance As Single) As Boolean
```

Draws a coin flip biased by a probability.

| parameter | what it is |
|---|---|
| `chance` | The probability of a True result, from 0 to 1. |

**Returns.** True when the draw falls under the given chance.

### ColorPack

```vba
Public Function ColorPack(ByVal r As Long, ByVal g As Long, ByVal b As Long) As Long
```

Builds a PowerPoint colour long from three channels, clamping each one.

| parameter | what it is |
|---|---|
| `r` | The red channel, 0 to 255. |
| `g` | The green channel, 0 to 255. |
| `b` | The blue channel, 0 to 255. |

**Returns.** The packed colour.

### ColorRed

```vba
Public Function ColorRed(ByVal col As Long) As Long
```

Extracts the red channel of a packed colour.

| parameter | what it is |
|---|---|
| `col` | The packed colour. |

**Returns.** The channel value, 0 to 255.

### ColorGreen

```vba
Public Function ColorGreen(ByVal col As Long) As Long
```

Extracts the green channel of a packed colour.

| parameter | what it is |
|---|---|
| `col` | The packed colour. |

**Returns.** The channel value, 0 to 255.

### ColorBlue

```vba
Public Function ColorBlue(ByVal col As Long) As Long
```

Extracts the blue channel of a packed colour.

| parameter | what it is |
|---|---|
| `col` | The packed colour. |

**Returns.** The channel value, 0 to 255.

### ColorUnpack

```vba
Public Sub ColorUnpack(ByVal col As Long, ByRef outR As Long, ByRef outG As Long, ByRef outB As Long)
```

Splits a packed colour into its three channels in one pass.

| parameter | what it is |
|---|---|
| `col` | The packed colour. |
| `outR` | Receives the red channel. |
| `outG` | Receives the green channel. |
| `outB` | Receives the blue channel. |

### ColorScale

```vba
Public Function ColorScale(ByVal col As Long, ByVal factor As Single) As Long
```

Multiplies every channel of a colour by a factor.

| parameter | what it is |
|---|---|
| `col` | The packed colour. |
| `factor` | The brightness multiplier; 1 leaves the colour untouched. |

**Returns.** The scaled colour.

### ColorMix

```vba
Public Function ColorMix(ByVal colA As Long, ByVal colB As Long, ByVal t As Single) As Long
```

Linearly blends two colours.

| parameter | what it is |
|---|---|
| `colA` | The colour returned at t = 0. |
| `colB` | The colour returned at t = 1. |
| `t` | The blend factor, clamped to 0..1. |

**Returns.** The mixed colour.

### ColorLighten

```vba
Public Function ColorLighten(ByVal col As Long, ByVal amount As Single) As Long
```

Pushes a colour towards white.

| parameter | what it is |
|---|---|
| `col` | The packed colour. |
| `amount` | How far to travel, from 0 to 1. |

**Returns.** The lightened colour.

### ColorDarken

```vba
Public Function ColorDarken(ByVal col As Long, ByVal amount As Single) As Long
```

Pushes a colour towards black.

| parameter | what it is |
|---|---|
| `col` | The packed colour. |
| `amount` | How far to travel, from 0 to 1. |

**Returns.** The darkened colour.

### ColorLuminance

```vba
Public Function ColorLuminance(ByVal col As Long) As Single
```

Computes the perceived brightness of a colour.

| parameter | what it is |
|---|---|
| `col` | The packed colour. |

**Returns.** The luminance, from 0 to 1.

### ColorGrayscale

```vba
Public Function ColorGrayscale(ByVal col As Long) As Long
```

Converts a colour to its grey equivalent.

| parameter | what it is |
|---|---|
| `col` | The packed colour. |

**Returns.** The desaturated colour.

### ColorFromHex

```vba
Public Function ColorFromHex(ByVal hexText As String) As Long
```

Parses an HTML style hex string into a packed colour.

| parameter | what it is |
|---|---|
| `hexText` | The colour text, with or without a leading hash, in RGB or RRGGBB form. |

**Returns.** The packed colour, or black when the text cannot be parsed.

> Every character is validated by hand, so a malformed string returns black instead of raising.

### ColorToHex

```vba
Public Function ColorToHex(ByVal col As Long) As String
```

Renders a packed colour as an HTML style hex string.

| parameter | what it is |
|---|---|
| `col` | The packed colour. |

**Returns.** The colour text in #RRGGBB form.

### Seconds

```vba
Public Function Seconds() As Double
```

Reads the clock.

**Returns.** Seconds, with a resolution far finer than a frame.

> Returns zero when the platform has no performance counter, which lets a caller subtract two readings and simply get zero rather than a wrong number.

### Available

```vba
Public Function Available() As Boolean
```

Reports whether the clock is usable.

**Returns.** True when the performance counter answered.

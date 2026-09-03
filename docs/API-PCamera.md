# PCamera

**Eye, yaw and pitch**

Holds the eye position and orientation of a view, caches the trigonometry the transform needs, and turns world points into the forward, side and up triple the canvas projects. Also answers the visibility question used to reject whole objects before any face work happens.

> Yaw turns around the Z axis with zero looking down positive X; pitch is positive when looking up. The cached sines and cosines are refreshed on assignment, so reading them in a tight loop costs nothing.

> **Scope.** Private to this VBA project: class modules carry VB_Exposed = False, which is the class level equivalent of Option Private Module.

## Index

**Lifetime.** [`Init`](#init)

**Position.** [`SetPosition`](#setposition), [`Translate`](#translate), [`X`](#x), [`Y`](#y), [`Z`](#z)

**Orientation.** [`SetAngles`](#setangles), [`AddAngles`](#addangles), [`Yaw`](#yaw), [`Pitch`](#pitch), [`SetPitchLimits`](#setpitchlimits), [`LookAt`](#lookat), [`CosYaw`](#cosyaw), [`SinYaw`](#sinyaw), [`CosPitch`](#cospitch), [`SinPitch`](#sinpitch), [`ForwardX`](#forwardx), [`ForwardY`](#forwardy), [`RightX`](#rightx), [`RightY`](#righty), [`ScreenShiftY`](#screenshifty)

**Transforms.** [`WorldToView`](#worldtoview), [`ViewDepth`](#viewdepth), [`SphereVisible`](#spherevisible), [`Describe`](#describe)

## Members

### Init

```vba
Public Function Init(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, Optional ByVal yawRad As Single = 0!, Optional ByVal pitchRad As Single = 0!) As PCamera
```

Places and aims the camera in one call.

| parameter | what it is |
|---|---|
| `X` | The eye X coordinate. |
| `Y` | The eye Y coordinate. |
| `Z` | The eye Z coordinate. |
| `yawRad` | The heading in radians. |
| `pitchRad` | The elevation in radians. |

**Returns.** The camera itself, so it can be created and aimed in one expression.

### SetPosition

```vba
Public Sub SetPosition(ByVal X As Single, ByVal Y As Single, ByVal Z As Single)
```

Moves the eye to an absolute position.

| parameter | what it is |
|---|---|
| `X` | The eye X coordinate. |
| `Y` | The eye Y coordinate. |
| `Z` | The eye Z coordinate. |

### Translate

```vba
Public Sub Translate(ByVal dx As Single, ByVal dy As Single, ByVal dz As Single)
```

Shifts the eye by an offset.

| parameter | what it is |
|---|---|
| `dx` | The change along X. |
| `dy` | The change along Y. |
| `dz` | The change along Z. |

### X

```vba
Public Property Get X() As Single
Public Property Let X(ByVal value As Single)
```

Reads the eye X coordinate.

**Write.** Sets the eye X coordinate.

| parameter | what it is |
|---|---|
| `value` | The world X position. |

**Returns.** The world X position.

### Y

```vba
Public Property Get Y() As Single
Public Property Let Y(ByVal value As Single)
```

Reads the eye Y coordinate.

**Write.** Sets the eye Y coordinate.

| parameter | what it is |
|---|---|
| `value` | The world Y position. |

**Returns.** The world Y position.

### Z

```vba
Public Property Get Z() As Single
Public Property Let Z(ByVal value As Single)
```

Reads the eye Z coordinate.

**Write.** Sets the eye Z coordinate.

| parameter | what it is |
|---|---|
| `value` | The world Z position. |

**Returns.** The world Z position.

### SetAngles

```vba
Public Sub SetAngles(ByVal yawRad As Single, ByVal pitchRad As Single)
```

Aims the camera at absolute angles.

| parameter | what it is |
|---|---|
| `yawRad` | The heading in radians. |
| `pitchRad` | The elevation in radians, clamped to the configured limits. |

### AddAngles

```vba
Public Sub AddAngles(ByVal dYaw As Single, ByVal dPitch As Single)
```

Turns the camera by a relative amount, the shape mouse look input arrives in.

| parameter | what it is |
|---|---|
| `dYaw` | The heading change in radians. |
| `dPitch` | The elevation change in radians. |

### Yaw

```vba
Public Property Get Yaw() As Single
Public Property Let Yaw(ByVal value As Single)
```

Reads the heading.

**Write.** Sets the heading and refreshes the cached trigonometry.

| parameter | what it is |
|---|---|
| `value` | The yaw in radians; values outside a single turn are wrapped. |

**Returns.** The yaw in radians.

### Pitch

```vba
Public Property Get Pitch() As Single
Public Property Let Pitch(ByVal value As Single)
```

Reads the elevation.

**Write.** Sets the elevation and refreshes the cached trigonometry.

| parameter | what it is |
|---|---|
| `value` | The pitch in radians, clamped to the configured limits. |

**Returns.** The pitch in radians.

### SetPitchLimits

```vba
Public Sub SetPitchLimits(ByVal minRad As Single, ByVal maxRad As Single)
```

Restricts how far the camera may look up and down.

| parameter | what it is |
|---|---|
| `minRad` | The lowest pitch allowed, in radians. |
| `maxRad` | The highest pitch allowed, in radians. |

### LookAt

```vba
Public Sub LookAt(ByVal tx As Single, ByVal ty As Single, ByVal tz As Single)
```

Aims the camera at a world point.

| parameter | what it is |
|---|---|
| `tx` | The target X coordinate. |
| `ty` | The target Y coordinate. |
| `tz` | The target Z coordinate. |

### CosYaw

```vba
Public Property Get CosYaw() As Single
```

Reads the cached cosine of the yaw.

**Returns.** The cosine, refreshed whenever the yaw changes.

### SinYaw

```vba
Public Property Get SinYaw() As Single
```

Reads the cached sine of the yaw.

**Returns.** The sine, refreshed whenever the yaw changes.

### CosPitch

```vba
Public Property Get CosPitch() As Single
```

Reads the cached cosine of the pitch.

**Returns.** The cosine, refreshed whenever the pitch changes.

### SinPitch

```vba
Public Property Get SinPitch() As Single
```

Reads the cached sine of the pitch.

**Returns.** The sine, refreshed whenever the pitch changes.

### ForwardX

```vba
Public Property Get ForwardX() As Single
```

Reads the X component of the flat forward vector, the direction a walking body advances in.

**Returns.** The forward X component.

### ForwardY

```vba
Public Property Get ForwardY() As Single
```

Reads the Y component of the flat forward vector.

**Returns.** The forward Y component.

### RightX

```vba
Public Property Get RightX() As Single
```

Reads the X component of the flat right vector, used for strafing.

**Returns.** The right X component.

### RightY

```vba
Public Property Get RightY() As Single
```

Reads the Y component of the flat right vector.

**Returns.** The right Y component.

### ScreenShiftY

```vba
Public Property Get ScreenShiftY() As Single
Public Property Let ScreenShiftY(ByVal value As Single)
```

Reads the vertical screen offset applied at projection time.

**Write.** Sets a vertical screen offset applied at projection time, which is how head bob and recoil are expressed.

| parameter | what it is |
|---|---|
| `value` | The offset in slide points. |

**Returns.** The offset in slide points.

### WorldToView

```vba
Public Sub WorldToView(ByVal wx As Single, ByVal wy As Single, ByVal wz As Single, ByRef outFwd As Single, ByRef outSide As Single, ByRef outUp As Single)
```

Converts a world point into the view space triple the canvas projects.

| parameter | what it is |
|---|---|
| `wx` | The world X coordinate. |
| `wy` | The world Y coordinate. |
| `wz` | The world Z coordinate. |
| `outFwd` | Receives the distance in front of the eye. |
| `outSide` | Receives the distance to the right of the eye. |
| `outUp` | Receives the distance above the eye. |

### ViewDepth

```vba
Public Function ViewDepth(ByVal wx As Single, ByVal wy As Single, ByVal wz As Single) As Single
```

Measures how far in front of the eye a world point sits.

| parameter | what it is |
|---|---|
| `wx` | The world X coordinate. |
| `wy` | The world Y coordinate. |
| `wz` | The world Z coordinate. |

**Returns.** The view depth, negative when the point is behind the camera.

### SphereVisible

```vba
Public Function SphereVisible(ByVal cv As PCanvas, ByVal wx As Single, ByVal wy As Single, ByVal wz As Single, ByVal radius As Single, ByRef outNear As Single, ByRef outMid As Single) As Boolean
```

Rejects an object against the view frustum using its bounding sphere.

| parameter | what it is |
|---|---|
| `cv` | The canvas whose frustum is tested against. |
| `wx` | The sphere centre X coordinate. |
| `wy` | The sphere centre Y coordinate. |
| `wz` | The sphere centre Z coordinate. |
| `radius` | The sphere radius in world units. |
| `outNear` | Receives the depth of the closest point of the sphere, for sorting. |
| `outMid` | Receives the depth of the sphere centre. |

**Returns.** True when any part of the sphere can reach the canvas.

> This is the cheap early exit that keeps whole objects out of the face pipeline.

### Describe

```vba
Public Function Describe() As String
```

Renders the camera state as readable text for logging and debugging.

**Returns.** A one line summary of the position and orientation.

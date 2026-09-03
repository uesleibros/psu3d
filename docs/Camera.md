# Camera

A position in the world plus two rotations.

```vba
Dim cam As PCamera
Set cam = New PCamera

cam.Init 0, -10, 1.7, P_HALF_PI, 0    ' x, y, z, yaw, pitch
cam.SetPosition 0, -10, 1.7
cam.SetAngles P_HALF_PI, 0
cam.AddAngles 0.02, -0.01
cam.LookAt 0, 0, 0
```

`Yaw` is which way you face in the ground plane, in radians. `Pitch` is how far up or down you are looking.

Pitch is limited, and the limit is adjustable through `PitchLimit`. Without a limit you go over the top and yaw starts turning the wrong way, which is disorienting and never what anybody wanted.

The sine and cosine of both angles are cached and recomputed only when the angle changes. A frame reads `CosYaw` about a thousand times, and calling `Cos` a thousand times with the same argument is work thrown away.

## Directions

`ForwardX` and `ForwardY` are where the camera points in the ground plane. `RightX` and `RightY` are its right hand side. Together they become movement:

```vba
wishX = cam.ForwardX * fwd + cam.RightX * strafe
wishY = cam.ForwardY * fwd + cam.RightY * strafe
```

Note there are only two components. Walking should not climb when you look up, so the forward vector is the flattened one.

## Transform

`WorldToView` takes a world point into camera space, giving back forward, side and up. That is the library's view matrix, and it is spelled out in [How the 3D works](How-The-3D-Works.md).

`ViewDepth` returns only the depth component, which is what the sorting uses. `SphereVisible` answers whether a sphere fits inside the frustum.

In practice neither `PScene` nor `PRenderer` calls those on the hot path. They read the sine, the cosine and the position once per frame and write the arithmetic out inline, because a method call per vertex is the most expensive thing in a frame that is almost entirely arithmetic.

## Bob and recoil

`ScreenShiftY` offsets the image vertically after projection. It is for head bob while walking and for weapon recoil. Do not use it to look up: pitch is there for that, and pitch is a real rotation.

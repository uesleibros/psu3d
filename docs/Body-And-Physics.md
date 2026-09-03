# Body and physics

`PBody` is an upright cylinder that walks through a scene and is stopped by it. It is optional: anybody who only draws never imports the module.

```vba
Dim body As PBody
Set body = Psu3D.CreateBody()      ' already bound to the booted scene

body.SetSize 0.34, 1.72            ' radius, height
body.SetPosition 0, 0, 4
body.WalkSpeed = 5.4
body.JumpSpeed = 6.4
body.StepHeight = 0.45             ' fallback: a material may say otherwise
body.KillZ = -12
```

## The step

```vba
body.Advance dt, wishX, wishY, up, down
```

One call per frame. `wishX` and `wishY` are the requested direction in world coordinates, and do not need to be normalised. The two booleans are the up and down controls, which mean jump, swim or climb depending on where the body is.

The order inside `Advance` is the order things happen in the world:

```
carried by whatever is underneath
accelerated by what you asked for
pushed out of anything solid
and only then does gravity get a say
```

Doing gravity before the push out would let the body sink into the floor for a frame before being lifted back, which reads as a stutter.

## Reading the result

```vba
body.X, body.Y, body.Z
body.VelX, body.VelY, body.VelZ
body.OnGround
body.GroundMaterial
body.GroundObject
body.Submersion         ' 0 to 1
body.IsSwimming
body.IsClimbing
body.LandingSpeed       ' impact of the last landing, for camera shake
body.CarryYaw           ' how far a platform turned you
body.FellOut            ' below KillZ
```

## Triggers

The body does not know what a coin is. It reports what it touched and refuses to guess what it was for:

```vba
For k = 0 To body.TouchCount - 1
    idx = body.TouchAt(k)

    If sc.TagOf(idx) = TAG_CHECKPOINT Then
        ' your rule
    Else
        sc.SetActive idx, False
        score = score + 1
    End If
Next k
```

That is why the same class serves a collecting level and a mine dodging level without knowing either word.

`Reach` is how far past its own edge the body notices a trigger. A coin the width of a hand, collected only when the body's own edge crosses it, is a coin players walk through and do not collect.

## The material decides

The body has no rules of its own about surfaces. It asks the material:

| property | effect |
|---|---|
| `Collision` | `solid` stops, `ghost` lets through, `oneway` holds only from above, `trigger` only reports |
| `Friction` | how quickly a body standing on it sheds speed |
| `SpeedMultiplier` | the speed ceiling while walking on it |
| `Bounce` | a hard landing is thrown back, a soft one is absorbed |
| `DamagePerSecond` | the game reads it and decides |
| `StepHeight` | the tallest lip that surface allows; zero means the body decides |
| `Climbable` | a ladder |
| `Buoyancy`, `Drag` | on a trigger, they make it a fluid |

## Being carried

The body remembers which object it is standing on, and at the start of the next step takes that object's displacement and goes along. That is why it stores the ground **id**, not just the material.

For a turning platform that means being swung around its centre, not standing still while the floor rotates underneath. `CarryYaw` gives back how far you turned, to add to the camera yaw:

```vba
If body.CarryYaw <> 0 Then cam.AddAngles body.CarryYaw, 0
```

Without it the disc turns you around its centre while you keep facing the same way, which is the one thing standing on a turntable never does.

For a platform driven by hand, outside the motion system, there is `Nudge`, which shifts without touching velocity, because being carried is not being accelerated:

```vba
sc.SetTopZ lift, newZ
If body.GroundObject = lift Then body.Nudge 0, 0, newZ - oldZ
```

**The scene has to move before the body does.** A platform that has already moved is one that can carry; one that moves afterwards is one the body spends a frame standing beside.

## Ladders

A ladder is a material with `Climbable`. While touching it, holding up or down moves the body vertically and switches gravity off.

`Climbable` and `Collision` are independent on purpose. A `ghost` ladder is entered and climbed from inside, a `solid` one is scaled from outside, and both are written the same way.

## Water

A trigger with `Buoyancy` above zero is a fluid. Buoyancy cancels part of gravity in proportion to how much of the body is submerged, and drag bleeds all three components of velocity.

Past a threshold of submersion the body starts swimming: up becomes a stroke rather than a jump, and down becomes a dive.

## Tuning

All of these have sensible defaults and all are properties:

`Gravity`, `WalkSpeed`, `JumpSpeed`, `GroundAccel`, `GroundFriction`, `AirAccel`, `SwimSpeed`, `DiveSpeed`, `ClimbSpeed`, `StepHeight`, `KillZ`, `Reach`.

## Writing another solver

`PBody` is written entirely against the scene's public API: `QueryBox`, `TopZAt`, `Blocks`, `GetOrientedBox`, `GetSpanZ`, `ContainsXY`, `GetMotionDelta` and `SpinDelta`. Anybody wanting a different solver writes it the same way and loses nothing.

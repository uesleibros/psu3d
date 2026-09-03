# PBody

**Corpo que anda e colide**

Everything that moves through a scene and is stopped by it: walking, jumping, climbing stairs, sliding on ice, swimming, riding platforms and reading what it touched. The material system describes how a surface behaves; this is the part that believes it.

> The body is a vertical cylinder, which is why rotation never needs to be applied to it: only its centre is carried into an object's frame, and a circle looks the same from every angle. Advance is the whole step; call it once a frame with what the player asked for and read the state back off the properties.

> **Escopo.** Private to this VBA project: class modules carry VB_Exposed = False, which is the class level equivalent of Option Private Module.

## Indice

**Lifetime.** [`Attach`](#attach), [`SetSize`](#setsize), [`SetPosition`](#setposition)

**The step.** [`Nudge`](#nudge), [`Advance`](#advance)

**Result.** [`X`](#x), [`Y`](#y), [`Z`](#z), [`VelX`](#velx), [`VelY`](#vely), [`VelZ`](#velz), [`OnGround`](#onground), [`GroundMaterial`](#groundmaterial), [`GroundObject`](#groundobject), [`Submersion`](#submersion), [`IsSwimming`](#isswimming), [`IsClimbing`](#isclimbing), [`LandingSpeed`](#landingspeed), [`CarryYaw`](#carryyaw), [`FellOut`](#fellout), [`TouchCount`](#touchcount), [`TouchAt`](#touchat)

**Tuning.** [`Gravity`](#gravity), [`WalkSpeed`](#walkspeed), [`JumpSpeed`](#jumpspeed), [`StepHeight`](#stepheight), [`KillZ`](#killz), [`ClimbSpeed`](#climbspeed), [`SwimSpeed`](#swimspeed), [`DiveSpeed`](#divespeed), [`Reach`](#reach), [`GroundAccel`](#groundaccel), [`GroundFriction`](#groundfriction), [`AirAccel`](#airaccel)

## Membros

### Attach

```vba
Public Function Attach(ByVal sc As PScene) As PBody
```

Binds the body to the scene it moves through.

| parametro | o que e |
|---|---|
| `sc` | The scene to collide against. |

**Devolve.** The body itself, so it can be created and bound in one expression.

### SetSize

```vba
Public Sub SetSize(ByVal radius As Single, ByVal height As Single)
```

Sets the shape of the body.

| parametro | o que e |
|---|---|
| `radius` | Half its width. |
| `height` | Its height from feet to crown. |

### SetPosition

```vba
Public Sub SetPosition(ByVal X As Single, ByVal Y As Single, ByVal Z As Single)
```

Puts the body somewhere and stops it dead.

| parametro | o que e |
|---|---|
| `X` | The world X coordinate of the feet. |
| `Y` | The world Y coordinate of the feet. |
| `Z` | The height of the feet. |

### Nudge

```vba
Public Sub Nudge(ByVal dx As Single, ByVal dy As Single, ByVal dz As Single)
```

Shifts the body without disturbing what it was doing.

| parametro | o que e |
|---|---|
| `dx` | The offset along X. |
| `dy` | The offset along Y. |
| `dz` | The offset along Z. |

> For a platform driven by hand rather than by the motion system: the body has to travel with it, and travelling is not the same as being accelerated, so velocity is deliberately left alone.

### Advance

```vba
Public Sub Advance(ByVal dt As Single, ByVal wishX As Single, ByVal wishY As Single, ByVal up As Boolean, ByVal down As Boolean)
```

Advances the body by one frame.

| parametro | o que e |
|---|---|
| `dt` | The elapsed frame time in seconds. |
| `wishX` | How hard the body is being pushed along world X, from -1 to 1. |
| `wishY` | How hard it is being pushed along world Y. |
| `up` | True while the jump, swim or climb up control is held. |
| `down` | True while the dive or climb down control is held. |

The order matters and is the order things happen in the world. Platforms have already moved by the time this runs, so the body is carried first; then it is accelerated, then pushed out of anything solid, and only then does gravity get a say. Doing gravity before the push-out would let the body sink into a floor for a frame before being lifted back, which reads as a stutter.

### X

```vba
Public Property Get X() As Single
```

Reads the X coordinate of the feet.

**Devolve.** The world X position.

### Y

```vba
Public Property Get Y() As Single
```

Reads the Y coordinate of the feet.

**Devolve.** The world Y position.

### Z

```vba
Public Property Get Z() As Single
```

Reads the height of the feet.

**Devolve.** The world Z position.

### VelX

```vba
Public Property Get VelX() As Single
```

Reads the horizontal speed along X.

**Devolve.** The velocity component.

### VelY

```vba
Public Property Get VelY() As Single
```

Reads the horizontal speed along Y.

**Devolve.** The velocity component.

### VelZ

```vba
Public Property Get VelZ() As Single
```

Reads the vertical speed.

**Devolve.** The velocity component.

### OnGround

```vba
Public Property Get OnGround() As Boolean
```

Reports whether the body is standing on something.

**Devolve.** True when it has ground under it.

### GroundMaterial

```vba
Public Property Get GroundMaterial() As Long
```

Reads the material of whatever the body is standing on.

**Devolve.** The material id, or the default when it is airborne.

### GroundObject

```vba
Public Property Get GroundObject() As Long
```

Reads which object the body is standing on.

**Devolve.** The object id, or P_INVALID_ID when it is airborne.

### Submersion

```vba
Public Property Get Submersion() As Single
```

Reads how much of the body is under a fluid surface.

**Devolve.** A fraction from 0 to 1.

### IsSwimming

```vba
Public Property Get IsSwimming() As Boolean
```

Reports whether the body is deep enough in a fluid to swim.

**Devolve.** True when swimming rather than wading.

### IsClimbing

```vba
Public Property Get IsClimbing() As Boolean
```

Reports whether the body is climbing a climbable surface.

**Devolve.** True while it is on a ladder and asking to move along it.

### LandingSpeed

```vba
Public Property Get LandingSpeed() As Single
```

Reads how hard the body hit the ground on the step just taken.

**Devolve.** The impact speed, or zero when it did not land.

> What a camera reads to decide how much to flinch.

### CarryYaw

```vba
Public Property Get CarryYaw() As Single
```

Reads how far a turning platform swung the body on the step just taken.

**Devolve.** The change in heading, in radians.

> Add it to the camera yaw and the view turns with the platform, which is what makes standing on a spinning disc feel like standing on it rather than beside it.

### FellOut

```vba
Public Property Get FellOut() As Boolean
```

Reports whether the body has fallen past the height the world ends at.

**Devolve.** True when it is lost.

### TouchCount

```vba
Public Property Get TouchCount() As Long
```

Reads how many triggers the body touched on the step just taken.

**Devolve.** The count.

> A trigger is reported, never acted on. What a coin or a checkpoint means is a rule about the game, and the body has no business knowing it.

### TouchAt

```vba
Public Function TouchAt(ByVal slot As Long) As Long
```

Reads one of the triggers touched on the step just taken.

| parametro | o que e |
|---|---|
| `slot` | Its position in the list. |

**Devolve.** The object id, or P_INVALID_ID when the slot is out of range.

### Gravity

```vba
Public Property Get Gravity() As Single
Public Property Let Gravity(ByVal value As Single)
```

Reads the downward acceleration.

Escrita: Sets the downward acceleration.

| parametro | o que e |
|---|---|
| `value` | The value in world units per second squared. |

**Devolve.** The value in world units per second squared.

### WalkSpeed

```vba
Public Property Get WalkSpeed() As Single
Public Property Let WalkSpeed(ByVal value As Single)
```

Reads the top walking speed on a normal surface.

Escrita: Sets the top walking speed on a normal surface.

| parametro | o que e |
|---|---|
| `value` | The speed in world units per second, before the ground material scales it. |

**Devolve.** The speed in world units per second.

### JumpSpeed

```vba
Public Property Get JumpSpeed() As Single
Public Property Let JumpSpeed(ByVal value As Single)
```

Reads the upward speed a jump gives.

Escrita: Sets the upward speed a jump gives.

| parametro | o que e |
|---|---|
| `value` | The speed in world units per second. |

**Devolve.** The speed in world units per second.

### StepHeight

```vba
Public Property Get StepHeight() As Single
Public Property Let StepHeight(ByVal value As Single)
```

Reads the tallest lip the body walks over without jumping.

Escrita: Sets the fallback height for a lip the body walks over.

| parametro | o que e |
|---|---|
| `value` | The height in world units. |

**Devolve.** The height in world units.

> Only the fallback. A surface that names its own StepHeight overrides this, which is how a kerb and a cliff can be made of the same shape.

### KillZ

```vba
Public Property Get KillZ() As Single
Public Property Let KillZ(ByVal value As Single)
```

Reads the height below which the body counts as lost.

Escrita: Sets the height below which the body counts as lost.

| parametro | o que e |
|---|---|
| `value` | The height in world units. |

**Devolve.** The height in world units.

### ClimbSpeed

```vba
Public Property Get ClimbSpeed() As Single
Public Property Let ClimbSpeed(ByVal value As Single)
```

Reads the speed the body climbs a ladder at.

Escrita: Sets the speed the body climbs a ladder at.

| parametro | o que e |
|---|---|
| `value` | The speed in world units per second. |

**Devolve.** The speed in world units per second.

### SwimSpeed

```vba
Public Property Get SwimSpeed() As Single
Public Property Let SwimSpeed(ByVal value As Single)
```

Reads the speed the body swims upwards at.

Escrita: Sets the speed the body swims upwards at.

| parametro | o que e |
|---|---|
| `value` | The speed in world units per second. |

**Devolve.** The speed in world units per second.

### DiveSpeed

```vba
Public Property Get DiveSpeed() As Single
Public Property Let DiveSpeed(ByVal value As Single)
```

Reads the speed the body dives at.

Escrita: Sets the speed the body dives at.

| parametro | o que e |
|---|---|
| `value` | The speed in world units per second. |

**Devolve.** The speed in world units per second.

### Reach

```vba
Public Property Get Reach() As Single
Public Property Let Reach(ByVal value As Single)
```

Reads how far past its own edge the body notices a trigger.

Escrita: Sets how far past its own edge the body notices a trigger.

| parametro | o que e |
|---|---|
| `value` | The extra distance in world units. |

**Devolve.** The extra distance in world units.

A coin the width of a hand, touched only when the body's own edge crosses it, is a coin players walk through and do not collect. Reach is the difference between what the body occupies and what it can reach for, and it applies to triggers alone: nothing solid uses it.

### GroundAccel

```vba
Public Property Get GroundAccel() As Single
Public Property Let GroundAccel(ByVal value As Single)
```

Reads how quickly the body reaches walking speed.

Escrita: Sets how quickly the body reaches walking speed.

| parametro | o que e |
|---|---|
| `value` | The acceleration in world units per second squared. |

**Devolve.** The acceleration in world units per second squared.

### GroundFriction

```vba
Public Property Get GroundFriction() As Single
Public Property Let GroundFriction(ByVal value As Single)
```

Reads how quickly the body sheds speed with nothing held.

Escrita: Sets how quickly the body sheds speed with nothing held.

| parametro | o que e |
|---|---|
| `value` | The damping in units per second, before the ground material scales it. |

**Devolve.** The damping in units per second.

### AirAccel

```vba
Public Property Get AirAccel() As Single
Public Property Let AirAccel(ByVal value As Single)
```

Reads how much control the body keeps in mid air.

Escrita: Sets how much control the body keeps in mid air.

| parametro | o que e |
|---|---|
| `value` | The acceleration in world units per second squared. |

**Devolve.** The acceleration in world units per second squared.

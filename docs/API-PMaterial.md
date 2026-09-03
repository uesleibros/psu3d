# PMaterial

**One surface**

A single surface description shared by the renderer and the physics layer: what the surface looks like, and whether a body can walk on it, fall through it or pass straight through. One material is meant to be authored once and referenced by many faces through its registry id.

> Every property setter that affects shading notifies PMaterials so cached shading tables are rebuilt on the next draw.

> **Scope.** Private to this VBA project: class modules carry VB_Exposed = False, which is the class level equivalent of Option Private Module.

## Index

**Identity.** [`Id`](#id), [`Name`](#name), [`Tag`](#tag)

**Appearance.** [`Color`](#color), [`Unlit`](#unlit), [`Fogged`](#fogged), [`Visible`](#visible), [`TwoSided`](#twosided), [`Transparency`](#transparency), [`EdgeVisible`](#edgevisible), [`EdgeColor`](#edgecolor), [`EdgeWeight`](#edgeweight), [`NeedsShapeFormat`](#needsshapeformat)

**Physics.** [`Collision`](#collision), [`IsSolid`](#issolid), [`IsGhost`](#isghost), [`IsOneWay`](#isoneway), [`IsTrigger`](#istrigger), [`BlocksMovement`](#blocksmovement), [`Friction`](#friction), [`Bounce`](#bounce), [`SpeedMultiplier`](#speedmultiplier), [`DamagePerSecond`](#damagepersecond), [`Climbable`](#climbable), [`StepHeight`](#stepheight), [`Buoyancy`](#buoyancy), [`Drag`](#drag), [`IsFluid`](#isfluid)

**Copying.** [`CopyFrom`](#copyfrom), [`Clone`](#clone), [`Describe`](#describe)

## Members

### Id

```vba
Public Property Get Id() As Long
```

Reads the registry id assigned to this material.

**Returns.** The id, or P_INVALID_ID while the material is unregistered.

### Name

```vba
Public Property Get Name() As String
Public Property Let Name(ByVal value As String)
```

Reads the lookup name of the material.

**Write.** Sets the lookup name of the material.

| parameter | what it is |
|---|---|
| `value` | The new name. |

**Returns.** The name used as the registry key.

> Renaming a material already held by the registry does not move its key; use PMaterials.Rename instead.

### Tag

```vba
Public Property Get Tag() As String
Public Property Let Tag(ByVal value As String)
```

Reads the free form payload attached to this material.

**Write.** Attaches a free form payload to this material, such as a sound name or a gameplay flag.

| parameter | what it is |
|---|---|
| `value` | The caller defined tag. |

**Returns.** The caller defined tag.

### Color

```vba
Public Property Get Color() As Long
Public Property Let Color(ByVal value As Long)
```

Reads the unlit base colour of the surface.

**Write.** Sets the unlit base colour of the surface.

| parameter | what it is |
|---|---|
| `value` | The packed colour, as produced by PCore.ColorPack. |

**Returns.** The packed colour.

### Unlit

```vba
Public Property Get Unlit() As Boolean
Public Property Let Unlit(ByVal value As Boolean)
```

Reports whether the surface ignores the directional light.

**Write.** Chooses whether the surface ignores the directional light.

| parameter | what it is |
|---|---|
| `value` | True to skip shading, which suits coins, HUD panels and emissive props. |

**Returns.** True when the material renders at full brightness.

### Fogged

```vba
Public Property Get Fogged() As Boolean
Public Property Let Fogged(ByVal value As Boolean)
```

Reports whether distance fog is applied to the surface.

**Write.** Chooses whether distance fog is applied to the surface.

| parameter | what it is |
|---|---|
| `value` | True to fade with depth, False to keep the colour constant at any distance. |

**Returns.** True when the material fades into the fog colour.

### Visible

```vba
Public Property Get Visible() As Boolean
Public Property Let Visible(ByVal value As Boolean)
```

Reports whether the surface is drawn at all.

**Write.** Chooses whether the surface is drawn at all.

| parameter | what it is |
|---|---|
| `value` | False to build an invisible collider that still blocks movement. |

**Returns.** True when the material produces geometry.

### TwoSided

```vba
Public Property Get TwoSided() As Boolean
Public Property Let TwoSided(ByVal value As Boolean)
```

Reports whether both sides of a face are drawn.

**Write.** Chooses whether both sides of a face are drawn.

| parameter | what it is |
|---|---|
| `value` | True for flat props such as flags and billboards, which have no inside. |

**Returns.** True when backface culling is skipped.

### Transparency

```vba
Public Property Get Transparency() As Single
Public Property Let Transparency(ByVal value As Single)
```

Reads how see-through the surface is.

**Write.** Sets how see-through the surface is.

| parameter | what it is |
|---|---|
| `value` | The transparency, from 0 for opaque to 1 for invisible. |

**Returns.** The transparency, from 0 for opaque to 1 for invisible.

> Any value above zero forces the renderer onto a slower path, since the fill has to be formatted per shape.

### EdgeVisible

```vba
Public Property Get EdgeVisible() As Boolean
Public Property Let EdgeVisible(ByVal value As Boolean)
```

Reports whether faces are outlined.

**Write.** Chooses whether faces are outlined, which is useful for wireframe and toon looks.

| parameter | what it is |
|---|---|
| `value` | True to draw the outline. |

**Returns.** True when the shape outline is drawn.

### EdgeColor

```vba
Public Property Get EdgeColor() As Long
Public Property Let EdgeColor(ByVal value As Long)
```

Reads the outline colour.

**Write.** Sets the outline colour.

| parameter | what it is |
|---|---|
| `value` | The packed colour. |

**Returns.** The packed colour.

### EdgeWeight

```vba
Public Property Get EdgeWeight() As Single
Public Property Let EdgeWeight(ByVal value As Single)
```

Reads the outline thickness.

**Write.** Sets the outline thickness.

| parameter | what it is |
|---|---|
| `value` | The weight in points. |

**Returns.** The weight in points.

### NeedsShapeFormat

```vba
Public Property Get NeedsShapeFormat() As Boolean
```

Reports whether the renderer must format each emitted shape individually.

**Returns.** True when transparency or outlines are in play.

> Materials that answer False are drawn through the fast path, which only assigns a fill colour.

### Collision

```vba
Public Property Get Collision() As PCollision
Public Property Let Collision(ByVal value As PCollision)
```

Reads how the surface answers collision queries.

**Write.** Sets how the surface answers collision queries.

| parameter | what it is |
|---|---|
| `value` | pcGhost to let bodies pass, pcSolid to block from every side, pcOneWay to block only a landing from above, or pcTrigger to report the overlap without blocking. |

**Returns.** The PCollision mode.

### IsSolid

```vba
Public Property Get IsSolid() As Boolean
```

Reports whether bodies are stopped from every direction.

**Returns.** True for pcSolid materials.

### IsGhost

```vba
Public Property Get IsGhost() As Boolean
```

Reports whether bodies pass straight through with no notification.

**Returns.** True for pcGhost materials.

### IsOneWay

```vba
Public Property Get IsOneWay() As Boolean
```

Reports whether the surface can only be landed on from above.

**Returns.** True for pcOneWay materials.

### IsTrigger

```vba
Public Property Get IsTrigger() As Boolean
```

Reports whether the surface only signals overlaps, as pickups and hazard volumes do.

**Returns.** True for pcTrigger materials.

### BlocksMovement

```vba
Public Function BlocksMovement(Optional ByVal landingFromAbove As Boolean = False) As Boolean
```

Answers the core physics question: does this surface stop the body that just touched it?

| parameter | what it is |
|---|---|
| `landingFromAbove` | True when the body is descending onto the top face of the surface. |

**Returns.** True when the motion must be resolved instead of ignored.

### Friction

```vba
Public Property Get Friction() As Single
Public Property Let Friction(ByVal value As Single)
```

Reads how strongly the surface slows a body sliding across it.

**Write.** Sets how strongly the surface slows a body sliding across it.

| parameter | what it is |
|---|---|
| `value` | The friction multiplier, clamped to zero or above. |

**Returns.** The friction multiplier; 1 is the engine default, lower values feel like ice.

### Bounce

```vba
Public Property Get Bounce() As Single
Public Property Let Bounce(ByVal value As Single)
```

Reads how much vertical speed a landing body keeps.

**Write.** Sets how much vertical speed a landing body keeps, turning the surface into a trampoline.

| parameter | what it is |
|---|---|
| `value` | The restitution, clamped to zero or above. |

**Returns.** The restitution; 0 absorbs the impact, 1 bounces back at full speed.

### SpeedMultiplier

```vba
Public Property Get SpeedMultiplier() As Single
Public Property Let SpeedMultiplier(ByVal value As Single)
```

Reads the walking speed multiplier applied while standing on the surface.

**Write.** Sets the walking speed multiplier applied while standing on the surface, for mud or conveyors.

| parameter | what it is |
|---|---|
| `value` | The multiplier, clamped to zero or above. |

**Returns.** The multiplier; 1 is normal speed.

### DamagePerSecond

```vba
Public Property Get DamagePerSecond() As Single
Public Property Let DamagePerSecond(ByVal value As Single)
```

Reads the damage dealt per second of contact.

**Write.** Sets the damage dealt per second of contact, which is how lava and spikes are described.

| parameter | what it is |
|---|---|
| `value` | The damage rate. |

**Returns.** The damage rate; zero for a harmless surface.

### Climbable

```vba
Public Property Get Climbable() As Boolean
Public Property Let Climbable(ByVal value As Boolean)
```

Reports whether a body may climb the vertical faces of the surface.

**Write.** Chooses whether a body may climb the vertical faces of the surface.

| parameter | what it is |
|---|---|
| `value` | True to allow climbing. |

**Returns.** True for ladders and similar geometry.

### StepHeight

```vba
Public Property Get StepHeight() As Single
Public Property Let StepHeight(ByVal value As Single)
```

Reads the tallest lip a walking body steps over instead of colliding with.

**Write.** Sets the tallest lip a walking body steps over instead of colliding with.

| parameter | what it is |
|---|---|
| `value` | The step height in world units; zero, the default, means the body decides. |

**Returns.** The step height in world units, or zero to leave the decision to the body.

This is what separates a kerb from a cliff when both are the same shape. A material that names a height overrides whatever the body would have allowed, so a low wall can be made unclimbable and a tall stair can be made walkable without either of them changing size.

### Buoyancy

```vba
Public Property Get Buoyancy() As Single
Public Property Let Buoyancy(ByVal value As Single)
```

Reads how strongly the surface holds a body up against gravity.

**Write.** Turns the material into a fluid a body can be inside of.

| parameter | what it is |
|---|---|
| `value` | How much of gravity the fluid cancels: 0 leaves the body falling normally, 1 makes it float in place, above 1 pushes it to the surface. |

**Returns.** The fraction of gravity cancelled while submerged; zero for anything that is not a fluid.

> Read together with Collision. A fluid is a trigger, because it must be entered rather than landed on, and it is the buoyancy that tells the physics this particular trigger is something to swim in rather than a coin to pick up.

### Drag

```vba
Public Property Get Drag() As Single
Public Property Let Drag(ByVal value As Single)
```

Reads how quickly the material bleeds off the speed of a body inside it.

**Write.** Sets how quickly the material bleeds off the speed of a body inside it.

| parameter | what it is |
|---|---|
| `value` | The damping in units per second; this is what makes water feel thick and a jump land softly. |

**Returns.** The damping in units per second.

### IsFluid

```vba
Public Property Get IsFluid() As Boolean
```

Reports whether the material is something a body can be immersed in.

**Returns.** True for a trigger that carries buoyancy.

### CopyFrom

```vba
Public Sub CopyFrom(ByVal src As PMaterial)
```

Overwrites this material with the settings of another, leaving the registry id alone.

| parameter | what it is |
|---|---|
| `src` | The material to copy from. |

### Clone

```vba
Public Function Clone(Optional ByVal newName As String = vbNullString) As PMaterial
```

Produces an unregistered copy of this material.

| parameter | what it is |
|---|---|
| `newName` | The name to give the copy; the source name is reused when omitted. |

**Returns.** The cloned material, ready to be handed to PMaterials.Add.

### Describe

```vba
Public Function Describe() As String
```

Renders the material as readable text for logging and debugging.

**Returns.** A one line summary of the identity, colour and collision mode.

## `Friend` members

Visible to the whole VBA project but not outside it. A public member of a class cannot take a type declared in a standard module, which is why these are `Friend` rather than `Public`.

### Id

```vba
Friend Property Let Id(ByVal value As Long)
```

Stamps the registry id onto the material.

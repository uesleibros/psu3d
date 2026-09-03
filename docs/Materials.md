# Materials

A material describes a surface: how it is drawn and how it behaves. Two objects with the same material behave the same, and swapping an object's material changes both at once.

```vba
Dim m As PMaterial
Set m = PMaterials.Create("ice", PCore.ColorPack(170, 220, 245), pcSolid)

m.Friction = 0.08
m.SpeedMultiplier = 1.25
m.EdgeVisible = True
m.EdgeColor = PCore.ColorPack(220, 245, 255)
```

The registry keeps the material and hands back an id, which is the number the scene uses:

```vba
sc.AddBox PMaterials.IdOf("ice"), -4, -4, 4, 4, 0, 0.3
```

## Collision

| value | effect |
|---|---|
| `pcSolid` | blocks from every side |
| `pcGhost` | drawn, and you pass through |
| `pcOneWay` | holds only a body arriving from above; from below you pass |
| `pcTrigger` | never blocks, only reports the contact |

## Appearance

| property | what it does |
|---|---|
| `Color` | the base colour, before light and fog |
| `Unlit` | ignores the light and stays at full colour |
| `Fogged` | whether fog acts on it |
| `Visible` | false makes an invisible collider |
| `TwoSided` | drawn even when facing away |
| `Transparency` | 0 to 1 |
| `EdgeVisible`, `EdgeColor`, `EdgeWeight` | outline |

A material with `Visible = False` is not merely transparent: `PScene` skips an invisible object during collection, so it is never culled, never sorted, never drawn, and never spends budget. That is how you make a guard rail that holds you and does not show.

## Physics

| property | what it does |
|---|---|
| `Friction` | how quickly a body standing on it sheds speed; ice is 0.08 |
| `Bounce` | a hard landing is thrown back, a soft one is absorbed |
| `SpeedMultiplier` | the speed ceiling while walking on it |
| `DamagePerSecond` | the game reads it and decides |
| `StepHeight` | the tallest lip *this surface* allows; zero means the body decides |
| `Climbable` | a ladder: while touching it, up and down move you instead of falling |
| `Buoyancy`, `Drag` | on a trigger, they make it a fluid |

`StepHeight` and `Climbable` are what separate a kerb from a cliff of the same shape. And `Climbable` and `Collision` are independent questions on purpose: a `ghost` ladder is entered and climbed from inside, a `solid` one is scaled from outside, and both are written the same way.

## Fluid

A trigger with `Buoyancy` above zero is a fluid. The body measures how much of itself is under the surface and applies everything in proportion:

```vba
Set water = PMaterials.Create("water", PCore.ColorPack(40, 110, 190), pcTrigger)
water.Buoyancy = 1.06
water.Drag = 3.4
water.SpeedMultiplier = 0.45
water.Transparency = 0.4
water.TwoSided = True
```

Buoyancy above 1 means letting go floats you back up. Drag bleeds all three components of velocity, and is what makes a fall into water land softly instead of hitting the bottom.

## The stock palette

`PMaterials.CreateDefaults` builds sixteen ready made materials: `default`, `stone`, `grass`, `brick`, `metal`, `ice`, `mud`, `rubber`, `glass`, `lava`, `water`, `platform`, `ladder`, `decor`, `pickup` and `clip`. `Psu3D.Boot` calls it for you.

They exist to get started quickly and to give the self test something known to measure against. Nothing requires their use.

## Precomputed shading

The colour of a face is material times face direction times fog band. That is a table, not a calculation, and `PMaterials` keeps it ready:

```vba
col = PMaterials.ShadeBand(matId, dirKey, band)
```

Two bounds checks and one array read. When the light, the fog or a material changes, a revision counter invalidates the table and it is rebuilt the next time anybody asks.

Only axis aligned faces use the table. A tilted face, such as a ramp, falls to the slow path that computes from the real normal.

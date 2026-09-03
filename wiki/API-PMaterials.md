# PMaterials

**Registro de materiais e tabela de sombreamento**

Owns every PMaterial in the project, hands out the small integer ids faces carry, and keeps a baked shading table so the renderer resolves a face colour with one array read instead of per-face lighting maths.

> The table is rebuilt lazily whenever a material changes or PLighting reports a new revision, so callers never have to think about invalidation.

## Constantes

| nome | tipo | valor | o que e |
|---|---|---|---|
| `P_DEFAULT_MATERIAL` | String | `"default"` | Name of the material that always occupies id 0. |

## Indice

**Registry.** [`Create`](#create), [`Add`](#add), [`ByName`](#byname), [`ById`](#byid), [`IdOf`](#idof), [`Exists`](#exists), [`Rename`](#rename), [`Count`](#count), [`Clear`](#clear)

**Stock materials.** [`CreateDefaults`](#createdefaults)

**Shading.** [`ShadeColor`](#shadecolor), [`ShadeBand`](#shadeband), [`Sync`](#sync), [`ShadeColorDynamic`](#shadecolordynamic), [`ColorOf`](#colorof), [`NotifyChanged`](#notifychanged), [`Invalidate`](#invalidate)

## Membros

### Create

```vba
Public Function Create(ByVal matName As String, Optional ByVal col As Long = P_WHITE, Optional ByVal mode As PCollision = pcSolid) As PMaterial
```

Creates a material, registers it and returns it for further tuning.

| parametro | o que e |
|---|---|
| `matName` | The lookup name, which must be unique. |
| `col` | The base colour; white when omitted. |
| `mode` | The collision mode; pcSolid when omitted. |

**Devolve.** The registered material, or the existing one when the name is already taken.

### Add

```vba
Public Function Add(ByVal mat As PMaterial) As Long
```

Registers a material that was built by hand.

| parametro | o que e |
|---|---|
| `mat` | The material to store; its Name is used as the key. |

**Devolve.** The id assigned to the material, or the id it already holds.

### ByName

```vba
Public Function ByName(ByVal matName As String) As PMaterial
```

Looks a material up by name.

| parametro | o que e |
|---|---|
| `matName` | The registered name; the search is case insensitive. |

**Devolve.** The material, or Nothing when the name is unknown.

### ById

```vba
Public Function ById(ByVal id As Long) As PMaterial
```

Looks a material up by id.

| parametro | o que e |
|---|---|
| `id` | The registry id carried by a face. |

**Devolve.** The material, falling back to the default material when the id is out of range.

### IdOf

```vba
Public Function IdOf(ByVal matName As String) As Long
```

Resolves a name into the id faces are tagged with.

| parametro | o que e |
|---|---|
| `matName` | The registered name. |

**Devolve.** The id, or P_INVALID_ID when the name is unknown.

> A linear scan over a handful of materials, and only ever used at build time; the render loop works with ids.

### Exists

```vba
Public Function Exists(ByVal matName As String) As Boolean
```

Reports whether a name is already registered.

| parametro | o que e |
|---|---|
| `matName` | The name to test. |

**Devolve.** True when a material answers to that name.

### Rename

```vba
Public Function Rename(ByVal oldName As String, ByVal newName As String) As Boolean
```

Moves a material to a new lookup name, keeping its id and every face that references it.

| parametro | o que e |
|---|---|
| `oldName` | The current name. |
| `newName` | The name to move it to. |

**Devolve.** True when the rename succeeded.

### Count

```vba
Public Property Get Count() As Long
```

Reads how many materials are registered.

**Devolve.** The material count, always at least one.

### Clear

```vba
Public Sub Clear()
```

Drops every material and restores the lone default entry.

> Any id held by existing geometry becomes meaningless, so rebuild the world after calling this.

### CreateDefaults

```vba
Public Sub CreateDefaults()
```

Registers a small palette of ready to use materials covering the common surface behaviours.

Creates ground, ice, mud, metal, glass, lava, water, one-way platform, decorative and pickup materials. Existing names are left untouched, so it is safe to call after your own setup.

### ShadeColor

```vba
Public Function ShadeColor(ByVal matId As Long, ByVal dirKey As PDirection, ByVal depth As Single) As Long
```

Resolves the final colour of an axis-aligned face straight from the baked table.

| parametro | o que e |
|---|---|
| `matId` | The material id carried by the face. |
| `dirKey` | The PDirection entry naming the face orientation. |
| `depth` | The view depth of the face, used to pick the fog band. |

**Devolve.** The shaded and fogged colour.

### ShadeBand

```vba
Public Function ShadeBand(ByVal matId As Long, ByVal dirKey As Long, ByVal band As Long) As Long
```

Reads a baked colour straight from its fog band.

| parametro | o que e |
|---|---|
| `matId` | The material id carried by the face. |
| `dirKey` | The PDirection entry naming the face orientation. |
| `band` | The fog band the caller already worked out. |

**Devolve.** The shaded and fogged colour.

> The hot path of the renderer. ShadeColor has to ask PLighting which band a depth lands in, which is several cross module calls per face; a renderer that already caches the fog settings computes the band itself and lands here, where the work is two bounds checks and one array read.

### Sync

```vba
Public Sub Sync()
```

Rebuilds the shading table now, so a frame can be drawn without any table checks.

### ShadeColorDynamic

```vba
Public Function ShadeColorDynamic(ByVal matId As Long, ByVal nx As Single, ByVal ny As Single, ByVal nz As Single, ByVal depth As Single) As Long
```

Resolves the final colour of a face whose normal is not axis-aligned.

| parametro | o que e |
|---|---|
| `matId` | The material id carried by the face. |
| `nx` | The X component of the unit normal. |
| `ny` | The Y component of the unit normal. |
| `nz` | The Z component of the unit normal. |
| `depth` | The view depth of the face. |

**Devolve.** The shaded and fogged colour.

> This path costs a dot product and a colour blend, which is why axis-aligned faces should always carry a real dirKey.

### ColorOf

```vba
Public Function ColorOf(ByVal matId As Long) As Long
```

Reads the unlit base colour of a material by id.

| parametro | o que e |
|---|---|
| `matId` | The material id. |

**Devolve.** The packed colour.

### NotifyChanged

```vba
Public Sub NotifyChanged()
```

Marks the baked shading table as stale.

> PMaterial calls this for you whenever a shading property changes; call it by hand only after editing lighting state through a route the engine cannot see.

### Invalidate

```vba
Public Sub Invalidate()
```

Forces the shading table to be rebuilt on the next draw.

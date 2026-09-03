# PLevel

**Read and write a scene as JSON**

Turns a level written as JSON into a built PScene. A level stops being code and becomes data: text anyone can edit without opening the VBE, without a compile step between an idea and seeing it, and readable by any tool that reads JSON.

> Parsing is done by the JSON class; this module only knows the schema, in both directions: Parse reads a document into a scene and Save writes a scene back out as one. Every field is optional and every missing field has a stated default, so a half written level still loads and shows you what you did write instead of refusing the whole file.

## Index

**Entry points.** [`Parse`](#parse), [`ParseFile`](#parsefile)

**Authoring.** [`SetSpawn`](#setspawn), [`SetBudget`](#setbudget), [`SetRules`](#setrules), [`SetKillZ`](#setkillz)

**Results.** [`Error`](#error), [`ObjectCount`](#objectcount), [`SpawnX`](#spawnx), [`SpawnY`](#spawny), [`SpawnZ`](#spawnz), [`SpawnYaw`](#spawnyaw), [`BudgetMin`](#budgetmin), [`BudgetMax`](#budgetmax), [`Gravity`](#gravity), [`Jump`](#jump), [`Walk`](#walk), [`KillZ`](#killz), [`HasKillZ`](#haskillz)

**Writing.** [`Save`](#save), [`SaveFile`](#savefile)

## Members

### Parse

```vba
Public Function Parse(ByVal source As String, ByVal target As PScene) As Boolean
```

Builds a level from JSON text.

| parameter | what it is |
|---|---|
| `source` | The level document. |
| `target` | The scene to build into; it is cleared first. |

**Returns.** True when the document held a usable level.

> On failure Error says what was wrong. A document that parses but describes nothing is a failure too, since a silently empty world is the least useful answer available.

### ParseFile

```vba
Public Function ParseFile(ByVal path As String, ByVal target As PScene) As Boolean
```

Builds a level from a JSON file on disk.

| parameter | what it is |
|---|---|
| `path` | The full path to the level file. |
| `target` | The scene to build into. |

**Returns.** True when the file was found and held a usable level.

> Existence is checked before opening rather than trapped, in keeping with the rest of the library.

### SetSpawn

```vba
Public Sub SetSpawn(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, Optional ByVal yawDeg As Single = 0!)
```

Sets where the player starts.

| parameter | what it is |
|---|---|
| `X` | The spawn X coordinate. |
| `Y` | The spawn Y coordinate. |
| `Z` | The spawn Z coordinate. |
| `yawDeg` | Which way they face, in degrees. |

### SetBudget

```vba
Public Sub SetBudget(ByVal lo As Long, ByVal hi As Long)
```

Sets the polygon budget the level asks for.

| parameter | what it is |
|---|---|
| `lo` | The floor. |
| `hi` | The ceiling; zero on both leaves the renderer to decide. |

### SetRules

```vba
Public Sub SetRules(ByVal gravityValue As Single, ByVal jumpValue As Single, ByVal walkValue As Single)
```

Sets the movement rules the level asks for.

| parameter | what it is |
|---|---|
| `gravityValue` | The downward acceleration; zero keeps the body default. |
| `jumpValue` | The jump speed; zero keeps the body default. |
| `walkValue` | The walking speed; zero keeps the body default. |

### SetKillZ

```vba
Public Sub SetKillZ(ByVal Z As Single)
```

Sets the height below which the player is lost.

| parameter | what it is |
|---|---|
| `Z` | The height in world units. |

### Error

```vba
Public Property Get Error() As String
```

Reads why the last load failed.

**Returns.** The message, or an empty string when it succeeded.

### ObjectCount

```vba
Public Property Get ObjectCount() As Long
```

Reads how many objects the last level built.

**Returns.** The object count.

### SpawnX

```vba
Public Property Get SpawnX() As Single
```

Reads the spawn X coordinate the level asked for.

**Returns.** The world X position.

### SpawnY

```vba
Public Property Get SpawnY() As Single
```

Reads the spawn Y coordinate the level asked for.

**Returns.** The world Y position.

### SpawnZ

```vba
Public Property Get SpawnZ() As Single
```

Reads the spawn Z coordinate the level asked for.

**Returns.** The world Z position.

### SpawnYaw

```vba
Public Property Get SpawnYaw() As Single
```

Reads the heading the level wants the body to start facing.

**Returns.** The yaw in radians.

### BudgetMin

```vba
Public Property Get BudgetMin() As Long
```

Reads the lower polygon budget the level asked for.

**Returns.** The budget floor, or zero when the level did not say.

### BudgetMax

```vba
Public Property Get BudgetMax() As Long
```

Reads the upper polygon budget the level asked for.

**Returns.** The budget ceiling, or zero when the level did not say.

### Gravity

```vba
Public Property Get Gravity() As Single
```

Reads the downward acceleration the level asked for.

**Returns.** The value, or zero when the level did not say.

### Jump

```vba
Public Property Get Jump() As Single
```

Reads the jump speed the level asked for.

**Returns.** The value, or zero when the level did not say.

### Walk

```vba
Public Property Get Walk() As Single
```

Reads the walking speed the level asked for.

**Returns.** The value, or zero when the level did not say.

### KillZ

```vba
Public Property Get KillZ() As Single
```

Reads the height below which the level considers a body lost.

**Returns.** The height; meaningful only when HasKillZ is True.

> Zero is a perfectly sensible floor for a level to choose, so absence is reported separately rather than encoded as a value.

### HasKillZ

```vba
Public Property Get HasKillZ() As Boolean
```

Reports whether the level named a height below which a body is lost.

**Returns.** True when it did.

### Save

```vba
Public Function Save(ByVal source As PScene) As String
```

Writes a scene back out as a level document.

| parameter | what it is |
|---|---|
| `source` | The scene to describe. |

**Returns.** The JSON text.

The round trip is the point: a level loaded, edited in place and saved is the same level plus the edit. What does not survive is the shorthand, since repeat and step are a way of writing a staircase, not a thing the world remembers; the fifty steps they stood for all come back, one entry each. The document is longer than the one that was read and describes exactly the same world.

### SaveFile

```vba
Public Function SaveFile(ByVal path As String, ByVal source As PScene) As Boolean
```

Writes a scene to a file as a level document.

| parameter | what it is |
|---|---|
| `path` | Where to write it. |
| `source` | The scene to describe. |

**Returns.** True when the file was written.

> An existing file is replaced, with no undo, so a caller pointing this at an authored level should be sure that is what it means.

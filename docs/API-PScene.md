# PScene

**Object store, spatial index and draw order**

Holds every object in a world as parallel arrays rather than objects, so adding a thousand boxes costs a thousand array slots and no allocations per frame. Answers the two questions a frame asks: which objects can the camera see, and which objects sit inside this box of space.

> There is no object ceiling; the arrays double when they fill. Ids stay valid for the life of an object because nothing ever reshuffles the arrays, but a slot given back by Remove is handed out again, so an id kept past a Remove names whatever took its place.

> **Scope.** Private to this VBA project: class modules carry VB_Exposed = False, which is the class level equivalent of Option Private Module.

## Index

**Object arrays.** [`LodSize`](#lodsize), [`Reserve`](#reserve), [`Clear`](#clear), [`Count`](#count), [`LiveCount`](#livecount)

**Building.** [`AddBox`](#addbox), [`AddRotatedBox`](#addrotatedbox), [`AddRamp`](#addramp), [`AddBillboard`](#addbillboard), [`AddSpinner`](#addspinner)

**Editing.** [`IsActive`](#isactive), [`SetActive`](#setactive), [`Remove`](#remove), [`MaterialOf`](#materialof), [`SetMaterial`](#setmaterial), [`KindOf`](#kindof), [`MoveBy`](#moveby), [`SetTopZ`](#settopz), [`SetAngle`](#setangle), [`AngleOf`](#angleof), [`SetPhase`](#setphase), [`SpinAll`](#spinall)

**Motion.** [`SetMotion`](#setmotion), [`SetSpin`](#setspin), [`UpdateMotion`](#updatemotion), [`GetMotionDelta`](#getmotiondelta), [`SpinDelta`](#spindelta), [`GetSize`](#getsize), [`GetCenter`](#getcenter), [`ThickOf`](#thickof), [`TopOf`](#topof), [`HighOf`](#highof), [`AxisOf`](#axisof), [`MotionOf`](#motionof), [`SetTag`](#settag), [`TagOf`](#tagof)

**Bounds and physics.** [`GetBoundsXY`](#getboundsxy), [`GetSpanZ`](#getspanz), [`TopZAt`](#topzat), [`Blocks`](#blocks), [`GetOrientedBox`](#getorientedbox), [`ContainsXY`](#containsxy)

**Spatial query.** [`QueryBox`](#querybox), [`QueryRadius`](#queryradius), [`ResultAt`](#resultat), [`ResultCount`](#resultcount)

**Spatial index.** [`DynamicCount`](#dynamiccount), [`CellSize`](#cellsize)

**Rendering.** [`Render`](#render), [`DrawnCount`](#drawncount), [`WasReduced`](#wasreduced), [`DrawnAt`](#drawnat), [`DrawObject`](#drawobject)

## Members

### LodSize

```vba
Public Property Get LodSize() As Single
Public Property Let LodSize(ByVal value As Single)
```

Reads the screen size under which a box drops to a single face.

**Write.** Sets the screen size under which a box drops to a single face.

| parameter | what it is |
|---|---|
| `value` | The projected diameter in slide points; zero or less turns the reduction off. |

**Returns.** The threshold in slide points.

A box shows at most three faces, and the two minor ones shrink to slivers long before the dominant one does. Past that point they cost a full shape each to contribute a few points of colour, which on a renderer whose frame is spent inside PowerPoint is the most expensive kind of detail there is. Raise the threshold to buy frame rate, lower it to buy silhouette.

### Reserve

```vba
Public Sub Reserve(ByVal slots As Long)
```

Grows the arrays up front so no allocation happens while the world is being built.

| parameter | what it is |
|---|---|
| `slots` | How many object slots to make room for. |

### Clear

```vba
Public Sub Clear()
```

Empties the scene without releasing the arrays, so the next world reuses the same memory.

### Count

```vba
Public Property Get Count() As Long
```

Reads how many object slots have ever been used, which is also the highest id plus one.

**Returns.** The slot count.

### LiveCount

```vba
Public Property Get LiveCount() As Long
```

Reads how many objects are currently active.

**Returns.** The live object count.

### AddBox

```vba
Public Function AddBox(ByVal matId As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal zTop As Single, ByVal thickness As Single) As Long
```

Adds an axis-aligned box.

| parameter | what it is |
|---|---|
| `matId` | The material the box is drawn and collided with. |
| `x1` | The lower X bound. |
| `y1` | The lower Y bound. |
| `x2` | The upper X bound. |
| `y2` | The upper Y bound. |
| `zTop` | The height of the top face. |
| `thickness` | How far the box extends below its top face. |

**Returns.** The id of the new object.

### AddRotatedBox

```vba
Public Function AddRotatedBox(ByVal matId As Long, ByVal cx As Single, ByVal cy As Single, ByVal halfW As Single, ByVal halfH As Single, ByVal zTop As Single, ByVal thickness As Single, Optional ByVal angleRad As Single = 0!) As Long
```

Adds a box that spins around its own vertical axis.

| parameter | what it is |
|---|---|
| `matId` | The material the box is drawn and collided with. |
| `cx` | The centre X coordinate. |
| `cy` | The centre Y coordinate. |
| `halfW` | Half the extent along the local X axis. |
| `halfH` | Half the extent along the local Y axis. |
| `zTop` | The height of the top face. |
| `thickness` | How far the box extends below its top face. |
| `angleRad` | The starting rotation in radians. |

**Returns.** The id of the new object.

### AddRamp

```vba
Public Function AddRamp(ByVal matId As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal zLow As Single, ByVal zHigh As Single, ByVal thickness As Single, Optional ByVal axis As PAxis = paX) As Long
```

Adds a sloped slab.

| parameter | what it is |
|---|---|
| `matId` | The material the ramp is drawn and collided with. |
| `x1` | The lower X bound. |
| `y1` | The lower Y bound. |
| `x2` | The upper X bound. |
| `y2` | The upper Y bound. |
| `zLow` | The height at the low end. |
| `zHigh` | The height at the high end. |
| `thickness` | How far the slab extends below its sloped face. |
| `axis` | paX to climb along X, paY to climb along Y. |

**Returns.** The id of the new object.

### AddBillboard

```vba
Public Function AddBillboard(ByVal matId As Long, ByVal X As Single, ByVal Y As Single, ByVal Z As Single, ByVal Width As Single, ByVal Height As Single) As Long
```

Adds a camera facing quad.

| parameter | what it is |
|---|---|
| `matId` | The material the billboard is drawn with. |
| `X` | The centre X coordinate. |
| `Y` | The centre Y coordinate. |
| `Z` | The centre Z coordinate. |
| `Width` | The width in world units. |
| `Height` | The height in world units. |

**Returns.** The id of the new object.

### AddSpinner

```vba
Public Function AddSpinner(ByVal matId As Long, ByVal X As Single, ByVal Y As Single, ByVal Z As Single, ByVal radius As Single) As Long
```

Adds a spinning double sided plate, the shape pickups usually take.

| parameter | what it is |
|---|---|
| `matId` | The material the plate is drawn with. |
| `X` | The centre X coordinate. |
| `Y` | The centre Y coordinate. |
| `Z` | The centre Z coordinate. |
| `radius` | The half size of the plate. |

**Returns.** The id of the new object.

### IsActive

```vba
Public Function IsActive(ByVal idx As Long) As Boolean
```

Reports whether an id refers to a live object.

| parameter | what it is |
|---|---|
| `idx` | The object id. |

**Returns.** True when the object exists and is active.

### SetActive

```vba
Public Sub SetActive(ByVal idx As Long, ByVal value As Boolean)
```

Turns an object on or off without disturbing any other id.

| parameter | what it is |
|---|---|
| `idx` | The object id. |
| `value` | True to draw and collide with the object. |

### Remove

```vba
Public Sub Remove(ByVal idx As Long)
```

Removes an object for good and returns its slot to the store.

| parameter | what it is |
|---|---|
| `idx` | The object id. |

> The id may be handed to a different object later, so drop any copy of it. To hide something you intend to bring back, use SetActive instead, which keeps the slot reserved.

### MaterialOf

```vba
Public Function MaterialOf(ByVal idx As Long) As Long
```

Reads the material an object carries.

| parameter | what it is |
|---|---|
| `idx` | The object id. |

**Returns.** The material id, or P_INVALID_ID for an unknown object.

### SetMaterial

```vba
Public Sub SetMaterial(ByVal idx As Long, ByVal matId As Long)
```

Swaps the material of an object, which is how surfaces change state at runtime.

| parameter | what it is |
|---|---|
| `idx` | The object id. |
| `matId` | The new material id. |

### KindOf

```vba
Public Function KindOf(ByVal idx As Long) As PObjectKind
```

Reads the primitive kind of an object.

| parameter | what it is |
|---|---|
| `idx` | The object id. |

**Returns.** The PObjectKind entry.

### MoveBy

```vba
Public Sub MoveBy(ByVal idx As Long, ByVal dx As Single, ByVal dy As Single, ByVal dz As Single)
```

Shifts an object by an offset, moving both its geometry and its bounds.

| parameter | what it is |
|---|---|
| `idx` | The object id. |
| `dx` | The change along X. |
| `dy` | The change along Y. |
| `dz` | The change along Z. |

### SetTopZ

```vba
Public Sub SetTopZ(ByVal idx As Long, ByVal zTop As Single)
```

Places the top face of an object at an absolute height.

| parameter | what it is |
|---|---|
| `idx` | The object id. |
| `zTop` | The new top height. |

### SetAngle

```vba
Public Sub SetAngle(ByVal idx As Long, ByVal angleRad As Single)
```

Sets the spin of a rotated box.

| parameter | what it is |
|---|---|
| `idx` | The object id. |
| `angleRad` | The rotation in radians. |

### AngleOf

```vba
Public Function AngleOf(ByVal idx As Long) As Single
```

Reads the spin of a rotated box.

| parameter | what it is |
|---|---|
| `idx` | The object id. |

**Returns.** The rotation in radians.

### SetPhase

```vba
Public Sub SetPhase(ByVal idx As Long, ByVal phase As Single)
```

Sets the animation phase used by spinners.

| parameter | what it is |
|---|---|
| `idx` | The object id. |
| `phase` | The phase in radians. |

### SpinAll

```vba
Public Sub SpinAll(ByVal dt As Single, Optional ByVal speed As Single = 2.6!)
```

Advances the animation phase of every spinner in the scene.

| parameter | what it is |
|---|---|
| `dt` | The elapsed frame time in seconds. |
| `speed` | How fast the plates turn, in radians per second. |

### SetMotion

```vba
Public Sub SetMotion(ByVal idx As Long, ByVal axisX As Single, ByVal axisY As Single, ByVal axisZ As Single, ByVal amp As Single, ByVal speed As Single, Optional ByVal phase As Single = 0!)
```

Makes an object slide back and forth along a direction.

| parameter | what it is |
|---|---|
| `idx` | The object id. |
| `axisX` | The X component of the direction it travels along. |
| `axisY` | The Y component. |
| `axisZ` | The Z component. |
| `amp` | How far it reaches from its resting place, in world units. |
| `speed` | How fast it oscillates, in radians per second. |
| `phase` | Where in its cycle it starts, so two platforms can be set out of step. |

> The offset is a sine, so the platform slows at each end instead of snapping around, which is what makes it possible to step onto one on purpose.

### SetSpin

```vba
Public Sub SetSpin(ByVal idx As Long, ByVal speed As Single, Optional ByVal phase As Single = 0!)
```

Makes an object turn on its own vertical axis.

| parameter | what it is |
|---|---|
| `idx` | The object id. |
| `speed` | The turn rate in radians per second. |
| `phase` | The angle it starts at. |

### UpdateMotion

```vba
Public Function UpdateMotion(ByVal dt As Single) As Long
```

Advances every animated object.

| parameter | what it is |
|---|---|
| `dt` | The elapsed frame time in seconds. |

**Returns.** How many objects moved, so a caller that only redraws on change knows to redraw.

> Each mover records the step it just took, which is what a body standing on it needs in order to travel with it. Carrying is left to the caller because it is a rule about bodies, not about scenery, and only the caller knows which object the body is standing on.

### GetMotionDelta

```vba
Public Sub GetMotionDelta(ByVal idx As Long, ByRef outDx As Single, ByRef outDy As Single, ByRef outDz As Single)
```

Reads the step an object took on the last update.

| parameter | what it is |
|---|---|
| `idx` | The object id. |
| `outDx` | Receives the change along X. |
| `outDy` | Receives the change along Y. |
| `outDz` | Receives the change along Z. |

### SpinDelta

```vba
Public Function SpinDelta(ByVal idx As Long) As Single
```

Reads how far an object turned on the last update.

| parameter | what it is |
|---|---|
| `idx` | The object id. |

**Returns.** The change in angle, in radians.

> A body standing on a turning platform has to be swung around its centre by this much, or it stands still while the floor rotates out from under it.

### GetSize

```vba
Public Sub GetSize(ByVal idx As Long, ByRef outA As Single, ByRef outB As Single)
```

Reads the two size numbers of an object.

| parameter | what it is |
|---|---|
| `idx` | The object id. |
| `outA` | Half the local X extent of a turned box, the width of a billboard, the radius of a spinner. |
| `outB` | The matching second number. |

### GetCenter

```vba
Public Sub GetCenter(ByVal idx As Long, ByRef outX As Single, ByRef outY As Single, ByRef outZ As Single)
```

Reads the centre of an object.

| parameter | what it is |
|---|---|
| `idx` | The object id. |
| `outX` | The centre X coordinate. |
| `outY` | The centre Y coordinate. |
| `outZ` | The centre Z coordinate. |

### ThickOf

```vba
Public Function ThickOf(ByVal idx As Long) As Single
```

Reads how far a slab hangs below its face.

| parameter | what it is |
|---|---|
| `idx` | The object id. |

**Returns.** The thickness in world units.

### TopOf

```vba
Public Function TopOf(ByVal idx As Long) As Single
```

Reads the height an object was built at.

| parameter | what it is |
|---|---|
| `idx` | The object id. |

**Returns.** The top face for a box, the low end for a ramp.

### HighOf

```vba
Public Function HighOf(ByVal idx As Long) As Single
```

Reads the second height of an object.

| parameter | what it is |
|---|---|
| `idx` | The object id. |

**Returns.** The high end for a ramp, and the same as TopOf for anything level.

### AxisOf

```vba
Public Function AxisOf(ByVal idx As Long) As PAxis
```

Reads which way a ramp climbs.

| parameter | what it is |
|---|---|
| `idx` | The object id. |

**Returns.** paX or paY.

### MotionOf

```vba
Public Function MotionOf(ByVal idx As Long, ByRef outAx As Single, ByRef outAy As Single, ByRef outAz As Single, ByRef outAmp As Single, ByRef outSpeed As Single, ByRef outPhase As Single) As Long
```

Reads the motion an object was given.

| parameter | what it is |
|---|---|
| `idx` | The object id. |
| `outAx` | The X part of the travel direction. |
| `outAy` | The Y part. |
| `outAz` | The Z part. |
| `outAmp` | How far it travels, or nothing for a spin. |
| `outSpeed` | Its rate, in cycles a second for travel and radians a second for a spin. |
| `outPhase` | Where in the cycle it starts. |

**Returns.** 0 when it is still, 1 when it travels, 2 when it spins.

### SetTag

```vba
Public Sub SetTag(ByVal idx As Long, ByVal value As Long)
```

Attaches a number of your own to an object.

| parameter | what it is |
|---|---|
| `idx` | The object id. |
| `value` | Whatever the number means to you; zero, the default, means nothing was said. |

The scene deliberately does not know what a checkpoint, a door or a waypoint is. It carries one number per object and lets the caller decide what it stands for, which is the difference between a library you can build a game on and a library you can only build this game on.

### TagOf

```vba
Public Function TagOf(ByVal idx As Long) As Long
```

Reads the number attached to an object.

| parameter | what it is |
|---|---|
| `idx` | The object id. |

**Returns.** The tag, or zero when none was set.

### GetBoundsXY

```vba
Public Sub GetBoundsXY(ByVal idx As Long, ByRef outX1 As Single, ByRef outY1 As Single, ByRef outX2 As Single, ByRef outY2 As Single)
```

Reads the horizontal bounds of an object.

| parameter | what it is |
|---|---|
| `idx` | The object id. |
| `outX1` | Receives the lower X bound. |
| `outY1` | Receives the lower Y bound. |
| `outX2` | Receives the upper X bound. |
| `outY2` | Receives the upper Y bound. |

### GetSpanZ

```vba
Public Sub GetSpanZ(ByVal idx As Long, ByRef outBottom As Single, ByRef outTop As Single)
```

Reads the vertical span of an object.

| parameter | what it is |
|---|---|
| `idx` | The object id. |
| `outBottom` | Receives the lowest point. |
| `outTop` | Receives the highest point. |

### TopZAt

```vba
Public Function TopZAt(ByVal idx As Long, ByVal X As Single, ByVal Y As Single) As Single
```

Measures the height of the walkable surface of an object under a point.

| parameter | what it is |
|---|---|
| `idx` | The object id. |
| `X` | The world X coordinate to sample. |
| `Y` | The world Y coordinate to sample. |

**Returns.** The surface height, interpolated across the slope for ramps.

### Blocks

```vba
Public Function Blocks(ByVal idx As Long, Optional ByVal landingFromAbove As Boolean = False) As Boolean
```

Answers whether an object stops a body that just touched it.

| parameter | what it is |
|---|---|
| `idx` | The object id. |
| `landingFromAbove` | True when the body is descending onto the top surface. |

**Returns.** True when the contact must be resolved.

### GetOrientedBox

```vba
Public Sub GetOrientedBox(ByVal idx As Long, ByRef outCx As Single, ByRef outCy As Single, ByRef outHalfW As Single, ByRef outHalfH As Single, ByRef outAngle As Single)
```

Reads the oriented footprint of an object.

| parameter | what it is |
|---|---|
| `idx` | The object id. |
| `outCx` | Receives the centre X coordinate. |
| `outCy` | Receives the centre Y coordinate. |
| `outHalfW` | Receives the half extent along the object's own X axis. |
| `outHalfH` | Receives the half extent along its own Y axis. |
| `outAngle` | Receives the rotation in radians. |

> For everything except a rotated box the angle is zero and the half extents come straight from the bounds, so a caller can use this one shape of answer for every object it meets.

### ContainsXY

```vba
Public Function ContainsXY(ByVal idx As Long, ByVal X As Single, ByVal Y As Single) As Boolean
```

Reports whether a point falls inside the real footprint of an object.

| parameter | what it is |
|---|---|
| `idx` | The object id. |
| `X` | The world X coordinate to test. |
| `Y` | The world Y coordinate to test. |

**Returns.** True when the point is inside, taking rotation into account.

> The bounds of a rotated box are the square that encloses it at any angle, so a point can be inside those bounds and well outside the box itself. Anything that decides what a body is standing on has to ask this rather than read the bounds.

### QueryBox

```vba
Public Function QueryBox(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Long
```

Collects every active object whose bounds overlap a rectangle of the world.

| parameter | what it is |
|---|---|
| `x1` | The lower X bound of the query. |
| `y1` | The lower Y bound of the query. |
| `x2` | The upper X bound of the query. |
| `y2` | The upper Y bound of the query. |

**Returns.** How many objects were found; read them back through ResultAt.

> The results live in a reused array, so a second query overwrites the first.

### QueryRadius

```vba
Public Function QueryRadius(ByVal X As Single, ByVal Y As Single, ByVal radius As Single) As Long
```

Collects every active object within a radius of a point.

| parameter | what it is |
|---|---|
| `X` | The centre X coordinate. |
| `Y` | The centre Y coordinate. |
| `radius` | The search radius. |

**Returns.** How many objects were found.

### ResultAt

```vba
Public Function ResultAt(ByVal slot As Long) As Long
```

Reads one entry of the last query result.

| parameter | what it is |
|---|---|
| `slot` | The position in the result list. |

**Returns.** The object id, or P_INVALID_ID when the slot is out of range.

### ResultCount

```vba
Public Property Get ResultCount() As Long
```

Reads how many objects the last query found.

**Returns.** The result count.

### DynamicCount

```vba
Public Property Get DynamicCount() As Long
```

Reads how many objects are swept rather than looked up.

**Returns.** The count of objects that moved at least once, plus those too big to be worth indexing.

### CellSize

```vba
Public Property Get CellSize() As Single
```

Reads how wide one cell of the lookup grid is.

**Returns.** The cell size in world units, or zero when nothing is indexed.

### Render

```vba
Public Function Render(ByVal rd As PRenderer) As Long
```

Draws the whole scene through a renderer, frustum culling, budget spending and painter ordering included.

| parameter | what it is |
|---|---|
| `rd` | The renderer to draw with. |

**Returns.** How many objects were submitted.

Runs in three stages. The budget is spent nearest first, so whatever gets dropped is always the far background and never the wall in front of the player. The survivors are seeded farthest centre first, and then reordered by a real occlusion sort: for every pair that can be separated by an axis-aligned plane, the side the eye sits on decides which one has to be painted second. Guessing from a single depth number cannot do that, which is what makes a small object shine through a wide wall or a ground plane paint over everything standing on it.

### DrawnCount

```vba
Public Property Get DrawnCount() As Long
```

Reads how many objects the last Render painted.

**Returns.** The count, which is what Render returned.

### WasReduced

```vba
Public Function WasReduced(ByVal slot As Long) As Boolean
```

Reports whether the object painted at a given position was drawn reduced.

| parameter | what it is |
|---|---|
| `slot` | The position in the last frame. |

**Returns.** True when only its most visible face was submitted.

> Paired with DrawnAt so a test can see the level of detail decision, which is otherwise invisible: a box drawn with one face and a box drawn with three look identical from the front, and the difference only shows as a flicker when the choice changes between two almost identical frames.

### DrawnAt

```vba
Public Function DrawnAt(ByVal slot As Long) As Long
```

Reads which object was painted at a given position of the last frame.

| parameter | what it is |
|---|---|
| `slot` | The position, zero being the first thing painted and therefore the furthest back. |

**Returns.** The object id, or P_INVALID_ID when the slot is out of range.

> The order the painter's algorithm settled on, which is the one thing about a frame that cannot be checked by looking at the shapes afterwards. Reading it is how a test proves that a coin resting on a platform is painted before the platform when you are underneath it.

### DrawObject

```vba
Public Sub DrawObject(ByVal rd As PRenderer, ByVal idx As Long, Optional ByVal reduced As Boolean = False)
```

Draws a single object through a renderer.

| parameter | what it is |
|---|---|
| `rd` | The renderer to draw with. |
| `idx` | The object id. |

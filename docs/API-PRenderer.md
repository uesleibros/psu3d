# PRenderer

**Pipeline de face e primitivas**

Turns world space faces into PowerPoint polylines: backface rejection, view frustum clipping, perspective projection into a PCanvas, material shading and the pooled shape emission that keeps the slide from thrashing. One renderer drives exactly one canvas, so a slide can host several independent views at once.

> Shapes are named from a per renderer prefix and deleted as a single Range on the next frame, which is by far the cheapest way to clear a slide from VBA.

> **Escopo.** Private to this VBA project: class modules carry VB_Exposed = False, which is the class level equivalent of Option Private Module.

## Indice

**Instance state.** [`Attach`](#attach)

**Bindings.** [`Canvas`](#canvas), [`Camera`](#camera), [`Target`](#target), [`Prefix`](#prefix)

**Frame control.** [`BeginFrame`](#beginframe), [`EndFrame`](#endframe), [`DoubleBuffer`](#doublebuffer), [`RefreshState`](#refreshstate), [`ClearShapes`](#clearshapes), [`Purge`](#purge)

**Budget.** [`SetPool`](#setpool), [`SetBudgetRange`](#setbudgetrange), [`SetFrameTargets`](#setframetargets), [`AdaptBudget`](#adaptbudget), [`AutoBudget`](#autobudget), [`Remaining`](#remaining), [`IsFull`](#isfull), [`PolyCount`](#polycount), [`PolyBudget`](#polybudget), [`PoolSize`](#poolsize)

**Quality tuning.** [`ComSeconds`](#comseconds), [`ComOps`](#comops), [`Profiling`](#profiling), [`DryRun`](#dryrun), [`SeamFill`](#seamfill), [`EdgeInflate`](#edgeinflate), [`MinPolyArea`](#minpolyarea), [`SetFarCulling`](#setfarculling)

**Drawing.** [`DrawQuad`](#drawquad), [`DrawTriangle`](#drawtriangle), [`DrawPolygon2D`](#drawpolygon2d)

**Primitives.** [`DrawBox`](#drawbox), [`DrawBoxLod`](#drawboxlod), [`DrawBoxRotated`](#drawboxrotated), [`DrawRamp`](#drawramp), [`DrawWall`](#drawwall), [`DrawFloor`](#drawfloor), [`DrawBillboard`](#drawbillboard), [`DrawSpinner`](#drawspinner)

## Membros

### Attach

```vba
Public Function Attach(ByVal target As Shapes, ByVal cv As PCanvas, ByVal cam As PCamera) As PRenderer
```

Binds the renderer to a slide, a canvas and a camera.

| parametro | o que e |
|---|---|
| `target` | The Shapes collection of the slide being drawn into. |
| `cv` | The canvas describing where and how to project. |
| `cam` | The camera describing where the view is taken from. |

**Devolve.** The renderer itself, so it can be created and bound in one expression.

### Canvas

```vba
Public Property Get Canvas() As PCanvas
Public Property Set Canvas(ByVal value As PCanvas)
```

Reads the canvas this renderer draws into.

| parametro | o que e |
|---|---|
| `value` | The canvas to draw into. |

**Devolve.** The bound canvas.

### Camera

```vba
Public Property Get Camera() As PCamera
Public Property Set Camera(ByVal value As PCamera)
```

Reads the camera driving this renderer.

| parametro | o que e |
|---|---|
| `value` | The camera to render from. |

**Devolve.** The bound camera.

### Target

```vba
Public Property Get Target() As Shapes
Public Property Set Target(ByVal value As Shapes)
```

Reads the shape collection being drawn into.

| parametro | o que e |
|---|---|
| `value` | The Shapes collection of the target slide. |

**Devolve.** The bound Shapes collection.

### Prefix

```vba
Public Property Get Prefix() As String
Public Property Let Prefix(ByVal value As String)
```

Reads the prefix given to every shape this renderer creates.

Escrita: Sets the prefix given to every shape this renderer creates.

| parametro | o que e |
|---|---|
| `value` | The new prefix; give each renderer on a slide its own so cleanup never crosses over. |

**Devolve.** The shape name prefix.

### BeginFrame

```vba
Public Sub BeginFrame()
```

Opens a frame by wiping the shapes drawn last time and re-reading the canvas and camera.

> Call once per frame before any draw call; the wipe is a single Range delete, not one delete per shape.

### EndFrame

```vba
Public Sub EndFrame()
```

Retires the frame that was on the slide before this one.

> Only does anything while double buffered. Call it once per frame, after every draw call, or the retired frame stays on the slide and the shape count keeps climbing.

### DoubleBuffer

```vba
Public Property Get DoubleBuffer() As Boolean
Public Property Let DoubleBuffer(ByVal value As Boolean)
```

Reports whether the previous frame is kept on the slide while the next one is drawn.

Escrita: Chooses whether the previous frame is kept while the next is drawn.

| parametro | o que e |
|---|---|
| `value` | True to trade a moment of doubled shape count for a slide that is never seen blank. |

**Devolve.** True when double buffered.

> Turning this off restores the clear-then-draw order and removes the need to call EndFrame.

### RefreshState

```vba
Public Sub RefreshState()
```

Re-reads the cached canvas and camera state.

> Only needed when the canvas or camera is moved in the middle of a frame, since BeginFrame already does it.

### ClearShapes

```vba
Public Sub ClearShapes()
```

Deletes every shape emitted during the current frame.

### Purge

```vba
Public Sub Purge()
```

Deletes every shape on the slide that carries this renderer prefix, including leftovers from an interrupted run.

> Worth calling once at start up, since a show stopped mid frame leaves its polygons behind.

### SetPool

```vba
Public Sub SetPool(ByVal maxPolys As Long)
```

Reserves shape name slots up front so no growth happens mid frame.

| parametro | o que e |
|---|---|
| `maxPolys` | How many polygon slots to reserve. |

> The pool is not a ceiling: it grows on demand. Reserving simply moves the allocation out of the render loop.

### SetBudgetRange

```vba
Public Sub SetBudgetRange(ByVal minPolys As Long, ByVal maxPolys As Long)
```

Sets the range the adaptive budget may travel inside.

| parametro | o que e |
|---|---|
| `minPolys` | The floor the budget never drops under, protecting the silhouette of the scene. |
| `maxPolys` | The ceiling the budget never climbs over. |

### SetFrameTargets

```vba
Public Sub SetFrameTargets(ByVal fastSeconds As Single, ByVal slowSeconds As Single)
```

Sets the frame times the adaptive budget aims between.

| parametro | o que e |
|---|---|
| `fastSeconds` | Frames quicker than this earn more polygons. |
| `slowSeconds` | Frames slower than this give polygons back. |

### AdaptBudget

```vba
Public Sub AdaptBudget(ByVal dt As Single)
```

Nudges the polygon budget towards the frame time targets.

| parametro | o que e |
|---|---|
| `dt` | The duration of the frame just finished, in seconds. |

Moves only after several frames agree, because the budget decides where the draw list is cut, and a budget that changes every frame moves that cut every frame: objects sitting on the boundary then blink in and out while the camera is barely moving. A single slow frame is noise, not a trend.

> Does nothing while AutoBudget is False, which is the right setting for a hand authored scene whose worst case fits the budget.

### AutoBudget

```vba
Public Property Get AutoBudget() As Boolean
Public Property Let AutoBudget(ByVal value As Boolean)
```

Reports whether the budget follows the frame time.

Escrita: Chooses whether the budget follows the frame time.

| parametro | o que e |
|---|---|
| `value` | False to pin the budget where it is, which removes the last source of objects popping in and out. |

**Devolve.** True when AdaptBudget is allowed to move it.

### Remaining

```vba
Public Property Get Remaining() As Long
```

Reads how many polygons may still be drawn this frame.

**Devolve.** The remaining budget.

### IsFull

```vba
Public Property Get IsFull() As Boolean
```

Reports whether the frame budget is spent.

**Devolve.** True when further draw calls will be ignored.

### PolyCount

```vba
Public Property Get PolyCount() As Long
```

Reads how many polygons have been emitted this frame.

**Devolve.** The current shape count.

### PolyBudget

```vba
Public Property Get PolyBudget() As Long
Public Property Let PolyBudget(ByVal value As Long)
```

Reads the current polygon budget.

Escrita: Overrides the polygon budget, pinning it until AdaptBudget moves it again.

| parametro | o que e |
|---|---|
| `value` | The new budget. |

**Devolve.** The number of polygons allowed this frame.

### PoolSize

```vba
Public Property Get PoolSize() As Long
```

Reads how many shape name slots are currently reserved.

**Devolve.** The reserved slot count, which grows on demand.

### ComSeconds

```vba
Public Property Get ComSeconds() As Double
```

Reads how long this frame spent inside PowerPoint.

**Devolve.** Seconds spent creating, formatting and deleting shapes.

> Compare against the whole frame time and the answer to where the frame went stops being a guess. Everything not counted here is arithmetic the engine itself did.

### ComOps

```vba
Public Property Get ComOps() As Long
```

Reads how many shapes this frame created.

**Devolve.** The emission count.

### Profiling

```vba
Public Property Get Profiling() As Boolean
Public Property Let Profiling(ByVal value As Boolean)
```

Reports whether COM timing is being collected.

Escrita: Chooses whether COM timing is collected.

| parametro | o que e |
|---|---|
| `value` | False to drop the two clock reads per polygon, which cost well under a percent of a frame. |

**Devolve.** True when profiling.

### DryRun

```vba
Public Property Get DryRun() As Boolean
Public Property Let DryRun(ByVal value As Boolean)
```

Reports whether the pipeline is running without touching the slide.

Escrita: Runs the whole pipeline but creates no shapes, so the cost of the slide can be measured.

| parametro | o que e |
|---|---|
| `value` | True to suppress shape creation. |

**Devolve.** True while shape creation is suppressed.

Culling, clipping, projection, shading and ordering all still run and PolyCount still reports what would have been drawn; only the COM calls are skipped. Time a run with this on and a run with it off, and the difference is what PowerPoint itself costs, which is the number worth knowing before optimising anything else. Toggle it between frames, never inside one.

### SeamFill

```vba
Public Property Get SeamFill() As Boolean
Public Property Let SeamFill(ByVal value As Boolean)
```

Reports how the gap between neighbouring faces is closed.

Escrita: Chooses how the gap between neighbouring faces is closed.

| parametro | o que e |
|---|---|
| `value` | True to stroke each polygon in its own fill colour, False to inflate its outline geometrically. |

**Devolve.** True when the outline is painted in the fill colour, False when polygons are inflated instead.

Two ways to hide the hairline seam PowerPoint leaves between adjacent polygons. Inflating pushes every vertex away from the centre, which distorts small polygons badly because the push is a fixed number of points regardless of how big the polygon is. Stroking the outline in the fill colour expands the shape uniformly by half a line weight, costs exactly the same single COM call, and leaves the geometry untouched. Set EdgeInflate to zero when turning this on, or the two stack.

### EdgeInflate

```vba
Public Property Get EdgeInflate() As Single
Public Property Let EdgeInflate(ByVal value As Single)
```

Reads how far polygon borders are pushed outwards.

Escrita: Sets how far polygon borders are pushed outwards, which hides the hairline gaps between neighbouring faces.

| parametro | o que e |
|---|---|
| `value` | The inflation in points; around half a point is usually enough. |

**Devolve.** The inflation in points.

### MinPolyArea

```vba
Public Property Get MinPolyArea() As Single
Public Property Let MinPolyArea(ByVal value As Single)
```

Reads the smallest screen area a polygon may cover.

Escrita: Sets the smallest screen area a polygon may cover before it is dropped.

| parametro | o que e |
|---|---|
| `value` | The area threshold in square points. |

**Devolve.** The area threshold in square points.

### SetFarCulling

```vba
Public Sub SetFarCulling(ByVal area As Single, ByVal depth As Single)
```

Sets the stricter area threshold applied to distant polygons.

| parametro | o que e |
|---|---|
| `area` | The area threshold in square points. |
| `depth` | The view depth beyond which the stricter threshold applies. |

### DrawQuad

```vba
Public Sub DrawQuad(ByVal x0 As Single, ByVal y0 As Single, ByVal z0 As Single, ByVal x1 As Single, ByVal y1 As Single, ByVal z1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal z2 As Single, ByVal x3 As Single, ByVal y3 As Single, ByVal z3 As Single, ByVal matId As Long, Optional ByVal dirKey As PDirection = pdNone)
```

Draws a quad straight from its four corners.

| parametro | o que e |
|---|---|
| `x0` | The first corner X coordinate. |
| `y0` | The first corner Y coordinate. |
| `z0` | The first corner Z coordinate. |
| `x1` | The second corner X coordinate. |
| `y1` | The second corner Y coordinate. |
| `z1` | The second corner Z coordinate. |
| `x2` | The third corner X coordinate. |
| `y2` | The third corner Y coordinate. |
| `z2` | The third corner Z coordinate. |
| `x3` | The fourth corner X coordinate. |
| `y3` | The fourth corner Y coordinate. |
| `z3` | The fourth corner Z coordinate. |
| `matId` | The material id to shade the quad with. |
| `dirKey` | The axis-aligned normal key, or pdNone when the face is tilted. |

### DrawTriangle

```vba
Public Sub DrawTriangle(ByVal x0 As Single, ByVal y0 As Single, ByVal z0 As Single, ByVal x1 As Single, ByVal y1 As Single, ByVal z1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal z2 As Single, ByVal matId As Long, Optional ByVal dirKey As PDirection = pdNone)
```

Draws a triangle straight from its three corners.

| parametro | o que e |
|---|---|
| `x0` | The first corner X coordinate. |
| `y0` | The first corner Y coordinate. |
| `z0` | The first corner Z coordinate. |
| `x1` | The second corner X coordinate. |
| `y1` | The second corner Y coordinate. |
| `z1` | The second corner Z coordinate. |
| `x2` | The third corner X coordinate. |
| `y2` | The third corner Y coordinate. |
| `z2` | The third corner Z coordinate. |
| `matId` | The material id to shade the triangle with. |
| `dirKey` | The axis-aligned normal key, or pdNone when the face is tilted. |

### DrawPolygon2D

```vba
Public Sub DrawPolygon2D(ByRef ax() As Single, ByRef ay() As Single, ByVal count As Long, ByVal col As Long, Optional ByVal depth As Single = 0!, Optional ByVal matId As Long = P_INVALID_ID)
```

Draws a polygon whose points are already in slide coordinates, bypassing the 3D pipeline.

| parametro | o que e |
|---|---|
| `ax` | The X coordinates of the polygon. |
| `ay` | The Y coordinates of the polygon. |
| `count` | How many points the arrays hold. |
| `col` | The fill colour. |
| `depth` | The view depth used only for the distance based area culling. |
| `matId` | The material whose transparency and outline settings are applied, or P_INVALID_ID for a plain fill. |

> Overlay polygons ignore the frame budget on purpose: a crosshair that vanishes exactly when the scene gets busy is worse than one extra shape, and the caller controls how many of these it asks for.

### DrawBox

```vba
Public Sub DrawBox(ByVal matId As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal zTop As Single, ByVal thickness As Single)
```

Draws an axis-aligned box, submitting only the faces the eye can see.

| parametro | o que e |
|---|---|
| `matId` | The material id of every face. |
| `x1` | The lower X bound. |
| `y1` | The lower Y bound. |
| `x2` | The upper X bound. |
| `y2` | The upper Y bound. |
| `zTop` | The height of the top face. |
| `thickness` | How far the box extends below its top face. |

### DrawBoxLod

```vba
Public Sub DrawBoxLod(ByVal matId As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal zTop As Single, ByVal thickness As Single)
```

Draws an axis-aligned box as the single face that covers the most of it from here.

| parametro | o que e |
|---|---|
| `matId` | The material id of the face. |
| `x1` | The lower X bound. |
| `y1` | The lower Y bound. |
| `x2` | The upper X bound. |
| `y2` | The upper Y bound. |
| `zTop` | The height of the top face. |
| `thickness` | How far the box extends below its top face. |

The reduced form used once a box is small on screen. A box shows at most three faces, and at distance the two minor ones are a few points wide and cost exactly as much to draw as the dominant one. Which face dominates is decided by the eye offset measured in units of the box's own half extents, so a wide flat slab keeps its top and a tall thin post keeps its side.

### DrawBoxRotated

```vba
Public Sub DrawBoxRotated(ByVal matId As Long, ByVal cx As Single, ByVal cy As Single, ByVal halfW As Single, ByVal halfH As Single, ByVal zTop As Single, ByVal thickness As Single, ByVal angleRad As Single)
```

Draws a box spun around its own vertical axis.

| parametro | o que e |
|---|---|
| `matId` | The material id of every face. |
| `cx` | The centre X coordinate. |
| `cy` | The centre Y coordinate. |
| `halfW` | Half the extent along the local X axis. |
| `halfH` | Half the extent along the local Y axis. |
| `zTop` | The height of the top face. |
| `thickness` | How far the box extends below its top face. |
| `angleRad` | The rotation around the vertical axis, in radians. |

> The four side walls carry no direction key, so they are lit from their real normals as the box turns.

### DrawRamp

```vba
Public Sub DrawRamp(ByVal matId As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal zLow As Single, ByVal zHigh As Single, ByVal thickness As Single, Optional ByVal axis As PAxis = paX)
```

Draws a sloped slab climbing along one axis.

| parametro | o que e |
|---|---|
| `matId` | The material id of every face. |
| `x1` | The lower X bound. |
| `y1` | The lower Y bound. |
| `x2` | The upper X bound. |
| `y2` | The upper Y bound. |
| `zLow` | The height at the low end of the slope. |
| `zHigh` | The height at the high end of the slope. |
| `thickness` | How far the slab extends below its sloped face. |
| `axis` | paX to climb along X, paY to climb along Y. |

### DrawWall

```vba
Public Sub DrawWall(ByVal matId As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal zBottom As Single, ByVal zTop As Single)
```

Draws a single upright wall between two ground points.

| parametro | o que e |
|---|---|
| `matId` | The material id of the wall. |
| `x1` | The X coordinate of the first end. |
| `y1` | The Y coordinate of the first end. |
| `x2` | The X coordinate of the second end. |
| `y2` | The Y coordinate of the second end. |
| `zBottom` | The height of the base. |
| `zTop` | The height of the crown. |

> Pair this with a two sided material when the wall must be visible from both faces.

### DrawFloor

```vba
Public Sub DrawFloor(ByVal matId As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal Z As Single)
```

Draws a flat ground quad.

| parametro | o que e |
|---|---|
| `matId` | The material id of the quad. |
| `x1` | The lower X bound. |
| `y1` | The lower Y bound. |
| `x2` | The upper X bound. |
| `y2` | The upper Y bound. |
| `Z` | The height of the quad. |

### DrawBillboard

```vba
Public Sub DrawBillboard(ByVal matId As Long, ByVal X As Single, ByVal Y As Single, ByVal Z As Single, ByVal Width As Single, ByVal Height As Single)
```

Draws an upright quad that always turns its face to the camera.

| parametro | o que e |
|---|---|
| `matId` | The material id of the billboard. |
| `X` | The centre X coordinate. |
| `Y` | The centre Y coordinate. |
| `Z` | The centre Z coordinate. |
| `Width` | The width of the quad in world units. |
| `Height` | The height of the quad in world units. |

> This is how coins, pickups and simple sprites are drawn without building real geometry.

### DrawSpinner

```vba
Public Sub DrawSpinner(ByVal matId As Long, ByVal X As Single, ByVal Y As Single, ByVal Z As Single, ByVal radius As Single, ByVal phase As Single)
```

Draws a spinning double sided plate, the classic collectible look.

| parametro | o que e |
|---|---|
| `matId` | The material id of the plate. |
| `X` | The centre X coordinate. |
| `Y` | The centre Y coordinate. |
| `Z` | The centre Z coordinate. |
| `radius` | The half size of the plate. |
| `phase` | The spin angle in radians. |

## Membros `Friend`

Visiveis para o projeto VBA inteiro, mas nao para fora dele. Um membro publico de classe nao pode receber um tipo declarado em modulo padrao, e e por isso que estes sao `Friend` em vez de `Public`.

### EyeX

```vba
Friend Property Get EyeX() As Single
```

Reads the eye X coordinate cached for this frame.

### EyeY

```vba
Friend Property Get EyeY() As Single
```

Reads the eye Y coordinate cached for this frame.

### EyeZ

```vba
Friend Property Get EyeZ() As Single
```

Reads the eye Z coordinate cached for this frame.

### DrawFace

```vba
Friend Sub DrawFace(ByRef f As PFace, Optional ByVal offX As Single = 0!, Optional ByVal offY As Single = 0!, Optional ByVal offZ As Single = 0!)
```

Draws a single world space face.

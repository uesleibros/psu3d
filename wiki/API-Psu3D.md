# Psu3D

**Fachada**

The one entry point a slide needs: it boots a canvas, a camera, a renderer and a scene, wires them together, and keeps them reachable as engine wide defaults. Everything it hands back is an ordinary object, so a project that outgrows the defaults can build its own set and ignore this module entirely.

> Import order inside the VBE does not matter, but the module list is PCore, PLighting, PMaterial, PMaterials, PCanvas, PCamera, PRenderer, PScene, PBody, PLevel and finally this facade.

## Indice

**Identity.** [`Version`](#version)

**Boot.** [`Boot`](#boot), [`Shutdown`](#shutdown), [`IsReady`](#isready)

**Defaults.** [`Canvas`](#canvas), [`Camera`](#camera), [`Renderer`](#renderer), [`Scene`](#scene)

**Factories.** [`CreateCanvas`](#createcanvas), [`CreateCamera`](#createcamera), [`CreateRenderer`](#createrenderer), [`CreateScene`](#createscene), [`CreateBody`](#createbody)

**Environment shortcuts.** [`SetLight`](#setlight), [`SetFog`](#setfog)

**Frame.** [`BeginFrame`](#beginframe), [`EndFrame`](#endframe), [`RenderScene`](#renderscene)

## Membros

### Version

```vba
Public Function Version() As String
```

Reads the library name and version.

**Devolve.** A short identification string.

### Boot

```vba
Public Sub Boot(ByVal target As Shapes, Optional ByVal X As Single = 0!, Optional ByVal Y As Single = 0!, Optional ByVal Width As Single = 0!, Optional ByVal Height As Single = 0!)
```

Builds the default canvas, camera, renderer and scene and binds them to a slide.

| parametro | o que e |
|---|---|
| `target` | The Shapes collection of the slide that will be drawn into. |
| `X` | The left edge of the view, in slide points; the whole slide is used when width and height are left at zero. |
| `Y` | The top edge of the view, in slide points. |
| `Width` | The view width, in slide points. |
| `Height` | The view height, in slide points. |

> Also registers the stock material palette and sweeps away polygons left behind by an interrupted run.

### Shutdown

```vba
Public Sub Shutdown()
```

Wipes every polygon this engine drew and releases the default objects.

> Call it when leaving the slide, so nothing is left on it once the show moves on.

### IsReady

```vba
Public Property Get IsReady() As Boolean
```

Reports whether Boot has run.

**Devolve.** True when the default objects exist.

### Canvas

```vba
Public Property Get Canvas() As PCanvas
```

Reads the default canvas.

**Devolve.** The canvas built by Boot, or Nothing before it runs.

### Camera

```vba
Public Property Get Camera() As PCamera
```

Reads the default camera.

**Devolve.** The camera built by Boot.

### Renderer

```vba
Public Property Get Renderer() As PRenderer
```

Reads the default renderer.

**Devolve.** The renderer built by Boot.

### Scene

```vba
Public Property Get Scene() As PScene
```

Reads the default scene.

**Devolve.** The scene built by Boot.

### CreateCanvas

```vba
Public Function CreateCanvas(ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, Optional ByVal cvName As String = vbNullString) As PCanvas
```

Creates a canvas, which is how a second view is added to a slide.

| parametro | o que e |
|---|---|
| `X` | The left edge in slide points. |
| `Y` | The top edge in slide points. |
| `Width` | The view width in slide points. |
| `Height` | The view height in slide points. |
| `cvName` | The canvas name, which also names its backdrop shape. |

**Devolve.** The new canvas.

### CreateCamera

```vba
Public Function CreateCamera(Optional ByVal X As Single = 0!, Optional ByVal Y As Single = 0!, Optional ByVal Z As Single = 0!, Optional ByVal yawRad As Single = 0!, Optional ByVal pitchRad As Single = 0!) As PCamera
```

Creates a camera.

| parametro | o que e |
|---|---|
| `X` | The eye X coordinate. |
| `Y` | The eye Y coordinate. |
| `Z` | The eye Z coordinate. |
| `yawRad` | The heading in radians. |
| `pitchRad` | The elevation in radians. |

**Devolve.** The new camera.

### CreateRenderer

```vba
Public Function CreateRenderer(ByVal target As Shapes, ByVal cv As PCanvas, ByVal cam As PCamera, Optional ByVal prefix As String = vbNullString) As PRenderer
```

Creates a renderer already bound to a slide, a canvas and a camera.

| parametro | o que e |
|---|---|
| `target` | The Shapes collection to draw into. |
| `cv` | The canvas to project onto. |
| `cam` | The camera to render from. |
| `prefix` | The shape name prefix; give every renderer on a slide its own. |

**Devolve.** The new renderer.

### CreateScene

```vba
Public Function CreateScene(Optional ByVal slots As Long = 0) As PScene
```

Creates a scene.

| parametro | o que e |
|---|---|
| `slots` | How many object slots to reserve up front. |

**Devolve.** The new scene.

### CreateBody

```vba
Public Function CreateBody(Optional ByVal sc As PScene) As PBody
```

Creates a body bound to a scene.

| parametro | o que e |
|---|---|
| `sc` | The scene it moves through; the booted one when none is named. |

**Devolve.** The new body, ready to be tuned and advanced.

### SetLight

```vba
Public Sub SetLight(ByVal dx As Single, ByVal dy As Single, ByVal dz As Single)
```

Points the global light in a new direction.

| parametro | o que e |
|---|---|
| `dx` | The X component of the direction. |
| `dy` | The Y component of the direction. |
| `dz` | The Z component of the direction. |

### SetFog

```vba
Public Sub SetFog(ByVal startDist As Single, ByVal endDist As Single, ByVal col As Long)
```

Configures the distance fog, and repaints the canvas backdrop to match.

| parametro | o que e |
|---|---|
| `startDist` | The depth at which fog begins. |
| `endDist` | The depth at which geometry disappears into the fog. |
| `col` | The fog colour. |

### BeginFrame

```vba
Public Sub BeginFrame(ByVal dt As Single)
```

Opens a frame on the default renderer and lets the polygon budget follow the frame time.

| parametro | o que e |
|---|---|
| `dt` | The duration of the previous frame, in seconds. |

### EndFrame

```vba
Public Sub EndFrame()
```

Closes the frame, retiring the shapes the previous one left on the slide.

> Required when the renderer is double buffered, which it is by default. Call it after every draw call of the frame.

### RenderScene

```vba
Public Function RenderScene() As Long
```

Draws the default scene through the default renderer.

**Devolve.** How many objects were submitted.

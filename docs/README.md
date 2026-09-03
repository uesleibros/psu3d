# Psu3D

A 3D engine written in pure VBA, running inside PowerPoint.

Psu3D is not a game engine. It is a 3D engine that games can be built on, in the same sense that three.js is. The core, meaning canvas, camera, renderer, scene, material and light, does not know the words *player*, *health* or *level*. Terrain, bar charts, data visualisation and a 3D figure on a slide all use the same core and never import the physics.

## Where to start

| if you want to | go to |
|---|---|
| put the library in your file | [Installation](Installation.md) |
| see something on screen in twenty lines | [Getting started](Getting-Started.md) |
| understand how the pieces fit | [Concepts](Concepts.md) |
| know whether this is real 3D | [How the 3D works](How-The-3D-Works.md) |
| build a game | [Body and physics](Body-And-Physics.md) |
| write levels as files | [Levels in JSON](Levels-In-JSON.md) |
| copy a level and change it | [Examples](Examples.md) |
| something is not working | [Troubleshooting](Troubleshooting.md) |
| a quick answer | [FAQ](FAQ.md) |
| the complete list of everything | [API reference](API-PCore.md) |

## The modules

The first eight are the core. `Psu3D` is a convenience layer. `PBody`, `PLevel`, `PSelfTest` and `PDemo` are imported only if you want them.

| module | what it does | guide |
|---|---|---|
| `PCore.bas` | types, enums, maths, deterministic random, colour, clock | [API](API-PCore.md) |
| `PLighting.bas` | directional light and global fog | [Lighting and fog](Lighting-And-Fog.md) |
| `PMaterial.cls` | the definition of one surface | [Materials](Materials.md) |
| `PMaterials.bas` | material registry and precomputed shading table | [Materials](Materials.md) |
| `PCanvas.cls` | where and how to project | [Canvas](Canvas.md) |
| `PCamera.cls` | eye, yaw, pitch | [Camera](Camera.md) |
| `PRenderer.cls` | face pipeline and the primitives | [Renderer and primitives](Renderer-And-Primitives.md) |
| `PScene.cls` | object store, spatial index, draw order | [Scene](Scene.md) |
| `PBody.cls` | a body that walks, collides, steps up, climbs and swims | [Body and physics](Body-And-Physics.md) |
| `PLevel.bas` | read and write a scene as JSON | [Levels in JSON](Levels-In-JSON.md) |
| `Psu3D.bas` | facade | [API](API-Psu3D.md) |
| `PSelfTest.bas` | library self test | [Self test](Self-Test.md) |
| `PDemo.bas` | playable example level | [Recipes](Recipes.md) |

Nothing third party is bundled, and **the core calls nothing outside itself**. `PLevel` needs a JSON parser, which lives at [vbacollective/json](https://github.com/vbacollective/json). Everything else, including the keyboard and the mouse used by the demo, is declared inside the modules that use it.

## The rules the library follows

These are not preferences. They hold in every file of the repository.

**No `On Error` anywhere.** Every path that could raise, a missing shape name, an invalid hex string, an id out of range, a duplicate key, is checked by hand. `On Error` in VBA hides the fault instead of handling it, and then loses the stack.

**`Option Private Module` in every `.bas`.** Classes cannot carry that statement, so they carry the `VB_Exposed = False` attribute instead, which is the class level equivalent. Two deliberate exceptions: `PSelfTest` and `PDemo`, because the macro dialog only lists public entry points.

**Plain arrays, no COM on the hot path.** COM appears only where it is physically unavoidable: `AddPolyline`, the fill colour of each polygon, and the single `Range(...).Delete` that retires a whole frame.

**No fixed ceilings.** The material registry, the shape pool and the object store all grow by doubling. There is no `MAX_WORLD_PLATS`.

**A UDT in a class is `Friend`, never `Public`.** VBA refuses a user defined type from a standard module in the public signature of a class module. Enums are fine as public; the rule applies only to `Type`.

## Status

178 self test assertions, 13 modules, roughly 11,000 lines of VBA. Run `PSelfTest.Psu3DSelfTest` after importing.

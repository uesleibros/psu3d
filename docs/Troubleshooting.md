# Troubleshooting

## Nothing appears on the slide

**The refresh pump is missing or throttled.** Writing to a WordArt shape is what forces PowerPoint to repaint during a show. It has to fire every frame, and the value has to actually change. See [Performance](Performance.md#the-refresh-pump).

**There is no `DoEvents`.** The repaint request sits in the queue until one runs.

**`Boot` was never called, or was called with the wrong `Shapes`.** In a slide module `Shapes` resolves to that slide. In a standard module it does not, and you need `Slide1.Shapes`.

**The budget is zero or tiny.** `Psu3D.Renderer.PolyBudget` decides how many polygons a frame may spend. A budget under the cost of one object draws one object.

**The camera is inside the geometry, or facing away.** Try `Psu3D.Camera.LookAt` aimed at something you know exists.

## It compiles but the self test fails

The report names every check that failed. The most common cause is a module that did not come across during the import, and the failing checks cluster in that module's block.

If nothing at all runs, the project did not compile, and the usual reason is `PLevel` being imported without a JSON parser. See [Installation](Installation.md).

## The VBE refuses to compile

**"User defined type not defined"** on a `JSON` reference: `PLevel` is imported and `JSON.cls` is not. Either import the parser or remove `PLevel`.

**"Only public user defined types ... can be used as parameters"**: a class is exposing a `Type` from a standard module in a `Public` signature. It must be `Friend`. The library follows that rule everywhere; if you hit it, it is in your own code.

**"Ambiguous name detected"**: the same module was imported twice, usually because a `.cls` was renamed. The VBE reads the class name from the `VB_Name` attribute inside the file, not from the file name.

## Macros will not run at all

PowerPoint blocks macros in a file downloaded from the internet. Right click the `.pptm`, choose **Properties**, and tick **Unblock** at the bottom of the General tab.

## Objects flicker or blink

**Two objects genuinely interpenetrate.** Without a Z buffer there is no correct order for geometry that passes through other geometry. Move them apart so they touch rather than overlap. See [Known limits](Known-Limits.md#no-z-buffer).

**The budget is right on the edge.** If the scene needs 100 polygons and the budget is 100, an object sits on the cut. Pin the budget above the worst case: `tools/check-level.py` reports the worst case for a level file.

Level of detail and budget cuts both have hysteresis, so neither should chatter on its own. If something still blinks, it is the first cause.

## The body falls through the floor

**The floor is a `trigger` or a `ghost`.** Only `solid` blocks from every side. `oneway` holds a body arriving from above and lets it through from below.

**The frame took too long.** A body moving at 5 units a second with a frame of half a second moves 2.5 units at once, and can cross a thin floor entirely. Clamp `dt`:

```vba
If dt > 0.05 Then dt = 0.05
```

**The floor is thinner than the body travels in one frame.** Give thin platforms more `thick`. The thickness is what a falling body has to be caught by.

## The body will not step onto something

Compare the rise against `StepHeight`. A material with its own `step` overrides the body's, and `step` of zero means the body decides. A lip taller than the step has to be jumped.

## Mouse look drifts or spins

The pointer is warped back to a fixed anchor every frame. If something else moves the pointer, or the show does not have the screen, the measured travel is wrong. Run the demo from a running slide show.

## A level file will not load

`PLevel.Error` names the reason. The parser is deliberately strict about anything that would silence a mistake: an unknown material, an unknown type, a document with no objects, and a `repeat` that can only be a typo.

Everything else is optional with a stated default, so a half written level loads and shows you what you did write.

## Frame rate is poor

Measure before changing anything:

```vba
rd.Profiling = True
rd.DryRun = True     ' the whole pipeline with no shapes created
```

The difference between the time with and without `DryRun` is what PowerPoint charges. If that dominates, no amount of shaving the pipeline will help, and the lever is how much the repaint has to draw: a lower `PolyBudget`, a higher `LodSize`, or fewer objects on screen.

If your own code dominates instead, look at the object count reaching the frustum cull, which is the one pass that is still linear.

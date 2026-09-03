# Self test

```vba
PSelfTest.Psu3DSelfTest        ' shows the result in a MsgBox
txt = PSelfTest.RunAll()       ' returns the report as a string
```

178 assertions against known answers, running entirely in memory. The renderer is put into `DryRun`, so the whole pipeline executes with no slide and leaves nothing behind. It can be run from the VBE with no presentation open.

It exists because the failures that matter in a hand imported library are silent ones: a module left out, a class that did not come across, a parser that is missing. The report names every check that failed.

## What each block covers

| block | what it proves |
|---|---|
| `CheckMath` | clamp, lerp, angles, overlap, deterministic random |
| `CheckColor` | packing, blending, hex round trip |
| `CheckLighting` | direction, intensity, fog bands |
| `CheckMaterials` | registry, ids, shading table, invalidation |
| `CheckCanvas` | rectangle, projection, local coordinates |
| `CheckCamera` | yaw, pitch, limits, `LookAt`, transform |
| `CheckScene` | adding, editing, span, ramp, turned box, slot reuse |
| `CheckIndex` | the grid answers exactly what a full sweep would answer |
| `CheckMotion` | oscillation, reported step, spin |
| `CheckBody` | gravity, walls, steps, ladders, water, triggers, bounce, being carried |
| `CheckOrder` | the coin is painted before the platform from below, and after it from above |
| `CheckRenderer` | the whole pipeline with no slide |
| `CheckLevel` | parsing, `repeat`, spawn, budget, named errors |
| `CheckSave` | writing and reading back gives the same world |
| `CheckHardening` | a typo in a file and an absurd coordinate break nothing |

## Four blocks written with extra care

The material's test colour is light enough that the top and the bottom stay distinguishable after shading. With a dark colour both round to the same value and the test would pass by accident.

In `CheckBody`, everything meant to be stood on sits above the floor. Two surfaces at the same height would make the test depend on which one the query reached first, and a test like that passes and fails without anyone changing anything.

In `CheckOrder`, besides the coin and the platform there is a distant decoy that shares no pixel with either. It is the point of the test: it was the decoy that closed the cycle that sacrificed the coin.

`CheckSave` verifies the saved document uses a dot for the decimal point. It is the only saving failure that does not show up as a wrong number: on a machine configured for commas, `Format` would write a file no JSON reader will take back.

## If the JSON parser is missing

`CheckLevel`, `CheckSave` and `CheckHardening` need `JSON.cls`. Without it the project does not compile and the self test never runs at all. That is the symptom of "I forgot to import the parser".

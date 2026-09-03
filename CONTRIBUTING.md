# Contributing

## The rules that are not negotiable

These hold in every file. A change that breaks one of them will be asked to change, however good the rest of it is.

**No `On Error`, anywhere.** Every path that could raise is checked by hand: a missing shape name, an invalid hex string, an id out of range, a duplicate key. `On Error Resume Next` hides the fault and loses the stack, and `On Error GoTo` in VBA does not compose. If a path can fail, prove it cannot, or return a value that says it did.

**`Option Private Module` in every `.bas`, `VB_Exposed = False` in every `.cls`.** The library must not leak into the rest of somebody's file. Two deliberate exceptions, `PSelfTest` and `PDemo`, exist because the macro dialog only lists public entry points.

**No COM on the hot path.** The per frame path never reads a shape property. Everything it would ask about is already ours, in an array. The only shape reads in the engine run at boot, and a change that adds one will be caught by the audit in the pull request.

**No fixed ceilings.** Stores grow by doubling. There is no `MAX_WORLD_PLATS`.

**A UDT in a class is `Friend`, never `Public`.** VBA refuses a user defined type from a standard module in the public signature of a class module. Enums are fine as public.

**No em dashes and no horizontal rules in the documentation.** A checker runs over the markdown.

## Measure anything that can be wrong

Most of this library can only be slow. Two parts can be **wrong**, and both are held to a different standard:

- **Depth sorting.** A wrong order is a visible artefact. It is measured against an exact oracle that runs a separating axis test on the true oriented shapes, counting only pairs that actually overlap on screen.
- **The spatial index.** A query that misses an object is a body falling through the floor. It is compared against the full sweep it replaces, on scenes built out of what breaks grids.

If you change either, bring the numbers. Several plausible improvements to both have been measured and rejected because they made things worse, and the only way anyone knew was the measurement.

The same applies to a fix for something visual. "It looks better" is not a report; "four level of detail pops over 260 frames became none" is.

## Style

The code reads like prose and the comments explain *why*, not *what*. A comment that restates the line above it is noise. A comment that says why the obvious approach was rejected is the most valuable thing in the file.

Docstrings follow the `UCursor` convention already used throughout:

```vba
'/**
' * @brief One line, what it does.
' * @param name What the parameter is.
' * @return What comes back.
' * @description The longer explanation, when there is one.
' * @remarks The caveat, or the reason it is written this way.
' */
```

The API reference is generated from these, so a missing `@param` shows up as a missing row in a published table.

## Before opening a pull request

```
python tools/check-level.py examples/*.json
python tools/wiki-sync.py ../psu3d-wiki
```

And in the VBE, `PSelfTest.Psu3DSelfTest`. It runs entirely in memory, with the renderer in `DryRun`, so it needs no slide and leaves nothing behind.

If you added or changed a public member, regenerate the reference so the documentation matches the code.

## Adding a self test

A new behaviour should come with an assertion. Two things make a test worth having:

**It has to be able to fail.** A check that passes with the change reverted is measuring nothing. Revert the change and watch it fail before you trust it.

**It must not depend on something another block might have changed.** Blocks run in order and share global state: the material registry, the light, the fog. `CheckStability` pins the far plane for exactly that reason, because it would otherwise inherit the end of the fog from whatever ran before it.

## Reporting a bug

The most useful report has the smallest level file, or the shortest Sub, that shows it. If it is visual, say what you expected to see and what you saw instead, and from roughly where the camera was: a lot of ordering behaviour depends on which side of an object the eye is on.

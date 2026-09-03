# Examples

Four level files ship with the library, in [`examples/`](https://github.com/uesleibros/psu3d/tree/main/examples). Each one is a plain `.json` document, editable in any text editor, and each is validated by `tools/check-level.py` before release.

```vba
PDemo.RunFile "C:\path\to\psu3d\examples\showcase.json"
```

The slide needs a WordArt shape named `timer`. It is the refresh pump, not a clock. See [Performance](Performance.md#the-refresh-pump).

## `minimal.json`

Three objects: a ground plane, one block to jump onto, and one coin. Twenty five lines including the level block.

This is the file to copy when starting your own. It uses no feature that needs explaining, which is the point: it shows the shape of a document and nothing else.

## `pool.json`

A swimming pool, for fluids.

| what to look at | why |
|---|---|
| the `water` material | `buoyancy` above 1 floats you back up, `drag` bleeds all three velocity components, `speed` slows walking through it |
| the ladders on both walls | `ghost` plus `climbable`: you enter them and climb, rather than being blocked |
| the diving board and its ramp | a `ramp` running along `y` with `low` at the far end, so you walk up it |
| three coins on the pool floor | they are only reachable by diving, which is `Shift` while swimming |

Jumping in from the board is the fastest way to feel buoyancy and drag: the fall is fast, the entry is soft, and you float back up without touching a key.

## `showcase.json`

One of everything. If you want to see what a material property actually does, this is where to look.

| section | what it demonstrates |
|---|---|
| stairs on the left | `repeat` and `step` writing a staircase as one entry |
| the ramp | `TopZAt` interpolating a slope, walked down rather than up |
| the ice strip | `friction` 0.08 and `speed` 1.25 |
| the kerb and the wall | the same height, one with `step` 0.9 and one with `step` 0.05 |
| the trampoline | `bounce` 1.45 reaching a ledge a plain jump cannot |
| the lava pool and its stepping stones | `damage`, and `repeat` again |
| two stacked one way platforms | `oneway`: you rise through them and land on top |
| the ladder to the high walkway | `climbable` reaching a place with no jump to it |
| the purple pillar | `ghost`: drawn and walked straight through |
| the glass pane and the panel beside it | `alpha` with an `edge`, next to a `clip` collider that is solid and invisible |
| two moving platforms | `move` along `x` and along `z` |
| the turntable | `spin`, which carries you around its centre and turns your view with it |
| the white banner | `bill`, a quad that always faces the camera |

## `obby.json`

A linear open air obstacle course, and the level the depth sorting was tuned against.

There is no ground. Falling kills you and returns you to the last checkpoint, of which there are five. 38 objects: a beam to walk like a tightrope, two moving platforms out of phase with each other, three turntables, an ice run, a trampoline, one way platforms and a lift.

Every gap in it was audited against the physics the level declares. Two jumps were originally impossible, rising 1.30 against an apex of 1.14, and were lowered.

## Checking your own

```
python tools/check-level.py mylevel.json
```

It reports two kinds of mistake. The first is a level that will not parse: an unknown material, an unknown type, a field that is not in the schema. `PLevel` reports those too, but only after a round trip through the VBE.

The second is the one nothing else reports: a level that parses perfectly and cannot be finished, because a jump is taller than the level's own jump height. The checker floods outwards from the spawn, walking, jumping, riding lifts and climbing ladders, using the gravity and jump speed the level declares, and names every surface it could not reach.

```
ok   showcase.json       35 objects  17 materials  100 polys worst case  jump clears 1.14  gap 3.84
```

It also warns when the budget ceiling is under the worst case, which is the difference between a level that never cuts a polygon and one that might.

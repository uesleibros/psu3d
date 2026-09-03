# Example levels

Four level documents. Each is plain JSON, editable in any text editor, and each is validated by `tools/check-level.py` before release.

```vba
PDemo.RunFile "C:\path\to\psu3d\examples\showcase.json"
```

The slide needs a WordArt shape named `timer`. It is the refresh pump, not a clock.

| file | objects | what it is for |
|---|---|---|
| [`minimal.json`](minimal.json) | 3 | the smallest document that is still a level, to copy when starting your own |
| [`pool.json`](pool.json) | 24 | fluids: buoyancy, drag, swimming, diving, and ladders out of the water |
| [`showcase.json`](showcase.json) | 35 | one of every material behaviour and every primitive |
| [`obby.json`](obby.json) | 38 | a linear open air obstacle course, and the level the depth sorting was tuned against |

## What each one shows

**`minimal.json`** is a ground plane, one block to jump onto, and one coin. It uses no feature that needs explaining, which is the point.

**`pool.json`** is a swimming pool. The `water` material is a trigger with `buoyancy` above 1, so letting go floats you back up, and `drag`, which is what makes a fall into water land softly. The ladders on both walls are `ghost` plus `climbable`, so you enter them rather than being blocked by them. Three coins sit on the pool floor and are only reachable by diving.

**`showcase.json`** has one of everything: a staircase written with `repeat` and `step`, a ramp, an ice strip, a kerb and a wall of identical height that differ only in their `step`, a trampoline reaching a ledge a plain jump cannot, a lava pool with stepping stones, two stacked one way platforms, a ladder to a high walkway, a pillar you walk straight through, a glass pane beside an invisible collider, two moving platforms, a turntable that carries you around and turns your view with it, and a billboard.

**`obby.json`** has no ground at all. Falling kills you and returns you to the last of five checkpoints. Every gap in it was audited against the physics the level declares; two jumps were originally impossible, rising 1.30 against an apex of 1.14, and were lowered.

## Checking a level

```
python ../tools/check-level.py mylevel.json
```

It reports two kinds of mistake. The first is a level that will not parse: an unknown material, an unknown type, a field that is not in the schema. `PLevel` reports those too, but only after a round trip through the VBE.

The second is the one nothing else reports: a level that parses perfectly and cannot be finished. The checker floods outwards from the spawn, walking, jumping, riding lifts and climbing ladders, using the gravity and jump speed the level itself declares, and names every surface it could not reach.

```
ok   showcase.json       35 objects  17 materials  100 polys worst case  jump clears 1.14  gap 3.84
```

It also warns when the budget ceiling is under the worst case, which is the difference between a level that never cuts a polygon and one that might.

## The format

The full schema is documented in [Levels in JSON](../docs/Levels-In-JSON.md). The short version:

```json
{
  "level":     { "spawn": [0, -6, 1.2], "yaw": 90, "gravity": 18, "jump": 6.4, "walk": 5.4 },
  "materials": { "stone": { "color": "#C8C8C8" } },
  "objects":   [ { "type": "box", "mat": "stone", "from": [-2, -2], "to": [2, 2],
                   "top": 1, "thick": 1 } ]
}
```

Every field is optional and every default is stated, so a half written level loads and shows you what you did write.

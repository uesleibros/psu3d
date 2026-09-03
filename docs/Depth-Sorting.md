# Depth sorting

With no Z buffer, the order polygons are painted in *is* the depth. This is the hardest part of the library and worth understanding.

## The four stages of `Render`

**1. Spend the budget.** Visible objects are sorted by the near edge of their bounding sphere, and the cost per primitive is charged until the ceiling is reached. What falls off is always the background, never the wall in front of the player.

| primitive | cost |
|---|---|
| box | 3 |
| rotated box | 6 |
| ramp | 4 |
| billboard | 1 |
| spinner | 2 |

An object sitting exactly on the budget line is allowed to overspend a little rather than disappear, and only goes when it is properly past the line. Without that it is cut on one frame and drawn on the next, vanishing and returning while nothing about it changed.

**2. Seed the order.** The survivors are sorted by centre depth, furthest first.

**3. Filter by screen space.** The eight corners of each survivor's box are projected into a rectangle. Two objects whose rectangles do not touch generate no constraint at all.

**4. Order by occlusion.** For each remaining pair, look for an axis aligned plane that separates the two boxes. If one exists, the side the eye is on decides exactly which is behind. Each relation becomes an edge, and a topological sort builds the sequence that satisfies all of them at once.

## Why a separating plane, and not depth

No single depth number solves this.

Sorting by centre makes the floor, which is wide and has a near centre, paint over everything standing on it. Sorting by far edge makes a small distant object cut through the wide wall in front of it.

The separating plane test gets both right, which is what block geometry needs.

## When no sequence exists

Three objects can each be provably behind the next, forming a cycle, and then no order satisfies everything. Something has to give.

The rule is whoever owes least: the object with the fewest outstanding constraints is forced into place, and ties go to the furthest, because a tangle no rule can order exactly is still best painted back to front.

A pair that overlaps on all three axes has no exact answer, because the two genuinely interpenetrate, and is decided by centre depth with size as the final tiebreak.

## The coin through the platform

Symptom: looking up from below, a coin resting on a platform stayed visible through it, and turning objects painted over things clearly in front of them.

The comparator was right the whole time. Measured against an exact oracle, 628 of 638 violations had the pair decided correctly. What threw the decision away was the cycle tiebreak. The cycle was this:

| object | constraint |
|---|---|
| platform `z[8.95, 9.25]` | before the coin, by a Z plane, because the eye is below |
| coin `z[9.73, 10.37]` | before the moving platform, by an X plane |
| moving platform `y[52, 53.6]` | before the first, by a Y plane |

All three correct, and impossible together. But the first and the third did not share a single pixel: one appeared low on screen, the other high. An invisible edge, a real cycle, and the old tiebreak, which chose the object whose blocker was smallest, sacrificed the coin every time, because a coin is always the small thing.

Four fixes, each measured before being written:

| fix | why |
|---|---|
| screen space filter | kills the invisible edges, which were the source of the cycles |
| rectangle rather than circle | the bounding sphere is too loose: it discarded 15% of pairs, the rectangle discards 60% |
| tiebreak by least debt, ties to the furthest | the old rule picked the victim by the size of its blocker, which always condemned the small one |
| the turned box AABB follows its angle | it used to be the sweep circle, which claims ground the slab does not occupy |

Result against the exact oracle, counting only pairs that actually overlap on screen:

| scene | before | after |
|---|---|---|
| `obby.json`, 1743 cameras across 4 spin angles | 1.846 % | 0.000 % |
| adversarial scene: giant floor, slabs sunk into it, coins, discs, ramps, 40 seeds | 6.152 % | 1.014 % |

Of the remaining 1.01%, 34 of 36 cases are still genuine cycles and 2 are geometry that really does interpenetrate, where no answer is correct without splitting polygons.

Cycle breaking by strongly connected component and by least visible damage were both tried. Neither improved anything: the simple rule already picks inside the cycle. That is the floor of this approach.

## Temporal stability

A frame can be correct and still look wrong if the next frame disagrees with it. Along a smooth camera path over 260 frames, measuring only pairs that overlap on screen and whose true order did not change between frames:

| source | before | after |
|---|---|---|
| level of detail popping | 4 | 0 |
| objects cut by the budget and returning | 14 per 400 frames | 0 |
| order flips from genuine cycles | 5 | 5 |

The first two were fixed with hysteresis. The third is the same cycle residue as above, and five alternative rules were measured against it without improvement.

## Inspecting

```vba
sc.Render rd
For i = 0 To sc.DrawnCount - 1
    Debug.Print i, sc.DrawnAt(i)
Next i
```

It is the one thing about a frame that cannot be checked by looking at the shapes afterwards.

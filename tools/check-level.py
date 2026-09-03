# -*- coding: utf-8 -*-
"""Checks a Psu3D level file before you ever open PowerPoint.

Two kinds of mistake cost the most time. The first is a level that will not parse:
an unknown material, an unknown type, a field that does not exist in the schema.
PLevel reports those, but only after a round trip through the VBE. The second is a
level that parses perfectly and cannot be finished, because a jump is taller than
the level's own jump height. Nothing reports that at all.

This checks both, using the physics the level itself declares in its `level` block,
so a level with a heavier gravity is judged against that gravity.

Usage:

    python tools/check-level.py examples/obby.json
    python tools/check-level.py examples/*.json
"""
import glob
import io
import json
import math
import os
import sys

TYPES = {"box", "ramp", "rot", "bill", "spin"}
COLLISIONS = {"solid", "ghost", "oneway", "trigger"}
MAT_FIELDS = {"color", "edge", "collision", "alpha", "friction", "bounce", "speed", "damage",
              "buoyancy", "drag", "step", "unlit", "twosided", "climbable", "fog", "visible",
              "invisible"}
OBJ_FIELDS = {"type", "mat", "from", "to", "at", "half", "size", "top", "thick", "low", "high",
              "angle", "radius", "repeat", "step", "move", "spin", "tag", "axis"}
LEVEL_FIELDS = {"spawn", "yaw", "budget", "lod", "fog", "light", "killz", "gravity", "jump", "walk"}
COST = {"box": 3, "rot": 6, "ramp": 4, "bill": 1, "spin": 2}

MAX_REPEAT = 100000
NARROW = 0.8          # a footprint thinner than this is a wall or a rail, not a floor


class Surface(object):
    """A place a body could stand, with the two heights a ramp has."""

    def __init__(self, x1, y1, x2, y2, low, high, mat, lift):
        self.x1, self.y1, self.x2, self.y2 = x1, y1, x2, y2
        self.low = low                 # the height you can arrive at
        self.high = high               # the height it takes you to
        self.mat = mat
        self.lift = lift               # how far it travels upwards on its own
        self.seen = False

    @property
    def narrow(self):
        return (self.x2 - self.x1) < NARROW or (self.y2 - self.y1) < NARROW

    def __repr__(self):
        return ("%s at z %.2f, x[%.1f,%.1f] y[%.1f,%.1f]"
                % (self.mat, self.low, self.x1, self.x2, self.y1, self.y2))


def expand(doc):
    """Applies repeat and step exactly as PLevel.ReadObject does.

    Capped at the same limit the parser enforces. A document asking for a billion copies
    is a typo, and a checker that tries to honour it before reporting it is no better than
    the parser that would have hung on the same number.
    """
    out = []
    for it in doc.get("objects", []):
        rep = max(1, min(int(it.get("repeat", 1)), MAX_REPEAT))
        st = it.get("step", {})
        cur = {k: (list(v) if isinstance(v, list) else v)
               for k, v in it.items() if k not in ("repeat", "step")}
        for _ in range(rep):
            out.append({k: (list(v) if isinstance(v, list) else v) for k, v in cur.items()})
            for key in ("from", "to", "at", "half", "size"):
                if key in st:
                    base = cur.get(key, [0, 0, 0])
                    cur[key] = [base[i] + (st[key][i] if i < len(st[key]) else 0)
                                for i in range(len(base))]
            for key in ("top", "thick", "low", "high", "angle", "radius"):
                if key in st:
                    cur[key] = cur.get(key, 0) + st[key]
    return out


def lift_of(o):
    """How far a moving object rises above its resting height."""
    mv = o.get("move")
    if not mv:
        return 0.0
    axis = mv.get("axis", [0, 0, 1])
    az = axis[2] if len(axis) > 2 else 0.0
    mag = math.sqrt(sum(a * a for a in axis)) or 1.0
    return abs(mv.get("amp", 0.0) * az / mag)


def surfaces(objs, mats):
    out = []
    for o in objs:
        mat = mats.get(o.get("mat"), {})
        if mat.get("collision") in ("ghost", "trigger"):
            continue
        t = o.get("type")
        lift = lift_of(o)
        if t == "box":
            x1, y1 = o.get("from", [0, 0])[:2]
            x2, y2 = o.get("to", [0, 0])[:2]
            z = o.get("top", 0.0)
            out.append(Surface(min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2), z, z,
                               o["mat"], lift))
        elif t == "ramp":
            x1, y1 = o.get("from", [0, 0])[:2]
            x2, y2 = o.get("to", [0, 0])[:2]
            lo, hi = o.get("low", 0.0), o.get("high", 1.0)
            # you get on at the low end and it takes you to the high end
            out.append(Surface(min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2),
                               min(lo, hi), max(lo, hi), o["mat"], lift))
        elif t == "rot":
            cx, cy = o.get("at", [0, 0])[:2]
            hw, hh = o.get("half", [1, 1])[:2]
            r = math.hypot(hw, hh)
            z = o.get("top", 0.0)
            out.append(Surface(cx - r, cy - r, cx + r, cy + r, z, z, o["mat"], lift))
    return out


def ladders(objs, mats):
    out = []
    for o in objs:
        if not mats.get(o.get("mat"), {}).get("climbable"):
            continue
        if o.get("type") != "box":
            continue
        x1, y1 = o.get("from", [0, 0])[:2]
        x2, y2 = o.get("to", [0, 0])[:2]
        top = o.get("top", 0.0)
        out.append((min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2),
                    top - o.get("thick", 0.5), top))
    return out


def gap(a, b):
    dx = max(a.x1 - b.x2, b.x1 - a.x2, 0.0)
    dy = max(a.y1 - b.y2, b.y1 - a.y2, 0.0)
    return math.hypot(dx, dy)


def check(path):
    problems = []
    doc = json.loads(io.open(path, encoding="utf-8").read())

    lv = doc.get("level", {})
    for k in lv:
        if k not in LEVEL_FIELDS:
            problems.append("level: unknown field %r" % k)

    mats = doc.get("materials", {})
    for name, d in mats.items():
        for k in d:
            if k not in MAT_FIELDS:
                problems.append("material %r: unknown field %r" % (name, k))
        if "collision" in d and d["collision"] not in COLLISIONS:
            problems.append("material %r: collision must be one of %s"
                            % (name, ", ".join(sorted(COLLISIONS))))
        if d.get("buoyancy", 0) > 0 and d.get("collision") != "trigger":
            problems.append("material %r: buoyancy only acts on a trigger" % name)

    raw_objs = doc.get("objects", [])
    if not raw_objs:
        problems.append("the document describes no objects")

    for i, it in enumerate(raw_objs):
        for k in it:
            if k not in OBJ_FIELDS:
                problems.append("object %d: unknown field %r" % (i, k))
        if it.get("type") not in TYPES:
            problems.append("object %d: type must be one of %s" % (i, ", ".join(sorted(TYPES))))
        if it.get("mat") not in mats:
            problems.append("object %d: unknown material %r" % (i, it.get("mat")))
        if it.get("type") == "ramp" and it.get("axis", "x") not in ("x", "y"):
            problems.append("object %d: axis must be x or y" % i)
        if int(it.get("repeat", 1)) > MAX_REPEAT:
            problems.append("object %d: repeat is past the parser limit of %d" % (i, MAX_REPEAT))

    if problems:
        # the schema is wrong, so anything measured on top of it would be measured on sand
        return problems, None

    objs = expand(doc)

    if len(objs) > 4000:
        problems.append("%d objects after repeat: too many to check reachability, which is "
                        "quadratic. The level may still be fine." % len(objs))
        return problems, None

    used = set(o.get("mat") for o in objs)
    for name in sorted(set(mats) - used):
        problems.append("material %r is declared and never used" % name)

    worst = sum(COST.get(o.get("type"), 0) for o in objs)
    budget = lv.get("budget", [0, 0])
    if len(budget) > 1 and budget[1] and worst > budget[1]:
        problems.append("budget ceiling %d is under the worst case of %d polygons, so the draw "
                        "list can be cut" % (budget[1], worst))

    g = lv.get("gravity", 18.0)
    jump = lv.get("jump", 6.4)
    walk = lv.get("walk", 5.4)
    apex = jump * jump / (2.0 * g)
    reach = walk * (2.0 * jump / g)

    surf = surfaces(objs, mats)
    lads = ladders(objs, mats)
    springy = set(n for n, d in mats.items() if d.get("bounce", 0) > 0)

    sx, sy, sz = (list(lv.get("spawn", [0, 0, 1])) + [0, 0, 0])[:3]
    kill = lv.get("killz")
    if kill is not None and sz <= kill:
        problems.append("the spawn is at or below killz")

    start = None
    for s in surf:
        if s.x1 - 0.4 <= sx <= s.x2 + 0.4 and s.y1 - 0.4 <= sy <= s.y2 + 0.4 and s.low <= sz + 0.15:
            if start is None or s.low > start.low:
                start = s
    if start is None:
        problems.append("the spawn at (%g, %g, %g) has nothing under it" % (sx, sy, sz))

    def linked(a, b):
        if gap(a, b) > reach:
            return False
        rise = b.low - a.high
        if rise <= 0.05:
            return True
        ceiling = apex + a.lift
        if a.mat in springy:
            ceiling = apex * 2.4 + a.lift
        return rise <= ceiling

    if start is not None:
        start.seen = True
        stack = [start]
        while stack:
            cur = stack.pop()
            for s in surf:
                if s.seen:
                    continue
                if linked(cur, s):
                    s.seen = True
                    stack.append(s)
                    continue
                for lx1, ly1, lx2, ly2, lz1, lz2 in lads:
                    near_cur = (max(cur.x1 - lx2, lx1 - cur.x2, 0) < 1.2
                                and max(cur.y1 - ly2, ly1 - cur.y2, 0) < 1.2
                                and lz1 - 0.6 <= cur.high <= lz2 + 0.6)
                    near_s = (max(s.x1 - lx2, lx1 - s.x2, 0) < 1.2
                              and max(s.y1 - ly2, ly1 - s.y2, 0) < 1.2
                              and lz1 - 0.6 <= s.low <= lz2 + 0.6)
                    if near_cur and near_s:
                        s.seen = True
                        stack.append(s)
                        break

        for s in surf:
            # a rail or a wall is not somewhere anybody was meant to stand
            if not s.seen and not s.narrow:
                problems.append("%s cannot be reached: the level declares a jump that clears "
                                "%.2f and a gap of %.2f" % (s, apex, reach))

    return problems, {
        "objects": len(objs), "materials": len(mats), "worst": worst,
        "apex": apex, "reach": reach,
    }


def main(paths):
    bad = 0
    for path in paths:
        problems, info = check(path)
        mark = "ok " if not problems else "BAD"

        if info is None:
            print("%s  %s" % (mark, os.path.basename(path)))
        else:
            print("%s  %-18s %3d objects  %2d materials  %3d polys worst case  "
                  "jump clears %.2f  gap %.2f"
                  % (mark, os.path.basename(path), info["objects"], info["materials"],
                     info["worst"], info["apex"], info["reach"]))
        for p in problems:
            print("       %s" % p)
        if problems:
            bad += 1
    print()
    print("all levels valid" if not bad else "%d of %d levels have problems" % (bad, len(paths)))
    return 1 if bad else 0


if __name__ == "__main__":
    args = sys.argv[1:]
    if not args:
        print(__doc__)
        sys.exit(2)
    files = []
    for a in args:
        files.extend(sorted(glob.glob(a)) or [a])
    sys.exit(main(files))

# Levels in JSON

A level stops being code. A `.json` file is plain JSON: any editor validates it, any tool can generate it, and nobody has to open the VBE to change a map.

```vba
If PLevel.ParseFile("C:\path\level.json", sc) Then
    ' PLevel.SpawnX, SpawnY, SpawnZ, SpawnYaw
    ' PLevel.Gravity, Jump, Walk, KillZ, HasKillZ
    ' PLevel.BudgetMin, BudgetMax
Else
    MsgBox PLevel.Error
End If
```

Parsing is done by `JSON.cls`. `PLevel` only knows the schema.

## A document

```json
{
  "level": { "spawn": [0, -12, 0.3], "yaw": 90, "budget": [120, 150],
             "fog": [9, 32, "#A0B9D2"], "light": [0.38, 0.42, 0.82], "lod": 44,
             "killz": -6, "gravity": 18, "jump": 6.4, "walk": 5.4 },

  "materials": {
    "ice":    { "color": "#AADCF5", "friction": 0.12, "speed": 1.15 },
    "water":  { "color": "#286EBE", "collision": "trigger", "alpha": 0.4,
                "buoyancy": 1.06, "drag": 3.4 },
    "beam":   { "color": "#E0DCC8", "step": 0.6 },
    "ladder": { "color": "#8C6E3C", "collision": "ghost", "climbable": true }
  },

  "objects": [
    { "type": "box", "mat": "stone", "from": [-4,-12], "to": [-2.8,-8], "top": 0.4, "thick": 0.4,
      "repeat": 6, "step": { "from": [1.2,0], "to": [1.2,0], "top": 0.4, "thick": 0.4 } },

    { "type": "rot", "mat": "beam", "at": [0, 5.4], "half": [3.6, 0.85], "top": 6.4,
      "thick": 0.4, "spin": 0.95 },

    { "type": "box", "mat": "ladder", "from": [4,4], "to": [4.4,5], "top": 6, "thick": 6 }
  ]
}
```

## The `level` block

| field | what it is |
|---|---|
| `spawn` | `[x, y, z]` where the player starts |
| `yaw` | which way they face, in degrees |
| `budget` | `[minimum, maximum]` polygons |
| `lod` | projected size below which a box drops to one face |
| `fog` | `[start, end, "#hex"]` |
| `light` | `[x, y, z]` of the light direction |
| `killz` | the height below which the player is lost |
| `gravity`, `jump`, `walk` | the level's physics |

`killz` is reported separately from its value, through `HasKillZ`, because zero is a perfectly reasonable death height and cannot be allowed to mean "not stated".

## The `materials` block

Names are the keys of the block, so a material is declared once and referred to by name everywhere after, exactly as the registry works underneath.

Accepted fields: `color`, `edge`, `collision` (`solid`, `ghost`, `oneway`, `trigger`), `alpha`, `friction`, `bounce`, `speed`, `damage`, `buoyancy`, `drag`, `step`, `unlit`, `twosided`, `climbable`, `fog`, `visible`, `invisible`.

## The `objects` block

| type | fields |
|---|---|
| `box` | `mat`, `from` [x,y], `to` [x,y], `top`, `thick` |
| `ramp` | `mat`, `from`, `to`, `low`, `high`, `thick`, `axis` `"x"` or `"y"` |
| `rot` | `mat`, `at` [x,y], `half` [hw,hh], `top`, `thick`, `angle` |
| `bill` | `mat`, `at` [x,y,z], `size` [w,h] |
| `spin` | `mat`, `at` [x,y,z], `radius` |
| any | `repeat` n, `step` {the same fields, as deltas} |
| any | `move` {axis [x,y,z], amp, speed, phase, stagger}, `spin` rad/s, `tag` n |

## `repeat` and `step`

This is how a staircase is written without betraying the format: one entry describes the first tread and what changes from one to the next.

```json
{ "type": "box", "mat": "stone", "from": [-1,-29.4], "to": [1,-27.4], "top": 4, "thick": 0.4,
  "repeat": 4, "step": { "from": [0,3.1], "to": [0,3.1], "top": 0.35 } }
```

A loop with an expression inside would turn the document back into code, and JSON has no loops, it has data.

## `tag`

A number of yours, stored on the object and returned by `PScene.TagOf`. The bundled `obby.json` uses `"tag": 1` for a checkpoint because `PDemo` decided that, in a `Const TAG_CHECKPOINT As Long = 1`. The library does not decide.

## What is tolerated and what is not

Every field is optional and every default is stated, so a half written level loads and shows what you did write instead of refusing the whole file.

What is not tolerated is anything that would silence a mistake: an unknown material, an unknown type, a document with no objects at all, and a `repeat` of a size that can only be a typo. Those stop the load and say why, in `PLevel.Error`.

## Saving

```vba
PLevel.SaveFile "C:\path\level.json", sc      ' or: txt = PLevel.Save(sc)
```

Loading a level, editing it in memory and saving gives back the same level plus the edit. What does not survive is the shorthand: `repeat` and `step` are a way of writing a staircase, not something the world remembers. The fifty steps they stood for all come back, one entry each. The document is longer than the one that went in and describes exactly the same world.

Anyone building a level without having loaded one first has `SetSpawn`, `SetBudget`, `SetRules` and `SetKillZ` to fill in the `level` block.

Numbers are written with `Str`, never with `Format`. `Format` follows the machine's regional settings, and a level saved on a Brazilian PowerPoint would come back full of commas that no JSON reader on earth will accept. Material names are escaped, because a material called `a"b` would produce a file that will not open again.

# Fases em JSON

Fase deixa de ser código. Um `.json` é JSON puro: qualquer editor valida, qualquer ferramenta gera, e ninguém precisa abrir o VBE para mudar um mapa.

```vba
If PLevel.ParseFile("C:\caminho\fase.json", sc) Then
    ' PLevel.SpawnX, SpawnY, SpawnZ, SpawnYaw
    ' PLevel.Gravity, Jump, Walk, KillZ, HasKillZ
    ' PLevel.BudgetMin, BudgetMax
Else
    MsgBox PLevel.Error
End If
```

O parsing é do `JSON.cls`. O `PLevel` só conhece o esquema.

## Um documento

```json
{
  "level": { "spawn": [0, -12, 0.3], "yaw": 90, "budget": [120, 150],
             "fog": [9, 32, "#A0B9D2"], "light": [0.38, 0.42, 0.82], "lod": 44,
             "killz": -6, "gravity": 18, "jump": 6.4, "walk": 5.4 },

  "materials": {
    "gelo":   { "color": "#AADCF5", "friction": 0.12, "speed": 1.15 },
    "agua":   { "color": "#286EBE", "collision": "trigger", "alpha": 0.4,
                "buoyancy": 1.06, "drag": 3.4 },
    "trave":  { "color": "#E0DCC8", "step": 0.6 },
    "escada": { "color": "#8C6E3C", "collision": "ghost", "climbable": true }
  },

  "objects": [
    { "type": "box", "mat": "pedra", "from": [-4,-12], "to": [-2.8,-8], "top": 0.4, "thick": 0.4,
      "repeat": 6, "step": { "from": [1.2,0], "to": [1.2,0], "top": 0.4, "thick": 0.4 } },

    { "type": "rot", "mat": "trave", "at": [0, 5.4], "half": [3.6, 0.85], "top": 6.4,
      "thick": 0.4, "spin": 0.95 },

    { "type": "box", "mat": "escada", "from": [4,4], "to": [4.4,5], "top": 6, "thick": 6 }
  ]
}
```

## O bloco `level`

| campo | o que é |
|---|---|
| `spawn` | `[x, y, z]` onde o jogador começa |
| `yaw` | para onde ele olha, em graus |
| `budget` | `[mínimo, máximo]` de polígonos |
| `lod` | tamanho projetado abaixo do qual a caixa vira uma face |
| `fog` | `[início, fim, "#hex"]` |
| `light` | `[x, y, z]` da direção da luz |
| `killz` | altura abaixo da qual o jogador é perdido |
| `gravity`, `jump`, `walk` | a física da fase |

`killz` é reportado à parte do valor, com `HasKillZ`, porque zero é uma altura de morte perfeitamente razoável e não pode significar "não disse".

## O bloco `materials`

Os nomes são as chaves do bloco, então um material é declarado uma vez e referido por nome em todo o resto, exatamente como o registro funciona por baixo.

Aceita: `color`, `edge`, `collision` (`solid`, `ghost`, `oneway`, `trigger`), `alpha`, `friction`, `bounce`, `speed`, `damage`, `buoyancy`, `drag`, `step`, `unlit`, `twosided`, `climbable`, `fog`, `visible`, `invisible`.

## O bloco `objects`

| tipo | campos |
|---|---|
| `box` | `mat`, `from` [x,y], `to` [x,y], `top`, `thick` |
| `ramp` | `mat`, `from`, `to`, `low`, `high`, `thick`, `axis` `"x"` ou `"y"` |
| `rot` | `mat`, `at` [x,y], `half` [hw,hh], `top`, `thick`, `angle` |
| `bill` | `mat`, `at` [x,y,z], `size` [w,h] |
| `spin` | `mat`, `at` [x,y,z], `radius` |
| qualquer um | `repeat` n, `step` {mesmos campos, como delta} |
| qualquer um | `move` {axis [x,y,z], amp, speed, phase, stagger}, `spin` rad/s, `tag` n |

## `repeat` e `step`

É como se escreve uma escada sem trair o formato: uma entrada descreve o primeiro degrau e o que muda de um para o outro.

```json
{ "type": "box", "mat": "pedra", "from": [-1,-29.4], "to": [1,-27.4], "top": 4, "thick": 0.4,
  "repeat": 4, "step": { "from": [0,3.1], "to": [0,3.1], "top": 0.35 } }
```

Um laço com expressão dentro transformaria o documento em código de novo, e JSON não tem laço, tem dado.

## `tag`

Um número seu, gravado no objeto e devolvido por `PScene.TagOf`. O `obby.json` usa `"tag": 1` para checkpoint porque foi o `PDemo` que decidiu isso, num `Const TAG_CHECKPOINT As Long = 1`. A lib não decide.

## O que é tolerado e o que não é

Todo campo é opcional e tem default declarado, então uma fase pela metade carrega e mostra o que você já escreveu, em vez de recusar o arquivo inteiro.

O que não é tolerado é o que silenciaria um erro: material inexistente, tipo desconhecido, documento sem objeto nenhum, e `repeat` num tamanho que só pode ser typo. Esses param a carga e dizem o motivo em `PLevel.Error`.

## Salvando

```vba
PLevel.SaveFile "C:\caminho\fase.json", sc      ' ou: txt = PLevel.Save(sc)
```

Carregar, editar em memória e salvar devolve a mesma fase mais a edição. O que não volta é o atalho: `repeat` e `step` são um jeito de escrever uma escada, não uma coisa que o mundo lembra. Os cinquenta degraus que eles descreviam voltam todos, um por entrada. O arquivo sai maior que o que entrou e descreve exatamente o mesmo mundo.

Quem vai montar fase sem ter carregado uma antes tem `SetSpawn`, `SetBudget`, `SetRules` e `SetKillZ` para preencher o bloco `level`.

Os números saem por `Str`, nunca por `Format`. `Format` segue a configuração da máquina, e uma fase salva num PowerPoint brasileiro voltaria cheia de vírgulas que nenhum leitor de JSON aceita. Nomes de material saem escapados, porque um material chamado `a"b` produziria um arquivo que não reabre.

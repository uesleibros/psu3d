# Changelog

## 1.1.0

### Flicker fixes

Two things could make an object blink or vanish between two frames that were otherwise identical. Both were found by measuring temporal stability along a smooth camera path, not by looking at a single frame.

**Level of detail had a single threshold.** An object drifting across it flipped between three faces and one on alternate frames, which reads as a shape blinking. It now has two thresholds: it has to fall well under the line to lose its faces and climb well over it to get them back. Measured over 260 frames: four pops became none.

**The budget cut had no hysteresis.** An object sitting exactly on the polygon budget was cut on one frame and drawn on the next, disappearing and returning while nothing about it changed. An object already on screen may now overspend by a small, bounded margin rather than vanish. Measured with a deliberately tight budget over 400 frames: fourteen disappearances became none.

The residual five order flips per 260 frames come from genuine ordering cycles. Five alternative tiebreak rules were measured against them, including breaking by strongly connected component and by least visible damage, and none improved on the current one.

`PScene.WasReduced(slot)` was added alongside `DrawnAt`, so the level of detail decision can be observed. It is what the new `CheckStability` block tests.

### The demo declares its own input

`PDemo` now declares `GetCursorPos`, `SetCursorPos`, `ShowCursor` and `GetSystemMetrics` at the top of the file, next to the keyboard declaration it already had. Mouse look reads the pointer, measures its travel from a fixed anchor, and warps it back, which is relative mouse without capturing the device. Because the same anchor is used to read and to warp, its exact position never matters, so no conversion between screen pixels and slide points is needed.

Psu3D now has exactly one external requirement, and only for one module: `PLevel` needs a JSON parser.

### Documentation

The whole documentation set was rewritten in English and reorganised into nineteen hand written pages plus twelve generated reference pages, one per module, produced from the docstrings in the source so no signature in the documentation can drift from the signature in the code.

## 1.0.1

### O núcleo deixou de depender do UCursor

`PCanvas` tinha quatro membros que chamavam um módulo de cursor externo: `CursorX`, `CursorY`, `CursorInside` e `CenterCursor`. Como o VBA compila o projeto inteiro, isso obrigava **todo** projeto que importasse o núcleo a importar um módulo de cursor junto, querendo mouse ou não. Era uma dependência de verdade, escondida atrás de quatro conveniências.

Os quatro foram removidos. A canvas mapeia coordenadas e não lê hardware: quem tem um ponteiro subtrai `CenterX` e `CenterY`, que é a mesma conta e serve para qualquer fonte de posição.

Agora só o `PDemo` menciona um módulo de cursor, porque é um demo em primeira pessoa. Os outros doze módulos compilam sozinhos.

**Migração.** Se você usava algum dos quatro:

| antes | agora |
|---|---|
| `cv.CursorX` | `MeuCursorX - cv.X` |
| `cv.CursorY` | `MeuCursorY - cv.Y` |
| `cv.CursorInside` | `cv.Contains(MeuCursorX, MeuCursorY)` |
| `cv.CenterCursor` | `MoveCursorPara cv.CenterX, cv.CenterY` |

E para mouse look, direto:

```vba
dx = MeuCursorX - cv.CenterX
dy = MeuCursorY - cv.CenterY
```

## 1.0.0

Primeira versão pública.

### Núcleo

Canvas com posição e tamanho no slide, câmera com yaw e pitch reais, renderer com recorte de frustum em cinco planos e backface culling, cena em arrays paralelos, materiais com sombreamento pré-calculado, luz direcional e névoa em faixas.

### Física

`PBody` como classe da engine: andar, pular, subir degrau por material, escalar, nadar com empuxo e arrasto, quicar, e ser carregado por plataforma que anda **e** por plataforma que gira. O corpo reporta os gatilhos que encostou e não decide o que eles significam.

### Fases

`PLevel` lê e escreve cena em JSON. O round trip devolve a mesma fase mais a edição. Números saem por `Str` para não depender da configuração regional da máquina.

### Ordenação por profundidade

Algoritmo do pintor com plano separador e sort topológico. Filtro de espaço de tela que descarta 60% dos pares antes do comparador, AABB de caixa girada que segue o ângulo, e desempate de ciclo por menor dívida.

Medido contra um oráculo exato: 0,00% de artefatos visíveis no `obby.json` contra 1,85% antes, e 1,01% num cenário adversário contra 6,15% antes.

### Índice espacial

Grade uniforme em XY para as consultas de física, com o que se move e o que é grande demais ficando de fora e sendo varrido. De 19 a 79 vezes menos objetos tocados por consulta em cenas grandes. Provado equivalente à varredura completa em 120 000 consultas.

### Robustez

Quatro caminhos que terminavam em erro de execução foram fechados: estouro de `Long` na grade com coordenada absurda, `repeat` gigante no JSON, nome de material com aspas quebrando o arquivo salvo, e `budget` e `tag` fora do alcance de um `Long`.

### Autoteste

174 asserções cobrindo matemática, cor, luz, materiais, canvas, câmera, cena, índice, movimento, corpo, ordem de desenho, renderer, fases, salvamento e robustez.

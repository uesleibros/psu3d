# Changelog

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

# Psu3D

Motor 3D escrito em VBA puro, rodando dentro do PowerPoint. Transformação de view com yaw e pitch reais, projeção perspectiva, recorte de frustum em cinco planos, backface culling, iluminação direcional, névoa em faixas e algoritmo do pintor com plano separador.

Psu3D não é uma engine de jogo. É um motor 3D que serve para fazer jogo, do mesmo jeito que o three.js serve. O núcleo não conhece a palavra "jogador", "vida" nem "fase". Quem quer terreno, gráfico de barras ou uma figura 3D num slide usa o mesmo núcleo e nunca importa a física.

## Instalando

Abra o VBE com `Alt+F11`, `Ctrl+M`, e importe os arquivos de `engine/`. Não há DLL, referência a marcar nem instalador.

Depois rode `PSelfTest.Psu3DSelfTest`. São 174 asserções que rodam em memória, sem slide, e dizem se algum módulo ficou para trás.

## O menor exemplo possível

Uma figura estática num slide. Sem loop, sem física, sem teclado. As shapes ficam lá quando o Sub termina.

```vba
Public Sub Terreno()
    Dim x As Long, y As Long, h As Single, terra As Long

    Psu3D.Boot Shapes, 40, 40, 640, 400
    terra = PMaterials.Create("terra", PCore.ColorPack(96, 140, 90)).Id

    Psu3D.Camera.SetPosition -16, -16, 13
    Psu3D.Camera.LookAt 0, 0, 0
    Psu3D.Renderer.PolyBudget = 3000

    For y = 0 To 23
        For x = 0 To 23
            h = 2 + Sin(x * 0.4) * Cos(y * 0.35) * 1.6
            Psu3D.Scene.AddBox terra, x - 12, y - 12, x - 11, y - 11, h, h
        Next x
    Next y

    Psu3D.BeginFrame 0
    Psu3D.RenderScene
    Psu3D.EndFrame
End Sub
```

576 caixas num retângulo de 640 por 400 pontos do slide. Troque o `h` por um valor da planilha e vira gráfico de barras. Troque por ruído e vira terreno. A engine não sabe a diferença.

## Os módulos

Os oito primeiros são o núcleo. `Psu3D` é conveniência. Os quatro últimos você importa só se quiser.

| módulo | o que faz |
|---|---|
| `PCore.bas` | tipos, enums, matemática, random determinístico, cor, relógio |
| `PLighting.bas` | luz direcional e névoa global |
| `PMaterial.cls` | definição de uma superfície |
| `PMaterials.bas` | registro de materiais e tabela de sombreamento pré-calculada |
| `PCanvas.cls` | onde e como projetar: x, y, largura, altura, FOV |
| `PCamera.cls` | olho, yaw, pitch, teste de visibilidade |
| `PRenderer.cls` | pipeline de face para polyline, e as primitivas |
| `PScene.cls` | store de objetos em arrays paralelos, índice espacial, ordem de desenho |
| `PBody.cls` | corpo que anda, colide, sobe degrau, escala, nada e é carregado |
| `PLevel.bas` | ler e escrever cena em JSON |
| `Psu3D.bas` | fachada |
| `PSelfTest.bas` | autoteste da biblioteca |
| `PDemo.bas` | fase jogável de exemplo |

## Módulos de fora

Psu3D não embute código de terceiros.

`PLevel` precisa de um parser de JSON. Use o [vbacollective/json](https://github.com/vbacollective/json). Se você não vai ler fase de arquivo, não importe o `PLevel` e o assunto some.

Mouse é opcional e **nenhum módulo do núcleo o menciona**. Só o `PDemo` chama um módulo `UCursor`, porque é um demo em primeira pessoa. A canvas entrega `CenterX` e `CenterY` e quem tem um ponteiro faz a subtração, seja qual for a fonte. Os detalhes estão em [Instalação](docs/Instalacao.md).

## Documentação

Tudo está em [`docs/`](docs/), e o mesmo conteúdo vai para a [wiki](https://github.com/uesleibros/psu3d/wiki): guias por assunto, como o 3D funciona por dentro, e a referência completa dos 342 membros públicos, gerada a partir dos docstrings do próprio código.

| página | assunto |
|---|---|
| [Instalação](docs/Instalacao.md) | importar, e os dois módulos de fora |
| [Primeiros passos](docs/Primeiros-passos.md) | ver algo na tela |
| [Conceitos](docs/Conceitos.md) | como as peças se encaixam |
| [Como o 3D funciona](docs/Como-o-3D-funciona.md) | é 3D de verdade, e onde está a fronteira |
| [Corpo e física](docs/Corpo-e-fisica.md) | andar, pular, escalar, nadar, plataforma |
| [Fases em JSON](docs/Fases-em-JSON.md) | o formato completo |
| [Receitas](docs/Receitas.md) | exemplos prontos para copiar |
| [Ordenação por profundidade](docs/Ordenacao-por-profundidade.md) | o algoritmo do pintor aqui dentro |
| [Índice espacial](docs/Indice-espacial.md) | a grade que deixa a física barata |
| [Performance](docs/Performance.md) | a bomba de refresh, duplo buffer, orçamento |
| [Limites conhecidos](docs/Limites-conhecidos.md) | o que a lib não faz, e por quê |

## Exemplo pronto

`examples/obby.json` é um obby linear a céu aberto: não existe chão, cair é morrer e voltar ao último checkpoint. 38 objetos, 5 checkpoints, plataformas móveis, discos girando, gelo, trampolim, plataformas de mão única e elevador.

```vba
PDemo.RunFile "C:\caminho\obby.json"
```

## As regras que a lib segue

**Zero `On Error`.** Todo caminho que poderia estourar é validado à mão.

**`Option Private Module` em todo `.bas`**, e `VB_Exposed = False` em toda classe, para a lib não vazar para o resto do arquivo.

**Array puro, sem COM no caminho quente.** As únicas leituras de shape em toda a engine rodam no boot.

**Sem limite fixo.** Tudo cresce dobrando de tamanho.

## Licença

MIT. Veja [LICENSE](LICENSE).

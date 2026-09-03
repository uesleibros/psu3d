# Psu3D

Motor 3D escrito em VBA puro, rodando dentro do PowerPoint.

Psu3D não é uma engine de jogo. É um motor 3D que serve para fazer jogo, do mesmo jeito que o three.js serve. O núcleo (canvas, câmera, renderer, cena, material, luz) não conhece a palavra "jogador", "vida" nem "fase". Quem quer terreno, gráfico de barras, visualização de dados ou uma figura 3D num slide usa o mesmo núcleo e nunca importa a física.

## Por onde começar

| se você quer | vá para |
|---|---|
| pôr a lib no seu arquivo | [Instalação](Instalacao.md) |
| ver alguma coisa na tela em 20 linhas | [Primeiros passos](Primeiros-passos.md) |
| entender como as peças se encaixam | [Conceitos](Conceitos.md) |
| saber se isso é 3D de verdade | [Como o 3D funciona](Como-o-3D-funciona.md) |
| fazer um jogo | [Corpo e física](Corpo-e-fisica.md) |
| escrever fase em arquivo | [Fases em JSON](Fases-em-JSON.md) |
| a lista completa de tudo | [Referência da API](API-PCore.md) |

## O mapa dos módulos

Os oito primeiros são o núcleo. `Psu3D` é conveniência. `PBody`, `PLevel`, `PSelfTest` e `PDemo` você importa só se quiser.

| módulo | o que faz | página |
|---|---|---|
| `PCore.bas` | tipos, enums, matemática, random determinístico, cor, relógio | [API](API-PCore.md) |
| `PLighting.bas` | luz direcional e névoa global | [Luz e névoa](Luz-e-nevoa.md) |
| `PMaterial.cls` | definição de uma superfície | [Materiais](Materiais.md) |
| `PMaterials.bas` | registro de materiais e tabela de sombreamento | [Materiais](Materiais.md) |
| `PCanvas.cls` | onde e como projetar | [Canvas](Canvas.md) |
| `PCamera.cls` | olho, yaw, pitch | [Câmera](Camera.md) |
| `PRenderer.cls` | pipeline de face e as primitivas | [Renderer e primitivas](Renderer-e-primitivas.md) |
| `PScene.cls` | store de objetos, índice espacial, ordem de desenho | [Cena](Cena.md) |
| `PBody.cls` | corpo que anda, colide, sobe degrau, escala, nada | [Corpo e física](Corpo-e-fisica.md) |
| `PLevel.bas` | ler e escrever cena em JSON | [Fases em JSON](Fases-em-JSON.md) |
| `Psu3D.bas` | fachada | [API](API-Psu3D.md) |
| `PSelfTest.bas` | autoteste da biblioteca | [Autoteste](Autoteste.md) |
| `PDemo.bas` | fase jogável de exemplo | [Receitas](Receitas.md) |

Nada de terceiros vem embutido. `PLevel` precisa de um parser de JSON, que está em [vbacollective/json](https://github.com/vbacollective/json). Mouse é opcional e sai por `UCursor`, que também não faz parte da lib. Os detalhes estão em [Instalação](Instalacao.md).

## As regras que a lib segue

Estas não são preferências, são restrições que valem em todo arquivo do repositório.

**Zero `On Error`.** Todo caminho que poderia estourar (nome de shape inexistente, hex inválido, id fora do range, chave duplicada) é validado à mão. `On Error` no VBA esconde o erro em vez de tratar, e depois some com a pilha.

**`Option Private Module` em todo `.bas`.** As classes não aceitam essa instrução; nelas o equivalente é o atributo `VB_Exposed = False`, que já está no cabeçalho de cada arquivo. Duas exceções deliberadas: `PSelfTest` e `PDemo`, porque o diálogo de macros só lista ponto de entrada público.

**Array puro, sem COM no caminho quente.** COM só aparece onde é fisicamente inevitável: `AddPolyline`, a cor do preenchimento por polígono, e o `Range(...).Delete` que limpa o frame inteiro numa chamada só.

**Sem limite fixo.** Registro de materiais, pool de shapes e store de objetos crescem dobrando de tamanho. Nada de `MAX_WORLD_PLATS`.

**UDT em classe é `Friend`, nunca `Public`.** O VBA recusa tipo definido pelo usuário de módulo padrão na assinatura pública de um módulo de classe. Enum pode ser público normalmente; a regra vale só para `Type`.

## Estado

174 asserções de autoteste, 13 módulos, cerca de 11 mil linhas de VBA. Rode `PSelfTest.Psu3DSelfTest` depois de importar.

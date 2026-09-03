# Materiais

Um material descreve uma superfície: como ela é desenhada e como ela se comporta. Dois objetos com o mesmo material se comportam igual, e trocar o material de um objeto muda as duas coisas na mesma linha.

```vba
Dim m As PMaterial
Set m = PMaterials.Create("gelo", PCore.ColorPack(170, 220, 245), pcSolid)

m.Friction = 0.08
m.SpeedMultiplier = 1.25
m.EdgeVisible = True
m.EdgeColor = PCore.ColorPack(220, 245, 255)
```

O registro guarda o material e devolve um id, que é o número que a cena usa:

```vba
sc.AddBox PMaterials.IdOf("gelo"), -4, -4, 4, 4, 0, 0.3
```

## Colisão

| valor | efeito |
|---|---|
| `pcSolid` | barra por todos os lados |
| `pcGhost` | é desenhado e você atravessa |
| `pcOneWay` | só segura quem vem de cima; por baixo você passa |
| `pcTrigger` | nunca barra, só avisa que houve contato |

## Aparência

| propriedade | o que faz |
|---|---|
| `Color` | a cor base, antes de luz e névoa |
| `Unlit` | ignora a luz, fica sempre na cor cheia |
| `Fogged` | se a névoa age sobre ele |
| `Visible` | falso faz um colisor invisível |
| `TwoSided` | desenha mesmo virado de costas |
| `Transparency` | de 0 a 1 |
| `EdgeVisible`, `EdgeColor`, `EdgeWeight` | contorno |

Um material com `Visible = False` não é apenas transparente: o `PScene` pula objeto invisível já na coleta, então ele não é cullado, nem ordenado, nem desenhado, e não gasta orçamento. É assim que se faz guarda-corpo que segura e não aparece.

## Física

| propriedade | o que faz |
|---|---|
| `Friction` | quanto o corpo freia parado em cima; gelo é 0,08 |
| `Bounce` | queda forte é devolvida, queda fraca é absorvida |
| `SpeedMultiplier` | teto de velocidade andando naquilo |
| `DamagePerSecond` | o jogo lê e decide o que fazer |
| `StepHeight` | o degrau máximo daquela superfície; zero significa que o corpo decide |
| `Climbable` | escada: encostado nela, sobe e desce em vez de cair |
| `Buoyancy`, `Drag` | num gatilho, fazem dele fluido |

`StepHeight` e `Climbable` são o que separa um meio fio de um paredão com a mesma forma. E `Climbable` e `Collision` são perguntas separadas: escada `ghost` você entra e sobe, escada `solid` você escala por fora, e as duas se escrevem igual.

## Fluido

Um gatilho com `Buoyancy` maior que zero é fluido. O corpo mede quanto de si está abaixo da superfície e aplica na proporção:

```vba
Set agua = PMaterials.Create("agua", PCore.ColorPack(40, 110, 190), pcTrigger)
agua.Buoyancy = 1.06
agua.Drag = 3.4
agua.SpeedMultiplier = 0.45
agua.Transparency = 0.4
agua.TwoSided = True
```

Empuxo acima de 1 significa que soltar o controle faz você boiar de volta. Arrasto sangra as três componentes da velocidade, e é o que faz uma queda na água aterrissar macio em vez de bater no fundo.

## A paleta padrão

`PMaterials.CreateDefaults` cria quinze materiais prontos: `default`, `stone`, `grass`, `brick`, `metal`, `ice`, `mud`, `rubber`, `glass`, `lava`, `water`, `platform`, `ladder`, `decor`, `pickup`, `clip`. `Psu3D.Boot` já chama isso.

Eles servem para começar rápido e para o autoteste ter contra o que medir. Nada obriga a usá-los.

## Sombreamento pré-calculado

A cor de uma face é material vezes direção da face vezes faixa de névoa. Isso é uma tabela, não uma conta, e o `PMaterials` a mantém pronta:

```vba
cor = PMaterials.ShadeBand(matId, dirKey, faixa)
```

Duas checagens de limite e uma leitura de array. Quando a luz, a névoa ou um material mudam, um contador de revisão invalida a tabela e ela é refeita na próxima vez que alguém pedir.

Só faces alinhadas aos eixos usam a tabela. Uma face inclinada, como a de uma rampa, cai no caminho lento que calcula a partir da normal real.

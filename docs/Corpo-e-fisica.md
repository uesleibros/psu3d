# Corpo e física

`PBody` é um cilindro em pé que anda pela cena e é parado por ela. Ele é opcional: quem só desenha nunca importa esse módulo.

```vba
Dim body As PBody
Set body = Psu3D.CreateBody()      ' já ligado à cena bootada

body.SetSize 0.34, 1.72            ' raio, altura
body.SetPosition 0, 0, 4
body.WalkSpeed = 5.4
body.JumpSpeed = 6.4
body.StepHeight = 0.45             ' fallback: o material pode dizer outra coisa
body.KillZ = -12
```

## O passo

```vba
body.Advance dt, wishX, wishY, pulando, agachando
```

Uma chamada por frame. `wishX` e `wishY` são a direção pedida em coordenadas do mundo, e não precisam estar normalizados. Os dois booleanos são as teclas de subir e descer, que significam pular, nadar ou escalar dependendo de onde o corpo está.

A ordem dentro de `Advance` é a ordem em que as coisas acontecem no mundo:

```
carregado pelo que está embaixo
acelerado pelo que você pediu
empurrado para fora do que é sólido
e só então a gravidade opina
```

Fazer a gravidade antes do empurrão deixaria o corpo afundar um frame no chão antes de ser levantado, e isso se lê como tremida.

## Lendo o resultado

```vba
body.X, body.Y, body.Z
body.VelX, body.VelY, body.VelZ
body.OnGround
body.GroundMaterial
body.GroundObject
body.Submersion         ' 0 a 1
body.IsSwimming
body.IsClimbing
body.LandingSpeed       ' o quanto bateu no último pouso, para tremer a câmera
body.CarryYaw           ' quanto a plataforma girou você
body.FellOut            ' passou do KillZ
```

## Gatilhos

O corpo não sabe o que é uma moeda. Ele reporta o que encostou e recusa adivinhar para que servia:

```vba
For k = 0 To body.TouchCount - 1
    idx = body.TouchAt(k)

    If sc.TagOf(idx) = TAG_CHECKPOINT Then
        ' regra sua
    Else
        sc.SetActive idx, False
        pontos = pontos + 1
    End If
Next k
```

É por isso que a mesma classe serve para uma fase de coletar e para uma de desviar de mina, sem saber nenhuma das duas palavras.

`Reach` é a distância além da própria borda em que o corpo percebe um gatilho. Moeda do tamanho de uma mão, coletada só quando a borda do corpo cruza ela, é moeda que o jogador atravessa sem pegar.

## O material manda

O corpo não tem regra própria sobre superfície. Ele pergunta ao material:

| propriedade | efeito |
|---|---|
| `Collision` | `solid` para, `ghost` deixa passar, `oneway` só segura por cima, `trigger` só avisa |
| `Friction` | quanto o corpo freia parado em cima |
| `SpeedMultiplier` | teto de velocidade andando naquilo |
| `Bounce` | queda forte é devolvida, queda fraca é absorvida |
| `DamagePerSecond` | o jogo lê e decide |
| `StepHeight` | o degrau máximo daquela superfície; zero significa que o corpo decide |
| `Climbable` | escada |
| `Buoyancy`, `Drag` | num gatilho, fazem dele fluido |

## Carona

O corpo guarda em qual objeto está apoiado, e no começo do passo seguinte pega o deslocamento daquele objeto e vai junto. Por isso ele guarda o id do chão, não só o material.

No giro isso quer dizer girar em volta do centro da plataforma, não ficar parado enquanto o chão roda embaixo. E `CarryYaw` devolve quanto você girou, para somar no yaw da câmera:

```vba
If body.CarryYaw <> 0 Then cam.AddAngles body.CarryYaw, 0
```

Sem isso o disco gira você em volta do centro dele enquanto você continua olhando para o mesmo lado, que é a única coisa que estar em cima de um carrossel nunca faz.

Para plataforma movida na mão, fora do sistema de movimento, existe `Nudge`, que empurra sem mexer na velocidade, porque ser carregado não é ser acelerado:

```vba
sc.SetTopZ elevador, novoZ
If body.GroundObject = elevador Then body.Nudge 0, 0, novoZ - zAntigo
```

**A cena tem que se mover antes do corpo.** Uma plataforma que já se moveu é uma que consegue carregar; uma que se move depois é uma que o corpo passa um frame ao lado.

## Escada

Escada é um material com `Climbable`. Encostado nela, segurar subir ou descer move o corpo verticalmente e desliga a gravidade.

`Climbable` e `Collision` são independentes de propósito. Escada `ghost` você entra e escala por dentro, escada `solid` você escala por fora agarrado, e as duas se escrevem igual.

## Água

Um gatilho com `Buoyancy` maior que zero é fluido. O empuxo cancela parte da gravidade na proporção de quanto do corpo está submerso, e o arrasto sangra as três componentes da velocidade.

Passado `SwimSpeed` de submersão o corpo passa a nadar: subir vira braçada em vez de pulo, e descer vira mergulho.

## Ajustando

Todos têm padrão sensato e todos são propriedades:

`Gravity`, `WalkSpeed`, `JumpSpeed`, `GroundAccel`, `GroundFriction`, `AirAccel`, `SwimSpeed`, `DiveSpeed`, `ClimbSpeed`, `StepHeight`, `KillZ`, `Reach`.

## Outro solver

`PBody` é escrito só em cima da API pública da cena: `QueryBox`, `TopZAt`, `Blocks`, `GetOrientedBox`, `GetSpanZ`, `ContainsXY`, `GetMotionDelta`, `SpinDelta`. Quem quiser um solver diferente escreve o dele do mesmo jeito, e não perde nada.

# Performance

## A bomba de refresh

Isto não é otimização, é a condição para qualquer coisa aparecer. Leia antes de mexer no loop.

```vba
timerShp.TextEffect.Text = m_statText & "  " & Format$(Timer, "0.00")
DoEvents
```

Escrever num WordArt é o que força o PowerPoint a repintar o slide durante uma apresentação, e o repintar é como o frame se torna visível. Pule isso, ou escreva a mesma string duas vezes seguidas, e a imagem para de atualizar mesmo com a engine rodando normalmente.

Três consequências:

1. **Tem que disparar todo frame.** Não dá para limitar a cada 100 ms.
2. **O valor tem que mudar de verdade.** Por isso o relógio cru vai concatenado no fim do texto: ele garante que a string seja diferente.
3. **`DoEvents` é onde o repintar acontece.** Sem ele o pedido fica na fila.

A bomba e o `DoEvents` são medidos juntos como um balde só, porque são um evento só: a escrita pede o repintar e o `DoEvents` é onde ele ocorre.

## Duplo buffer

O PowerPoint não tem `ScreenUpdating`, então não dá para congelar a tela enquanto se desenha. O que dá para fazer é nunca deixar a tela vazia.

O renderer mantém dois bancos de nomes de shape. As shapes do frame novo entram com nomes de um banco enquanto as do frame anterior ainda estão na tela com nomes do outro. No `EndFrame` o banco velho é apagado numa chamada só:

```vba
target.Range(nomesDoBancoVelho).Delete
```

Uma chamada de COM para apagar cento e quarenta shapes, em vez de cento e quarenta.

## Orçamento de polígonos

```vba
rd.SetBudgetRange 120, 150
rd.PolyBudget = 140
rd.AutoBudget = False
```

`AdaptBudget dt` sobe o teto devagar, mais 2, quando o frame vem rápido, e derruba rápido, menos 8, quando atrasa. A histerese é assimétrica de propósito: são precisos três frames lentos seguidos para encolher e dez rápidos para crescer, senão o orçamento oscila e a geometria do fundo pisca.

**Para uma cena de tamanho conhecido, fixe.** Se o pior caso, com tudo na tela ao mesmo tempo, custa 100 polígonos, fixar em 140 significa que a lista de desenho nunca é cortada, e um corte que nunca acontece é um corte que nunca pode piscar.

`AutoBudget` serve para cena grande demais para qualquer número fixo, onde trocar geometria distante por frame rate é o negócio certo.

## Nada é lido de volta das shapes

Regra da lib: o caminho por frame nunca consulta uma propriedade de shape. Objeto COM é caro de interrogar, e tudo que precisaríamos perguntar já é nosso: posição, tamanho, cor, ângulo, tudo vive em array.

Auditado: as únicas leituras de shape em toda a engine são `PRenderer.Purge` e `PCanvas.FindShape`, e as duas só rodam no boot. Nem o `PScene`, nem o `PMaterials`, nem as primitivas do `PRenderer` tocam numa shape.

## Frame que não mudou

Um corpo que terminou o frame onde começou não muda nada na tela. O demo mantém um flag:

```vba
If m_dirty Then
    Psu3D.BeginFrame dt
    Psu3D.RenderScene
    Psu3D.EndFrame
    m_dirty = False
End If
```

Pular o frame pula toda chamada de COM dentro dele: nada de apagar, nada de `AddPolyline`, nada de preencher. Um jogador parado custa zero, que é a otimização mais barata que existe aqui.

Cuidado: a bomba de refresh **não** entra nesse `If`. Ela roda sempre.

## Sombreamento pré-calculado

A cor de uma face é material vezes direção vezes faixa de névoa, e isso é uma tabela, não uma conta. `PMaterials.ShadeBand` é duas checagens de limite e uma leitura de array.

Só faces alinhadas aos eixos usam a tabela. Rampa cai no caminho lento, que é o preço de ela poder apontar para qualquer lado.

## Medindo de verdade

```vba
rd.Profiling = True
rd.DryRun = True        ' roda tudo sem tocar em shape nenhuma
```

Com `DryRun` ligado você mede só o seu código. A diferença entre o tempo com e sem `DryRun` é o que o PowerPoint cobra, e essa é a conta que decide onde vale otimizar.

O demo mostra os três números na barra de status: `render`, `com` e `doevents`. Se `doevents` domina, encolher o pipeline não move nada e a alavanca é quanto o repintar tem que desenhar.

## Nível de detalhe

`sc.LodSize = 44` faz uma caixa cujo diâmetro projetado ficou abaixo de 44 pontos desenhar só a face mais visível. Uma caixa mostra no máximo três faces, e as duas menores viram tiras finas muito antes da dominante.

Sobe o número para comprar frame rate, desce para comprar silhueta.

## Índice espacial

Consultas de física passam por uma grade uniforme. Veja [Índice espacial](Indice-espacial.md).

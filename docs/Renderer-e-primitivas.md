# Renderer e primitivas

O renderer é onde uma face 3D vira uma shape do PowerPoint. É também onde moram as primitivas, porque aqui não existe objeto malha: uma caixa não é uma coisa guardada, é uma chamada que acontece, igual a um `fillRect` num canvas 2D.

```vba
Dim rd As PRenderer
Set rd = New PRenderer

rd.Attach Shapes, cv, cam
rd.PolyBudget = 140
rd.SeamFill = True
```

## Primitivas

Todas desenham direto, sem nada guardado e sem alocar:

```vba
rd.DrawBox         matId, x1, y1, x2, y2, topo, espessura
rd.DrawBoxLod      matId, x1, y1, x2, y2, topo, espessura
rd.DrawBoxRotated  matId, cx, cy, meiaW, meiaH, topo, espessura, angulo
rd.DrawRamp        matId, x1, y1, x2, y2, zBaixo, zAlto, espessura, paX
rd.DrawFloor       matId, x1, y1, x2, y2, z
rd.DrawWall        matId, x1, y1, x2, y2, z, altura
rd.DrawBillboard   matId, x, y, z, largura, altura
rd.DrawSpinner     matId, x, y, z, raio, fase
```

`DrawBox` submete só as faces que o olho pode ver, no máximo três das seis. `DrawBoxLod` submete só a face mais visível, e é o que a cena usa quando a caixa ficou pequena demais na tela para as outras duas valerem uma shape cada.

`DrawBillboard` é um quad que sempre vira para a câmera. `DrawSpinner` é uma placa de duas faces girando no próprio eixo, que é a forma que coletável costuma ter.

## Mais baixo ainda

Se as primitivas não servem, você entrega a geometria:

```vba
rd.DrawQuad     x0,y0,z0, x1,y1,z1, x2,y2,z2, x3,y3,z3, matId
rd.DrawTriangle x0,y0,z0, x1,y1,z1, x2,y2,z2, matId
rd.DrawPolygon2D ...     ' em espaço de canvas, para HUD
```

`DrawPolygon2D` pula toda a parte 3D e desenha direto em pontos do canvas. É o que o HUD do demo usa, e é de propósito: shape de texto obriga a mexer no z-order todo frame, e no PowerPoint isso repinta o slide inteiro.

## O frame

```vba
rd.BeginFrame dt      ' troca o banco de nomes, ajusta o orçamento
' ... desenhe ...
rd.EndFrame           ' apaga o banco anterior de uma vez só
```

Entre `BeginFrame` e `EndFrame` as shapes novas são criadas com nomes de um banco, e as do frame anterior continuam na tela com nomes do outro banco. No `EndFrame` o banco velho é apagado numa chamada só, com `Range(...).Delete`.

Isso é duplo buffer. O PowerPoint não tem `ScreenUpdating`, então não dá para congelar a tela enquanto se desenha; o que dá para fazer é nunca deixar a tela vazia.

## Orçamento

`PolyBudget` é o teto de polígonos que um frame pode gastar. `AdaptBudget dt` sobe o teto devagar quando o frame vem rápido e derruba rápido quando atrasa, entre os limites de `SetBudgetRange`. `AutoBudget` liga e desliga isso.

Para uma cena de tamanho conhecido, fixar o orçamento acima do pior caso é melhor: um corte que nunca acontece é um corte que nunca pode piscar.

## Costura

Dois polígonos vizinhos deixam um fio de fundo aparecendo entre eles, porque o antialiasing do PowerPoint não fecha a junta. `SeamFill = True` contorna cada polígono com a cor dele mesmo, o que fecha a junta sem custo extra de COM, já que a cor da borda entra na mesma chamada.

`EdgeInflate` faz a mesma coisa empurrando os vértices para fora. Funciona, mas distorce face pequena, então `SeamFill` é a escolha melhor.

## Medindo

```vba
rd.Profiling = True
Debug.Print rd.ComSeconds, rd.ComOps, rd.PolyCount
```

`DryRun = True` roda o pipeline inteiro sem tocar em shape nenhuma. É o que o autoteste usa, e é como você separa o custo do seu código do custo do PowerPoint.

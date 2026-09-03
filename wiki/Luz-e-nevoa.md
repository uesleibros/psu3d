# Luz e névoa

Uma luz direcional global e uma névoa global. Ambas valem para a cena inteira.

```vba
PLighting.SetLight 0.38, 0.42, 0.82
PLighting.SetLightIntensity 0.55, 0.45      ' ambiente, difusa
PLighting.SetFog 8, 34, PCore.ColorPack(160, 185, 210)
PLighting.SetFogSteps 24
PLighting.EnableFog True
```

A direção da luz é normalizada na entrada, então você pode passar qualquer vetor.

`SetLightIntensity` divide a iluminação em duas partes: a ambiente é o quanto uma face recebe estando de costas para a luz, e a difusa é o quanto ela ganha por estar virada para ela. As duas somando perto de 1 dá contraste natural; ambiente alta demais deixa a cena chapada.

## Névoa em faixas

`SetFog` recebe a distância em que a névoa começa, a distância em que ela é total, e a cor. Fora do intervalo não há mistura.

`SetFogSteps` é o detalhe que importa para performance. A névoa não é contínua: ela é quantizada num número de faixas, e a cor de cada faixa é pré-calculada por material e por direção de face. Com 24 faixas, a cor de qualquer face é uma leitura de array em vez de uma interpolação.

Poucas faixas fazem aparecer bandas visíveis no chão. Muitas faixas fazem a tabela crescer. O teto é `P_MAX_FOG_STEPS`, que é 64.

Combine a cor da névoa com a cor de fundo do canvas. Se elas forem diferentes, o horizonte vira uma linha.

```vba
cv.BackColor = PCore.ColorPack(160, 185, 210)
PLighting.SetFog 8, 34, PCore.ColorPack(160, 185, 210)
```

## Revisão

`PLighting` mantém um contador. Toda mudança o incrementa e avisa o `PMaterials`, que joga fora a tabela de sombreamento. Você nunca precisa invalidar nada à mão: mudar a luz no meio do jogo simplesmente funciona, e custa uma remontagem de tabela no frame seguinte.

## Direto

Se você quiser a cor de uma face por conta própria:

```vba
f = PLighting.ShadeFactor(nx, ny, nz)              ' 0 a 1
cor = PLighting.ShadeColor(corBase, nx, ny, nz)
cor = PLighting.ShadeColorByDir(corBase, pdPosZ)   ' face alinhada ao eixo
cor = PLighting.ApplyFog(cor, distancia)
```

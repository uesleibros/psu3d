# Canvas

O canvas é o retângulo do slide onde a cena é desenhada, mais a lente com que ela é vista.

```vba
Dim cv As PCanvas
Set cv = New PCanvas

cv.Init 40, 40, 640, 400          ' x, y, largura, altura em pontos do slide
cv.FieldOfViewDeg = 46
cv.NearPlane = 0.05
cv.FarPlane = 60
```

Quatro atalhos cobrem os casos comuns:

```vba
cv.FillSlide                      ' o slide inteiro
cv.CenterOnSlide 800, 450         ' esse tamanho, centralizado
cv.FitToSlide 1.7778              ' essa proporção, cabendo no slide
cv.SetFromShape Shapes("visor")   ' usa um placeholder que já existe no design
```

`SetFromShape` é o mais prático quando alguém desenhou a moldura no PowerPoint: você posiciona o retângulo com o mouse e a lib obedece.

## Por que o canvas tem posição

Num motor comum a tela é a tela. Aqui você quase nunca quer ocupar o slide inteiro: quer um visor ao lado de um texto, dentro de uma moldura, no canto.

E como o canvas tem posição, dois canvas cabem no mesmo slide. Duas câmeras, dois renderers, uma cena:

```vba
Set cvMini = New PCanvas
cvMini.Init 20, 20, 250, 150

Set rdMini = New PRenderer
rdMini.Prefix = "mini_"           ' prefixo próprio, senão um apaga o frame do outro
rdMini.Attach Shapes, cvMini, camMini
rdMini.PolyBudget = 48            ' segunda vista é segunda passada pela mesma geometria
```

## Lente

`FieldOfViewDeg` é o campo de visão horizontal. O vertical é derivado do formato do canvas, então mudar a altura do retângulo não distorce a imagem, só corta ou revela.

`FocalLength`, `TanFovH`, `TanFovV`, `NormH` e `NormV` são valores derivados que o renderer lê uma vez por frame. Você não precisa deles a não ser que esteja escrevendo seu próprio pipeline.

`NearPlane` é a distância mínima em que alguma coisa ainda é desenhada. Muito perto de zero e a projeção explode em números gigantes; o padrão é seguro.

## Fundo

O canvas pode manter um retângulo de fundo atrás da geometria:

```vba
cv.BackVisible = True
cv.BackColor = PCore.ColorPack(160, 185, 210)
cv.EnsureBackdrop Shapes
```

Esse retângulo é uma shape chamada `p3dbg_<nome>`. O prefixo é diferente do `p3d_` das shapes de geometria de propósito: a limpeza do renderer apaga tudo que começa com `p3d_`, e o fundo precisa sobreviver a ela.

## Coordenadas

`ToLocal` e `ToSlide` convertem entre o espaço do canvas e o espaço do slide. `Contains` responde se um ponto do slide caiu dentro do canvas.

**A canvas não lê hardware.** Ela mapeia coordenadas, e nada mais. Quem tem um ponteiro faz a conta:

```vba
dx = MeuCursorX - cv.CenterX
dy = MeuCursorY - cv.CenterY
cam.AddAngles dx * 0.0026, -dy * 0.0026
MoveMeuCursorPara cv.CenterX, cv.CenterY
```

`CenterX` e `CenterY` são o centro da canvas em pontos do slide, então a subtração dá a distância do ponteiro até o meio, que é exatamente o que mouse look quer. Recentralizar depois é o que faz mouse relativo sem captura.

Existiram aqui um `CursorX`, um `CursorY`, um `CursorInside` e um `CenterCursor`. Eram convenientes, e faziam todo projeto que importasse o núcleo ter que importar um módulo de cursor junto, querendo mouse ou não. Foram removidos, e o núcleo passou a não depender de nada.

# Cena

`PScene` guarda o mundo e responde duas perguntas por frame: o que a câmera vê, e o que está dentro deste pedaço de espaço.

## Adicionando

```vba
Dim caixa As Long, rampa As Long, disco As Long

caixa = sc.AddBox(pedra, -2, -2, 2, 2, 1, 0.5)
'                       x1  y1  x2 y2 topo espessura

rampa = sc.AddRamp(pedra, 0, 0, 6, 3, 0, 2, 0.4, paX)
'                        x1 y1 x2 y2 baixo alto espessura eixo

disco = sc.AddRotatedBox(metal, 0, 5, 3.6, 0.85, 1, 0.4, 0)
'                              cx cy meiaW meiaH topo espessura ângulo

placa = sc.AddBillboard(enfeite, 0, 8, 2, 3, 2)
'                               x  y  z  larg alt

moeda = sc.AddSpinner(ouro, 0, 3, 1.2, 0.32)
'                          x  y   z   raio
```

Todos devolvem um `Long`, que é o id do objeto. O id vale pela vida do objeto porque nada nunca reordena os arrays.

## Editando

```vba
sc.SetMaterial caixa, gelo
sc.MoveBy caixa, 0, 0, 3
sc.SetTopZ caixa, 4.2
sc.SetAngle disco, 1.2
sc.SetActive caixa, False        ' some da tela e da colisão, mas o slot fica seu
sc.Remove caixa                  ' some de vez, e o slot volta para a loja
```

A diferença entre `SetActive False` e `Remove` importa: depois do `Remove`, o id nomeia quem tomou o lugar. Use `SetActive` para esconder algo que você vai trazer de volta.

## Movimento

Duas formas, e a cena cuida das duas:

```vba
sc.SetMotion plataforma, 1, 0, 0, 5.5, 0.85, 0
'                       eixo x,y,z  amp  vel  fase

sc.SetSpin disco, 1.15
'                 radianos por segundo
```

`SetMotion` faz a plataforma oscilar em senoide, o que a faz desacelerar nas pontas, que é o que permite subir nela de propósito. Chame `sc.UpdateMotion dt` uma vez por frame; ela devolve quantos objetos se mexeram, então quem só redesenha quando algo muda sabe quando redesenhar.

Quem estava em cima vai junto. Veja [Corpo e física](Corpo-e-fisica#carona).

## Consultas

```vba
n = sc.QueryBox(x1, y1, x2, y2)
n = sc.QueryRadius(x, y, raio)

For k = 0 To n - 1
    idx = sc.ResultAt(k)
Next k
```

As consultas passam pelo [índice espacial](Indice-espacial), então continuam baratas com milhares de objetos.

Para perguntar sobre um objeto específico:

```vba
z = sc.TopZAt(idx, x, y)              ' altura da superfície ali, interpolando rampa
b = sc.Blocks(idx, True)              ' o material barra, vindo de cima?
ok = sc.ContainsXY(idx, x, y)         ' o ponto está dentro, respeitando o giro
sc.GetSpanZ idx, baixo, alto
sc.GetBoundsXY idx, x1, y1, x2, y2
sc.GetOrientedBox idx, cx, cy, hw, hh, ang
```

`GetOrientedBox` devolve a caixa no referencial dela mesma, que é o que um solver precisa para tratar uma laje girada como a laje que ela é, e não como o quadrado que a envolve.

## Tag

Um número seu por objeto:

```vba
sc.SetTag portao, 7
If sc.TagOf(idx) = 7 Then ...
```

A cena deliberadamente não sabe o que é checkpoint, porta ou waypoint. Ela carrega um número e deixa você decidir o que ele significa, que é a diferença entre uma lib em que dá para fazer um jogo e uma lib em que só dá para fazer este jogo.

## Desenhando

```vba
desenhados = sc.Render(rd)
```

`Render` faz tudo: cull de frustum, gasto do orçamento, ordem de pintura e desenho. Devolve quantos objetos foram pintados.

Depois dele, `DrawnCount` e `DrawnAt(i)` devolvem a ordem que o frame usou. É a única coisa de um frame que não dá para conferir olhando as shapes depois.

## Nível de detalhe

`LodSize` é o tamanho projetado, em pontos do slide, abaixo do qual uma caixa passa a desenhar só a face mais visível. Uma caixa mostra no máximo três faces, e as duas menores viram tiras finas muito antes da dominante. Passado esse ponto elas custam uma shape cada para contribuir alguns pontos de cor.

Zero desliga.

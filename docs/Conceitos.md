# Conceitos

Cinco objetos, cada um com uma pergunta só. Se você souber qual pergunta cada um responde, sabe onde mexer.

| objeto | pergunta que responde |
|---|---|
| `PCanvas` | onde na tela, e com que abertura de lente |
| `PCamera` | de onde estou olhando, e para onde |
| `PRenderer` | como uma face vira uma shape do PowerPoint |
| `PScene` | o que existe no mundo, e em que ordem desenhar |
| `PBody` | o que se move e é parado pelo que existe |

## O canvas não é a tela

Num motor 3D comum a tela é a tela. Aqui não: o canvas é um retângulo qualquer do slide, com x, y, largura e altura em pontos. Isso existe porque num slide você quase nunca quer ocupar tudo. Quer um visor no canto, ao lado de um texto, dentro de uma moldura que já está no design.

Consequência direta: dois canvas, duas câmeras e dois renderers podem dividir uma cena e um slide. Cada renderer nomeia suas shapes com prefixo próprio, então nenhum apaga o frame do outro. É assim que a arena do demo tem uma minivista no canto.

O FOV vertical é derivado do formato do canvas, então mudar a altura não distorce a imagem.

## A cena guarda dados, não objetos

`PScene` não guarda uma lista de objetos. Guarda arrays paralelos: um array de tipo, um de material, um de x1, um de y1, e assim por diante. Adicionar mil caixas custa mil posições de array e zero alocações por frame.

O preço é que você não tem "a caixa" na mão, tem um `Long` que é o índice dela. Toda a API trabalha assim:

```vba
Dim caixa As Long
caixa = sc.AddBox(pedra, -2, -2, 2, 2, 1, 0.5)
sc.SetMaterial caixa, gelo
sc.MoveBy caixa, 0, 0, 3
```

O id vale pela vida do objeto porque nada nunca reordena os arrays. A exceção é `Remove`, que devolve o slot para a loja: depois de removê-lo, o id nomeia quem tomou o lugar. Para esconder algo que você pretende trazer de volta existe `SetActive`, que mantém o slot reservado.

## Não existe malha

Uma caixa não é uma coisa guardada. É uma chamada que acontece, do mesmo jeito que um retângulo num canvas 2D. Não há vértice guardado, não há buffer, não há transformação de modelo.

É por isso que as primitivas são comandos do renderer:

```vba
rd.DrawBox matId, x1, y1, x2, y2, topo, espessura
```

e não métodos de um objeto malha. E é por isso que desenhar mil caixas não aloca nada.

## O material decide quase tudo

Cor, se é sólido, se atravessa, atrito, quique, empuxo, se dá para escalar: tudo mora no material, não no objeto. Dois objetos com o mesmo material se comportam igual, e trocar o material de um objeto muda como ele é desenhado e como ele colide na mesma linha.

```vba
sc.SetMaterial ponte, quebradica
```

Leia [Materiais](Materiais.md).

## O corpo reporta, não decide

`PBody` sabe andar, pular, subir degrau, escalar, nadar e ser carregado por plataforma. O que ele não sabe é o que as coisas significam. Ele devolve a lista de gatilhos que encostou e recusa adivinhar se aquilo era moeda, checkpoint ou mina.

```vba
For k = 0 To body.TouchCount - 1
    idx = body.TouchAt(k)
    If sc.TagOf(idx) = MEU_CHECKPOINT Then
        ' regra sua
    End If
Next k
```

Pelo mesmo motivo a cena guarda uma `tag` por objeto, que é um número seu, em vez de um `IsCheckpoint`. As regras do seu domínio ficam no seu código.

## A ordem de um frame

```
UpdateMotion       o que se move na cena se move
body.Advance       o corpo é carregado, acelerado, empurrado, e cai
camera.SetPosition a câmera segue o corpo
BeginFrame         troca o banco de shapes e ajusta o orçamento
RenderScene        cull, orçamento, ordem, desenho
EndFrame           apaga o frame anterior
bomba de refresh   o PowerPoint repinta
DoEvents           e o repintar acontece
```

A ordem entre os dois primeiros importa. Uma plataforma que já se moveu é uma plataforma que consegue carregar o corpo; uma que se move depois é uma que o corpo passa um frame ao lado.

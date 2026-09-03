# Ordenação por profundidade

Sem Z buffer, a ordem em que os polígonos são pintados **é** a profundidade. Essa é a parte mais difícil da lib e vale entender por quê.

## Os quatro estágios do `Render`

**1. Gastar o orçamento.** Os objetos visíveis são ordenados pela borda próxima da esfera envolvente, e o custo por primitiva é cobrado até estourar o teto. O que cai fora é sempre o fundo, nunca a parede na frente do jogador.

| primitiva | custo |
|---|---|
| caixa | 3 |
| caixa girada | 6 |
| rampa | 4 |
| billboard | 1 |
| spinner | 2 |

**2. Semear a ordem.** Os sobreviventes são ordenados por profundidade de centro, do mais longe para o mais perto.

**3. Filtrar por espaço de tela.** Os oito cantos da caixa de cada sobrevivente são projetados e viram um retângulo. Dois objetos cujos retângulos não se tocam não geram restrição nenhuma.

**4. Ordenar por oclusão.** Para cada par restante, procura-se um plano alinhado aos eixos que separe as duas caixas. Existindo o plano, o lado em que o olho está decide exatamente quem é o de trás. Cada relação vira uma aresta, e um sort topológico monta a sequência que satisfaz todas de uma vez.

## Por que plano separador, e não profundidade

Nenhum número único de profundidade resolve isso.

Ordenar por centro faz o chão, que é largo e tem centro perto, pintar por cima de tudo que está apoiado nele. Ordenar pela borda distante faz um objeto pequeno e longe atravessar a parede larga que está na frente dele.

O teste de plano separador acerta os dois casos, que é o que geometria de blocos precisa.

## Quando não existe sequência

Três objetos podem cada um estar provadamente atrás do próximo, formando um ciclo, e aí nenhuma ordem satisfaz tudo. Alguma restrição tem que cair.

O critério é quem deve menos: força-se o objeto com o menor número de restrições pendentes, e no empate o mais distante, porque um emaranhado que nenhuma regra ordena exato ainda fica melhor pintado de trás para frente.

Par que se sobrepõe nos três eixos não tem resposta exata, porque os dois realmente se interpenetram, e é decidido por profundidade de centro com o tamanho como último desempate.

## O bug da moeda atravessando a plataforma

Sintoma: olhando de baixo para cima, uma moeda apoiada numa plataforma continuava visível através dela, e objetos girando passavam por cima de coisas claramente na frente.

O comparador estava certo o tempo todo. Medindo contra um oráculo exato, 628 das 638 violações tinham o par decidido corretamente. Quem jogava fora a decisão era o desempate de ciclo. O ciclo era este:

| objeto | restrição |
|---|---|
| plataforma `z[8.95, 9.25]` | antes da moeda, por plano Z, porque o olho está embaixo |
| moeda `z[9.73, 10.37]` | antes da plataforma móvel, por plano X |
| plataforma móvel `y[52, 53.6]` | antes da primeira, por plano Y |

As três corretas, e impossíveis juntas. Mas a primeira e a terceira não dividiam um único pixel: uma aparecia embaixo na tela, a outra em cima. Aresta invisível, ciclo real, e o desempate antigo, que escolhia o objeto cujo bloqueador é menor, sacrificava sempre a moeda, porque moeda é sempre a coisa pequena.

Três correções, cada uma medida antes de escrever:

| correção | por quê |
|---|---|
| filtro de espaço de tela | mata as arestas invisíveis, que eram a origem dos ciclos |
| retângulo em vez de círculo | a esfera envolvente é folgada demais: descartava 15% dos pares, o retângulo descarta 60% |
| desempate por menor dívida, empate no mais longe | o critério antigo escolhia a vítima pelo tamanho do bloqueador, o que sempre condenava o pequeno |
| AABB da caixa girada segue o ângulo | era o círculo de varredura, que reivindica chão que a laje não ocupa |

Resultado contra o oráculo exato, contando só pares que realmente se sobrepõem na tela:

| cenário | antes | depois |
|---|---|---|
| `obby.json`, 1743 câmeras vezes 4 ângulos de giro | 1,846 % | 0,000 % |
| cenário adversário: chão gigante, lajes afundadas nele, moedas, discos, rampas, 40 seeds | 6,152 % | 1,014 % |

Do 1,01 % restante, 34 dos 36 casos ainda são ciclo genuíno e 2 são geometria que de fato se interpenetra, onde não existe resposta certa sem cortar polígono.

Foram testadas ainda a quebra de ciclo por componente fortemente conexo e por menor dano na tela. Nenhuma das duas melhorou: a regra simples já escolhe dentro do ciclo. É o piso desta abordagem.

## Inspecionando

```vba
sc.Render rd
For i = 0 To sc.DrawnCount - 1
    Debug.Print i, sc.DrawnAt(i)
Next i
```

É a única coisa de um frame que não dá para conferir olhando as shapes depois.

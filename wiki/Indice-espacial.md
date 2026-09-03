# Índice espacial

`QueryBox` era uma varredura sobre todos os objetos, e o `PBody` a chama uma vez por passe de colisão, umas cinco vezes por frame. Com 2500 objetos isso é 12 500 iterações por frame só para descobrir o que está debaixo dos pés.

Agora existe uma grade uniforme em XY, montada por contagem mais soma de prefixos num array plano: sem lista por célula, sem alocação com o frame rodando.

| objetos | varredura toca | grade toca |
|---|---|---|
| 40 | 40 | 24 |
| 576 | 576 | 30 |
| 2500 | 2500 | 31 |

## A célula é dimensionada pelo objeto médio

Um terreno de tiles de 1 metro indexado em células de 4 metros colocaria dezesseis tiles em cada balde e não indexaria nada.

## Duas coisas ficam de fora e são varridas

**O que se move.** Quem chamou `MoveBy`, `SetTopZ`, `SetAngle`, `SetMotion` ou `SetSpin` sai da grade e não volta. Uma coisa que se moveu vai se mover de novo, e remontar a grade todo frame por causa de três plataformas custa mais do que ela economiza.

**O que é grande demais.** Um plano de chão aparece em todas as células. Indexá-lo enche cada balde com o único objeto que toda consulta ia alcançar de qualquer jeito, então acima de 32 células ele é varrido junto com as plataformas. Num cenário com chão gigante isso derruba as entradas do índice de 5879 para 2158.

`DynamicCount` e `CellSize` contam o que ficou de fora e quão fina a grade ficou.

## Um índice pode estar errado, não só lento

Uma consulta que perde um objeto é um corpo caindo pelo chão. Então ele não foi amostrado, foi comparado: 120 000 consultas aleatórias contra a varredura completa, em cenas feitas de propósito com o que quebra grade, ou seja objetos muito maiores que a célula, objetos fora da área indexada, objetos desligados, e objetos que se moveram. Zero divergências.

O mesmo teste roda no VBA, em `CheckIndex`.

## O que continua linear

O cull de frustum do `Render`, uma passada por frame. O frustum é um cone, e o retângulo que o cobre pega tantas células que a busca sairia mais cara que a varredura. É decisão, não esquecimento.

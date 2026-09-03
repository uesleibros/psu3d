# Limites conhecidos

O que a lib não faz, e por quê. Nada aqui é bug, é fronteira.

## Sem Z buffer

Não existe profundidade por pixel. Polígonos inteiros são ordenados e pintados de trás para frente.

**Geometria que se interpenetra não tem ordem correta.** Duas caixas que se atravessam não têm "quem está na frente", e o melhor que a lib faz é escolher a ordem menos visível.

Na prática, medindo contra um oráculo exato: 0,00 % de artefatos visíveis numa fase real, 1,01 % num cenário construído de propósito para ser hostil, e desses a maior parte é ciclo genuíno. Detalhe em [Ordenação por profundidade](Ordenacao-por-profundidade.md).

Contorno prático: evite geometria que se atravessa. Encostada é fácil de ordenar, atravessada não é.

## Sem rasterização própria

O polígono já transformado, recortado e colorido é entregue ao PowerPoint. O preenchimento é dele. Isso é o que impede o Z buffer, e é também o que faz a lib caber em VBA.

## Sem textura

Cor chapada por face. Sem UV, sem interpolação perspectiva correta. Contorno por face é o que existe para dar detalhe.

## Flat shading

Uma cor por face, da tabela pré-calculada. Sem Gouraud, sem Phong, sem normal por vértice.

## Rotação só em yaw

`AddRotatedBox` gira no eixo vertical. Não há pitch nem roll de objeto, e a câmera não tem roll.

`DrawQuad` e `DrawTriangle` aceitam três ou quatro vértices arbitrários em 3D, então orientação qualquer dá para desenhar na mão. O que falta é malha guardada com matriz de modelo.

## Sem malha

Não há carregamento de OBJ nem buffer de vértice. As primitivas são geradas por chamada. Isso é decisão de projeto: sem malha guardada, desenhar mil caixas não aloca nada.

## O cull de frustum é linear

Uma passada por todos os objetos por frame. O índice espacial cobre as consultas de física, não o cull. O frustum é um cone, e o retângulo que o cobre pega tantas células que a busca sairia mais cara que a varredura.

Com 2500 objetos essa passada é real. Se isso te incomodar, o caminho é hierarquia de volumes, não grade.

## Grade limitada a 4096 células por eixo

Um mundo absurdamente grande degrada para busca lenta em vez de travar. É clamp de propósito.

## O texto do HUD não é texto

A engine não desenha texto. Shape de texto obriga a mexer no z-order todo frame, e no PowerPoint isso repinta o slide inteiro. O HUD do demo é feito de retângulos por `DrawPolygon2D`.

A única shape de texto do loop é a bomba de refresh, e ela existe exatamente porque repinta o slide.

## Single, não Double

Toda a geometria é `Single`. São cerca de 7 dígitos significativos, o que é bastante para um mundo de algumas centenas de unidades e insuficiente para coordenadas na casa dos milhões. Se a sua cena é enorme, mova a origem em vez de aumentar os números.

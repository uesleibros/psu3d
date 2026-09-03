# Como o 3D funciona

Uma pergunta justa sobre um motor 3D dentro do PowerPoint é se ele é 3D de verdade ou uma imitação com escala por distância. É 3D de verdade, e dá para apontar exatamente onde.

## As quatro etapas que definem 3D

### 1. Transformação de view

`PCamera.WorldToView` é uma mudança de base ortonormal com yaw e pitch:

```
fwd  = (dx*cosYaw + dy*sinYaw) * cosPitch + dz * sinPitch
side = -dx*sinYaw + dy*cosYaw
up   = dz*cosPitch - (dx*cosYaw + dy*sinYaw) * sinPitch
```

Isso é matriz de view escrita à mão. O ponto importante: **o pitch é rotação de verdade**, não deslocamento vertical da imagem. Deslocar a imagem é o truque barato que motor falso usa, e ele se entrega quando você olha muito para cima ou para baixo, porque a perspectiva não acompanha.

Existe um `ScreenShiftY` separado no `PCamera`, mas ele serve para balanço de cabeça e coice de arma, não para olhar para cima.

### 2. Projeção perspectiva

`PCanvas.Project`:

```
invD = focal / fwd
x = centroX + side * invD
y = centroY - up   * invD
```

Câmera pinhole de livro texto. É por isso que paralelas convergem certo e a distorção nas bordas é a correta.

### 3. Recorte de frustum

Cinco planos em espaço de view (near mais os quatro laterais), recorte estilo Sutherland e Hodgman. Um polígono que atravessa o plano near é **recortado**, gerando vértices novos, não descartado.

Esse detalhe é o divisor de águas. Quem só escala sprite por distância não tem isso, e se entrega quando uma parede passa pelo olho.

### 4. Backface culling

Sinal do produto escalar entre a normal da face e o vetor do ponto até o olho, com a normal vinda de um produto vetorial. A normal fica sem normalizar de propósito: dividir pelo próprio comprimento custa uma raiz e três divisões por face, e o teste só precisa do sinal, que escala positiva não muda.

Mais iluminação direcional por normal e névoa por profundidade.

## Onde está a fronteira

Duas coisas que um motor 3D completo tem e Psu3D não.

### Não existe Z buffer

Não há profundidade por pixel. Polígonos inteiros são ordenados e pintados de trás para frente, que é o algoritmo do pintor.

A consequência concreta: **geometria que se interpenetra não tem ordem correta**, porque não existe "quem está na frente" para duas caixas que se atravessam. Foi por isso que a ordenação virou a parte mais difícil da lib, e por isso ela quebra ciclo em vez de resolver. Leia [Ordenação por profundidade](Ordenacao-por-profundidade).

Um Z buffer mataria isso em uma linha, e é impossível aqui pelo motivo seguinte.

### Não rasterizamos

O polígono já transformado, recortado e colorido é entregue ao PowerPoint, e ele preenche. Psu3D faz geometria e iluminação; o preenchimento é dele. É o mesmo contrato que um renderer SVG tem.

### O resto

Cor chapada por face, sem textura e sem UV. Flat shading, sem Gouraud nem Phong. Rotação de objeto só em yaw, embora `DrawQuad` e `DrawTriangle` aceitem três ou quatro vértices arbitrários em 3D, então orientação qualquer dá para fazer na mão; o que falta é malha guardada com matriz.

## A comparação honesta

Não é three.js. O three.js tem GPU, Z buffer, textura e malha.

Psu3D é um renderer 3D por software da geração anterior à placa aceleradora: transformação e recorte corretos, algoritmo do pintor, flat shading, primitivas geradas em vez de malhas carregadas. Que é literalmente como o 3D funcionava antes de existir placa 3D.

Então: 3D de verdade, sim. De 1996, rodando dentro do PowerPoint.

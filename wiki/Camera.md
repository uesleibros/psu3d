# Câmera

Posição no mundo mais duas rotações.

```vba
Dim cam As PCamera
Set cam = New PCamera

cam.Init 0, -10, 1.7, P_HALF_PI, 0    ' x, y, z, yaw, pitch
cam.SetPosition 0, -10, 1.7
cam.SetAngles P_HALF_PI, 0
cam.AddAngles 0.02, -0.01
cam.LookAt 0, 0, 0
```

`Yaw` é para onde você está virado no plano do chão, em radianos. `Pitch` é o quanto está olhando para cima ou para baixo.

O pitch é limitado, e o limite é ajustável com `PitchLimit`. Sem limite você vira de cabeça para baixo e o yaw passa a girar ao contrário, que é desorientador e nunca é o que alguém quis.

O seno e o cosseno dos dois ângulos ficam guardados e só são recalculados quando o ângulo muda. Um frame lê `CosYaw` umas mil vezes, e chamar `Cos` mil vezes com o mesmo argumento é trabalho jogado fora.

## Direções

`ForwardX` e `ForwardY` são para onde a câmera aponta no plano do chão. `RightX` e `RightY` são o lado direito dela. Juntos, viram movimento:

```vba
wishX = cam.ForwardX * frente + cam.RightX * lado
wishY = cam.ForwardY * frente + cam.RightY * lado
```

Repare que são só duas componentes. Andar não deve subir quando você olha para cima, então o vetor de frente é o achatado.

## Transformação

`WorldToView` leva um ponto do mundo para o espaço da câmera, devolvendo frente, lado e cima. Essa é a matriz de view da lib, e está detalhada em [Como o 3D funciona](Como-o-3D-funciona).

`ViewDepth` devolve só a componente de profundidade, que é o que a ordenação usa. `SphereVisible` responde se uma esfera cabe no frustum.

Na prática nem o `PScene` nem o `PRenderer` chamam esses métodos no caminho quente: eles leem o seno, o cosseno e a posição uma vez por frame e escrevem a conta na mão, porque uma chamada de método por vértice é a coisa mais cara de um frame que é quase todo aritmética.

## Balanço e coice

`ScreenShiftY` desloca a imagem verticalmente depois da projeção. É para balanço de cabeça andando e para coice de arma. Não use isso para olhar para cima: para isso existe o pitch, que é rotação de verdade.

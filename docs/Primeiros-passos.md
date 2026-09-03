# Primeiros passos

## Uma figura estática, sem jogo nenhum

Nada de loop, nada de física, nada de teclado. As shapes ficam no slide quando o Sub termina, e você pode fechar o VBE e continuar editando a apresentação normalmente.

Cole isto no módulo do slide (`Slide1`, não num módulo padrão, porque `Shapes` precisa resolver):

```vba
Public Sub Terreno()
    Dim x As Long, y As Long, h As Single, terra As Long

    Psu3D.Boot Shapes, 40, 40, 640, 400
    terra = PMaterials.Create("terra", PCore.ColorPack(96, 140, 90)).Id

    Psu3D.Camera.SetPosition -16, -16, 13
    Psu3D.Camera.LookAt 0, 0, 0
    Psu3D.Renderer.PolyBudget = 3000

    For y = 0 To 23
        For x = 0 To 23
            h = 2 + Sin(x * 0.4) * Cos(y * 0.35) * 1.6
            Psu3D.Scene.AddBox terra, x - 12, y - 12, x - 11, y - 11, h, h
        Next x
    Next y

    Psu3D.BeginFrame 0
    Psu3D.RenderScene
    Psu3D.EndFrame
End Sub
```

São 576 caixas num retângulo de 640 por 400 pontos do slide. Troque o `h` por um valor da planilha e vira gráfico de barras 3D. Troque por ruído e vira terreno. A engine não sabe a diferença.

## O que cada linha faz

`Psu3D.Boot Shapes, 40, 40, 640, 400` monta canvas, câmera, renderer e cena de uma vez, cria a paleta de materiais padrão, e limpa qualquer shape que tenha sobrado de uma execução anterior. Os quatro números são x, y, largura e altura do retângulo do slide onde vai desenhar. Sem eles o canvas ocupa o slide inteiro.

`PMaterials.Create` registra uma superfície e devolve o objeto. O `.Id` é o número que a cena usa daí em diante, porque um `Long` num array é mais barato que um ponteiro de objeto.

`Psu3D.Camera.LookAt 0, 0, 0` aponta a câmera para um ponto do mundo, calculando yaw e pitch.

`Psu3D.Renderer.PolyBudget = 3000` diz quantos polígonos o frame pode gastar. 576 caixas custam 3 cada, então 1728, e o teto de 3000 garante que nada é cortado. Sem isso o orçamento padrão cortaria o fundo do terreno.

`BeginFrame`, `RenderScene`, `EndFrame` é o frame: troca de banco de shapes, faz cull mais ordenação mais desenho, e aposenta o frame anterior.

## Um loop de verdade

Quando você quer movimento, o desenho vai para dentro de um laço e aparecem duas coisas que só existem no PowerPoint.

```vba
Public Sub Main()
    Dim dt As Single, t0 As Single
    Dim timerShp As Shape

    Set timerShp = Slide1.Shapes("timer")

    Psu3D.Boot Shapes
    Psu3D.Camera.SetPosition 0, 0, 1.7
    Psu3D.Scene.AddBox PMaterials.IdOf("grass"), -20, -20, 20, 20, 0, 1

    t0 = Timer

    Do While ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 1
        dt = Timer - t0
        If dt < 0.001 Then dt = 0.001
        If dt > 0.05 Then dt = 0.05
        t0 = Timer

        Psu3D.Camera.AddAngles (UCursor.CursorX - Psu3D.Canvas.CenterX) * 0.0026, 0
        UCursor.SetCurPos Psu3D.Canvas.CenterX, Psu3D.Canvas.CenterY

        Psu3D.BeginFrame dt
        Psu3D.RenderScene
        Psu3D.EndFrame

        timerShp.TextEffect.Text = Format$(Timer, "0.00")
        DoEvents
    Loop

    Psu3D.Shutdown
End Sub
```

A primeira coisa estranha é o `timerShp`. Ele não é um relógio. Escrever num WordArt é o que força o PowerPoint a repintar o slide durante uma apresentação, e o repintar é como o frame se torna visível. Leia [a bomba de refresh](Performance.md#a-bomba-de-refresh) antes de mexer nele.

A segunda é o `dt` limitado entre 1 ms e 50 ms. Um frame que demorou meio segundo, porque o Windows resolveu fazer outra coisa, moveria tudo meio segundo de uma vez e atravessaria parede. O limite superior transforma isso em câmera lenta, que é sempre preferível.

## Sem a fachada

`Psu3D` é conveniência. Nada obriga a usá-la:

```vba
Dim cv As PCanvas, cam As PCamera, rd As PRenderer, sc As PScene

Set cv = New PCanvas
cv.Init 40, 40, 640, 400

Set cam = New PCamera
cam.Init 0, 0, 2, P_HALF_PI, 0

Set rd = New PRenderer
rd.Attach Shapes, cv, cam

Set sc = New PScene
```

E nada obriga a usar a cena. O renderer desenha primitiva direto, sem nada guardado. Veja [Renderer e primitivas](Renderer-e-primitivas.md).

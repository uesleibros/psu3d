# Receitas

## Um cubo girando num slide

```vba
Public Sub Cubo()
    Dim i As Long, t As Single, m As Long

    Psu3D.Boot Shapes, 100, 60, 500, 320
    m = PMaterials.Create("cubo", PCore.ColorPack(220, 120, 60)).Id
    Psu3D.Camera.SetPosition -6, -6, 4
    Psu3D.Camera.LookAt 0, 0, 0.5

    i = Psu3D.Scene.AddRotatedBox(m, 0, 0, 1, 1, 1, 2)

    For t = 0 To 6.2 Step 0.05
        Psu3D.Scene.SetAngle i, t
        Psu3D.BeginFrame 0.05
        Psu3D.RenderScene
        Psu3D.EndFrame
        Slide1.Shapes("timer").TextEffect.Text = Format$(Timer, "0.000")
        DoEvents
    Next t

    Psu3D.Shutdown
End Sub
```

## Gráfico de barras 3D a partir de células

```vba
Public Sub Grafico(ByVal dados As Variant)
    Dim i As Long, alt As Single, m As Long

    Psu3D.Boot Shapes, 40, 40, 640, 380
    m = PMaterials.Create("barra", PCore.ColorPack(70, 130, 200)).Id
    Psu3D.Camera.SetPosition -8, -14, 9
    Psu3D.Camera.LookAt 0, 0, 2
    Psu3D.Renderer.PolyBudget = 400

    For i = LBound(dados) To UBound(dados)
        alt = CSng(dados(i))
        Psu3D.Scene.AddBox m, i * 1.4, -0.5, i * 1.4 + 1, 0.5, alt, alt
    Next i

    Psu3D.BeginFrame 0
    Psu3D.RenderScene
    Psu3D.EndFrame
End Sub
```

O `thick` igual à altura faz a barra descer até o zero, que é o que uma barra deve fazer.

## Primeira pessoa mínima

```vba
Dim body As PBody
Set body = Psu3D.CreateBody()
body.SetPosition 0, 0, 2

Do While ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 1
    dt = Timer - t0
    If dt < 0.001 Then dt = 0.001
    If dt > 0.05 Then dt = 0.05
    t0 = Timer

    dx = Psu3D.Canvas.CursorX - Psu3D.Canvas.HalfWidth
    dy = Psu3D.Canvas.CursorY - Psu3D.Canvas.HalfHeight
    Psu3D.Camera.AddAngles dx * 0.0026, -dy * 0.0026
    Psu3D.Canvas.CenterCursor

    frente = 0: lado = 0
    If KeyDown(vbKeyW) Then frente = frente + 1
    If KeyDown(vbKeyS) Then frente = frente - 1
    If KeyDown(vbKeyD) Then lado = lado + 1
    If KeyDown(vbKeyA) Then lado = lado - 1

    wishX = Psu3D.Camera.ForwardX * frente + Psu3D.Camera.RightX * lado
    wishY = Psu3D.Camera.ForwardY * frente + Psu3D.Camera.RightY * lado

    Psu3D.Scene.UpdateMotion dt
    body.Advance dt, wishX, wishY, KeyDown(vbKeySpace), KeyDown(vbKeyShift)
    If body.CarryYaw <> 0 Then Psu3D.Camera.AddAngles body.CarryYaw, 0

    Psu3D.Camera.SetPosition body.X, body.Y, body.Z + 1.62

    Psu3D.BeginFrame dt
    Psu3D.RenderScene
    Psu3D.EndFrame

    timerShp.TextEffect.Text = Format$(Timer, "0.00")
    DoEvents
Loop
```

Isso depende do `UCursor` para o mouse e de uma função `KeyDown` com `GetAsyncKeyState`. O `PDemo.bas` tem as duas prontas.

## Plataforma que vai e volta, com você em cima

```vba
p = sc.AddBox(metal, -1.4, -1, 1.4, 1, 5.4, 0.4)
sc.SetMotion p, 1, 0, 0, 5.5, 0.85, 0
```

E no loop, antes do corpo:

```vba
sc.UpdateMotion dt
body.Advance dt, wishX, wishY, up, down
```

O corpo é carregado sozinho. A ordem é que importa: a cena se move antes.

## Disco que gira e leva você junto

```vba
d = sc.AddRotatedBox(metal, 0, 5, 3.6, 0.85, 6.4, 0.4, 0)
sc.SetSpin d, 0.95
```

E some `body.CarryYaw` no yaw da câmera, senão você gira em volta do centro do disco olhando sempre para o mesmo lado.

## Elevador movido na mão

```vba
zAntes = zAgora
zAgora = 3.6 + Sin(t * 0.7) * 3.2
sc.SetTopZ elevador, zAgora
If body.GroundObject = elevador Then body.Nudge 0, 0, zAgora - zAntes
```

`Nudge` empurra sem mexer na velocidade, porque ser carregado não é ser acelerado.

## Colisor invisível

```vba
Set m = PMaterials.Create("clip", 0, pcSolid)
m.Visible = False
sc.AddBox m.Id, -10, 8, 10, 8.4, 4, 4
```

Objeto invisível é pulado já na coleta: não é cullado, nem ordenado, nem desenhado, e não gasta orçamento.

## Duas vistas no mesmo slide

```vba
Set cvMini = New PCanvas
cvMini.Init 20, 20, 250, 150

Set camMini = New PCamera

Set rdMini = New PRenderer
rdMini.Prefix = "mini_"
rdMini.Attach Shapes, cvMini, camMini
rdMini.PolyBudget = 48

' por frame, depois da vista principal:
camMini.SetPosition body.X - cam.ForwardX * 9, body.Y - cam.ForwardY * 9, body.Z + 10
camMini.LookAt body.X, body.Y, body.Z + 1
rdMini.BeginFrame dt
sc.Render rdMini
rdMini.EndFrame
```

O prefixo diferente é obrigatório. Sem ele um renderer apaga as shapes do outro.

## Carregar uma fase de arquivo

```vba
If Not PLevel.ParseFile("C:\fases\obby.json", sc) Then
    MsgBox PLevel.Error
    Exit Sub
End If

body.SetPosition PLevel.SpawnX, PLevel.SpawnY, PLevel.SpawnZ
cam.SetAngles PLevel.SpawnYaw, 0
If PLevel.HasKillZ Then body.KillZ = PLevel.KillZ
If PLevel.Gravity > 0 Then body.Gravity = PLevel.Gravity
```

## Rodar o demo

```vba
PDemo.RunDemo                        ' obby.json, a fase difícil
PDemo.RunPool                        ' piscina, para testar água
PDemo.RunArena                       ' arena, com minivista e movedores
PDemo.RunFile "C:\fases\minha.json"  ' sua fase
```

O slide precisa de uma shape de WordArt chamada `timer`. Ela é a bomba de refresh, não um relógio.

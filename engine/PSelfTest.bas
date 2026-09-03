Attribute VB_Name = "PSelfTest"
'/**
' * PSelfTest - Psu3D Acceptance Check
' * @description Exercises every module of the library against known answers and reports what passed. Meant to be the first thing run after importing the engine into a project, because the failures worth catching there are silent ones: a module left out of the import, a class that did not come across, or the JSON reader missing.
' * @author UesleiDev
' * @version 1.0
' * @scope Deliberately NOT Option Private Module, for the same reason as PDemo: the macro dialog only lists public entry points, and this one has to be startable from there. Every engine module stays private.
' * @remarks Runs entirely in memory. The renderer is put in dry run so the whole pipeline executes without a slide, which means this can be run from the VBE with no presentation open and leaves nothing behind.
' */

Option Explicit

'/** @section Module state */

Private m_pass As Long
Private m_fail As Long
Private m_log As String

'/** @section Entry points */

'/**
' * @brief Runs every check and shows the result.
' * @remarks The one call to make after importing the library.
' */
Public Sub Psu3DSelfTest()
    MsgBox RunAll(), vbInformation, "Psu3D self test"
End Sub

'/**
' * @brief Runs every check and returns the report.
' * @return A line per failure, then a summary line.
' */
Public Function RunAll() As String
    m_pass = 0
    m_fail = 0
    m_log = vbNullString

    CheckMath
    CheckColor
    CheckLighting
    CheckMaterials
    CheckCanvas
    CheckCamera
    CheckScene
    CheckIndex
    CheckMotion
    CheckBody
    CheckRenderer
    CheckOrder
    CheckStability
    CheckLevel
    CheckSave
    CheckHardening

    If m_fail = 0 Then
        RunAll = "All " & m_pass & " checks passed." & vbLf & _
                 "The library is imported correctly and every module answers."
    Else
        RunAll = m_log & vbLf & m_fail & " of " & (m_pass + m_fail) & " checks failed."
    End If
End Function

'/** @section Checks */

'/**
' * @brief Checks the scalar helpers.
' */
Private Sub CheckMath()
    Ok "math.clamp", PCore.Clamp(9!, 0!, 5!) = 5!
    Ok "math.clamp01", PCore.Clamp01(-3!) = 0!
    Ok "math.lerp", Near(PCore.Lerp(10!, 20!, 0.25), 12.5)
    Ok "math.inverselerp", Near(PCore.InverseLerp(10!, 20!, 12.5), 0.25)
    Ok "math.inverselerp degenerate", PCore.InverseLerp(4!, 4!, 9!) = 0!
    Ok "math.wrap", Near(PCore.WrapAngle(P_TWO_PI + 1!), 1!)
    Ok "math.angledelta", Near(PCore.AngleDelta(3!, -3!), P_TWO_PI - 6!)
    Ok "math.boxesoverlap", PCore.BoxesOverlap(0!, 0!, 2!, 2!, 1!, 1!, 3!, 3!)
    Ok "math.boxesapart", Not PCore.BoxesOverlap(0!, 0!, 1!, 1!, 2!, 2!, 3!, 3!)

    PCore.SeedRandom 12345
    Dim a As Double
    a = PCore.RandomNext()
    PCore.SeedRandom 12345
    Ok "math.random repeats", PCore.RandomNext() = a
End Sub

'/**
' * @brief Checks colour packing and parsing.
' */
Private Sub CheckColor()
    Dim c As Long

    c = PCore.ColorPack(10, 20, 30)
    Ok "color.roundtrip", PCore.ColorRed(c) = 10 And PCore.ColorGreen(c) = 20 And PCore.ColorBlue(c) = 30
    Ok "color.clamps", PCore.ColorRed(PCore.ColorPack(999, 0, 0)) = 255
    Ok "color.fromhex", PCore.ColorFromHex("#0A141E") = c
    Ok "color.shorthex", PCore.ColorFromHex("#FFF") = PCore.ColorPack(255, 255, 255)
    Ok "color.badhex", PCore.ColorFromHex("#ZZZZZZ") = 0
    Ok "color.tohex", PCore.ColorToHex(c) = "#0A141E"
    Ok "color.mix", PCore.ColorMix(PCore.ColorPack(0, 0, 0), PCore.ColorPack(200, 200, 200), 0.5) = PCore.ColorPack(100, 100, 100)
End Sub

'/**
' * @brief Checks the light and fog environment.
' */
Private Sub CheckLighting()
    PLighting.ResetDefaults

    Ok "light.normalised", Near(PLighting.LightX * PLighting.LightX + _
                               PLighting.LightY * PLighting.LightY + _
                               PLighting.LightZ * PLighting.LightZ, 1!)

    PLighting.SetFog 10!, 20!, PCore.ColorPack(1, 2, 3)
    Ok "fog.near band", PLighting.FogStepOf(5!) = 0
    Ok "fog.far band", PLighting.FogStepOf(30!) = PLighting.FogSteps
    Ok "fog.midway", PLighting.FogStepOf(15!) > 0 And PLighting.FogStepOf(15!) < PLighting.FogSteps
    Ok "fog.solid at range", PLighting.ApplyFog(P_WHITE, 25!) = PLighting.FogColor
    Ok "fog.untouched near", PLighting.ApplyFog(P_WHITE, 5!) = P_WHITE

    PLighting.ResetDefaults
End Sub

'/**
' * @brief Checks the material registry and its baked shading.
' */
Private Sub CheckMaterials()
    Dim m As PMaterial
    Dim id As Long

    PMaterials.Clear
    Ok "mat.default exists", PMaterials.Count = 1
    Ok "mat.id zero valid", Not PMaterials.ById(0) Is Nothing
    Ok "mat.bad id falls back", PMaterials.ById(999).Id = 0

    Set m = PMaterials.Create("selftest_ice", PCore.ColorPack(120, 140, 160), pcSolid)
    m.Friction = 0.12
    id = m.Id

    Ok "mat.lookup by name", PMaterials.IdOf("selftest_ice") = id
    Ok "mat.lookup is case free", PMaterials.IdOf("SELFTEST_ICE") = id
    Ok "mat.unknown name", PMaterials.IdOf("nope") = P_INVALID_ID
    Ok "mat.property kept", Near(PMaterials.ById(id).Friction, 0.12)
    Ok "mat.create is idempotent", PMaterials.Create("selftest_ice").Id = id

    Set m = PMaterials.Create("selftest_one", P_WHITE, pcOneWay)
    Ok "mat.oneway blocks from above", m.BlocksMovement(True)
    Ok "mat.oneway passes from below", Not m.BlocksMovement(False)

    Set m = PMaterials.Create("selftest_water", P_WHITE, pcTrigger)
    m.Buoyancy = 1.06
    Ok "mat.fluid needs both", m.IsFluid
    m.Collision = pcSolid
    Ok "mat.solid is not fluid", Not m.IsFluid

    Ok "mat.shade band differs with fog", _
        PMaterials.ShadeBand(id, pdPosZ, 0) <> PMaterials.ShadeBand(id, pdPosZ, PLighting.FogSteps)
    Ok "mat.shade differs by face", _
        PMaterials.ShadeBand(id, pdPosZ, 0) <> PMaterials.ShadeBand(id, pdNegZ, 0)

    PMaterials.Clear
End Sub

'/**
' * @brief Checks the canvas rectangle and its projection.
' */
Private Sub CheckCanvas()
    Dim cv As PCanvas
    Dim sx As Single
    Dim sy As Single

    Set cv = New PCanvas
    cv.Init 100!, 50!, 640!, 360!

    Ok "canvas.right", cv.Right = 740!
    Ok "canvas.bottom", cv.Bottom = 410!
    Ok "canvas.centre", cv.CenterX = 420! And cv.CenterY = 230!
    Ok "canvas.aspect", Near(cv.Aspect, 640! / 360!)
    Ok "canvas.contains", cv.Contains(420!, 230!)
    Ok "canvas.excludes", Not cv.Contains(50!, 230!)

    '/* A point straight ahead lands dead centre, whatever the field of view. */
    Ok "canvas.projects centre", cv.Project(10!, 0!, 0!, sx, sy) And Near(sx, cv.CenterX) And Near(sy, cv.CenterY)

    '/* Half the focal length to the side at unit depth lands half a screen across. */
    Ok "canvas.projects side", cv.Project(1!, 1!, 0!, sx, sy) And sx > cv.CenterX
    Ok "canvas.projects up", cv.Project(1!, 0!, 1!, sx, sy) And sy < cv.CenterY
    Ok "canvas.rejects behind", Not cv.Project(-1!, 0!, 0!, sx, sy)

    cv.FieldOfViewDeg = 45
    Ok "canvas.fov clamped in", Near(cv.FieldOfViewDeg, 45!)
    Ok "canvas.vertical fov derived", cv.TanFovV < cv.TanFovH
End Sub

'/**
' * @brief Checks the camera transform.
' */
Private Sub CheckCamera()
    Dim cam As PCamera
    Dim f As Single
    Dim s As Single
    Dim u As Single

    Set cam = New PCamera
    cam.Init 0!, 0!, 0!, 0!, 0!

    cam.WorldToView 5!, 0!, 0!, f, s, u
    Ok "camera.forward is +x at yaw 0", Near(f, 5!) And Near(s, 0!) And Near(u, 0!)

    cam.WorldToView 0!, 5!, 0!, f, s, u
    Ok "camera.left is +y", Near(f, 0!) And Near(s, 5!)

    cam.WorldToView 0!, 0!, 5!, f, s, u
    Ok "camera.up is +z", Near(u, 5!)

    cam.Yaw = P_HALF_PI
    cam.WorldToView 0!, 5!, 0!, f, s, u
    Ok "camera.turns", Near(f, 5!)

    cam.SetPitchLimits -1!, 1!
    cam.Pitch = 99!
    Ok "camera.pitch clamped", Near(cam.Pitch, 1!)

    cam.SetAngles 0!, 0!
    Ok "camera.depth", Near(cam.ViewDepth(7!, 0!, 0!), 7!)

    cam.LookAt 0!, 10!, 0!
    Ok "camera.lookat turns to +y", Near(cam.Yaw, P_HALF_PI)
End Sub

'/**
' * @brief Checks the object store, its queries and its geometry answers.
' */
Private Sub CheckScene()
    Dim sc As PScene
    Dim box As Long
    Dim ramp As Long
    Dim disc As Long
    Dim b As Single
    Dim t As Single
    Dim n As Long

    PMaterials.Clear
    Set sc = New PScene

    box = sc.AddBox(0, -2!, -2!, 2!, 2!, 1!, 1!)
    Ok "scene.counts", sc.Count = 1 And sc.LiveCount = 1

    sc.GetSpanZ box, b, t
    Ok "scene.span", Near(t, 1!) And Near(b, 0!)
    Ok "scene.top of box is flat", Near(sc.TopZAt(box, 1.9, -1.9), 1!)
    Ok "scene.contains", sc.ContainsXY(box, 0!, 0!)
    Ok "scene.excludes", Not sc.ContainsXY(box, 9!, 0!)

    ramp = sc.AddRamp(0, 0!, 0!, 10!, 4!, 2!, 0!, 0.4, paX)
    Ok "scene.ramp high end", Near(sc.TopZAt(ramp, 0!, 2!), 2!)
    Ok "scene.ramp low end", Near(sc.TopZAt(ramp, 10!, 2!), 0!)
    Ok "scene.ramp middle", Near(sc.TopZAt(ramp, 5!, 2!), 1!)
    Ok "scene.ramp clamps past the end", Near(sc.TopZAt(ramp, 99!, 2!), 0!)

    '/* A slab far longer than it is wide: the bounds hold points the slab itself does not. */
    disc = sc.AddRotatedBox(0, 0!, 0!, 4!, 0.5, 1!, 0.4, 0!)
    Ok "scene.rot contains along its length", sc.ContainsXY(disc, 3.5, 0!)
    Ok "scene.rot excludes across it", Not sc.ContainsXY(disc, 0!, 3.5)

    sc.SetAngle disc, P_HALF_PI
    Ok "scene.rot follows the turn", sc.ContainsXY(disc, 0!, 3.5)
    Ok "scene.rot excludes after turning", Not sc.ContainsXY(disc, 3.5, 0!)

    Ok "scene.query finds", sc.QueryBox(-1!, -1!, 1!, 1!) > 0
    Ok "scene.query misses", sc.QueryBox(500!, 500!, 501!, 501!) = 0

    sc.SetActive box, False
    Ok "scene.deactivates", Not sc.IsActive(box) And sc.LiveCount = 2
    sc.SetActive box, True
    Ok "scene.reactivates", sc.IsActive(box) And sc.LiveCount = 3

    '/* Remove hands the slot back; the next object takes it instead of growing the store. */
    sc.Remove box
    Ok "scene.remove deactivates", Not sc.IsActive(box)

    n = sc.Count
    Ok "scene.reuses the freed slot", sc.AddBox(0, 0!, 0!, 1!, 1!, 1!, 1!) = box
    Ok "scene.reuse costs no slot", sc.Count = n

    sc.Clear
    Ok "scene.clears", sc.Count = 0
End Sub

'/**
' * @brief Checks that animated objects move and report the step they took.
' */
Private Sub CheckMotion()
    Dim sc As PScene
    Dim lift As Long
    Dim dx As Single
    Dim dy As Single
    Dim dz As Single
    Dim b As Single
    Dim t0 As Single
    Dim t1 As Single

    PMaterials.Clear
    Set sc = New PScene

    lift = sc.AddBox(0, -1!, -1!, 1!, 1!, 4!, 0.4)
    sc.GetSpanZ lift, b, t0

    sc.SetMotion lift, 0!, 0!, 1!, 2!, 1!, 0!
    sc.UpdateMotion 0.5
    sc.GetSpanZ lift, b, t1

    Ok "motion.moves", t1 <> t0
    sc.GetMotionDelta lift, dx, dy, dz
    Ok "motion.reports its step", Near(t1 - t0, dz)
    Ok "motion.step is on its axis", Near(dx, 0!) And Near(dy, 0!)

    Ok "motion.reports how many moved", sc.UpdateMotion(0.1) = 1
End Sub

'/**
' * @brief Checks the full render pipeline with the slide taken out of it.
' */
Private Sub CheckRenderer()
    Dim sc As PScene
    Dim cv As PCanvas
    Dim cam As PCamera
    Dim rd As PRenderer
    Dim drawn As Long

    PMaterials.Clear
    PMaterials.CreateDefaults

    Set sc = New PScene
    sc.AddBox PMaterials.IdOf("stone"), -6!, 4!, 6!, 12!, 0!, 1!
    sc.AddBox PMaterials.IdOf("brick"), -2!, 6!, 2!, 8!, 3!, 3!

    Set cv = New PCanvas
    cv.Init 0!, 0!, 960!, 540!

    Set cam = New PCamera
    cam.Init 0!, 0!, 2!, P_HALF_PI, 0!

    Set rd = New PRenderer
    Set rd.Canvas = cv
    Set rd.Camera = cam
    rd.DryRun = True
    rd.PolyBudget = 64

    rd.BeginFrame
    drawn = sc.Render(rd)

    Ok "render.draws what is in front", drawn = 2
    Ok "render.emits polygons", rd.PolyCount > 0
    Ok "render.stays inside budget", rd.PolyCount <= rd.PolyBudget

    '/* Turn right around: the same scene is now behind the camera. */
    cam.Yaw = -P_HALF_PI
    rd.BeginFrame
    Ok "render.culls what is behind", sc.Render(rd) = 0
End Sub

'/**
' * @brief Checks the JSON level reader end to end.
' */
Private Sub CheckLevel()
    Dim sc As PScene
    Dim src As String

    PMaterials.Clear
    Set sc = New PScene

    src = "{""level"":{""spawn"":[1,2,3],""yaw"":90,""budget"":[10,20],""killz"":-5}," & _
          """materials"":{""rock"":{""color"":""#204080"",""friction"":0.5}}," & _
          """objects"":[" & _
          "{""type"":""box"",""mat"":""rock"",""from"":[0,0],""to"":[2,2],""top"":1,""thick"":1," & _
          """repeat"":3,""step"":{""from"":[3,0],""to"":[3,0],""top"":0.5}}]}"

    Ok "level.parses", PLevel.Parse(src, sc)
    Ok "level.repeat expands", sc.Count = 3
    Ok "level.spawn read", Near(PLevel.SpawnX, 1!) And Near(PLevel.SpawnY, 2!) And Near(PLevel.SpawnZ, 3!)
    Ok "level.yaw in radians", Near(PLevel.SpawnYaw, P_HALF_PI)
    Ok "level.budget read", PLevel.BudgetMin = 10 And PLevel.BudgetMax = 20
    Ok "level.killz is reported apart", PLevel.HasKillZ And Near(PLevel.KillZ, -5!)
    Ok "level.material built", PMaterials.IdOf("rock") >= 0
    Ok "level.material colour", PMaterials.ById(PMaterials.IdOf("rock")).Color = PCore.ColorFromHex("#204080")
    Ok "level.step applied", Near(sc.TopZAt(2, 7!, 1!), 2!)

    '/* The failures worth reporting are the ones that would otherwise build a silently empty world. */
    Ok "level.rejects junk", Not PLevel.Parse("not json at all", sc)
    Ok "level.rejects empty", Not PLevel.Parse("{}", sc)
    Ok "level.names the unknown material", _
        Not PLevel.Parse("{""objects"":[{""type"":""box"",""mat"":""ghostmat""}]}", sc) And _
        InStr(PLevel.Error, "ghostmat") > 0

    PMaterials.Clear
End Sub

'/**
' * @brief Checks the body: gravity, walls, stairs, ladders, fluids, triggers and being carried.
' * @description Run headless, with no slide and no renderer, because none of this is about drawing.
' * One floor is built and every case is a body dropped somewhere else on it. Anything meant to be
' * landed on stands above that floor, since two surfaces at the same height make the test about which
' * one the query happened to reach first.
' */
Private Sub CheckBody()
    Dim sc As PScene
    Dim b As PBody
    Dim solid As Long
    Dim tall As Long
    Dim ladder As Long
    Dim waterMat As Long
    Dim coinMat As Long
    Dim springy As Long
    Dim coin As Long
    Dim disc As Long
    Dim lift As Long
    Dim i As Long
    Dim apex As Single
    Dim bounced As Boolean
    Dim mark As Single

    PMaterials.Clear
    Set sc = New PScene

    solid = PMaterials.Create("t_solid", PCore.ColorPack(120, 120, 120), pcSolid).Id
    tall = PMaterials.Create("t_tall", PCore.ColorPack(120, 120, 120), pcSolid).Id
    PMaterials.ById(tall).StepHeight = 1.2
    ladder = PMaterials.Create("t_ladder", PCore.ColorPack(90, 70, 40), pcGhost).Id
    PMaterials.ById(ladder).Climbable = True
    waterMat = PMaterials.Create("t_water", PCore.ColorPack(60, 120, 200), pcTrigger).Id
    PMaterials.ById(waterMat).Buoyancy = 1.15
    PMaterials.ById(waterMat).Drag = 4!
    coinMat = PMaterials.Create("t_coin", PCore.ColorPack(250, 205, 60), pcTrigger).Id
    springy = PMaterials.Create("t_spring", PCore.ColorPack(200, 60, 90), pcSolid).Id
    PMaterials.ById(springy).Bounce = 0.8

    sc.AddBox solid, -60!, -60!, 60!, 60!, 0!, 1!

    Set b = New PBody
    b.Attach sc

    '/* Falls, and stops when it meets the floor. */
    b.SetPosition 0!, 0!, 5!
    b.Advance 0.016, 0!, 0!, False, False
    Ok "body.gravity pulls down", b.VelZ < 0!

    For i = 1 To 120
        b.Advance 0.016, 0!, 0!, False, False
    Next i

    Ok "body.lands on the floor", b.OnGround And Near(b.Z, 0!)
    Ok "body.stops falling once landed", Near(b.VelZ, 0!)

    '/* Jumps, and only from the ground. */
    b.Advance 0.016, 0!, 0!, True, False
    Ok "body.jumps", b.VelZ > 0! And Not b.OnGround
    mark = b.VelZ
    b.Advance 0.016, 0!, 0!, True, False
    Ok "body.cannot jump twice", b.VelZ < mark

    '/* A wall it cannot climb over stops it, at exactly its own radius from the face. */
    sc.AddBox solid, 4!, -6!, 5!, 6!, 3!, 3!
    b.SetPosition 0!, 0!, 0!

    For i = 1 To 200
        b.Advance 0.016, 1!, 0!, False, False
    Next i

    Ok "body.is stopped by a wall", Near(b.X, 3.66)
    Ok "body.does not tunnel through", b.OnGround And Near(b.Z, 0!)

    '/* A kerb inside the step height it walks straight up. */
    sc.AddBox solid, -30!, -6!, -4!, 6!, 0.3, 0.3
    b.SetPosition 0!, 0!, 0!

    For i = 1 To 200
        b.Advance 0.016, -1!, 0!, False, False
    Next i

    Ok "body.steps up a kerb", Near(b.Z, 0.3) And b.X < -5!

    '/* Twice that height, in a material that says it may be stepped on: over that too. */
    sc.AddBox tall, -30!, 10!, -4!, 20!, 0.9, 0.9
    b.SetPosition 0!, 15!, 0!

    For i = 1 To 200
        b.Advance 0.016, -1!, 0!, False, False
    Next i

    Ok "body.material may raise the step", Near(b.Z, 0.9) And b.X < -5!

    '/* The same height in a plain material stops it dead. */
    sc.AddBox solid, -30!, -30!, -4!, -20!, 0.9, 0.9
    b.SetPosition 0!, -25!, 0!

    For i = 1 To 200
        b.Advance 0.016, -1!, 0!, False, False
    Next i

    Ok "body.plain material keeps the default step", Near(b.Z, 0!) And Near(b.X, -3.66)

    '/* A ladder: entered rather than blocked, and climbed while a direction is held. */
    sc.AddBox ladder, 20!, -1!, 21!, 1!, 6!, 6!
    b.SetPosition 20.5, 0!, 0!
    b.Advance 0.016, 0!, 0!, False, False
    mark = b.Z
    b.Advance 0.016, 0!, 0!, True, False

    Ok "body.notices a ladder", b.IsClimbing
    Ok "body.climbs it", b.Z > mark
    Ok "body.hangs there instead of falling", b.VelZ > 0!

    For i = 1 To 40
        b.Advance 0.016, 0!, 0!, True, False
    Next i

    Ok "body.keeps climbing", b.Z > 1.5

    b.Advance 0.016, 0!, 0!, False, True
    Ok "body.climbs back down", b.VelZ < 0!

    '/* Water: a trigger the body can be inside, which holds it up. */
    sc.AddBox waterMat, 30!, -4!, 38!, 4!, 3!, 3!
    b.SetPosition 34!, 0!, 0!
    b.Advance 0.016, 0!, 0!, False, False

    Ok "body.knows it is submerged", b.Submersion > 0.9
    Ok "body.swims when deep enough", b.IsSwimming

    b.Advance 0.016, 0!, 0!, True, False
    Ok "body.strokes upward", b.VelZ > 0!

    b.SetPosition 0!, 0!, 0!
    b.Advance 0.016, 0!, 0!, False, False
    Ok "body.is dry on land", Near(b.Submersion, 0!)

    '/* A trigger is reported and never acted on. */
    coin = sc.AddSpinner(coinMat, 45!, 0!, 0.6, 0.32)
    b.SetPosition 45!, 0!, 0!
    b.Advance 0.016, 0!, 0!, False, False

    Ok "body.reports the trigger it touched", b.TouchCount = 1 And b.TouchAt(0) = coin
    Ok "body.leaves the trigger alone", sc.IsActive(coin)

    b.SetPosition 0!, 0!, 0!
    b.Advance 0.016, 0!, 0!, False, False
    Ok "body.reports nothing when clear", b.TouchCount = 0

    '/* Falling out of the world. */
    b.KillZ = -5!
    b.SetPosition 0!, 0!, -10!
    b.Advance 0.016, 0!, 0!, False, False
    Ok "body.knows it fell out", b.FellOut

    b.SetPosition 0!, 0!, 1!
    b.Advance 0.016, 0!, 0!, False, False
    Ok "body.is fine above the floor", Not b.FellOut
    b.KillZ = -1E+30

    '/* Bounce: a hard landing on a springy surface is thrown back. */
    sc.AddBox springy, -50!, 30!, -44!, 36!, 0.5, 0.5
    b.SetPosition -47!, 33!, 4.5
    bounced = False
    apex = -99!

    For i = 1 To 300
        b.Advance 0.016, 0!, 0!, False, False
        If b.VelZ > 0.5 Then bounced = True
        If bounced And b.Z > apex Then apex = b.Z
    Next i

    Ok "body.is thrown back by a springy floor", bounced
    Ok "body.bounce keeps most of the drop", apex > 1.4

    '/* Carried by a platform that travels. */
    lift = sc.AddBox(solid, 8!, 20!, 12!, 24!, 1!, 1!)
    sc.SetMotion lift, 1!, 0!, 0!, 3!, 0.5, 0!
    b.SetPosition 10!, 22!, 1.5

    For i = 1 To 40
        b.Advance 0.016, 0!, 0!, False, False
    Next i

    Ok "body.stands on the platform", b.GroundObject = lift And Near(b.Z, 1!)

    mark = b.X
    sc.UpdateMotion 0.1
    b.Advance 0.001, 0!, 0!, False, False
    Ok "body.travels with the platform", Abs(b.X - mark) > 0.05

    '/* Carried around by a platform that turns. This is the case a travel-only carry gets wrong: the
    ' * floor rotates, the body keeps the world position it had, and it is left standing over nothing
    ' * without having moved an inch. */
    disc = sc.AddRotatedBox(solid, -20!, 40!, 3!, 1!, 1!, 1!, 0!)
    b.SetPosition -18!, 40!, 1.5

    For i = 1 To 40
        b.Advance 0.016, 0!, 0!, False, False
    Next i

    Ok "body.stands on the turning platform", b.GroundObject = disc

    sc.SetSpin disc, 1!, 0!
    sc.UpdateMotion 0.1
    b.Advance 0.001, 0!, 0!, False, False

    Ok "body.is swung around the centre", _
        Near(b.X, -20! + 2! * Cos(0.1)) And Near(b.Y, 40! + 2! * Sin(0.1))
    Ok "body.reports how far it was turned", Near(b.CarryYaw, 0.1)
    Ok "body.is still on the platform after turning", b.GroundObject = disc

    PMaterials.Clear
End Sub

'/**
' * @brief Checks that a level written out and read back describes the same world.
' */
Private Sub CheckSave()
    Dim sc As PScene
    Dim sc2 As PScene
    Dim src As String
    Dim txt As String
    Dim n As Long
    Dim i As Long

    PMaterials.Clear
    Set sc = New PScene
    Set sc2 = New PScene

    src = "{""level"":{""spawn"":[1,2,3],""yaw"":90,""killz"":-5}," & _
          """materials"":{""rock"":{""color"":""#204080"",""friction"":0.5}," & _
          """flag"":{""color"":""#20FF80"",""collision"":""trigger""}}," & _
          """objects"":[" & _
          "{""type"":""box"",""mat"":""rock"",""from"":[0,0],""to"":[2,2],""top"":1,""thick"":1}," & _
          "{""type"":""ramp"",""mat"":""rock"",""from"":[4,0],""to"":[8,2],""low"":0,""high"":2," & _
          """thick"":0.5,""axis"":""x""}," & _
          "{""type"":""rot"",""mat"":""rock"",""at"":[12,0],""half"":[3,1],""top"":1,""thick"":0.4," & _
          """angle"":30,""spin"":1.5}," & _
          "{""type"":""box"",""mat"":""flag"",""from"":[16,0],""to"":[18,2],""top"":1,""thick"":1," & _
          """tag"":7}," & _
          "{""type"":""spin"",""mat"":""flag"",""at"":[20,0,2],""radius"":0.4}," & _
          "{""type"":""box"",""mat"":""rock"",""from"":[24,0],""to"":[26,2],""top"":1,""thick"":1," & _
          """move"":{""axis"":[0,0,1],""amp"":2,""speed"":0.7}}]}"

    Ok "save.level loads first", PLevel.Parse(src, sc)
    n = sc.Count

    txt = PLevel.Save(sc)
    Ok "save.writes something", Len(txt) > 0

    '/* The one failure that would not show up as a wrong number: a machine set to a comma decimal
     ' * separator writing a document no JSON reader will take back. */
    Ok "save.writes a dot decimal point", InStr(txt, "0.5") > 0 And InStr(txt, "0,5") = 0

    Ok "save.reloads", PLevel.Parse(txt, sc2)
    Ok "save.same object count", sc2.Count = n
    Ok "save.spawn survives", Near(PLevel.SpawnX, 1!) And Near(PLevel.SpawnZ, 3!)
    Ok "save.yaw survives", Near(PLevel.SpawnYaw, P_HALF_PI)
    Ok "save.killz survives", PLevel.HasKillZ And Near(PLevel.KillZ, -5!)

    Ok "save.box survives", Near(sc2.TopZAt(0, 1!, 1!), 1!)
    Ok "save.ramp survives", Near(sc2.TopZAt(1, 8!, 1!), 2!) And Near(sc2.TopZAt(1, 4!, 1!), 0!)
    Ok "save.turned box keeps its angle", Near(sc2.AngleOf(2), 30! * P_DEG2RAD)
    Ok "save.tag survives", sc2.TagOf(3) = 7
    Ok "save.spinner survives", sc2.KindOf(4) = pkSpinner
    Ok "save.motion survives", sc2.UpdateMotion(0.1) = 2

    Ok "save.material colour survives", _
        PMaterials.ById(PMaterials.IdOf("rock")).Color = PCore.ColorFromHex("#204080")
    Ok "save.material friction survives", Near(PMaterials.ById(PMaterials.IdOf("rock")).Friction, 0.5)
    Ok "save.material collision survives", PMaterials.ById(PMaterials.IdOf("flag")).IsTrigger

    Ok "save.refuses an empty path", Not PLevel.SaveFile("", sc)

    PMaterials.Clear
End Sub

'/**
' * @brief Checks that the painter order puts the far thing first, from above and from below.
' * @description The bug this locks down: a coin resting on a platform stayed visible through the
' * platform when the camera was underneath it. The pair was ordered correctly all along; what went
' * wrong was that a third object on the far side of the level, sharing no pixel with either of them,
' * closed a cycle, and breaking the cycle sacrificed the coin. So the checks here are not only about
' * the coin and the platform: the decoy is the whole point of the test.
' */
Private Sub CheckOrder()
    Dim sc As PScene
    Dim rd As PRenderer
    Dim cv As PCanvas
    Dim cam As PCamera
    Dim solid As Long
    Dim shiny As Long
    Dim plat As Long
    Dim coin As Long
    Dim decoy As Long
    Dim i As Long
    Dim pPlat As Long
    Dim pCoin As Long
    Dim disc As Long
    Dim bx1 As Single, by1 As Single, bx2 As Single, by2 As Single

    PMaterials.Clear
    Set sc = New PScene
    Set cv = New PCanvas
    Set cam = New PCamera
    Set rd = New PRenderer

    cv.Init 0!, 0!, 720!, 405!
    Set rd.Canvas = cv
    Set rd.Camera = cam
    rd.DryRun = True
    rd.PolyBudget = 400

    solid = PMaterials.Create("o_plat", PCore.ColorPack(150, 150, 150), pcSolid).Id
    shiny = PMaterials.Create("o_coin", PCore.ColorPack(250, 205, 60), pcTrigger).Id

    plat = sc.AddBox(solid, -1.4, 45.6, 1.4, 48!, 9.25, 0.3)
    coin = sc.AddSpinner(shiny, 0!, 46.8, 10.05, 0.32)

    '/* Far away, well off to the side and much higher: it shares no canvas space with the other two,
    ' * so whatever order it takes is invisible, and it must not be allowed to disturb theirs. */
    decoy = sc.AddBox(solid, -1.3, 52!, 1.3, 53.6, 10.2, 0.4)

    '/* Underneath: the platform is nearer, so it must be painted last, over the coin. */
    cam.Init 0!, 30!, 4!, P_HALF_PI, 0.35
    sc.Render rd

    pPlat = -1
    pCoin = -1

    For i = 0 To sc.DrawnCount - 1
        If sc.DrawnAt(i) = plat Then pPlat = i
        If sc.DrawnAt(i) = coin Then pCoin = i
    Next i

    Ok "order.both are drawn from below", pPlat >= 0 And pCoin >= 0
    Ok "order.platform hides the coin from below", pCoin < pPlat

    '/* Above: the coin is nearer, so it must be painted last, over the platform. */
    cam.Init 0!, 30!, 16!, P_HALF_PI, -0.35
    sc.Render rd

    pPlat = -1
    pCoin = -1

    For i = 0 To sc.DrawnCount - 1
        If sc.DrawnAt(i) = plat Then pPlat = i
        If sc.DrawnAt(i) = coin Then pCoin = i
    Next i

    Ok "order.both are drawn from above", pPlat >= 0 And pCoin >= 0
    Ok "order.coin sits on the platform from above", pPlat < pCoin

    '/* A turning slab keeps a bounding box that follows its angle, or every plane test against a
    ' * neighbour fails on ground the slab does not occupy. */
    disc = sc.AddRotatedBox(solid, 0!, 20!, 3.6, 0.85, 2!, 0.4, 0!)
    sc.GetBoundsXY disc, bx1, by1, bx2, by2
    Ok "order.turned slab is not a circle", Near(by2 - by1, 1.7) And Near(bx2 - bx1, 7.2)

    sc.SetAngle disc, P_HALF_PI
    sc.GetBoundsXY disc, bx1, by1, bx2, by2
    Ok "order.turned slab follows the angle", Near(by2 - by1, 7.2) And Near(bx2 - bx1, 1.7)

    PMaterials.Clear
End Sub

'/**
' * @brief Proves the spatial index answers exactly what a full sweep would have answered.
' * @description The index is the one optimisation here that can be wrong instead of merely slow, and a
' * query that quietly misses an object is a body that falls through a floor. So it is not sampled, it is
' * compared: a few hundred random boxes against a scene deliberately built out of the shapes that break
' * grids, meaning objects far wider than a cell, objects outside the indexed area, objects switched off,
' * and objects that have moved and therefore left the index.
' */
Private Sub CheckIndex()
    Dim sc As PScene
    Dim i As Long
    Dim k As Long
    Dim trial As Long
    Dim mismatches As Long
    Dim fromGrid As Long
    Dim fromSweep As Long
    Dim found As Boolean
    Dim qx1 As Single, qy1 As Single, qx2 As Single, qy2 As Single
    Dim bx1 As Single, by1 As Single, bx2 As Single, by2 As Single
    Dim px As Single, py As Single
    Dim probe As PScene
    Dim hit(0 To 511) As Long

    PMaterials.Clear
    Set sc = New PScene
    PCore.SeedRandom 20260903

    '/* A field of small tiles: the case where a grid earns its keep. */
    For i = 0 To 13
        For k = 0 To 13
            sc.AddBox 0, i * 1! - 7!, k * 1! - 7!, i * 1! - 6!, k * 1! - 6!, 0!, 0.5
        Next k
    Next i

    '/* Slabs far wider than any cell, so they are listed in many cells at once. */
    For i = 0 To 5
        px = PCore.RandomRange(-14!, 6!)
        py = PCore.RandomRange(-14!, 6!)
        sc.AddBox 0, px, py, px + 12!, py + 9!, 1!, 0.4
    Next i

    '/* Far outside the indexed area, which the clamped border cell has to account for. */
    sc.AddBox 0, 60!, 60!, 64!, 64!, 0!, 0.5
    sc.AddBox 0, -90!, 12!, -86!, 16!, 0!, 0.5

    '/* Switched off: never an answer. */
    sc.SetActive sc.AddBox(0, 0!, 0!, 3!, 3!, 2!, 0.4), False

    '/* Moved, so they belong to the swept list rather than the grid. */
    For i = 0 To 4
        k = sc.AddBox(0, PCore.RandomRange(-12!, 4!), PCore.RandomRange(-12!, 4!), 0!, 0!, 1.5, 0.4)
        sc.MoveBy k, PCore.RandomRange(-3!, 3!), PCore.RandomRange(-3!, 3!), 0!
    Next i

    k = sc.AddBox(0, -4!, -4!, -1!, -1!, 3!, 0.4)
    sc.SetMotion k, 1!, 0!, 0!, 2!, 0.6

    '/* Two reasons an object is swept instead of indexed, and they are checked apart. First, it moved. */
    Set probe = New PScene

    For i = 0 To 8
        probe.AddBox 0, i * 1!, 0!, i * 1! + 1!, 1!, 0!, 0.5
    Next i

    Ok "index.indexes what stands still", probe.DynamicCount = 0
    probe.MoveBy 3, 0.5, 0!, 0!
    probe.SetMotion 5, 1!, 0!, 0!, 2!, 0.6
    Ok "index.a moved object leaves the grid", probe.DynamicCount = 2

    '/* Second, it is so big that indexing it would put it in every bucket anyway. */
    probe.AddBox 0, -400!, -400!, 400!, 400!, -1!, 0.5
    Ok "index.a huge object is swept, not indexed", probe.DynamicCount = 3

    Ok "index.keeps the movers apart", sc.DynamicCount >= 6
    Ok "index.sizes the cell from the objects", sc.CellSize > 0!

    For trial = 1 To 150
        qx1 = PCore.RandomRange(-20!, 14!)
        qy1 = PCore.RandomRange(-20!, 14!)
        qx2 = qx1 + PCore.RandomRange(0.05, 9!)
        qy2 = qy1 + PCore.RandomRange(0.05, 9!)

        fromGrid = sc.QueryBox(qx1, qy1, qx2, qy2)

        For k = 0 To fromGrid - 1
            If k <= UBound(hit) Then hit(k) = sc.ResultAt(k)
        Next k

        fromSweep = 0

        For i = 0 To sc.Count - 1
            If sc.IsActive(i) Then
                sc.GetBoundsXY i, bx1, by1, bx2, by2

                If Not (bx2 < qx1 Or bx1 > qx2 Or by2 < qy1 Or by1 > qy2) Then
                    fromSweep = fromSweep + 1
                    found = False

                    For k = 0 To fromGrid - 1
                        If k <= UBound(hit) Then
                            If hit(k) = i Then found = True
                        End If
                    Next k

                    If Not found Then mismatches = mismatches + 1
                End If
            End If
        Next i

        If fromGrid <> fromSweep Then mismatches = mismatches + 1
    Next trial

    Ok "index.answers exactly what a full sweep answers", mismatches = 0

    '/* A query nowhere near anything still has to come back empty rather than clamped onto a border cell. */
    Ok "index.finds nothing where there is nothing", sc.QueryBox(500!, 500!, 501!, 501!) = 0
    Ok "index.finds the far outlier", sc.QueryBox(61!, 61!, 62!, 62!) = 1

    '/* Adding after the index was built has to invalidate it. */
    sc.AddBox 0, 200!, 200!, 201!, 201!, 0!, 0.5
    Ok "index.notices a new object", sc.QueryBox(200!, 200!, 201!, 201!) = 1

    sc.Clear
    Ok "index.survives being emptied", sc.QueryBox(-5!, -5!, 5!, 5!) = 0

    PMaterials.Clear
End Sub

'/**
' * @brief Checks that the things a broken document or a broken world can do are survived, not crashed on.
' * @description Everything here was a real failure found by reading the code rather than by playing it,
' * and every one of them ended in a runtime error rather than a wrong picture, which in a slideshow means
' * the show stops. A library that reads files anyone can edit has to treat a typo as data.
' */
Private Sub CheckHardening()
    Dim sc As PScene
    Dim sc2 As PScene
    Dim m As PMaterial
    Dim txt As String
    Dim n As Long

    PMaterials.Clear
    Set sc = New PScene

    '/* A repeat past the limit is named, not obeyed. Obeying it builds until the machine gives up,
    ' * with a slideshow running and no way to interrupt. */
    Ok "hard.a runaway repeat is refused", _
        Not PLevel.Parse("{""materials"":{""a"":{}},""objects"":[" & _
                         "{""type"":""box"",""mat"":""a"",""repeat"":999999999}]}", sc)
    Ok "hard.and it says which field", InStr(PLevel.Error, "repeat") > 0

    '/* A repeat too big for a Long used to overflow on the conversion, before any check could run. */
    Ok "hard.a repeat past Long is refused", _
        Not PLevel.Parse("{""materials"":{""a"":{}},""objects"":[" & _
                         "{""type"":""box"",""mat"":""a"",""repeat"":1e12}]}", sc)

    '/* Same for the budget and the tag, which are read as numbers and used as counts. */
    Ok "hard.an absurd budget is clamped", _
        PLevel.Parse("{""level"":{""budget"":[0,1e12]},""materials"":{""a"":{}}," & _
                     """objects"":[{""type"":""box"",""mat"":""a""}]}", sc) _
        And PLevel.BudgetMax <= 1000000

    Ok "hard.an absurd tag is clamped", _
        PLevel.Parse("{""materials"":{""a"":{}},""objects"":[" & _
                     "{""type"":""box"",""mat"":""a"",""tag"":1e12}]}", sc) _
        And sc.TagOf(0) > 0

    '/* One object at a mistyped coordinate used to ask the lookup grid for two billion cells along an
    ' * axis, which overflowed the multiplication meant to catch it. */
    PMaterials.Clear
    Set sc = New PScene

    '/* Small tiles keep the cell small, and one object a billion units away then asks for a billion
    ' * cells across. The old code multiplied that by the row count before it clamped anything, which
    ' * overflowed a Long and stopped the show. */
    For n = 0 To 4
        For i = 0 To 4
            sc.AddBox 0, n * 1!, i * 5!, n * 1! + 1!, i * 5! + 1!, 0!, 0.5
        Next i
    Next n

    sc.AddBox 0, 1000000000!, 0!, 1000000001!, 1!, 0!, 0.5

    Ok "hard.a stray coordinate does not break the index", sc.QueryBox(0.2, 0.2, 0.8, 0.8) = 1
    Ok "hard.and the stray object is still findable", sc.QueryBox(1000000000!, 0!, 1000000000!, 1!) = 1
    Ok "hard.the grid stays within its ceiling", sc.CellSize > 0!

    '/* A material may be named anything, including something that would end the JSON string early. */
    PMaterials.Clear
    Set sc = New PScene
    Set sc2 = New PScene

    Set m = PMaterials.Create("a" & Chr$(34) & "b\c", PCore.ColorPack(10, 20, 30), pcSolid)
    sc.AddBox m.Id, 0!, 0!, 2!, 2!, 1!, 1!

    txt = PLevel.Save(sc)
    Ok "hard.an awkward name is written escaped", InStr(txt, "\" & Chr$(34)) > 0 And InStr(txt, "\\") > 0
    Ok "hard.and the document still reads back", PLevel.Parse(txt, sc2)
    Ok "hard.with the object intact", sc2.Count = sc.Count

    PMaterials.Clear
End Sub

'/**
' * @brief Checks that two nearly identical frames agree with each other.
' * @description A frame can be correct and still look wrong if the next one disagrees with it. An object
' * that changes level of detail, or drops out of the budget, between two frames that are almost the same
' * is a flicker, and a flicker is what a player actually notices. Both decisions carry hysteresis, and
' * hysteresis can only be tested by holding a value right on the threshold and jiggling it.
' */
Private Sub CheckStability()
    Dim sc As PScene
    Dim rd As PRenderer
    Dim cv As PCanvas
    Dim cam As PCamera
    Dim solid As Long
    Dim i As Long
    Dim edge As Single
    Dim dist As Single
    Dim flips As Long
    Dim wasLod As Boolean
    Dim isLod As Boolean
    Dim first As Boolean
    Dim n1 As Long
    Dim n2 As Long

    PMaterials.Clear
    Set sc = New PScene
    Set cv = New PCanvas
    Set cam = New PCamera
    Set rd = New PRenderer

    cv.Init 0!, 0!, 720!, 405!

    '/* Pinned rather than inherited. Left alone, the far plane falls back to the end of the fog, and
    ' * an earlier check may have moved that, which would cull the boxes this one depends on. */
    cv.FarPlane = 60!

    Set rd.Canvas = cv
    Set rd.Camera = cam
    rd.DryRun = True
    rd.PolyBudget = 400

    solid = PMaterials.Create("s_box", PCore.ColorPack(150, 150, 150), pcSolid).Id

    '/* One box, and a camera parked at exactly the distance where its projected size sits on the level
    ' * of detail line, jiggling by a few centimetres. The bounding sphere of a box spanning two by two
    ' * by one has a radius of 1.5, and a projected size is radius times focal length over distance, so
    ' * the distance that lands exactly on the threshold can be worked out rather than guessed. */
    sc.AddBox solid, -1!, 20!, 1!, 22!, 1!, 1!
    sc.LodSize = 200!

    edge = 1.5 * cv.FocalLength / (sc.LodSize * 0.5!)

    flips = 0
    first = True

    For i = 0 To 90
        dist = edge + Sin(i * 0.9) * 0.2
        cam.Init 0!, 21! - dist, 0.5, P_HALF_PI, 0!
        sc.Render rd

        If sc.DrawnCount > 0 Then
            isLod = sc.WasReduced(0)

            If Not first Then
                If isLod <> wasLod Then flips = flips + 1
            End If

            wasLod = isLod
            first = False
        End If
    Next i

    '/* With one threshold the answer changes on nearly every frame. With two it never changes at all,
    ' * because the jiggle stays inside the band between them. */
    Ok "stable.level of detail does not dither on the threshold", flips = 0

    '/* Ten boxes at three polygons each, and a budget that admits six of them. Shaving two polygons off
    ' * the budget would drop the sixth, which is the object sitting on the cut. It is already on screen,
    ' * so it is allowed to overspend by a little rather than blink out. */
    PMaterials.Clear
    Set sc = New PScene
    solid = PMaterials.Create("s_box", PCore.ColorPack(150, 150, 150), pcSolid).Id

    For i = 0 To 9
        sc.AddBox solid, -0.4, 6! + i * 2!, 0.4, 7! + i * 2!, 1!, 1!
    Next i

    sc.LodSize = 0!
    cam.Init 0!, 0!, 0.5, P_HALF_PI, 0!

    rd.PolyBudget = 18
    n1 = sc.Render(rd)

    rd.PolyBudget = 16
    n2 = sc.Render(rd)

    Ok "stable.the budget admits six of ten", n1 = 6
    Ok "stable.an object on screen is not cut for two polygons", n2 >= n1

    '/* And an object that was never drawn does not sneak in through the same door. */
    Ok "stable.the slack does not admit the rest", n2 < 10

    PMaterials.Clear
End Sub

'/** @section Plumbing */

'/**
' * @brief Records the outcome of one check.
' * @param name What was being checked.
' * @param passed Whether it held.
' */
Private Sub Ok(ByVal name As String, ByVal passed As Boolean)
    If passed Then
        m_pass = m_pass + 1
    Else
        m_fail = m_fail + 1
        m_log = m_log & "FAILED  " & name & vbLf
    End If
End Sub

'/**
' * @brief Compares two values with the slack a Single deserves.
' * @param a The first value.
' * @param b The second value.
' * @return True when they agree to within a thousandth.
' */
Private Function Near(ByVal a As Single, ByVal b As Single) As Boolean
    Near = (Abs(a - b) < 0.001)
End Function

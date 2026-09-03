Attribute VB_Name = "PDemo"
'/**
' * PDemo - Psu3D Playable Test Level
' * @description A small first person level that exercises every part of the engine: a canvas covering the whole slide, the stock material palette, ramps, one-way platforms, ice, glass, a walk-through pillar and pickups. Doubles as the reference for how the pieces are wired together in a real loop: the walking, colliding and swimming all belong to PBody, and what is left here is which keys this game uses and what it decides a coin is.
' * @author UesleiDev
' * @version 1.0
' * @remarks This module is deliberately NOT declared Option Private Module: the macro dialog and shape action buttons only see public entry points. The engine modules themselves stay private.
' */

Option Explicit

'/** @section Native declarations */

'/**
' * @description A point in screen pixels, which is the unit the cursor API speaks.
' */
Private Type POINTAPI
    X As Long
    Y As Long
End Type

#If VBA7 Then
    '/** @description Reports whether a key is down at this instant, independent of focus. */
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

    '/** @description Reads where the pointer is, in screen pixels. */
    Private Declare PtrSafe Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long

    '/** @description Warps the pointer, in screen pixels. */
    Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

    '/** @description Shows or hides the pointer. The count is internal, so every hide needs a show. */
    Private Declare PtrSafe Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

    '/** @description Reads a screen measurement, used here for the centre of the primary display. */
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#Else
    '/** @description Reports whether a key is down at this instant, independent of focus. */
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

    '/** @description Reads where the pointer is, in screen pixels. */
    Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long

    '/** @description Warps the pointer, in screen pixels. */
    Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

    '/** @description Shows or hides the pointer. The count is internal, so every hide needs a show. */
    Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

    '/** @description Reads a screen measurement, used here for the centre of the primary display. */
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If

'/** @description Index of the primary screen width in GetSystemMetrics. */
Private Const SM_CXSCREEN As Long = 0

'/** @description Index of the primary screen height in GetSystemMetrics. */
Private Const SM_CYSCREEN As Long = 1

'/** @section Tuning */

'/** @description Stock downward acceleration, in world units per second squared. */
Private Const GRAVITY_DEFAULT As Single = 18!

'/** @description Stock upward speed given by a jump. */
Private Const JUMP_DEFAULT As Single = 6.4

'/** @description Stock top walking speed on a normal surface. */
Private Const WALK_DEFAULT As Single = 5.4

'/** @description Stock height below which the body is considered lost. */
Private Const KILL_DEFAULT As Single = -12!

'/** @description Half width of the body used for collision. */
Private Const BODY_RADIUS As Single = 0.34

'/** @description Height of the body from feet to crown. */
Private Const BODY_HEIGHT As Single = 1.72

'/** @description Height of the eye above the feet. */
Private Const EYE_HEIGHT As Single = 1.62

'/** @description Tallest lip the body climbs without jumping. */
Private Const STEP_HEIGHT As Single = 0.45

'/** @description The tag this game gives a checkpoint. The scene only stores the number; what it
' * means is a decision of the game, and this is where that decision is written down. */
Private Const TAG_CHECKPOINT As Long = 1

'/** @description Radians of rotation per pixel of pointer travel. */
Private Const MOUSE_SENS As Single = 0.0022

'/** @description Reach of the pickup test, measured from the centre of the body. */
Private Const PICK_RADIUS As Single = 0.95

'/** @section Module state */

Private m_scene As PScene
Private m_cam As PCamera
Private m_canvas As PCanvas
Private m_rd As PRenderer

Private m_body As PBody
Private m_anchorX As Long
Private m_anchorY As Long
Private m_lookReady As Boolean
Private m_eyeZ As Single
Private m_map As Long
Private m_levelPath As String
Private m_spawnX As Single
Private m_spawnY As Single
Private m_spawnZ As Single

Private m_lift As Long
Private m_bridge As Long
Private m_disc As Long
Private m_animT As Single
Private m_bridgeY As Single

Private m_cvMini As PCanvas
Private m_camMini As PCamera
Private m_rdMini As PRenderer

Private m_pickupCount As Long
Private m_score As Long
Private m_dirty As Boolean
Private GRAVITY As Single
Private JUMP_SPEED As Single
Private WALK_SPEED As Single
Private KILL_Z As Single
Private m_lastPolys As Long
Private m_lastCom As Double
Private m_accRender As Double
Private m_accEvents As Double
Private m_statAcc As Single
Private m_statFrames As Long
Private m_statText As String

'/** @section Entry points */

'/**
' * @brief Builds the test level and runs it on slide one until Escape is pressed or the show moves on.
' * @remarks Meant to be launched from a running slide show. The pointer is warped back to a fixed anchor every frame, which is what relative mouse needs and is only reasonable while a show owns the screen.
' */
Public Sub RunDemo()
    m_map = 0
    RunLevel
End Sub

'/**
' * @brief Builds the pool and runs it, to exercise the fluid side of the material system.
' * @remarks Same loop and same physics as RunDemo; only the level differs.
' */
Public Sub RunPool()
    m_map = 1
    RunLevel
End Sub

'/**
' * @brief Loads a level written as JSON and runs it.
' * @param path The full path to a level file.
' * @remarks The proof that a level no longer has to be code. Edit the file, run again, no VBE and no
' * recompile. On a parse error nothing is run and the message names the line.
' */
Public Sub RunFile(ByVal path As String)
    m_map = 3
    m_levelPath = path
    RunLevel
End Sub

'/**
' * @brief Builds the showcase arena and runs it with a second camera on screen.
' * @remarks The one map that draws the scene twice, into two independent canvases.
' */
Public Sub RunArena()
    m_map = 2
    RunLevel
End Sub

'/**
' * @brief Shared runner behind every map, built in or loaded.
' */
Private Sub RunLevel()
    Dim target As Slide
    Dim dt As Single
    Dim t0 As Single
    Dim lastUndo As Single
    Dim timerShp As Shape
    Dim tMark As Double

    Set target = Slide1
    Set timerShp = target.Shapes("timer")

    '/* A level may override any of these; these are what it inherits when it says nothing. */
    GRAVITY = GRAVITY_DEFAULT
    JUMP_SPEED = JUMP_DEFAULT
    WALK_SPEED = WALK_DEFAULT
    KILL_Z = KILL_DEFAULT

    Psu3D.Boot target.Shapes

    Set m_canvas = Psu3D.Canvas
    Set m_cam = Psu3D.Camera
    Set m_rd = Psu3D.Renderer
    Set m_scene = Psu3D.Scene

    m_canvas.Name = "demo"
    m_canvas.FieldOfViewDeg = 46
    m_canvas.BackVisible = True
    m_canvas.BackColor = PCore.ColorPack(160, 185, 210)
    m_canvas.EnsureBackdrop target.Shapes

    Psu3D.SetLight 0.38, 0.42, 0.82
    Psu3D.SetFog 8!, 34!, PCore.ColorPack(160, 185, 210)

    '/* The whole level costs 100 polygons with every last object on screen at once, which never
    ' * actually happens, plus nine for the HUD. Pinning the budget above that worst case means the
    ' * draw list is never cut, and a cut that never happens is a cut that can never blink. Turn
    ' * AutoBudget back on for a scene too big for any fixed number, and it will trade distant
    ' * geometry for frame rate instead. */
    m_rd.SetBudgetRange 120, 140
    m_rd.PolyBudget = 140
    m_rd.AutoBudget = False

    '/* Close the seams by stroking each polygon in its own colour instead of pushing its vertices
    ' * outwards: same single COM call, no geometric distortion on small faces. */
    m_rd.SeamFill = True
    m_rd.EdgeInflate = 0!

    If m_map = 3 Then
        If Not BuildFromFile() Then
            Psu3D.Shutdown
            Exit Sub
        End If
    ElseIf m_map = 2 Then
        BuildArena
    ElseIf m_map = 1 Then
        BuildPool
    Else
        BuildLevel
    End If

    Set m_body = Psu3D.CreateBody()
    m_body.SetSize BODY_RADIUS, BODY_HEIGHT
    m_body.StepHeight = STEP_HEIGHT
    m_body.Reach = PICK_RADIUS - BODY_RADIUS
    m_body.Gravity = GRAVITY
    m_body.JumpSpeed = JUMP_SPEED
    m_body.WalkSpeed = WALK_SPEED
    m_body.KillZ = KILL_Z

    Respawn
    CountPickups
    SetupMiniView target.Shapes

    '/* The anchor is the middle of the primary screen. Nothing is measured against the slide, so
    ' * letterboxing, DPI and windowed shows all stop being anyone's problem. */
    m_anchorX = GetSystemMetrics(SM_CXSCREEN) \ 2
    m_anchorY = GetSystemMetrics(SM_CYSCREEN) \ 2
    m_lookReady = False

    ShowCursor 0
    SetCursorPos m_anchorX, m_anchorY

    t0 = Timer
    lastUndo = t0

    Do While ActivePresentation.SlideShowWindow.View.CurrentShowPosition = 1
        dt = Timer - t0
        If dt < 0.001 Then dt = 0.001
        If dt > 0.05 Then dt = 0.05
        t0 = Timer

        If KeyDown(vbKeyEscape) Then Exit Do

        UpdateStats dt
        UpdateMovers dt
        UpdateSceneMotion dt
        tMark = PCore.Seconds()
        UpdateLook
        UpdateMove dt

        '/* Nothing moved, so the shapes already on the slide are still correct. Skipping the frame
        ' * skips every COM call in it: no delete, no AddPolyline, no fill. A player standing still
        ' * costs nothing at all, which is the single cheapest optimisation available here. */
        If m_dirty Then
            m_scene.SpinAll dt, 2.4
            UpdateEye dt
            m_cam.SetPosition m_body.X, m_body.Y, m_eyeZ

            Psu3D.BeginFrame dt
            Psu3D.RenderScene
            DrawHud
            Psu3D.EndFrame

            RenderMiniView

            m_lastPolys = m_rd.PolyCount
            m_lastCom = m_rd.ComSeconds
            m_dirty = False
        End If

        m_accRender = m_accRender + (PCore.Seconds() - tMark)

        If t0 - lastUndo >= 1! Then
            Application.StartNewUndoEntry
            lastUndo = t0
        End If

        tMark = PCore.Seconds()

        '/* NOT a clock. Writing to this WordArt is what forces PowerPoint to repaint the slide during
        ' * a show, and the repaint is how the frame becomes visible at all. Skip it, or write the same
        ' * string twice, and the picture stops updating even though the engine keeps running. It must
        ' * fire every single frame, and the value must actually differ each time, which is why the raw
        ' * clock is appended to the readout. Do not throttle this: it is the refresh pump. */
        timerShp.TextEffect.Text = m_statText & "  " & Format$(Timer, "0.00")

        '/* The pump and the DoEvents are timed together as one bucket, because they are one event:
        ' * the write asks PowerPoint to repaint and the DoEvents is where it gets to. If the frame
        ' * turns out to be mostly this, no amount of shaving the pipeline will move it, and the lever
        ' * becomes how much the repaint has to draw. */
        DoEvents
        m_accEvents = m_accEvents + (PCore.Seconds() - tMark)
    Loop

    ShowCursor 1
    ReleaseMiniView
    Psu3D.Shutdown

    Set m_body = Nothing

    Set m_rd = Nothing
    Set m_scene = Nothing
    Set m_cam = Nothing
    Set m_canvas = Nothing
End Sub

'/** @section Level */

'/**
' * @brief Fills the scene with the test level.
' * @description Every material behaviour the engine ships with appears once: solid ground, a climbable stair, a ramp, a one-way platform reachable only from below, an ice patch, a transparent glass wall, a pillar you walk straight through, a hazard pool and a handful of pickups.
' */
Private Sub BuildLevel()
    Dim grass As Long
    Dim stone As Long
    Dim brick As Long
    Dim metal As Long
    Dim ice As Long
    Dim glass As Long
    Dim oneWay As Long
    Dim ghost As Long
    Dim lava As Long
    Dim rubber As Long
    Dim pick As Long
    Dim i As Long
    Dim z As Single
    Dim bx As Single
    Dim by As Single

    grass = PMaterials.IdOf("grass")
    stone = PMaterials.IdOf("stone")
    brick = PMaterials.IdOf("brick")
    metal = PMaterials.IdOf("metal")
    ice = PMaterials.IdOf("ice")
    glass = PMaterials.IdOf("glass")
    oneWay = PMaterials.IdOf("platform")
    ghost = PMaterials.IdOf("decor")
    lava = PMaterials.IdOf("lava")
    rubber = PMaterials.IdOf("rubber")
    pick = PMaterials.IdOf("pickup")

    m_scene.Clear
    m_scene.Reserve 128

    m_spawnX = 0!
    m_spawnY = -14!
    m_spawnZ = 1!

    '/* Ground and outer walls. The side walls stop where the end walls begin: overlapping boxes have no
    ' * correct painting order, so two of them sharing a corner is what makes that corner flicker. */
    m_scene.AddBox grass, -18!, -18!, 18!, 18!, 0!, 1!
    m_scene.AddBox brick, -18!, -18!, 18!, -17!, 3!, 3!
    m_scene.AddBox brick, -18!, 17!, 18!, 18!, 3!, 3!
    m_scene.AddBox brick, -18!, -17!, -17!, 17!, 3!, 3!
    m_scene.AddBox brick, 17!, -17!, 18!, 17!, 3!, 3!

    '/* Stair of four steps climbing to a landing. Each step is 0.4 tall, just under STEP_HEIGHT,
    ' * so the body walks up instead of being stopped by every single one. */
    For i = 0 To 3
        z = 0.4 * (i + 1)
        m_scene.AddBox stone, -4! + i * 1.2, -12!, -2.8 + i * 1.2, -8!, z, z
    Next i
    m_scene.AddBox metal, 0.8, -13!, 5.6, -7!, 1.6, 1.6

    '/* Ramp carrying you back down to the ground, flush with the landing. */
    m_scene.AddRamp stone, 5.6, -11.5, 10.6, -8.5, 1.6, 0.2, 0.4, paX

    '/* One-way platform, low enough that a jump clears it: you pass up through it and land on top. */
    m_scene.AddBox oneWay, -3!, 4!, 3!, 8!, 1!, 0.25

    '/* Ice patch, low friction. Thickness equals its height, so the slab rests exactly on the ground
    ' * instead of sinking into it: that leaves a real separating plane between the two, and a pair
    ' * decided by an exact plane can never be reordered by a cycle. */
    m_scene.AddBox ice, -13!, 2!, -6!, 9!, 0.14, 0.14

    '/* Glass wall, solid but see-through. */
    m_scene.AddBox glass, 7!, 1!, 7.4, 9!, 3!, 3!

    '/* Pillar with a ghost material: it is drawn, but you walk right through it. */
    m_scene.AddBox ghost, -1!, -4!, 1!, -2!, 4.5, 4.5

    '/* Hazard pool: standing on it sends you back to the spawn. */
    m_scene.AddBox lava, 10!, 5!, 15!, 10!, 0.12, 0.12

    '/* Obby: a zigzag of jump platforms climbing to a lookout. The rise per platform stays under the
    ' * jump apex and the diagonal step under the jump range, so every hop is reachable. */
    For i = 0 To 7
        z = 0.6 + i * 0.5

        If i Mod 2 = 0 Then bx = -12.6 Else bx = -10.4
        by = -15! + i * 1.7

        If i Mod 4 = 3 Then
            m_scene.AddBox metal, bx - 1!, by - 1!, bx + 1!, by + 1!, z, 0.3
        Else
            m_scene.AddBox stone, bx - 1!, by - 1!, bx + 1!, by + 1!, z, 0.3
        End If
    Next i

    '/* Lookout at the top of the climb. */
    m_scene.AddBox brick, -12.6, -1.8, -9.4, 1.4, 4.6, 0.4

    '/* Trampoline: a rubber pad tuned to throw the body higher than its own jump. */
    PMaterials.ByName("rubber").Bounce = 1.35
    m_scene.AddBox rubber, 2!, -4!, 4.4, -1.6, 0.35, 0.35
    m_scene.AddBox metal, 1.6, 1.2, 5!, 4.6, 2.6, 0.35

    '/* Pickups, floating clear of whatever carries them. */

    AddPickup pick, -6!, -10!, 0.9
    AddPickup pick, 0!, -10!, 2.2
    AddPickup pick, 3.5, -10!, 2.2
    AddPickup pick, 8!, -10!, 1.5
    AddPickup pick, -9.5, 5.5, 0.9
    AddPickup pick, 0!, 6!, 1.7
    AddPickup pick, -12.6, -11.6, 2.3
    AddPickup pick, -11!, -0.2, 5.3
    AddPickup pick, 3.3, 2.9, 3.4

    m_score = 0
End Sub

'/**
' * @brief Fills the scene with a pool, to test the fluid material.
' * @description The ground is four slabs with a hole in the middle rather than one plate, because a
' * pit has to be an absence of floor, not a dent in it. Their inner faces are the pool walls for free.
' * The water is a trigger volume whose base sits exactly on the pool floor, which gives the painter an
' * exact separating plane between the two and keeps the floor visible through the water.
' */
Private Sub BuildPool()
    Dim grass As Long
    Dim stone As Long
    Dim brick As Long
    Dim metal As Long
    Dim water As Long
    Dim pick As Long
    Dim i As Long
    Dim z As Single

    grass = PMaterials.IdOf("grass")
    stone = PMaterials.IdOf("stone")
    brick = PMaterials.IdOf("brick")
    metal = PMaterials.IdOf("metal")
    water = PMaterials.IdOf("water")
    pick = PMaterials.IdOf("pickup")

    m_scene.Clear
    m_scene.Reserve 64

    m_spawnX = 0!
    m_spawnY = -10!
    m_spawnZ = 0.2

    '/* Deck: four slabs around a six by six hole, 2.5 deep. */
    m_scene.AddBox grass, -14!, 6!, 14!, 14!, 0!, 2.5
    m_scene.AddBox grass, -14!, -14!, 14!, -6!, 0!, 2.5
    m_scene.AddBox grass, -14!, -6!, -6!, 6!, 0!, 2.5
    m_scene.AddBox grass, 6!, -6!, 14!, 6!, 0!, 2.5

    '/* Pool floor, and the water resting exactly on it. */
    m_scene.AddBox stone, -6!, -6!, 6!, 6!, -2.5, 0.5
    m_scene.AddBox water, -6!, -6!, 6!, 6!, -0.2, 2.3

    '/* Steps out of the shallow end. */
    For i = 0 To 2
        z = -1.9 + i * 0.62
        m_scene.AddBox stone, -6!, -1.2, -4.6 + i * 0.5, 1.2, z, 0.62
    Next i

    '/* Diving board over the deep end. */
    m_scene.AddBox metal, 7!, -1!, 8.4, 1!, 2.2, 2.2
    m_scene.AddBox metal, 2.2, -0.7, 7.2, 0.7, 2.4, 0.22

    '/* Outer wall. */
    m_scene.AddBox brick, -15!, -15!, 15!, -14!, 1.4, 1.4
    m_scene.AddBox brick, -15!, 14!, 15!, 15!, 1.4, 1.4
    m_scene.AddBox brick, -15!, -14!, -14!, 14!, 1.4, 1.4
    m_scene.AddBox brick, 14!, -14!, 15!, 14!, 1.4, 1.4

    m_pickupCount = 0

    '/* Two floating on the surface, three sunk, one on the board. */
    AddPickup pick, -3!, 3!, 0.1
    AddPickup pick, 3!, -3!, 0.1
    AddPickup pick, 0!, 0!, -1.6
    AddPickup pick, -4!, -4!, -2!
    AddPickup pick, 4.5, 4.5, -2!
    AddPickup pick, 4.5, 0!, 2.9

    m_score = 0
End Sub

'/**
' * @brief Builds the scene from the level file the run was started with.
' * @return True when the file parsed; False leaves the message on the slide clock.
' */
Private Function BuildFromFile() As Boolean
    Dim shp As Shape

    m_score = 0
    m_animT = 0!
    m_bridgeY = 0!
    m_lift = P_INVALID_ID
    m_bridge = P_INVALID_ID
    m_disc = P_INVALID_ID

    If Not PLevel.ParseFile(m_levelPath, m_scene) Then
        Set shp = Slide1.Shapes("timer")
        shp.TextEffect.Text = "level: " & PLevel.Error
        Exit Function
    End If

    m_spawnX = PLevel.SpawnX
    m_spawnY = PLevel.SpawnY
    m_spawnZ = PLevel.SpawnZ

    If PLevel.Gravity > 0! Then GRAVITY = PLevel.Gravity
    If PLevel.Jump > 0! Then JUMP_SPEED = PLevel.Jump
    If PLevel.Walk > 0! Then WALK_SPEED = PLevel.Walk
    If PLevel.HasKillZ Then KILL_Z = PLevel.KillZ

    If PLevel.BudgetMax > 0 Then
        m_rd.SetBudgetRange PLevel.BudgetMin, PLevel.BudgetMax
        m_rd.PolyBudget = PLevel.BudgetMax
    End If

    BuildFromFile = True
End Function

'/**
' * @brief Fills the scene with the showcase arena.
' * @description Built to exercise the parts of the engine the other two maps never touch: a box that
' * spins on its own axis, a lift and a bridge that carry the body, a billboard that always faces you,
' * an invisible collider, and a second camera rendering the same scene into its own canvas.
' */
Private Sub BuildArena()
    Dim grass As Long
    Dim stone As Long
    Dim brick As Long
    Dim metal As Long
    Dim glass As Long
    Dim rubber As Long
    Dim lava As Long
    Dim water As Long
    Dim ice As Long
    Dim clip As Long
    Dim decor As Long
    Dim pick As Long
    Dim i As Long

    grass = PMaterials.IdOf("grass")
    stone = PMaterials.IdOf("stone")
    brick = PMaterials.IdOf("brick")
    metal = PMaterials.IdOf("metal")
    glass = PMaterials.IdOf("glass")
    rubber = PMaterials.IdOf("rubber")
    lava = PMaterials.IdOf("lava")
    water = PMaterials.IdOf("water")
    ice = PMaterials.IdOf("ice")
    clip = PMaterials.IdOf("clip")
    decor = PMaterials.IdOf("decor")
    pick = PMaterials.IdOf("pickup")

    PMaterials.ByName("rubber").Bounce = 1.4

    m_scene.Clear
    m_scene.Reserve 64

    m_spawnX = 0!
    m_spawnY = -12!
    m_spawnZ = 0.2
    m_animT = 0!
    m_bridgeY = 0!

    '/* Arena floor and its wall. */
    m_scene.AddBox grass, -15!, -15!, 15!, 15!, 0!, 1!
    m_scene.AddBox brick, -15!, -15!, 15!, -14!, 2.2, 2.2
    m_scene.AddBox brick, -15!, 14!, 15!, 15!, 2.2, 2.2
    m_scene.AddBox brick, -15!, -14!, -14!, 14!, 2.2, 2.2
    m_scene.AddBox brick, 14!, -14!, 15!, 14!, 2.2, 2.2

    '/* Lift and the ledge it serves. */
    m_lift = m_scene.AddBox(metal, 10!, -2!, 13!, 2!, 3.6, 0.35)
    m_scene.AddBox stone, 13!, -3.5, 14!, 3.5, 6.9, 6.9
    m_scene.AddBox stone, 8!, -3.5, 13!, -2.6, 6.9, 0.4
    '/* An invisible rail so the top ledge cannot be walked off blind. */
    m_scene.AddBox clip, 7.6, -3.5, 8!, 3.5, 8.6, 1.7

    '/* Two towers and the bridge that slides between them. */
    m_scene.AddBox brick, -4!, -9!, -1!, -6!, 3.2, 3.2
    m_scene.AddBox brick, -4!, 6!, -1!, 9!, 3.2, 3.2
    m_bridge = m_scene.AddBox(metal, -3.4, -1.6, -1.6, 1.6, 3.2, 0.3)

    '/* A slab spinning on its own axis. */
    m_disc = m_scene.AddRotatedBox(stone, 5!, 7!, 3.4, 0.9, 1.1, 0.4, 0!)
    m_scene.SetSpin m_disc, 1.1

    '/* Bounce pad tuned to reach the tower tops. */
    m_scene.AddBox rubber, -1!, -13!, 1.6, -10.4, 0.4, 0.4

    '/* Pond: a trigger volume resting on the floor, walled by a low rim. */
    m_scene.AddBox water, -13!, -6!, -7!, 0!, 1.1, 1.1
    m_scene.AddBox stone, -13.6, -6.6, -6.4, -6!, 1.3, 1.3
    m_scene.AddBox stone, -13.6, 0!, -6.4, 0.6, 1.3, 1.3
    m_scene.AddBox stone, -13.6, -6!, -13!, 0!, 1.3, 1.3
    m_scene.AddBox stone, -7!, -6!, -6.4, 0!, 1.3, 1.3

    '/* Ice run. */
    m_scene.AddBox ice, -13!, 3!, -5!, 9!, 0.12, 0.12

    '/* Lava pool crossed by two lines of stepping stones. */
    m_scene.AddBox lava, 0.5, -9.5, 9.5, -3.5, 0.1, 0.1

    For i = 0 To 4
        m_scene.AddBox stone, 1.4 + i * 1.6, -8.4, 3! + i * 1.6, -7.2, 0.62, 0.3
        m_scene.AddBox stone, 1.4 + i * 1.6, -5.8, 3! + i * 1.6, -4.6, 0.62, 0.3
    Next i

    '/* Glass screen, to see the arena through it. */
    m_scene.AddBox glass, 9!, 5!, 9.4, 12!, 3.4, 3.4

    '/* Billboards: flat panels that always turn to the camera. */
    m_scene.AddBillboard decor, 0!, 12!, 2.6, 3.2, 1.8
    m_scene.AddBillboard decor, -12!, 12!, 2.2, 2.2, 1.4
    m_scene.AddBillboard decor, 12!, -12!, 2.2, 2.2, 1.4

    m_pickupCount = 0

    AddPickup pick, 0!, -6!, 0.9
    AddPickup pick, -2.5, 0!, 4!
    AddPickup pick, 11.5, 0!, 7.8
    AddPickup pick, 5!, 7!, 2!
    AddPickup pick, -10!, -3!, 0.5
    AddPickup pick, -9!, 6!, 0.9
    AddPickup pick, 11!, 9!, 0.9

    m_score = 0
End Sub

'/**
' * @brief Adds one spinning pickup and remembers its id.
' * @param matId The pickup material.
' * @param X The world X coordinate.
' * @param Y The world Y coordinate.
' * @param Z The world Z coordinate.
' */
Private Sub AddPickup(ByVal matId As Long, ByVal X As Single, ByVal Y As Single, ByVal Z As Single)
    m_scene.AddSpinner matId, X, Y, Z, 0.32
End Sub

'/**
' * @brief Puts the body back at the start of the level.
' */
Private Sub Respawn()
    m_body.SetPosition m_spawnX, m_spawnY, m_spawnZ
    m_eyeZ = m_spawnZ + EYE_HEIGHT
    If m_map = 3 Then
        m_cam.SetAngles PLevel.SpawnYaw, 0!
    Else
        m_cam.SetAngles PCore.P_HALF_PI, 0!
    End If
    m_dirty = True
End Sub

'/** @section Frame steps */

'/**
' * @brief Turns the camera from the pointer travel since the last frame and warps the pointer back.
' * @description Relative mouse without capturing the device: read where the pointer is, measure how far
' * it went from a fixed anchor, then put it back on the anchor. Because the same anchor is used to read
' * and to warp, its exact position never matters, only that it stays put, so the centre of the primary
' * screen does the job without any conversion between screen pixels and slide points.
' * @remarks The first frame only warps. Measuring on that one would turn the camera by however far the
' * pointer happened to be from the anchor when the show started, which is a spin, not a look.
' */
Private Sub UpdateLook()
    Dim p As POINTAPI
    Dim dx As Single
    Dim dy As Single

    GetCursorPos p

    If m_lookReady Then
        dx = CSng(p.X - m_anchorX)
        dy = CSng(p.Y - m_anchorY)

        If dx <> 0! Or dy <> 0! Then
            m_cam.AddAngles dx * MOUSE_SENS, -dy * MOUSE_SENS
            m_dirty = True
        End If
    Else
        m_lookReady = True
    End If

    SetCursorPos m_anchorX, m_anchorY
End Sub

'/**
' * @brief Folds this frame into the running readout of frame rate and polygon count.
' * @param dt The elapsed frame time in seconds.
' * @description Averaged over a quarter of a second and only then written out, because the text layer
' * only pays a COM call when the string actually changes, and a number that changes every frame would
' * pay it every frame. Toggle PRenderer.DryRun and watch this line to see what the slide itself costs.
' */
Private Sub UpdateStats(ByVal dt As Single)
    m_statAcc = m_statAcc + dt
    m_statFrames = m_statFrames + 1

    If m_statAcc < 0.25 Then Exit Sub

    m_statText = Format$(m_statFrames / m_statAcc, "0") & " fps " & _
                 Format$(m_statAcc / m_statFrames * 1000!, "0") & "ms" & _
                 "  |  render " & Format$(m_accRender / m_statFrames * 1000!, "0") & _
                 "  com " & Format$(m_lastCom * 1000!, "0") & _
                 "  doevents " & Format$(m_accEvents / m_statFrames * 1000!, "0") & _
                 "  |  " & m_lastPolys & " pol"

    m_accRender = 0#
    m_accEvents = 0#

    m_statAcc = 0!
    m_statFrames = 0
End Sub

'/**
' * @brief Slides the eye towards its target height instead of snapping to it.
' * @param dt The elapsed frame time in seconds.
' * @remarks A body that steps onto a stair gains its whole height in one frame. Following that instantly reads as the camera popping; easing it while grounded, and tracking it exactly while airborne, keeps a jump crisp and a staircase smooth.
' */
Private Sub UpdateEye(ByVal dt As Single)
    Dim target As Single

    target = m_body.Z + EYE_HEIGHT

    If m_body.OnGround Then
        m_eyeZ = PCore.Approach(m_eyeZ, target, 16!, dt)
    Else
        m_eyeZ = target
    End If
End Sub

'/**
' * @brief Turns the keyboard into a direction, hands the frame to the body and reads back what it found.
' * @param dt The elapsed frame time in seconds.
' * @description Everything that was here, the friction and the push-out and the landing, lives in PBody
' * now, where any level can have it. What is left is the part that is genuinely this demo: which keys
' * this game uses, and what a coin and a checkpoint mean to it. The body reports the triggers it touched
' * and refuses to guess what they were for, which is why a coin can be a coin here and a mine somewhere
' * else without the body knowing either word.
' */
Private Sub UpdateMove(ByVal dt As Single)
    Dim fwd As Single
    Dim strafe As Single
    Dim wishX As Single
    Dim wishY As Single
    Dim wasX As Single
    Dim wasY As Single
    Dim wasZ As Single
    Dim k As Long
    Dim idx As Long
    Dim spanBottom As Single
    Dim spanTop As Single

    wasX = m_body.X
    wasY = m_body.Y
    wasZ = m_body.Z

    If KeyDown(vbKeyW) Then fwd = fwd + 1!
    If KeyDown(vbKeyS) Then fwd = fwd - 1!
    If KeyDown(vbKeyD) Then strafe = strafe + 1!
    If KeyDown(vbKeyA) Then strafe = strafe - 1!

    wishX = m_cam.ForwardX * fwd + m_cam.RightX * strafe
    wishY = m_cam.ForwardY * fwd + m_cam.RightY * strafe

    m_body.Advance dt, wishX, wishY, KeyDown(vbKeySpace), KeyDown(vbKeyShift)

    '/* Turned by the floor. Without this the disc spins you around its centre while you keep facing
    ' * the way you were, which is the one thing standing on a turntable never does. */
    If m_body.CarryYaw <> 0! Then
        m_cam.AddAngles m_body.CarryYaw, 0!
        m_dirty = True
    End If

    For k = 0 To m_body.TouchCount - 1
        idx = m_body.TouchAt(k)
        m_scene.GetSpanZ idx, spanBottom, spanTop

        If m_scene.TagOf(idx) = TAG_CHECKPOINT Then
            '/* Reached by standing on it, not by brushing past it. */
            If m_body.Z >= spanTop - 0.25 And m_body.Z <= spanTop + 0.7 Then
                m_spawnX = CentreX(idx)
                m_spawnY = CentreY(idx)
                m_spawnZ = spanTop + 0.15
            End If
        Else
            '/* Anything else you can touch is a coin. A rule, not a list: levels arriving from a file
            ' * place their own coins, and a list built while placing them would come back empty. */
            m_scene.SetActive idx, False
            m_score = m_score + 1
            m_dirty = True
        End If
    Next k

    If m_body.FellOut Then Respawn

    If m_body.OnGround Then
        If PMaterials.ById(m_body.GroundMaterial).DamagePerSecond > 0! Then Respawn
    End If

    '/* A body that ended the frame where it started changes nothing on screen. */
    If m_body.X <> wasX Or m_body.Y <> wasY Or m_body.Z <> wasZ Then m_dirty = True
    If m_body.Submersion > 0! Then m_dirty = True
    If m_eyeZ <> m_body.Z + EYE_HEIGHT Then m_dirty = True
End Sub

'/** @section Second view */

'/**
' * @brief Builds the drone camera and the small canvas it draws into.
' * @param target The shape collection of the slide.
' * @description The whole point of a canvas being a rectangle rather than the screen. Two canvases, two
' * cameras and two renderers share one scene and one slide; each renderer names its shapes from its own
' * prefix, so neither one ever deletes the other one's frame. Its budget is set by hand and low,
' * because a second view is a second pass over the same geometry.
' */
Private Sub SetupMiniView(ByVal target As Shapes)
    If m_map <> 2 Then Exit Sub

    Set m_cvMini = Psu3D.CreateCanvas(686!, 24!, 250!, 150!, "mini")
    m_cvMini.FieldOfViewDeg = 52
    m_cvMini.BackVisible = True
    m_cvMini.BackColor = PCore.ColorPack(22, 30, 40)
    m_cvMini.EnsureBackdrop target

    Set m_camMini = Psu3D.CreateCamera()
    Set m_rdMini = Psu3D.CreateRenderer(target, m_cvMini, m_camMini, "p3dm_")

    m_rdMini.Purge
    m_rdMini.SetBudgetRange 40, 48
    m_rdMini.PolyBudget = 48
    m_rdMini.AutoBudget = False
End Sub

'/**
' * @brief Draws the scene a second time, from a camera trailing above the body.
' */
Private Sub RenderMiniView()
    If m_rdMini Is Nothing Then Exit Sub

    m_camMini.SetPosition m_body.X - m_cam.ForwardX * 9!, m_body.Y - m_cam.ForwardY * 9!, m_body.Z + 10!
    m_camMini.LookAt m_body.X, m_body.Y, m_body.Z + 1!

    '/* The main frame was drawn after this backdrop was created, so it is sitting on top of it. */
    m_cvMini.LiftBackdrop

    m_rdMini.BeginFrame
    m_scene.Render m_rdMini
    m_rdMini.EndFrame
End Sub

'/**
' * @brief Takes the second view off the slide.
' */
Private Sub ReleaseMiniView()
    If m_rdMini Is Nothing Then Exit Sub

    m_rdMini.ClearShapes
    m_rdMini.EndFrame
    m_rdMini.Purge
    m_cvMini.RemoveBackdrop m_rdMini.Target

    Set m_rdMini = Nothing
    Set m_camMini = Nothing
    Set m_cvMini = Nothing
End Sub

'/**
' * @brief Advances the moving parts of the arena and carries whatever is standing on them.
' * @param dt The elapsed frame time in seconds.
' * @description A platform that moves under a body without taking the body with it is a platform the
' * body slides off. The engine deliberately has no opinion about that, since carrying is a gameplay
' * decision, so the rule lives here: whatever moved the platform this frame is added to the body too,
' * as long as the body was standing on that exact object when the frame began.
' */
Private Sub UpdateMovers(ByVal dt As Single)
    Dim liftZ As Single
    Dim liftWas As Single
    Dim bridgeY As Single
    Dim b1 As Single

    If m_map <> 2 Then Exit Sub

    m_animT = m_animT + dt
    m_dirty = True

    '/* Lift: a slab that rides up and down its shaft. */
    m_scene.GetSpanZ m_lift, b1, liftWas
    liftZ = 3.6 + Sin(m_animT * 0.7) * 3.2
    m_scene.SetTopZ m_lift, liftZ
    If m_body.GroundObject = m_lift Then m_body.Nudge 0!, 0!, liftZ - liftWas

    '/* Bridge: a slab that slides across the gap. */
    bridgeY = Sin(m_animT * 0.55) * 6!
    m_scene.MoveBy m_bridge, 0!, bridgeY - m_bridgeY, 0!
    If m_body.GroundObject = m_bridge Then m_body.Nudge 0!, bridgeY - m_bridgeY, 0!
    m_bridgeY = bridgeY

    '/* The disc turns through the motion system rather than by hand, which is what lets it report how
    ' * far it moved and carry whoever is standing on it. */
End Sub

'/**
' * @brief Counts how many pickups the level holds, whoever placed them.
' */
Private Sub CountPickups()
    Dim i As Long
    Dim mat As PMaterial

    m_pickupCount = 0

    For i = 0 To m_scene.Count - 1
        Set mat = PMaterials.ById(m_scene.MaterialOf(i))

        If mat.IsTrigger And Not mat.IsFluid Then
            If m_scene.TagOf(i) <> TAG_CHECKPOINT Then m_pickupCount = m_pickupCount + 1
        End If
    Next i
End Sub

'/**
' * @brief Reads the centre X of an object.
' * @param idx The object id.
' * @return The world X coordinate.
' */
Private Function CentreX(ByVal idx As Long) As Single
    Dim x1 As Single, y1 As Single, x2 As Single, y2 As Single

    m_scene.GetBoundsXY idx, x1, y1, x2, y2
    CentreX = (x1 + x2) * 0.5!
End Function

'/**
' * @brief Reads the centre Y of an object.
' * @param idx The object id.
' * @return The world Y coordinate.
' */
Private Function CentreY(ByVal idx As Long) As Single
    Dim x1 As Single, y1 As Single, x2 As Single, y2 As Single

    m_scene.GetBoundsXY idx, x1, y1, x2, y2
    CentreY = (y1 + y2) * 0.5!
End Function

'/**
' * @brief Advances every object that moves on its own.
' * @param dt The elapsed frame time in seconds.
' * @remarks This has to run before the body does. A platform that has already moved is one the body can
' * be carried by; a platform that moves afterwards is one the body spends a frame standing beside.
' */
Private Sub UpdateSceneMotion(ByVal dt As Single)
    If m_scene.UpdateMotion(dt) > 0 Then m_dirty = True
End Sub

'/**
' * @brief Draws the pickup counter as one tick per coin, in canvas space.
' * @remarks Straight through DrawPolygon2D, so the HUD shares the pool, the double buffering and the
' * frame skip with the geometry. Nothing here touches a text shape.
' */
Private Sub DrawHud()
    Dim i As Long

    For i = 0 To m_score - 1
        FillRect m_canvas.X + 16! + i * 14!, m_canvas.Y + 16!, _
                 m_canvas.X + 27! + i * 14!, m_canvas.Y + 27!, PCore.ColorPack(250, 205, 60)
    Next i
End Sub

'/**
' * @brief Paints one axis-aligned rectangle in slide coordinates.
' * @param x1 The left edge in points.
' * @param y1 The top edge in points.
' * @param x2 The right edge in points.
' * @param y2 The bottom edge in points.
' * @param col The fill colour.
' */
Private Sub FillRect(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal col As Long)
    Dim ax(0 To 3) As Single
    Dim ay(0 To 3) As Single

    ax(0) = x1: ay(0) = y1
    ax(1) = x2: ay(1) = y1
    ax(2) = x2: ay(2) = y2
    ax(3) = x1: ay(3) = y2

    m_rd.DrawPolygon2D ax, ay, 4, col
End Sub

'/** @section Input */

'/**
' * @brief Reports whether a key is held down right now.
' * @param vk The virtual key code, such as vbKeyW.
' * @return True while the key is pressed.
' */
Private Function KeyDown(ByVal vk As Long) As Boolean
    KeyDown = (GetAsyncKeyState(vk) < 0)
End Function

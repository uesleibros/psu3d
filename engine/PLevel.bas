Attribute VB_Name = "PLevel"
'/**
' * PLevel - Psu3D Level Reader
' * @description Turns a level written as JSON into a built PScene. A level stops being code and becomes data: text anyone can edit without opening the VBE, without a compile step between an idea and seeing it, and readable by any tool that reads JSON.
' * @author UesleiDev
' * @version 1.0
' * @remarks Parsing is done by the JSON class; this module only knows the schema, in both directions: Parse reads a document into a scene and Save writes a scene back out as one. Every field is optional and every missing field has a stated default, so a half written level still loads and shows you what you did write instead of refusing the whole file.
' */

Option Explicit
Option Private Module

'/** @section Module state */

Private m_scene As PScene
Private m_error As String

Private m_spawnX As Single
Private m_spawnY As Single
Private m_spawnZ As Single
Private m_spawnYaw As Single
Private m_budgetMin As Long
Private m_budgetMax As Long
Private m_objects As Long
Private m_gravity As Single
Private m_jump As Single
Private m_walk As Single
Private m_killZ As Single
Private m_hasKillZ As Boolean

'/** @description Most copies one entry may stand for. Past this it is a typo, not a staircase. */
Private Const MAX_REPEAT As Long = 100000

'/** @description Widest polygon budget a level may ask the renderer for. */
Private Const MAX_BUDGET As Long = 1000000

'/** @section Entry points */

'/**
' * @brief Builds a level from JSON text.
' * @param source The level document.
' * @param target The scene to build into; it is cleared first.
' * @return True when the document held a usable level.
' * @remarks On failure Error says what was wrong. A document that parses but describes nothing is a failure too, since a silently empty world is the least useful answer available.
' */
Public Function Parse(ByVal source As String, ByVal target As PScene) As Boolean
    Dim doc As JSON

    Set m_scene = target
    m_error = vbNullString
    m_objects = 0

    m_spawnX = 0!
    m_spawnY = 0!
    m_spawnZ = 1!
    m_spawnYaw = 0!
    m_budgetMin = 0
    m_budgetMax = 0
    m_gravity = 0!
    m_jump = 0!
    m_walk = 0!
    m_killZ = 0!
    m_hasKillZ = False

    If target Is Nothing Then
        m_error = "no scene to build into"
        Exit Function
    End If

    If Len(Trim$(source)) = 0 Then
        m_error = "the level is empty"
        Exit Function
    End If

    Set doc = JSON.Parse(source)

    If doc Is Nothing Then
        m_error = "the level is not valid JSON"
        Exit Function
    End If

    If Not doc.IsObject Then
        m_error = "the level must be a JSON object, with level, materials and objects inside it"
        Exit Function
    End If

    target.Clear

    ReadLevel doc
    ReadMaterials doc
    ReadObjects doc

    If Len(m_error) > 0 Then Exit Function

    If m_objects = 0 Then
        m_error = "the level describes no objects"
        Exit Function
    End If

    Parse = True
End Function

'/**
' * @brief Builds a level from a JSON file on disk.
' * @param path The full path to the level file.
' * @param target The scene to build into.
' * @return True when the file was found and held a usable level.
' * @remarks Existence is checked before opening rather than trapped, in keeping with the rest of the library.
' */
Public Function ParseFile(ByVal path As String, ByVal target As PScene) As Boolean
    Dim f As Integer
    Dim body As String

    m_error = vbNullString

    If Len(path) = 0 Then
        m_error = "no path given"
        Exit Function
    End If

    If Len(Dir$(path)) = 0 Then
        m_error = "file not found: " & path
        Exit Function
    End If

    f = FreeFile
    Open path For Input As #f
    If LOF(f) > 0 Then body = Input$(LOF(f), #f)
    Close #f

    ParseFile = Parse(body, target)
End Function

'/** @section Authoring */

'/**
' * @brief Sets where the player starts.
' * @param X The spawn X coordinate.
' * @param Y The spawn Y coordinate.
' * @param Z The spawn Z coordinate.
' * @param yawDeg Which way they face, in degrees.
' */
Public Sub SetSpawn(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, Optional ByVal yawDeg As Single = 0!)
    m_spawnX = X
    m_spawnY = Y
    m_spawnZ = Z
    m_spawnYaw = yawDeg * P_DEG2RAD
End Sub

'/**
' * @brief Sets the polygon budget the level asks for.
' * @param lo The floor.
' * @param hi The ceiling; zero on both leaves the renderer to decide.
' */
Public Sub SetBudget(ByVal lo As Long, ByVal hi As Long)
    m_budgetMin = lo
    m_budgetMax = hi
End Sub

'/**
' * @brief Sets the movement rules the level asks for.
' * @param gravityValue The downward acceleration; zero keeps the body default.
' * @param jumpValue The jump speed; zero keeps the body default.
' * @param walkValue The walking speed; zero keeps the body default.
' */
Public Sub SetRules(ByVal gravityValue As Single, ByVal jumpValue As Single, ByVal walkValue As Single)
    m_gravity = gravityValue
    m_jump = jumpValue
    m_walk = walkValue
End Sub

'/**
' * @brief Sets the height below which the player is lost.
' * @param Z The height in world units.
' */
Public Sub SetKillZ(ByVal Z As Single)
    m_killZ = Z
    m_hasKillZ = True
End Sub

'/** @section Results */

'/**
' * @brief Reads why the last load failed.
' * @return The message, or an empty string when it succeeded.
' */
Public Property Get Error() As String
    Error = m_error
End Property

'/**
' * @brief Reads how many objects the last level built.
' * @return The object count.
' */
Public Property Get ObjectCount() As Long
    ObjectCount = m_objects
End Property

'/**
' * @brief Reads the spawn X coordinate the level asked for.
' * @return The world X position.
' */
Public Property Get SpawnX() As Single
    SpawnX = m_spawnX
End Property

'/**
' * @brief Reads the spawn Y coordinate the level asked for.
' * @return The world Y position.
' */
Public Property Get SpawnY() As Single
    SpawnY = m_spawnY
End Property

'/**
' * @brief Reads the spawn Z coordinate the level asked for.
' * @return The world Z position.
' */
Public Property Get SpawnZ() As Single
    SpawnZ = m_spawnZ
End Property

'/**
' * @brief Reads the heading the level wants the body to start facing.
' * @return The yaw in radians.
' */
Public Property Get SpawnYaw() As Single
    SpawnYaw = m_spawnYaw
End Property

'/**
' * @brief Reads the lower polygon budget the level asked for.
' * @return The budget floor, or zero when the level did not say.
' */
Public Property Get BudgetMin() As Long
    BudgetMin = m_budgetMin
End Property

'/**
' * @brief Reads the upper polygon budget the level asked for.
' * @return The budget ceiling, or zero when the level did not say.
' */
Public Property Get BudgetMax() As Long
    BudgetMax = m_budgetMax
End Property

'/**
' * @brief Reads the downward acceleration the level asked for.
' * @return The value, or zero when the level did not say.
' */
Public Property Get Gravity() As Single
    Gravity = m_gravity
End Property

'/**
' * @brief Reads the jump speed the level asked for.
' * @return The value, or zero when the level did not say.
' */
Public Property Get Jump() As Single
    Jump = m_jump
End Property

'/**
' * @brief Reads the walking speed the level asked for.
' * @return The value, or zero when the level did not say.
' */
Public Property Get Walk() As Single
    Walk = m_walk
End Property

'/**
' * @brief Reads the height below which the level considers a body lost.
' * @return The height; meaningful only when HasKillZ is True.
' * @remarks Zero is a perfectly sensible floor for a level to choose, so absence is reported separately rather than encoded as a value.
' */
Public Property Get KillZ() As Single
    KillZ = m_killZ
End Property

'/**
' * @brief Reports whether the level named a height below which a body is lost.
' * @return True when it did.
' */
Public Property Get HasKillZ() As Boolean
    HasKillZ = m_hasKillZ
End Property

'/** @section Sections */

'/**
' * @brief Reads the level block: where the body starts and what the world looks like.
' * @param doc The root document.
' */
Private Sub ReadLevel(ByVal doc As JSON)
    Dim lv As JSON

    If Not doc.ExistsKey("level") Then Exit Sub
    Set lv = doc.NodeKey("level")
    If lv Is Nothing Then Exit Sub

    m_spawnX = VecOf(lv, "spawn", 0, 0!)
    m_spawnY = VecOf(lv, "spawn", 1, 0!)
    m_spawnZ = VecOf(lv, "spawn", 2, 1!)
    m_spawnYaw = NumOf(lv, "yaw", 0!) * P_DEG2RAD

    m_budgetMin = WholeOf(VecOf(lv, "budget", 0, 0!), 0, MAX_BUDGET)
    m_budgetMax = WholeOf(VecOf(lv, "budget", 1, 0!), 0, MAX_BUDGET)

    If lv.ExistsKey("lod") Then m_scene.LodSize = NumOf(lv, "lod", 44!)

    m_gravity = NumOf(lv, "gravity", 0!)
    m_jump = NumOf(lv, "jump", 0!)
    m_walk = NumOf(lv, "walk", 0!)

    If lv.ExistsKey("killz") Then
        m_killZ = NumOf(lv, "killz", -12!)
        m_hasKillZ = True
    End If

    If lv.ExistsKey("fog") Then
        PLighting.SetFog VecOf(lv, "fog", 0, 8!), VecOf(lv, "fog", 1, 34!), _
                         PCore.ColorFromHex(VecStrOf(lv, "fog", 2, "#A0B9D2"))
    End If

    If lv.ExistsKey("light") Then
        PLighting.SetLight VecOf(lv, "light", 0, 0.38), VecOf(lv, "light", 1, 0.42), VecOf(lv, "light", 2, 0.82)
    End If
End Sub

'/**
' * @brief Reads the material table.
' * @param doc The root document.
' * @remarks Names are the keys of the block, so a material is declared once and referred to by name everywhere after, exactly as the registry works underneath.
' */
Private Sub ReadMaterials(ByVal doc As JSON)
    Dim mats As JSON
    Dim def As JSON
    Dim mat As PMaterial
    Dim i As Long
    Dim nm As String

    If Not doc.ExistsKey("materials") Then Exit Sub
    Set mats = doc.NodeKey("materials")
    If mats Is Nothing Then Exit Sub
    If Not mats.IsObject Then
        m_error = "materials must be an object of name to definition"
        Exit Sub
    End If

    For i = 0 To mats.Count - 1
        nm = mats.KeyAt(i)
        Set def = mats.NodeAt(i)
        If def Is Nothing Then
            m_error = "material '" & nm & "' has no definition"
            Exit Sub
        End If

        Set mat = PMaterials.Create(nm)

        If def.ExistsKey("color") Then mat.Color = PCore.ColorFromHex(def.StringKey("color"))
        If def.ExistsKey("alpha") Then mat.Transparency = NumOf(def, "alpha", 0!)
        If def.ExistsKey("friction") Then mat.Friction = NumOf(def, "friction", 1!)
        If def.ExistsKey("bounce") Then mat.Bounce = NumOf(def, "bounce", 0!)
        If def.ExistsKey("speed") Then mat.SpeedMultiplier = NumOf(def, "speed", 1!)
        If def.ExistsKey("damage") Then mat.DamagePerSecond = NumOf(def, "damage", 0!)
        If def.ExistsKey("buoyancy") Then mat.Buoyancy = NumOf(def, "buoyancy", 0!)
        If def.ExistsKey("drag") Then mat.Drag = NumOf(def, "drag", 0!)
        If def.ExistsKey("step") Then mat.StepHeight = NumOf(def, "step", 0.38)

        If def.ExistsKey("edge") Then
            mat.EdgeVisible = True
            mat.EdgeColor = PCore.ColorFromHex(def.StringKey("edge"))
        End If

        If BoolOf(def, "unlit") Then mat.Unlit = True
        If BoolOf(def, "twosided") Then mat.TwoSided = True
        If BoolOf(def, "climbable") Then mat.Climbable = True
        If def.ExistsKey("fog") Then mat.Fogged = BoolOf(def, "fog")
        If def.ExistsKey("visible") Then mat.Visible = BoolOf(def, "visible")
        If BoolOf(def, "invisible") Then mat.Visible = False

        If def.ExistsKey("collision") Then
            Select Case LCase$(def.StringKey("collision"))
                Case "solid":   mat.Collision = pcSolid
                Case "ghost":   mat.Collision = pcGhost
                Case "oneway":  mat.Collision = pcOneWay
                Case "trigger": mat.Collision = pcTrigger
                Case Else
                    m_error = "material '" & nm & "': collision must be solid, ghost, oneway or trigger"
                    Exit Sub
            End Select
        End If
    Next i
End Sub

'/**
' * @brief Reads the object list.
' * @param doc The root document.
' */
Private Sub ReadObjects(ByVal doc As JSON)
    Dim list As JSON
    Dim i As Long

    If Not doc.ExistsKey("objects") Then Exit Sub
    Set list = doc.NodeKey("objects")
    If list Is Nothing Then Exit Sub

    If Not list.IsArray Then
        m_error = "objects must be an array"
        Exit Sub
    End If

    For i = 0 To list.Count - 1
        ReadObject list.NodeAt(i), i
        If Len(m_error) > 0 Then Exit Sub
    Next i
End Sub

'/**
' * @brief Reads one entry of the object list and adds it to the scene.
' * @param it The entry.
' * @param at Its position in the list, for error reporting.
' * @description An entry may carry a repeat count and a step block, and that is how a stair is written:
' * one entry describing the first tread plus what changes from one to the next. It keeps the document
' * declarative, where a loop with an expression in it would have quietly turned the level back into
' * code.
' */
Private Sub ReadObject(ByVal it As JSON, ByVal at As Long)
    Dim kind As String
    Dim matName As String
    Dim matId As Long
    Dim stepNode As JSON
    Dim moveNode As JSON

    Dim x1 As Single, y1 As Single, x2 As Single, y2 As Single
    Dim ax As Single, ay As Single, az As Single
    Dim hw As Single, hh As Single
    Dim topZ As Single, thick As Single
    Dim lowZ As Single, highZ As Single
    Dim ang As Single, rad As Single
    Dim sw As Single, sh As Single

    Dim dx1 As Single, dy1 As Single, dx2 As Single, dy2 As Single
    Dim dax As Single, day As Single, daz As Single
    Dim dhw As Single, dhh As Single
    Dim dTop As Single, dThick As Single
    Dim dLow As Single, dHigh As Single
    Dim dAng As Single, dRad As Single
    Dim dsw As Single, dsh As Single

    Dim rep As Long
    Dim n As Long
    Dim idx As Long
    Dim axis As PAxis

    If it Is Nothing Then
        m_error = "object " & at & " is empty"
        Exit Sub
    End If

    If Not it.ExistsKey("type") Then
        m_error = "object " & at & " has no type"
        Exit Sub
    End If

    kind = LCase$(it.StringKey("type"))
    matName = StrOf(it, "mat", "default")
    matId = PMaterials.IdOf(matName)

    If matId < 0 Then
        m_error = "object " & at & ": unknown material '" & matName & "'"
        Exit Sub
    End If

    x1 = VecOf(it, "from", 0, 0!)
    y1 = VecOf(it, "from", 1, 0!)
    x2 = VecOf(it, "to", 0, 0!)
    y2 = VecOf(it, "to", 1, 0!)
    ax = VecOf(it, "at", 0, 0!)
    ay = VecOf(it, "at", 1, 0!)
    az = VecOf(it, "at", 2, 0!)
    hw = VecOf(it, "half", 0, 1!)
    hh = VecOf(it, "half", 1, 1!)
    sw = VecOf(it, "size", 0, 1!)
    sh = VecOf(it, "size", 1, 1!)
    topZ = NumOf(it, "top", 0!)
    thick = NumOf(it, "thick", 0.5)
    lowZ = NumOf(it, "low", 0!)
    highZ = NumOf(it, "high", 1!)
    ang = NumOf(it, "angle", 0!) * P_DEG2RAD
    rad = NumOf(it, "radius", 0.32)

    '/* Read wide and clamped, never converted straight to a Long. A repeat of ten billion overflows the
    ' * conversion before any sanity check can run, and one of a million would build until the machine
    ' * gave up, with a slideshow running and no way to interrupt it. Both are typos, so both are said
    ' * out loud rather than silently obeyed or silently trimmed. */
    If NumOf(it, "repeat", 1!) > CSng(MAX_REPEAT) Then
        m_error = "object " & at & ": repeat is " & NumText(NumOf(it, "repeat", 1!)) & ", which is past the limit of " & MAX_REPEAT
        Exit Sub
    End If

    rep = WholeOf(NumOf(it, "repeat", 1!), 1, MAX_REPEAT)

    If it.ExistsKey("step") Then
        Set stepNode = it.NodeKey("step")
        dx1 = VecOf(stepNode, "from", 0, 0!)
        dy1 = VecOf(stepNode, "from", 1, 0!)
        dx2 = VecOf(stepNode, "to", 0, 0!)
        dy2 = VecOf(stepNode, "to", 1, 0!)
        dax = VecOf(stepNode, "at", 0, 0!)
        day = VecOf(stepNode, "at", 1, 0!)
        daz = VecOf(stepNode, "at", 2, 0!)
        dhw = VecOf(stepNode, "half", 0, 0!)
        dhh = VecOf(stepNode, "half", 1, 0!)
        dsw = VecOf(stepNode, "size", 0, 0!)
        dsh = VecOf(stepNode, "size", 1, 0!)
        dTop = NumOf(stepNode, "top", 0!)
        dThick = NumOf(stepNode, "thick", 0!)
        dLow = NumOf(stepNode, "low", 0!)
        dHigh = NumOf(stepNode, "high", 0!)
        dAng = NumOf(stepNode, "angle", 0!) * P_DEG2RAD
        dRad = NumOf(stepNode, "radius", 0!)
    End If

    axis = paX
    If LCase$(StrOf(it, "axis", "x")) = "y" Then axis = paY

    For n = 1 To rep
        Select Case kind
            Case "box"
                idx = m_scene.AddBox(matId, x1, y1, x2, y2, topZ, thick)

            Case "ramp"
                idx = m_scene.AddRamp(matId, x1, y1, x2, y2, lowZ, highZ, thick, axis)

            Case "rot"
                idx = m_scene.AddRotatedBox(matId, ax, ay, hw, hh, topZ, thick, ang)

            Case "bill"
                idx = m_scene.AddBillboard(matId, ax, ay, az, sw, sh)

            Case "spin"
                idx = m_scene.AddSpinner(matId, ax, ay, az, rad)

            Case Else
                m_error = "object " & at & ": unknown type '" & kind & "'"
                Exit Sub
        End Select

        m_objects = m_objects + 1

        If it.ExistsKey("move") Then
            Set moveNode = it.NodeKey("move")
            m_scene.SetMotion idx, VecOf(moveNode, "axis", 0, 0!), VecOf(moveNode, "axis", 1, 0!), _
                                   VecOf(moveNode, "axis", 2, 1!), _
                                   NumOf(moveNode, "amp", 2!), NumOf(moveNode, "speed", 0.7), _
                                   NumOf(moveNode, "phase", 0!) + (n - 1) * NumOf(moveNode, "stagger", 0!)
        End If

        If it.ExistsKey("spin") Then m_scene.SetSpin idx, NumOf(it, "spin", 1!), ang
        If it.ExistsKey("tag") Then m_scene.SetTag idx, WholeOf(NumOf(it, "tag", 0!), -2000000000, 2000000000)

        x1 = x1 + dx1: y1 = y1 + dy1
        x2 = x2 + dx2: y2 = y2 + dy2
        ax = ax + dax: ay = ay + day: az = az + daz
        hw = hw + dhw: hh = hh + dhh
        sw = sw + dsw: sh = sh + dsh
        topZ = topZ + dTop
        thick = thick + dThick
        lowZ = lowZ + dLow
        highZ = highZ + dHigh
        ang = ang + dAng
        rad = rad + dRad
    Next n
End Sub

'/** @section Writing */

'/**
' * @brief Writes a scene back out as a level document.
' * @param source The scene to describe.
' * @return The JSON text.
' * @description The round trip is the point: a level loaded, edited in place and saved is the same level
' * plus the edit. What does not survive is the shorthand, since repeat and step are a way of writing a
' * staircase, not a thing the world remembers; the fifty steps they stood for all come back, one entry
' * each. The document is longer than the one that was read and describes exactly the same world.
' */
Public Function Save(ByVal source As PScene) As String
    Dim sb As String
    Dim i As Long
    Dim first As Boolean

    If source Is Nothing Then Exit Function

    sb = "{" & vbCrLf
    sb = sb & "  ""level"": {" & vbCrLf
    sb = sb & "    ""spawn"": [" & NumText(m_spawnX) & ", " & NumText(m_spawnY) & ", " & NumText(m_spawnZ) & "]," & vbCrLf
    sb = sb & "    ""yaw"": " & NumText(m_spawnYaw * P_RAD2DEG) & "," & vbCrLf
    sb = sb & "    ""budget"": [" & m_budgetMin & ", " & m_budgetMax & "]," & vbCrLf
    sb = sb & "    ""lod"": " & NumText(source.LodSize) & "," & vbCrLf

    If m_gravity <> 0! Then sb = sb & "    ""gravity"": " & NumText(m_gravity) & "," & vbCrLf
    If m_jump <> 0! Then sb = sb & "    ""jump"": " & NumText(m_jump) & "," & vbCrLf
    If m_walk <> 0! Then sb = sb & "    ""walk"": " & NumText(m_walk) & "," & vbCrLf
    If m_hasKillZ Then sb = sb & "    ""killz"": " & NumText(m_killZ) & "," & vbCrLf

    sb = sb & "    ""fog"": [" & NumText(PLighting.FogStart) & ", " & NumText(PLighting.FogEnd) & ", """ & _
              PCore.ColorToHex(PLighting.FogColor) & """]," & vbCrLf
    sb = sb & "    ""light"": [" & NumText(PLighting.LightX) & ", " & NumText(PLighting.LightY) & ", " & _
              NumText(PLighting.LightZ) & "]" & vbCrLf
    sb = sb & "  }," & vbCrLf

    sb = sb & "  ""materials"": {" & vbCrLf
    first = True

    For i = 0 To PMaterials.Count - 1
        If Not first Then sb = sb & "," & vbCrLf
        sb = sb & WriteMaterial(PMaterials.ById(i))
        first = False
    Next i

    sb = sb & vbCrLf & "  }," & vbCrLf

    sb = sb & "  ""objects"": [" & vbCrLf
    first = True

    For i = 0 To source.Count - 1
        If source.IsActive(i) Then
            If Not first Then sb = sb & "," & vbCrLf
            sb = sb & WriteObject(source, i)
            first = False
        End If
    Next i

    sb = sb & vbCrLf & "  ]" & vbCrLf & "}" & vbCrLf

    Save = sb
End Function

'/**
' * @brief Writes a scene to a file as a level document.
' * @param path Where to write it.
' * @param source The scene to describe.
' * @return True when the file was written.
' * @remarks An existing file is replaced, with no undo, so a caller pointing this at an authored level should be sure that is what it means.
' */
Public Function SaveFile(ByVal path As String, ByVal source As PScene) As Boolean
    Dim f As Integer
    Dim body As String

    m_error = vbNullString

    If Len(path) = 0 Then
        m_error = "no path given"
        Exit Function
    End If

    body = Save(source)

    If Len(body) = 0 Then
        m_error = "nothing to save"
        Exit Function
    End If

    f = FreeFile
    Open path For Output As #f
    Print #f, body;
    Close #f

    SaveFile = True
End Function

'/**
' * @brief Writes one material as a named block.
' * @param mat The material.
' * @return The block text, indented to sit inside the materials object.
' * @remarks Only what differs from a fresh material is written. A block listing every property would be
' * accurate and unreadable, and the reader fills in exactly these defaults anyway.
' */
Private Function WriteMaterial(ByVal mat As PMaterial) As String
    Dim sb As String

    sb = "    " & JsonText(mat.Name) & ": { ""color"": """ & PCore.ColorToHex(mat.Color) & """"

    Select Case mat.Collision
        Case pcSolid:   sb = sb & ", ""collision"": ""solid"""
        Case pcGhost:   sb = sb & ", ""collision"": ""ghost"""
        Case pcOneWay:  sb = sb & ", ""collision"": ""oneway"""
        Case pcTrigger: sb = sb & ", ""collision"": ""trigger"""
    End Select

    If mat.Transparency <> 0! Then sb = sb & ", ""alpha"": " & NumText(mat.Transparency)
    If mat.Friction <> 1! Then sb = sb & ", ""friction"": " & NumText(mat.Friction)
    If mat.Bounce <> 0! Then sb = sb & ", ""bounce"": " & NumText(mat.Bounce)
    If mat.SpeedMultiplier <> 1! Then sb = sb & ", ""speed"": " & NumText(mat.SpeedMultiplier)
    If mat.DamagePerSecond <> 0! Then sb = sb & ", ""damage"": " & NumText(mat.DamagePerSecond)
    If mat.Buoyancy <> 0! Then sb = sb & ", ""buoyancy"": " & NumText(mat.Buoyancy)
    If mat.Drag <> 0! Then sb = sb & ", ""drag"": " & NumText(mat.Drag)
    If mat.StepHeight <> 0! Then sb = sb & ", ""step"": " & NumText(mat.StepHeight)
    If mat.EdgeVisible Then sb = sb & ", ""edge"": """ & PCore.ColorToHex(mat.EdgeColor) & """"
    If mat.Unlit Then sb = sb & ", ""unlit"": true"
    If mat.TwoSided Then sb = sb & ", ""twosided"": true"
    If mat.Climbable Then sb = sb & ", ""climbable"": true"
    If Not mat.Fogged Then sb = sb & ", ""fog"": false"
    If Not mat.Visible Then sb = sb & ", ""visible"": false"

    WriteMaterial = sb & " }"
End Function

'/**
' * @brief Writes one object as a list entry.
' * @param source The scene it belongs to.
' * @param idx The object id.
' * @return The entry text, indented to sit inside the objects array.
' */
Private Function WriteObject(ByVal source As PScene, ByVal idx As Long) As String
    Dim sb As String
    Dim head As String
    Dim x1 As Single, y1 As Single, x2 As Single, y2 As Single
    Dim cx As Single, cy As Single, cz As Single
    Dim sizeA As Single, sizeB As Single
    Dim mode As Long
    Dim ax As Single, ay As Single, az As Single
    Dim amp As Single, spd As Single, ph As Single

    source.GetSize idx, sizeA, sizeB
    source.GetCenter idx, cx, cy, cz

    head = ", ""mat"": " & JsonText(PMaterials.ById(source.MaterialOf(idx)).Name)

    Select Case source.KindOf(idx)
        Case pkBox
            source.GetBoundsXY idx, x1, y1, x2, y2
            sb = "    { ""type"": ""box""" & head
            sb = sb & ", ""from"": [" & NumText(x1) & ", " & NumText(y1) & "]"
            sb = sb & ", ""to"": [" & NumText(x2) & ", " & NumText(y2) & "]"
            sb = sb & ", ""top"": " & NumText(source.TopOf(idx))
            sb = sb & ", ""thick"": " & NumText(source.ThickOf(idx))

        Case pkRamp
            source.GetBoundsXY idx, x1, y1, x2, y2
            sb = "    { ""type"": ""ramp""" & head
            sb = sb & ", ""from"": [" & NumText(x1) & ", " & NumText(y1) & "]"
            sb = sb & ", ""to"": [" & NumText(x2) & ", " & NumText(y2) & "]"
            sb = sb & ", ""low"": " & NumText(source.TopOf(idx))
            sb = sb & ", ""high"": " & NumText(source.HighOf(idx))
            sb = sb & ", ""thick"": " & NumText(source.ThickOf(idx))
            If source.AxisOf(idx) = paY Then
                sb = sb & ", ""axis"": ""y"""
            Else
                sb = sb & ", ""axis"": ""x"""
            End If

        Case pkRotatedBox
            sb = "    { ""type"": ""rot""" & head
            sb = sb & ", ""at"": [" & NumText(cx) & ", " & NumText(cy) & "]"
            sb = sb & ", ""half"": [" & NumText(sizeA) & ", " & NumText(sizeB) & "]"
            sb = sb & ", ""top"": " & NumText(source.TopOf(idx))
            sb = sb & ", ""thick"": " & NumText(source.ThickOf(idx))
            sb = sb & ", ""angle"": " & NumText(source.AngleOf(idx) * P_RAD2DEG)

        Case pkBillboard
            sb = "    { ""type"": ""bill""" & head
            sb = sb & ", ""at"": [" & NumText(cx) & ", " & NumText(cy) & ", " & NumText(cz) & "]"
            sb = sb & ", ""size"": [" & NumText(sizeA) & ", " & NumText(sizeB) & "]"

        Case pkSpinner
            sb = "    { ""type"": ""spin""" & head
            sb = sb & ", ""at"": [" & NumText(cx) & ", " & NumText(cy) & ", " & NumText(cz) & "]"
            sb = sb & ", ""radius"": " & NumText(sizeA)
    End Select

    mode = source.MotionOf(idx, ax, ay, az, amp, spd, ph)

    If mode = 1 Then
        sb = sb & ", ""move"": { ""axis"": [" & NumText(ax) & ", " & NumText(ay) & ", " & NumText(az) & "]"
        sb = sb & ", ""amp"": " & NumText(amp) & ", ""speed"": " & NumText(spd) & ", ""phase"": " & NumText(ph) & " }"
    ElseIf mode = 2 Then
        sb = sb & ", ""spin"": " & NumText(spd)
    End If

    If source.TagOf(idx) <> 0 Then sb = sb & ", ""tag"": " & source.TagOf(idx)

    WriteObject = sb & " }"
End Function

'/**
' * @brief Turns a string into a quoted JSON string, escaping what has to be escaped.
' * @param value The text.
' * @return The text in quotes, ready to drop into a document.
' * @remarks A material may be named anything at all, including something with a quote or a backslash in
' * it, and writing that straight out produces a file that will not parse back. The escape is the minimum
' * JSON demands: the two structural characters, and the control codes that may not appear raw.
' */
Private Function JsonText(ByVal value As String) As String
    Dim out As String
    Dim i As Long
    Dim ch As String
    Dim code As Long

    out = Chr$(34)

    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        code = AscW(ch)

        Select Case code
            Case 34:  out = out & "\" & Chr$(34)
            Case 92:  out = out & "\\"
            Case 8:   out = out & "\b"
            Case 9:   out = out & "\t"
            Case 10:  out = out & "\n"
            Case 12:  out = out & "\f"
            Case 13:  out = out & "\r"
            Case Else
                If code < 32 Then
                    out = out & "\u" & Right$("000" & Hex$(code), 4)
                Else
                    out = out & ch
                End If
        End Select
    Next i

    JsonText = out & Chr$(34)
End Function

'/**
' * @brief Turns a number from a document into a whole number inside a stated range.
' * @param value The number as it was written.
' * @param lo The lowest value allowed.
' * @param hi The highest value allowed.
' * @return The clamped whole number.
' * @remarks Clamped as a Double and only then made a Long. Converting first is what raises an overflow
' * on a value a document is perfectly free to contain, and a level file crashing the parser is the one
' * outcome a forgiving format must never have. Named apart from PCore.ClampLong on purpose: that one
' * takes a Long and so cannot be handed the very value this exists to survive.
' */
Private Function WholeOf(ByVal value As Single, ByVal lo As Long, ByVal hi As Long) As Long
    Dim v As Double

    v = CDbl(value)

    If v < CDbl(lo) Then v = CDbl(lo)
    If v > CDbl(hi) Then v = CDbl(hi)

    WholeOf = CLng(Int(v))
End Function

'/**
' * @brief Turns a number into text a JSON reader will take back.
' * @param v The value.
' * @return The text, always with a dot for the decimal point.
' * @remarks Str is used rather than Format because Format follows the machine locale, and a level saved on a Brazilian PowerPoint would come back full of commas that no JSON reader on earth will accept. Four decimals is past the precision a Single carries and well past the precision a level needs.
' */
Private Function NumText(ByVal v As Single) As String
    Dim r As Double

    r = CDbl(v) * 10000#

    If r >= 0# Then
        r = Int(r + 0.5) / 10000#
    Else
        r = -Int(-r + 0.5) / 10000#
    End If

    NumText = Trim$(Str$(r))
End Function

'/** @section Field readers */

'/**
' * @brief Reads a number, falling back when the field is absent.
' * @param n The object to read from.
' * @param key The field name.
' * @param dflt The value to use when it is missing.
' * @return The number.
' */
Private Function NumOf(ByVal n As JSON, ByVal key As String, ByVal dflt As Single) As Single
    NumOf = dflt

    If n Is Nothing Then Exit Function
    If Not n.ExistsKey(key) Then Exit Function

    NumOf = CSng(n.NumberKey(key))
End Function

'/**
' * @brief Reads a string, falling back when the field is absent.
' * @param n The object to read from.
' * @param key The field name.
' * @param dflt The value to use when it is missing.
' * @return The string.
' */
Private Function StrOf(ByVal n As JSON, ByVal key As String, ByVal dflt As String) As String
    StrOf = dflt

    If n Is Nothing Then Exit Function
    If Not n.ExistsKey(key) Then Exit Function

    StrOf = n.StringKey(key)
End Function

'/**
' * @brief Reads a flag, treating an absent field as False.
' * @param n The object to read from.
' * @param key The field name.
' * @return The flag.
' */
Private Function BoolOf(ByVal n As JSON, ByVal key As String) As Boolean
    If n Is Nothing Then Exit Function
    If Not n.ExistsKey(key) Then Exit Function

    BoolOf = n.BoolKey(key)
End Function

'/**
' * @brief Reads one number out of an array field.
' * @param n The object holding the array.
' * @param key The field name.
' * @param idx Which entry to read.
' * @param dflt The value to use when the field or the entry is missing.
' * @return The number.
' */
Private Function VecOf(ByVal n As JSON, ByVal key As String, ByVal idx As Long, ByVal dflt As Single) As Single
    Dim a As JSON

    VecOf = dflt

    If n Is Nothing Then Exit Function
    If Not n.ExistsKey(key) Then Exit Function

    Set a = n.NodeKey(key)
    If a Is Nothing Then Exit Function
    If Not a.IsArray Then Exit Function
    If idx >= a.Count Then Exit Function

    VecOf = CSng(a.NumberAt(idx))
End Function

'/**
' * @brief Reads one string out of an array field, which is how a colour rides along with numbers.
' * @param n The object holding the array.
' * @param key The field name.
' * @param idx Which entry to read.
' * @param dflt The value to use when the field or the entry is missing.
' * @return The string.
' */
Private Function VecStrOf(ByVal n As JSON, ByVal key As String, ByVal idx As Long, ByVal dflt As String) As String
    Dim a As JSON

    VecStrOf = dflt

    If n Is Nothing Then Exit Function
    If Not n.ExistsKey(key) Then Exit Function

    Set a = n.NodeKey(key)
    If a Is Nothing Then Exit Function
    If Not a.IsArray Then Exit Function
    If idx >= a.Count Then Exit Function

    VecStrOf = a.StringAt(idx)
End Function

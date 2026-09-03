VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PBody"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/**
' * PBody - Psu3D Moving Body
' * @description Everything that moves through a scene and is stopped by it: walking, jumping, climbing stairs, sliding on ice, swimming, riding platforms and reading what it touched. The material system describes how a surface behaves; this is the part that believes it.
' * @author UesleiDev
' * @version 1.0
' * @scope Private to this VBA project: class modules carry VB_Exposed = False, which is the class level equivalent of Option Private Module.
' * @remarks The body is a vertical cylinder, which is why rotation never needs to be applied to it: only its centre is carried into an object's frame, and a circle looks the same from every angle. Advance is the whole step; call it once a frame with what the player asked for and read the state back off the properties.
' */

Option Explicit

'/** @section Constants */

'/** @description Passes of push-out per step. Leaving one box can push into the next, which is what happens in a corner. */
Private Const SOLVE_PASSES As Long = 4

'/** @description Slack used when deciding whether the body is above a surface rather than against it. */
Private Const SURFACE_EPS As Single = 0.004

'/** @description Impact speed under which a bouncy surface simply catches the body instead of throwing it back. */
Private Const BOUNCE_FLOOR As Single = 2!

'/** @section Instance state */

Private m_scene As PScene

Private m_x As Single
Private m_y As Single
Private m_z As Single
Private m_vx As Single
Private m_vy As Single
Private m_vz As Single

Private m_radius As Single
Private m_height As Single

Private m_gravity As Single
Private m_walk As Single
Private m_groundAccel As Single
Private m_groundFriction As Single
Private m_airAccel As Single
Private m_jump As Single
Private m_swim As Single
Private m_dive As Single
Private m_climb As Single
Private m_stepHeight As Single
Private m_killZ As Single
Private m_swimAt As Single
Private m_reach As Single

Private m_onGround As Boolean
Private m_groundMat As Long
Private m_groundIdx As Long
Private m_climbing As Boolean
Private m_climbTouch As Boolean
Private m_submersion As Single
Private m_fluidMat As Long
Private m_landing As Single
Private m_carryYaw As Single
Private m_fellOut As Boolean

Private t_idx() As Long
Private t_count As Long

'/** @section Lifetime */

'/**
' * @brief Applies the stock build and tuning.
' */
Private Sub Class_Initialize()
    m_radius = 0.34
    m_height = 1.72

    m_gravity = 18!
    m_walk = 5.4
    m_groundAccel = 42!
    m_groundFriction = 9!
    m_airAccel = 8!
    m_jump = 6.4
    m_swim = 3.4
    m_dive = 3!
    m_climb = 3.2
    m_stepHeight = 0.45
    m_killZ = -1E+30
    m_swimAt = 0.3
    m_reach = 0.61

    m_groundMat = 0
    m_groundIdx = P_INVALID_ID
    m_fluidMat = P_INVALID_ID

    ReDim t_idx(0 To 15)
End Sub

'/**
' * @brief Binds the body to the scene it moves through.
' * @param sc The scene to collide against.
' * @return The body itself, so it can be created and bound in one expression.
' */
Public Function Attach(ByVal sc As PScene) As PBody
    Set m_scene = sc
    Set Attach = Me
End Function

'/**
' * @brief Sets the shape of the body.
' * @param radius Half its width.
' * @param height Its height from feet to crown.
' */
Public Sub SetSize(ByVal radius As Single, ByVal height As Single)
    If radius > 0! Then m_radius = radius
    If height > 0! Then m_height = height
End Sub

'/**
' * @brief Puts the body somewhere and stops it dead.
' * @param X The world X coordinate of the feet.
' * @param Y The world Y coordinate of the feet.
' * @param Z The height of the feet.
' */
Public Sub SetPosition(ByVal X As Single, ByVal Y As Single, ByVal Z As Single)
    m_x = X
    m_y = Y
    m_z = Z
    m_vx = 0!
    m_vy = 0!
    m_vz = 0!
    m_onGround = False
    m_climbing = False
    m_groundIdx = P_INVALID_ID
    m_submersion = 0!
    m_fellOut = False
    m_landing = 0!
End Sub

'/** @section The step */

'/**
' * @brief Shifts the body without disturbing what it was doing.
' * @param dx The offset along X.
' * @param dy The offset along Y.
' * @param dz The offset along Z.
' * @remarks For a platform driven by hand rather than by the motion system: the body has to travel with
' * it, and travelling is not the same as being accelerated, so velocity is deliberately left alone.
' */
Public Sub Nudge(ByVal dx As Single, ByVal dy As Single, ByVal dz As Single)
    m_x = m_x + dx
    m_y = m_y + dy
    m_z = m_z + dz
End Sub

'/**
' * @brief Advances the body by one frame.
' * @param dt The elapsed frame time in seconds.
' * @param wishX How hard the body is being pushed along world X, from -1 to 1.
' * @param wishY How hard it is being pushed along world Y.
' * @param up True while the jump, swim or climb up control is held.
' * @param down True while the dive or climb down control is held.
' * @description The order matters and is the order things happen in the world. Platforms have already
' * moved by the time this runs, so the body is carried first; then it is accelerated, then pushed out
' * of anything solid, and only then does gravity get a say. Doing gravity before the push-out would
' * let the body sink into a floor for a frame before being lifted back, which reads as a stutter.
' */
Public Sub Advance(ByVal dt As Single, ByVal wishX As Single, ByVal wishY As Single, _
                   ByVal up As Boolean, ByVal down As Boolean)
    Dim wasGround As Boolean
    Dim prevZ As Single

    If m_scene Is Nothing Then Exit Sub

    wasGround = m_onGround
    m_carryYaw = 0!
    m_landing = 0!

    CarryWithGround
    SampleTriggers

    '/* Decided once, before anything reads it. Whether the body is against a ladder is known from the
    ' * push-out at the end of the last step, so the control that sets the climbing speed and the gravity
    ' * that must stay switched off while climbing are answering the same question with the same answer.
    ' * Working it out again after the push-out, which is where it used to happen, left one frame able to
    ' * apply a climbing speed under full gravity, or no gravity to a body that was not climbing. */
    m_climbing = (m_climbTouch And (up Or down))

    ApplyWish dt, wishX, wishY
    ApplyControls up, down

    m_x = m_x + m_vx * dt
    m_y = m_y + m_vy * dt
    SolveLateral

    prevZ = m_z
    ApplyVertical dt, up, down
    ResolveVertical prevZ, wasGround

    m_fellOut = (m_z < m_killZ)
End Sub

'/** @section Result */

'/**
' * @brief Reads the X coordinate of the feet.
' * @return The world X position.
' */
Public Property Get X() As Single
    X = m_x
End Property

'/**
' * @brief Reads the Y coordinate of the feet.
' * @return The world Y position.
' */
Public Property Get Y() As Single
    Y = m_y
End Property

'/**
' * @brief Reads the height of the feet.
' * @return The world Z position.
' */
Public Property Get Z() As Single
    Z = m_z
End Property

'/**
' * @brief Reads the horizontal speed along X.
' * @return The velocity component.
' */
Public Property Get VelX() As Single
    VelX = m_vx
End Property

'/**
' * @brief Reads the horizontal speed along Y.
' * @return The velocity component.
' */
Public Property Get VelY() As Single
    VelY = m_vy
End Property

'/**
' * @brief Reads the vertical speed.
' * @return The velocity component.
' */
Public Property Get VelZ() As Single
    VelZ = m_vz
End Property

'/**
' * @brief Reports whether the body is standing on something.
' * @return True when it has ground under it.
' */
Public Property Get OnGround() As Boolean
    OnGround = m_onGround
End Property

'/**
' * @brief Reads the material of whatever the body is standing on.
' * @return The material id, or the default when it is airborne.
' */
Public Property Get GroundMaterial() As Long
    GroundMaterial = m_groundMat
End Property

'/**
' * @brief Reads which object the body is standing on.
' * @return The object id, or P_INVALID_ID when it is airborne.
' */
Public Property Get GroundObject() As Long
    GroundObject = m_groundIdx
End Property

'/**
' * @brief Reads how much of the body is under a fluid surface.
' * @return A fraction from 0 to 1.
' */
Public Property Get Submersion() As Single
    Submersion = m_submersion
End Property

'/**
' * @brief Reports whether the body is deep enough in a fluid to swim.
' * @return True when swimming rather than wading.
' */
Public Property Get IsSwimming() As Boolean
    IsSwimming = (m_submersion >= m_swimAt)
End Property

'/**
' * @brief Reports whether the body is climbing a climbable surface.
' * @return True while it is on a ladder and asking to move along it.
' */
Public Property Get IsClimbing() As Boolean
    IsClimbing = m_climbing
End Property

'/**
' * @brief Reads how hard the body hit the ground on the step just taken.
' * @return The impact speed, or zero when it did not land.
' * @remarks What a camera reads to decide how much to flinch.
' */
Public Property Get LandingSpeed() As Single
    LandingSpeed = m_landing
End Property

'/**
' * @brief Reads how far a turning platform swung the body on the step just taken.
' * @return The change in heading, in radians.
' * @remarks Add it to the camera yaw and the view turns with the platform, which is what makes standing on a spinning disc feel like standing on it rather than beside it.
' */
Public Property Get CarryYaw() As Single
    CarryYaw = m_carryYaw
End Property

'/**
' * @brief Reports whether the body has fallen past the height the world ends at.
' * @return True when it is lost.
' */
Public Property Get FellOut() As Boolean
    FellOut = m_fellOut
End Property

'/**
' * @brief Reads how many triggers the body touched on the step just taken.
' * @return The count.
' * @remarks A trigger is reported, never acted on. What a coin or a checkpoint means is a rule about the game, and the body has no business knowing it.
' */
Public Property Get TouchCount() As Long
    TouchCount = t_count
End Property

'/**
' * @brief Reads one of the triggers touched on the step just taken.
' * @param slot Its position in the list.
' * @return The object id, or P_INVALID_ID when the slot is out of range.
' */
Public Function TouchAt(ByVal slot As Long) As Long
    If slot < 0 Or slot >= t_count Then
        TouchAt = P_INVALID_ID
        Exit Function
    End If

    TouchAt = t_idx(slot)
End Function

'/** @section Tuning */

'/**
' * @brief Reads the downward acceleration.
' * @return The value in world units per second squared.
' */
Public Property Get Gravity() As Single
    Gravity = m_gravity
End Property

'/**
' * @brief Sets the downward acceleration.
' * @param value The value in world units per second squared.
' */
Public Property Let Gravity(ByVal value As Single)
    m_gravity = value
End Property

'/**
' * @brief Reads the top walking speed on a normal surface.
' * @return The speed in world units per second.
' */
Public Property Get WalkSpeed() As Single
    WalkSpeed = m_walk
End Property

'/**
' * @brief Sets the top walking speed on a normal surface.
' * @param value The speed in world units per second, before the ground material scales it.
' */
Public Property Let WalkSpeed(ByVal value As Single)
    m_walk = value
End Property

'/**
' * @brief Reads the upward speed a jump gives.
' * @return The speed in world units per second.
' */
Public Property Get JumpSpeed() As Single
    JumpSpeed = m_jump
End Property

'/**
' * @brief Sets the upward speed a jump gives.
' * @param value The speed in world units per second.
' */
Public Property Let JumpSpeed(ByVal value As Single)
    m_jump = value
End Property

'/**
' * @brief Reads the tallest lip the body walks over without jumping.
' * @return The height in world units.
' * @remarks Only the fallback. A surface that names its own StepHeight overrides this, which is how a kerb and a cliff can be made of the same shape.
' */
Public Property Get StepHeight() As Single
    StepHeight = m_stepHeight
End Property

'/**
' * @brief Sets the fallback height for a lip the body walks over.
' * @param value The height in world units.
' */
Public Property Let StepHeight(ByVal value As Single)
    m_stepHeight = value
End Property

'/**
' * @brief Reads the height below which the body counts as lost.
' * @return The height in world units.
' */
Public Property Get KillZ() As Single
    KillZ = m_killZ
End Property

'/**
' * @brief Sets the height below which the body counts as lost.
' * @param value The height in world units.
' */
Public Property Let KillZ(ByVal value As Single)
    m_killZ = value
End Property

'/**
' * @brief Reads the speed the body climbs a ladder at.
' * @return The speed in world units per second.
' */
Public Property Get ClimbSpeed() As Single
    ClimbSpeed = m_climb
End Property

'/**
' * @brief Sets the speed the body climbs a ladder at.
' * @param value The speed in world units per second.
' */
Public Property Let ClimbSpeed(ByVal value As Single)
    m_climb = value
End Property

'/**
' * @brief Reads the speed the body swims upwards at.
' * @return The speed in world units per second.
' */
Public Property Get SwimSpeed() As Single
    SwimSpeed = m_swim
End Property

'/**
' * @brief Sets the speed the body swims upwards at.
' * @param value The speed in world units per second.
' */
Public Property Let SwimSpeed(ByVal value As Single)
    m_swim = value
End Property

'/**
' * @brief Reads the speed the body dives at.
' * @return The speed in world units per second.
' */
Public Property Get DiveSpeed() As Single
    DiveSpeed = m_dive
End Property

'/**
' * @brief Sets the speed the body dives at.
' * @param value The speed in world units per second.
' */
Public Property Let DiveSpeed(ByVal value As Single)
    m_dive = value
End Property

'/**
' * @brief Reads how far past its own edge the body notices a trigger.
' * @return The extra distance in world units.
' */
Public Property Get Reach() As Single
    Reach = m_reach
End Property

'/**
' * @brief Sets how far past its own edge the body notices a trigger.
' * @param value The extra distance in world units.
' * @description A coin the width of a hand, touched only when the body's own edge crosses it, is a coin
' * players walk through and do not collect. Reach is the difference between what the body occupies and
' * what it can reach for, and it applies to triggers alone: nothing solid uses it.
' */
Public Property Let Reach(ByVal value As Single)
    m_reach = value
End Property

'/**
' * @brief Reads how quickly the body reaches walking speed.
' * @return The acceleration in world units per second squared.
' */
Public Property Get GroundAccel() As Single
    GroundAccel = m_groundAccel
End Property

'/**
' * @brief Sets how quickly the body reaches walking speed.
' * @param value The acceleration in world units per second squared.
' */
Public Property Let GroundAccel(ByVal value As Single)
    m_groundAccel = value
End Property

'/**
' * @brief Reads how quickly the body sheds speed with nothing held.
' * @return The damping in units per second.
' */
Public Property Get GroundFriction() As Single
    GroundFriction = m_groundFriction
End Property

'/**
' * @brief Sets how quickly the body sheds speed with nothing held.
' * @param value The damping in units per second, before the ground material scales it.
' */
Public Property Let GroundFriction(ByVal value As Single)
    m_groundFriction = value
End Property

'/**
' * @brief Reads how much control the body keeps in mid air.
' * @return The acceleration in world units per second squared.
' */
Public Property Get AirAccel() As Single
    AirAccel = m_airAccel
End Property

'/**
' * @brief Sets how much control the body keeps in mid air.
' * @param value The acceleration in world units per second squared.
' */
Public Property Let AirAccel(ByVal value As Single)
    m_airAccel = value
End Property

'/** @section Private steps */

'/**
' * @brief Travels with whatever the body is standing on.
' * @description A platform that moves out from under a body is a platform the body slides off, and a
' * turning one is worse: the floor rotates and the body stays put, facing the wrong way. Translation is
' * added straight on; rotation is applied by swinging the body around the object's centre, and the same
' * angle is offered back as CarryYaw so the view can turn with it.
' */
Private Sub CarryWithGround()
    Dim dx As Single
    Dim dy As Single
    Dim dz As Single
    Dim da As Single
    Dim ocx As Single
    Dim ocy As Single
    Dim ohw As Single
    Dim ohh As Single
    Dim oang As Single
    Dim relX As Single
    Dim relY As Single
    Dim c As Single
    Dim s As Single

    If m_groundIdx = P_INVALID_ID Then Exit Sub

    m_scene.GetMotionDelta m_groundIdx, dx, dy, dz
    m_x = m_x + dx
    m_y = m_y + dy
    m_z = m_z + dz

    da = m_scene.SpinDelta(m_groundIdx)
    If da = 0! Then Exit Sub

    m_scene.GetOrientedBox m_groundIdx, ocx, ocy, ohw, ohh, oang

    relX = m_x - ocx
    relY = m_y - ocy
    c = Cos(da)
    s = Sin(da)

    m_x = ocx + relX * c - relY * s
    m_y = ocy + relX * s + relY * c
    m_carryYaw = da
End Sub

'/**
' * @brief Turns the requested direction into speed, respecting the surface underfoot.
' * @param dt The elapsed frame time in seconds.
' * @param wishX How hard the body is being pushed along world X.
' * @param wishY How hard it is being pushed along world Y.
' */
Private Sub ApplyWish(ByVal dt As Single, ByVal wishX As Single, ByVal wishY As Single)
    Dim mgn As Single
    Dim accel As Single
    Dim cap As Single
    Dim drop As Single
    Dim mat As PMaterial

    Set mat = PMaterials.ById(m_groundMat)

    cap = m_walk
    If m_onGround Then cap = cap * mat.SpeedMultiplier

    If m_submersion > 0! Then
        cap = cap * PCore.Lerp(1!, PMaterials.ById(m_fluidMat).SpeedMultiplier, m_submersion)
    End If

    mgn = Sqr(wishX * wishX + wishY * wishY)

    If mgn > P_EPSILON Then
        wishX = wishX / mgn
        wishY = wishY / mgn
    End If

    If m_onGround Then accel = m_groundAccel Else accel = m_airAccel

    m_vx = m_vx + wishX * accel * dt
    m_vy = m_vy + wishY * accel * dt

    If m_onGround And mgn <= P_EPSILON Then
        drop = m_groundFriction * mat.Friction * dt
        If drop > 1! Then drop = 1!
        m_vx = m_vx - m_vx * drop
        m_vy = m_vy - m_vy * drop
    End If

    mgn = Sqr(m_vx * m_vx + m_vy * m_vy)

    If mgn > cap And mgn > P_EPSILON Then
        m_vx = m_vx * cap / mgn
        m_vy = m_vy * cap / mgn
    End If
End Sub

'/**
' * @brief Reads the jump, swim, dive and climb controls.
' * @param up True while the up control is held.
' * @param down True while the down control is held.
' */
Private Sub ApplyControls(ByVal up As Boolean, ByVal down As Boolean)
    If m_climbing Then
        If up Then m_vz = m_climb Else m_vz = -m_climb
        Exit Sub
    End If

    If m_submersion >= m_swimAt Then
        If up Then
            If m_vz < m_swim Then m_vz = m_swim
        ElseIf down Then
            If m_vz > -m_dive Then m_vz = -m_dive
        End If

        Exit Sub
    End If

    If m_onGround And up Then
        m_vz = m_jump
        m_onGround = False
    End If
End Sub

'/**
' * @brief Applies gravity, or what a fluid leaves of it.
' * @param dt The elapsed frame time in seconds.
' * @param up True while the up control is held.
' * @param down True while the down control is held.
' */
Private Sub ApplyVertical(ByVal dt As Single, ByVal up As Boolean, ByVal down As Boolean)
    Dim fluid As PMaterial
    Dim damp As Single

    If m_climbing Then
        m_z = m_z + m_vz * dt
        Exit Sub
    End If

    If m_submersion > 0! Then
        Set fluid = PMaterials.ById(m_fluidMat)

        m_vz = m_vz - m_gravity * (1! - fluid.Buoyancy * m_submersion) * dt

        damp = fluid.Drag * m_submersion * dt
        If damp > 0.9 Then damp = 0.9

        m_vx = m_vx - m_vx * damp
        m_vy = m_vy - m_vy * damp
        m_vz = m_vz - m_vz * damp
    Else
        m_vz = m_vz - m_gravity * dt
    End If

    m_z = m_z + m_vz * dt
End Sub

'/**
' * @brief Collects the fluid the body is in and every trigger it is touching.
' * @description A fluid has to actually contain the body, since being beside a pool is not being in it.
' * Everything else is measured to its centre and given Reach, because those are things to be picked up
' * or stepped onto and both want to be easier to hit than a wall is to bump into.
' */
Private Sub SampleTriggers()
    Dim n As Long
    Dim k As Long
    Dim idx As Long
    Dim mat As PMaterial
    Dim spanBottom As Single
    Dim spanTop As Single
    Dim best As Single
    Dim far As Single
    Dim x1 As Single, y1 As Single, x2 As Single, y2 As Single

    t_count = 0
    m_submersion = 0!
    m_fluidMat = P_INVALID_ID
    best = -1E+30
    far = m_radius + m_reach

    n = m_scene.QueryBox(m_x - far, m_y - far, m_x + far, m_y + far)

    For k = 0 To n - 1
        idx = m_scene.ResultAt(k)
        Set mat = PMaterials.ById(m_scene.MaterialOf(idx))

        If mat.IsTrigger Then
            m_scene.GetSpanZ idx, spanBottom, spanTop

            If mat.IsFluid Then
                If m_scene.ContainsXY(idx, m_x, m_y) Then
                    If m_z < spanTop And m_z + m_height > spanBottom Then
                        If spanTop > best Then
                            best = spanTop
                            m_fluidMat = m_scene.MaterialOf(idx)
                        End If
                    End If
                End If
            Else
                m_scene.GetBoundsXY idx, x1, y1, x2, y2

                If PCore.DistSq2D(m_x, m_y, (x1 + x2) * 0.5!, (y1 + y2) * 0.5!) < far * far Then
                    If m_z < spanTop + 0.4! And m_z + m_height > spanBottom - 0.4! Then
                        If t_count > UBound(t_idx) Then ReDim Preserve t_idx(0 To t_count * 2)
                        t_idx(t_count) = idx
                        t_count = t_count + 1
                    End If
                End If
            End If
        End If
    Next k

    If m_fluidMat = P_INVALID_ID Then Exit Sub

    m_submersion = PCore.Clamp01((best - m_z) / m_height)
End Sub

'/**
' * @brief Pushes the body out of everything solid it ended the step inside of.
' * @description Resolution happens along the axis of least penetration, in the object's own frame, so a
' * slab that has been turned is treated as the slab it is rather than as the square that encloses it.
' * A lip no taller than the surface allows is stepped over instead of blocked, which is the difference
' * between a stair and a wall.
' */
Private Sub SolveLateral()
    Dim pass As Long
    Dim n As Long
    Dim k As Long
    Dim idx As Long
    Dim mat As PMaterial
    Dim spanBottom As Single
    Dim spanTop As Single
    Dim surface As Single
    Dim ocx As Single, ocy As Single
    Dim ohw As Single, ohh As Single
    Dim oang As Single
    Dim ca As Single, sa As Single
    Dim lx As Single, ly As Single
    Dim ovLeft As Single, ovRight As Single
    Dim ovFront As Single, ovBack As Single
    Dim pushX As Single, pushY As Single
    Dim vLocal As Single
    Dim lip As Single
    Dim touched As Boolean

    m_climbTouch = False

    For pass = 1 To SOLVE_PASSES
        touched = False
        n = m_scene.QueryBox(m_x - m_radius, m_y - m_radius, m_x + m_radius, m_y + m_radius)

        For k = 0 To n - 1
            idx = m_scene.ResultAt(k)
            Set mat = PMaterials.ById(m_scene.MaterialOf(idx))

            If mat.IsSolid Or mat.Climbable Then
                m_scene.GetSpanZ idx, spanBottom, spanTop

                '/* Measured where the body stands, not at the highest corner: on a ramp the two differ,
                ' * and the corner would shove the body off the slope it is walking up. */
                surface = m_scene.TopZAt(idx, m_x, m_y)

                If m_z < surface - SURFACE_EPS And m_z + m_height > spanBottom + SURFACE_EPS Then
                    m_scene.GetOrientedBox idx, ocx, ocy, ohw, ohh, oang

                    If oang = 0! Then
                        ca = 1!
                        sa = 0!
                        lx = m_x - ocx
                        ly = m_y - ocy
                    Else
                        ca = Cos(oang)
                        sa = Sin(oang)
                        lx = (m_x - ocx) * ca + (m_y - ocy) * sa
                        ly = -(m_x - ocx) * sa + (m_y - ocy) * ca
                    End If

                    ovLeft = (lx + m_radius) + ohw
                    ovRight = ohw - (lx - m_radius)
                    ovFront = (ly + m_radius) + ohh
                    ovBack = ohh - (ly - m_radius)

                    If ovLeft > 0! And ovRight > 0! And ovFront > 0! And ovBack > 0! Then
                        '/* Climbable and solid are separate answers to separate questions: whether the
                        ' * body can go up here, and whether it can go through. A ghost ladder is entered
                        ' * and climbed, a solid one is scaled from the outside, and both are written the
                        ' * same way. */
                        If mat.Climbable Then m_climbTouch = True

                        If mat.IsSolid Then
                            lip = mat.StepHeight
                            If lip <= 0! Then lip = m_stepHeight

                            If (m_onGround Or m_vz <= 0!) And surface - m_z <= lip Then
                                m_z = surface
                                If m_vz < 0! Then m_vz = 0!
                                m_groundMat = m_scene.MaterialOf(idx)
                                touched = True
                            Else
                                If ovLeft < ovRight Then pushX = -ovLeft Else pushX = ovRight
                                If ovFront < ovBack Then pushY = -ovFront Else pushY = ovBack

                                If Abs(pushX) <= Abs(pushY) Then pushY = 0! Else pushX = 0!

                                m_x = m_x + pushX * ca - pushY * sa
                                m_y = m_y + pushX * sa + pushY * ca

                                If pushX <> 0! Then
                                    vLocal = m_vx * ca + m_vy * sa
                                    m_vx = m_vx - vLocal * ca
                                    m_vy = m_vy - vLocal * sa
                                Else
                                    vLocal = -m_vx * sa + m_vy * ca
                                    m_vx = m_vx + vLocal * sa
                                    m_vy = m_vy - vLocal * ca
                                End If

                                touched = True
                            End If
                        End If
                    End If
                End If
            End If
        Next k

        If Not touched Then Exit For
    Next pass
End Sub

'/**
' * @brief Lands the body on the highest surface it crossed, honouring one-way platforms and bounce.
' * @param prevZ The height of the feet before gravity was applied.
' * @param wasGround True when the body was already standing on something when the step began.
' * @remarks A body already on the ground is snapped to the surface under it even when that surface rose,
' * which is what happens every time it edges up a slope. Without that, drifting uphill leaves the body
' * airborne for a frame, and an airborne body is one the lateral solver may shove sideways out of the ramp.
' * This still runs while climbing, because it only ever acts on a body that is descending: climbing up
' * passes straight through it, and climbing down stops at the floor instead of going through it.
' */
Private Sub ResolveVertical(ByVal prevZ As Single, ByVal wasGround As Boolean)
    Dim n As Long
    Dim k As Long
    Dim idx As Long
    Dim mat As PMaterial
    Dim surface As Single
    Dim spanBottom As Single
    Dim spanTop As Single
    Dim bestZ As Single
    Dim bestMat As Long
    Dim bestIdx As Long
    Dim bounce As Single
    Dim landed As Boolean
    Dim lip As Single

    m_onGround = False
    bestZ = -1E+30
    bestMat = 0
    bestIdx = P_INVALID_ID

    n = m_scene.QueryBox(m_x - m_radius, m_y - m_radius, m_x + m_radius, m_y + m_radius)

    For k = 0 To n - 1
        idx = m_scene.ResultAt(k)
        Set mat = PMaterials.ById(m_scene.MaterialOf(idx))

        If mat.IsSolid Or mat.IsOneWay Then
            If m_scene.ContainsXY(idx, m_x, m_y) Then
                surface = m_scene.TopZAt(idx, m_x, m_y)

                lip = mat.StepHeight
                If lip <= 0! Then lip = m_stepHeight

                If m_vz <= 0! And m_z <= surface + 0.001 Then
                    If prevZ >= surface - 0.002 Or (wasGround And prevZ >= surface - lip) Then
                        If surface > bestZ Then
                            bestZ = surface
                            bestMat = m_scene.MaterialOf(idx)
                            bestIdx = idx
                            landed = True
                        End If
                    End If
                End If

                If mat.IsSolid And m_vz > 0! Then
                    m_scene.GetSpanZ idx, spanBottom, spanTop

                    If m_z + m_height > spanBottom And prevZ + m_height <= spanBottom + 0.002 Then
                        m_z = spanBottom - m_height
                        m_vz = 0!
                    End If
                End If
            End If
        End If
    Next k

    m_groundIdx = P_INVALID_ID

    If Not landed Then Exit Sub

    m_z = bestZ
    m_groundMat = bestMat
    m_groundIdx = bestIdx
    m_landing = Abs(m_vz)

    bounce = PMaterials.ById(bestMat).Bounce

    If bounce > 0! And m_vz < -BOUNCE_FLOOR Then
        m_vz = -m_vz * bounce
        m_onGround = False
    Else
        m_vz = 0!
        m_onGround = True
    End If
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PScene"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/**
' * PScene - Psu3D Object Store And Draw List
' * @description Holds every object in a world as parallel arrays rather than objects, so adding a thousand boxes costs a thousand array slots and no allocations per frame. Answers the two questions a frame asks: which objects can the camera see, and which objects sit inside this box of space.
' * @author UesleiDev
' * @version 1.0
' * @scope Private to this VBA project: class modules carry VB_Exposed = False, which is the class level equivalent of Option Private Module.
' * @remarks There is no object ceiling; the arrays double when they fill. Ids stay valid for the life of an object because nothing ever reshuffles the arrays, but a slot given back by Remove is handed out again, so an id kept past a Remove names whatever took its place.
' */

Option Explicit

'/** @section Constants */

'/** @description Slot count the arrays start at. */
Private Const SEED_CAPACITY As Long = 64

'/** @description Polygon cost charged per primitive kind, used to spend the renderer budget. */
Private Const COST_BOX As Long = 3
Private Const COST_ROTATED As Long = 6
Private Const COST_RAMP As Long = 4
Private Const COST_BILLBOARD As Long = 1
Private Const COST_SPINNER As Long = 2

'/** @section Object arrays */

Private o_kind() As Byte
Private o_active() As Boolean
Private o_mat() As Long
Private o_cost() As Long

Private o_x1() As Single
Private o_y1() As Single
Private o_x2() As Single
Private o_y2() As Single
Private o_zTop() As Single
Private o_zHigh() As Single
Private o_thick() As Single

Private o_cx() As Single
Private o_cy() As Single
Private o_cz() As Single
Private o_radius() As Single

Private o_angle() As Single
Private o_phase() As Single
Private o_axis() As Byte
Private o_sizeA() As Single
Private o_sizeB() As Single
Private o_invis() As Boolean


Private mo_mode() As Byte
Private mo_ax() As Single
Private mo_ay() As Single
Private mo_az() As Single
Private mo_amp() As Single
Private mo_speed() As Single
Private mo_phase() As Single
Private mo_t() As Single
Private mo_last() As Single
Private mo_dx() As Single
Private mo_dy() As Single
Private mo_dz() As Single
Private mo_dAng() As Single
Private o_tag() As Long
Private f_free() As Long
Private f_gone() As Boolean
Private o_dyn() As Boolean

Private g_start() As Long
Private g_items() As Long
Private g_stamp() As Long
Private dyn_idx() As Long
Private m_dynCount As Long
Private m_gridReady As Boolean
Private m_cell As Single
Private m_gx0 As Single
Private m_gy0 As Single
Private m_gw As Long
Private m_gh As Long
Private m_stampTick As Long
Private f_count As Long

Private m_count As Long
Private m_capacity As Long
Private m_live As Long
Private m_lodSize As Single
Private m_lodOn As Boolean
Private m_drawn As Long

'/** @description Slack allowed when deciding that two boxes are separated along an axis. */
Private Const SEPARATION_EPS As Single = 0.004

'/** @description Ceiling on the ordering constraints held for one frame. */
Private Const MAX_ORDER_EDGES As Long = 32768

'/** @description Ceiling on the number of cells the lookup grid may hold. */
Private Const MAX_GRID_CELLS As Long = 65536

'/** @description Ceiling on the cells along one axis, so the cell count never overflows a Long. */
Private Const MAX_GRID_SIDE As Long = 4096

'/** @description How many cells one object may occupy before it is swept instead of indexed. */
Private Const MAX_CELLS_PER_OBJECT As Long = 32

'/** @section Working arrays */

Private q_idx() As Long
Private q_count As Long

Private d_idx() As Long
Private d_near() As Single
Private d_far() As Single
Private d_mid() As Single
Private d_z1() As Single
Private d_z2() As Single
Private d_ramp() As Boolean
Private d_lod() As Boolean
Private d_rx1() As Single
Private d_ry1() As Single
Private d_rx2() As Single
Private d_ry2() As Single
Private d_count As Long

Private t_indeg() As Long
Private t_adjHead() As Long
Private t_placed() As Boolean
Private t_queue() As Long
Private t_order() As Long
Private e_to() As Long
Private e_next() As Long
Private e_count As Long
Private e_cap As Long

'/** @section Lifetime */

'/**
' * @brief Allocates the seed arrays.
' */
Private Sub Class_Initialize()
    m_lodSize = 44!
    m_lodOn = True
    Reserve SEED_CAPACITY
End Sub

'/**
' * @brief Reads the screen size under which a box drops to a single face.
' * @return The threshold in slide points.
' */
Public Property Get LodSize() As Single
    LodSize = m_lodSize
End Property

'/**
' * @brief Sets the screen size under which a box drops to a single face.
' * @param value The projected diameter in slide points; zero or less turns the reduction off.
' * @description A box shows at most three faces, and the two minor ones shrink to slivers long before
' * the dominant one does. Past that point they cost a full shape each to contribute a few points of
' * colour, which on a renderer whose frame is spent inside PowerPoint is the most expensive kind of
' * detail there is. Raise the threshold to buy frame rate, lower it to buy silhouette.
' */
Public Property Let LodSize(ByVal value As Single)
    m_lodSize = value
    m_lodOn = (value > 0!)
End Property

'/**
' * @brief Grows the arrays up front so no allocation happens while the world is being built.
' * @param slots How many object slots to make room for.
' */
Public Sub Reserve(ByVal slots As Long)
    Dim target As Long

    If slots <= m_capacity Then Exit Sub

    target = m_capacity
    If target < SEED_CAPACITY Then target = SEED_CAPACITY
    Do While target < slots
        target = target * 2
    Loop

    ReDim Preserve o_kind(0 To target - 1)
    ReDim Preserve o_active(0 To target - 1)
    ReDim Preserve o_mat(0 To target - 1)
    ReDim Preserve o_cost(0 To target - 1)

    ReDim Preserve o_x1(0 To target - 1)
    ReDim Preserve o_y1(0 To target - 1)
    ReDim Preserve o_x2(0 To target - 1)
    ReDim Preserve o_y2(0 To target - 1)
    ReDim Preserve o_zTop(0 To target - 1)
    ReDim Preserve o_zHigh(0 To target - 1)
    ReDim Preserve o_thick(0 To target - 1)

    ReDim Preserve o_cx(0 To target - 1)
    ReDim Preserve o_cy(0 To target - 1)
    ReDim Preserve o_cz(0 To target - 1)
    ReDim Preserve o_radius(0 To target - 1)

    ReDim Preserve o_angle(0 To target - 1)
    ReDim Preserve o_phase(0 To target - 1)
    ReDim Preserve o_axis(0 To target - 1)
    ReDim Preserve o_sizeA(0 To target - 1)
    ReDim Preserve o_sizeB(0 To target - 1)
    ReDim Preserve o_invis(0 To target - 1)


    ReDim Preserve mo_mode(0 To target - 1)
    ReDim Preserve mo_ax(0 To target - 1)
    ReDim Preserve mo_ay(0 To target - 1)
    ReDim Preserve mo_az(0 To target - 1)
    ReDim Preserve mo_amp(0 To target - 1)
    ReDim Preserve mo_speed(0 To target - 1)
    ReDim Preserve mo_phase(0 To target - 1)
    ReDim Preserve mo_t(0 To target - 1)
    ReDim Preserve mo_last(0 To target - 1)
    ReDim Preserve mo_dx(0 To target - 1)
    ReDim Preserve mo_dy(0 To target - 1)
    ReDim Preserve mo_dz(0 To target - 1)
    ReDim Preserve mo_dAng(0 To target - 1)
    ReDim Preserve o_tag(0 To target - 1)
    ReDim Preserve f_free(0 To target - 1)
    ReDim Preserve f_gone(0 To target - 1)
    ReDim Preserve o_dyn(0 To target - 1)
    ReDim Preserve g_stamp(0 To target - 1)
    ReDim Preserve dyn_idx(0 To target - 1)

    ReDim Preserve q_idx(0 To target - 1)
    ReDim Preserve d_idx(0 To target - 1)
    ReDim Preserve d_near(0 To target - 1)
    ReDim Preserve d_far(0 To target - 1)
    ReDim Preserve d_mid(0 To target - 1)
    ReDim Preserve d_z1(0 To target - 1)
    ReDim Preserve d_z2(0 To target - 1)
    ReDim Preserve d_ramp(0 To target - 1)
    ReDim Preserve d_lod(0 To target - 1)
    ReDim Preserve d_rx1(0 To target - 1)
    ReDim Preserve d_ry1(0 To target - 1)
    ReDim Preserve d_rx2(0 To target - 1)
    ReDim Preserve d_ry2(0 To target - 1)

    ReDim Preserve t_indeg(0 To target - 1)
    ReDim Preserve t_adjHead(0 To target - 1)
    ReDim Preserve t_placed(0 To target - 1)
    ReDim Preserve t_order(0 To target - 1)
    ReDim Preserve t_queue(0 To target * 2 - 1)

    m_capacity = target
End Sub

'/**
' * @brief Empties the scene without releasing the arrays, so the next world reuses the same memory.
' */
Public Sub Clear()
    m_count = 0
    m_drawn = 0
    m_live = 0
    f_count = 0
    m_gridReady = False
    q_count = 0
    d_count = 0
End Sub

'/**
' * @brief Reads how many object slots have ever been used, which is also the highest id plus one.
' * @return The slot count.
' */
Public Property Get Count() As Long
    Count = m_count
End Property

'/**
' * @brief Reads how many objects are currently active.
' * @return The live object count.
' */
Public Property Get LiveCount() As Long
    LiveCount = m_live
End Property

'/** @section Building */

'/**
' * @brief Adds an axis-aligned box.
' * @param matId The material the box is drawn and collided with.
' * @param x1 The lower X bound.
' * @param y1 The lower Y bound.
' * @param x2 The upper X bound.
' * @param y2 The upper Y bound.
' * @param zTop The height of the top face.
' * @param thickness How far the box extends below its top face.
' * @return The id of the new object.
' */
Public Function AddBox(ByVal matId As Long, _
                       ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, _
                       ByVal zTop As Single, ByVal thickness As Single) As Long
    Dim i As Long

    i = NewSlot(pkBox, matId, COST_BOX)

    o_x1(i) = PCore.MinOf(x1, x2)
    o_x2(i) = PCore.MaxOf(x1, x2)
    o_y1(i) = PCore.MinOf(y1, y2)
    o_y2(i) = PCore.MaxOf(y1, y2)
    o_zTop(i) = zTop
    o_zHigh(i) = zTop
    o_thick(i) = PCore.MaxOf(thickness, 0.01!)

    RefreshBounds i
    AddBox = i
End Function

'/**
' * @brief Adds a box that spins around its own vertical axis.
' * @param matId The material the box is drawn and collided with.
' * @param cx The centre X coordinate.
' * @param cy The centre Y coordinate.
' * @param halfW Half the extent along the local X axis.
' * @param halfH Half the extent along the local Y axis.
' * @param zTop The height of the top face.
' * @param thickness How far the box extends below its top face.
' * @param angleRad The starting rotation in radians.
' * @return The id of the new object.
' */
Public Function AddRotatedBox(ByVal matId As Long, _
                              ByVal cx As Single, ByVal cy As Single, _
                              ByVal halfW As Single, ByVal halfH As Single, _
                              ByVal zTop As Single, ByVal thickness As Single, _
                              Optional ByVal angleRad As Single = 0!) As Long
    Dim i As Long

    i = NewSlot(pkRotatedBox, matId, COST_ROTATED)

    o_cx(i) = cx
    o_cy(i) = cy
    o_sizeA(i) = halfW
    o_sizeB(i) = halfH
    o_zTop(i) = zTop
    o_zHigh(i) = zTop
    o_thick(i) = PCore.MaxOf(thickness, 0.01!)
    o_angle(i) = angleRad

    RefreshBounds i
    AddRotatedBox = i
End Function

'/**
' * @brief Adds a sloped slab.
' * @param matId The material the ramp is drawn and collided with.
' * @param x1 The lower X bound.
' * @param y1 The lower Y bound.
' * @param x2 The upper X bound.
' * @param y2 The upper Y bound.
' * @param zLow The height at the low end.
' * @param zHigh The height at the high end.
' * @param thickness How far the slab extends below its sloped face.
' * @param axis paX to climb along X, paY to climb along Y.
' * @return The id of the new object.
' */
Public Function AddRamp(ByVal matId As Long, _
                        ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, _
                        ByVal zLow As Single, ByVal zHigh As Single, ByVal thickness As Single, _
                        Optional ByVal axis As PAxis = paX) As Long
    Dim i As Long

    i = NewSlot(pkRamp, matId, COST_RAMP)

    o_x1(i) = PCore.MinOf(x1, x2)
    o_x2(i) = PCore.MaxOf(x1, x2)
    o_y1(i) = PCore.MinOf(y1, y2)
    o_y2(i) = PCore.MaxOf(y1, y2)
    o_zTop(i) = zLow
    o_zHigh(i) = zHigh
    o_thick(i) = PCore.MaxOf(thickness, 0.01!)
    o_axis(i) = axis

    RefreshBounds i
    AddRamp = i
End Function

'/**
' * @brief Adds a camera facing quad.
' * @param matId The material the billboard is drawn with.
' * @param X The centre X coordinate.
' * @param Y The centre Y coordinate.
' * @param Z The centre Z coordinate.
' * @param Width The width in world units.
' * @param Height The height in world units.
' * @return The id of the new object.
' */
Public Function AddBillboard(ByVal matId As Long, _
                             ByVal X As Single, ByVal Y As Single, ByVal Z As Single, _
                             ByVal Width As Single, ByVal Height As Single) As Long
    Dim i As Long

    i = NewSlot(pkBillboard, matId, COST_BILLBOARD)

    o_cx(i) = X
    o_cy(i) = Y
    o_cz(i) = Z
    o_sizeA(i) = Width
    o_sizeB(i) = Height

    RefreshBounds i
    AddBillboard = i
End Function

'/**
' * @brief Adds a spinning double sided plate, the shape pickups usually take.
' * @param matId The material the plate is drawn with.
' * @param X The centre X coordinate.
' * @param Y The centre Y coordinate.
' * @param Z The centre Z coordinate.
' * @param radius The half size of the plate.
' * @return The id of the new object.
' */
Public Function AddSpinner(ByVal matId As Long, _
                           ByVal X As Single, ByVal Y As Single, ByVal Z As Single, _
                           ByVal radius As Single) As Long
    Dim i As Long

    i = NewSlot(pkSpinner, matId, COST_SPINNER)

    o_cx(i) = X
    o_cy(i) = Y
    o_cz(i) = Z
    o_sizeA(i) = radius
    o_sizeB(i) = radius

    RefreshBounds i
    AddSpinner = i
End Function

'/** @section Editing */

'/**
' * @brief Reports whether an id refers to a live object.
' * @param idx The object id.
' * @return True when the object exists and is active.
' */
Public Function IsActive(ByVal idx As Long) As Boolean
    If idx < 0 Or idx >= m_count Then Exit Function
    IsActive = o_active(idx)
End Function

'/**
' * @brief Turns an object on or off without disturbing any other id.
' * @param idx The object id.
' * @param value True to draw and collide with the object.
' */
Public Sub SetActive(ByVal idx As Long, ByVal value As Boolean)
    If idx < 0 Or idx >= m_count Then Exit Sub
    If o_active(idx) = value Then Exit Sub

    o_active(idx) = value
    m_gridReady = False
    If value Then m_live = m_live + 1 Else m_live = m_live - 1
End Sub

'/**
' * @brief Removes an object for good and returns its slot to the store.
' * @param idx The object id.
' * @remarks The id may be handed to a different object later, so drop any copy of it. To hide something
' * you intend to bring back, use SetActive instead, which keeps the slot reserved.
' */
Public Sub Remove(ByVal idx As Long)
    If idx < 0 Or idx >= m_count Then Exit Sub

    '/* Tracked apart from the active flag, so removing something already switched off still frees it and
    ' * removing the same id twice does not put one slot on the free list two times. */
    If f_gone(idx) Then Exit Sub
    f_gone(idx) = True

    SetActive idx, False
    mo_mode(idx) = 0
    o_tag(idx) = 0

    f_free(f_count) = idx
    f_count = f_count + 1
End Sub

'/**
' * @brief Reads the material an object carries.
' * @param idx The object id.
' * @return The material id, or P_INVALID_ID for an unknown object.
' */
Public Function MaterialOf(ByVal idx As Long) As Long
    If idx < 0 Or idx >= m_count Then
        MaterialOf = P_INVALID_ID
        Exit Function
    End If

    MaterialOf = o_mat(idx)
End Function

'/**
' * @brief Swaps the material of an object, which is how surfaces change state at runtime.
' * @param idx The object id.
' * @param matId The new material id.
' */
Public Sub SetMaterial(ByVal idx As Long, ByVal matId As Long)
    If idx < 0 Or idx >= m_count Then Exit Sub
    If o_mat(idx) = matId Then Exit Sub

    o_mat(idx) = matId

    '/* Whether the material draws anything is settled here, once, so the gather loop can skip an
    ' * invisible object without asking the registry about it on every frame. */
    o_invis(idx) = Not PMaterials.ById(matId).Visible
End Sub

'/**
' * @brief Reads the primitive kind of an object.
' * @param idx The object id.
' * @return The PObjectKind entry.
' */
Public Function KindOf(ByVal idx As Long) As PObjectKind
    If idx < 0 Or idx >= m_count Then Exit Function
    KindOf = o_kind(idx)
End Function

'/**
' * @brief Shifts an object by an offset, moving both its geometry and its bounds.
' * @param idx The object id.
' * @param dx The change along X.
' * @param dy The change along Y.
' * @param dz The change along Z.
' */
Public Sub MoveBy(ByVal idx As Long, ByVal dx As Single, ByVal dy As Single, ByVal dz As Single)
    If idx < 0 Or idx >= m_count Then Exit Sub

    MarkDynamic idx

    o_x1(idx) = o_x1(idx) + dx
    o_x2(idx) = o_x2(idx) + dx
    o_y1(idx) = o_y1(idx) + dy
    o_y2(idx) = o_y2(idx) + dy
    o_cx(idx) = o_cx(idx) + dx
    o_cy(idx) = o_cy(idx) + dy
    o_cz(idx) = o_cz(idx) + dz
    o_zTop(idx) = o_zTop(idx) + dz
    o_zHigh(idx) = o_zHigh(idx) + dz
End Sub

'/**
' * @brief Places the top face of an object at an absolute height.
' * @param idx The object id.
' * @param zTop The new top height.
' */
Public Sub SetTopZ(ByVal idx As Long, ByVal zTop As Single)
    Dim delta As Single

    If idx < 0 Or idx >= m_count Then Exit Sub

    MarkDynamic idx

    delta = zTop - o_zTop(idx)
    o_zTop(idx) = zTop
    o_zHigh(idx) = o_zHigh(idx) + delta
    o_cz(idx) = o_cz(idx) + delta
End Sub

'/**
' * @brief Sets the spin of a rotated box.
' * @param idx The object id.
' * @param angleRad The rotation in radians.
' */
Public Sub SetAngle(ByVal idx As Long, ByVal angleRad As Single)
    If idx < 0 Or idx >= m_count Then Exit Sub

    MarkDynamic idx

    o_angle(idx) = angleRad
    If o_kind(idx) = pkRotatedBox Then RefreshBounds idx
End Sub

'/**
' * @brief Reads the spin of a rotated box.
' * @param idx The object id.
' * @return The rotation in radians.
' */
Public Function AngleOf(ByVal idx As Long) As Single
    If idx < 0 Or idx >= m_count Then Exit Function
    AngleOf = o_angle(idx)
End Function

'/**
' * @brief Sets the animation phase used by spinners.
' * @param idx The object id.
' * @param phase The phase in radians.
' */
Public Sub SetPhase(ByVal idx As Long, ByVal phase As Single)
    If idx < 0 Or idx >= m_count Then Exit Sub
    o_phase(idx) = phase
End Sub

'/**
' * @brief Advances the animation phase of every spinner in the scene.
' * @param dt The elapsed frame time in seconds.
' * @param speed How fast the plates turn, in radians per second.
' */
Public Sub SpinAll(ByVal dt As Single, Optional ByVal speed As Single = 2.6!)
    Dim i As Long
    Dim d As Single

    d = dt * speed

    For i = 0 To m_count - 1
        If o_active(i) Then
            If o_kind(i) = pkSpinner Then o_phase(i) = o_phase(i) + d
        End If
    Next i
End Sub

'/** @section Motion */

'/**
' * @brief Makes an object slide back and forth along a direction.
' * @param idx The object id.
' * @param axisX The X component of the direction it travels along.
' * @param axisY The Y component.
' * @param axisZ The Z component.
' * @param amp How far it reaches from its resting place, in world units.
' * @param speed How fast it oscillates, in radians per second.
' * @param phase Where in its cycle it starts, so two platforms can be set out of step.
' * @remarks The offset is a sine, so the platform slows at each end instead of snapping around, which
' * is what makes it possible to step onto one on purpose.
' */
Public Sub SetMotion(ByVal idx As Long, ByVal axisX As Single, ByVal axisY As Single, ByVal axisZ As Single, _
                     ByVal amp As Single, ByVal speed As Single, Optional ByVal phase As Single = 0!)
    Dim mgn As Single

    If idx < 0 Or idx >= m_count Then Exit Sub

    mgn = Sqr(axisX * axisX + axisY * axisY + axisZ * axisZ)
    If mgn < P_EPSILON Then Exit Sub

    MarkDynamic idx

    mo_mode(idx) = 1
    mo_ax(idx) = axisX / mgn
    mo_ay(idx) = axisY / mgn
    mo_az(idx) = axisZ / mgn
    mo_amp(idx) = amp
    mo_speed(idx) = speed
    mo_phase(idx) = phase
    mo_t(idx) = 0!
    mo_last(idx) = amp * Sin(phase)

    MoveBy idx, mo_ax(idx) * mo_last(idx), mo_ay(idx) * mo_last(idx), mo_az(idx) * mo_last(idx)
End Sub

'/**
' * @brief Makes an object turn on its own vertical axis.
' * @param idx The object id.
' * @param speed The turn rate in radians per second.
' * @param phase The angle it starts at.
' */
Public Sub SetSpin(ByVal idx As Long, ByVal speed As Single, Optional ByVal phase As Single = 0!)
    If idx < 0 Or idx >= m_count Then Exit Sub

    MarkDynamic idx

    mo_mode(idx) = 2
    mo_speed(idx) = speed
    mo_phase(idx) = phase
    mo_t(idx) = 0!
End Sub

'/**
' * @brief Advances every animated object.
' * @param dt The elapsed frame time in seconds.
' * @return How many objects moved, so a caller that only redraws on change knows to redraw.
' * @remarks Each mover records the step it just took, which is what a body standing on it needs in
' * order to travel with it. Carrying is left to the caller because it is a rule about bodies, not
' * about scenery, and only the caller knows which object the body is standing on.
' */
Public Function UpdateMotion(ByVal dt As Single) As Long
    Dim i As Long
    Dim offset As Single
    Dim delta As Single

    For i = 0 To m_count - 1
        mo_dx(i) = 0!
        mo_dy(i) = 0!
        mo_dz(i) = 0!
        mo_dAng(i) = 0!

        If mo_mode(i) = 1 Then
            mo_t(i) = mo_t(i) + dt
            offset = mo_amp(i) * Sin(mo_t(i) * mo_speed(i) + mo_phase(i))
            delta = offset - mo_last(i)
            mo_last(i) = offset

            mo_dx(i) = mo_ax(i) * delta
            mo_dy(i) = mo_ay(i) * delta
            mo_dz(i) = mo_az(i) * delta

            MoveBy i, mo_dx(i), mo_dy(i), mo_dz(i)
            UpdateMotion = UpdateMotion + 1

        ElseIf mo_mode(i) = 2 Then
            mo_t(i) = mo_t(i) + dt
            mo_dAng(i) = mo_speed(i) * dt
            SetAngle i, mo_t(i) * mo_speed(i) + mo_phase(i)
            UpdateMotion = UpdateMotion + 1
        End If
    Next i
End Function

'/**
' * @brief Reads the step an object took on the last update.
' * @param idx The object id.
' * @param outDx Receives the change along X.
' * @param outDy Receives the change along Y.
' * @param outDz Receives the change along Z.
' */
Public Sub GetMotionDelta(ByVal idx As Long, ByRef outDx As Single, ByRef outDy As Single, ByRef outDz As Single)
    outDx = 0!
    outDy = 0!
    outDz = 0!

    If idx < 0 Or idx >= m_count Then Exit Sub

    outDx = mo_dx(idx)
    outDy = mo_dy(idx)
    outDz = mo_dz(idx)
End Sub

'/**
' * @brief Reads how far an object turned on the last update.
' * @param idx The object id.
' * @return The change in angle, in radians.
' * @remarks A body standing on a turning platform has to be swung around its centre by this much, or it
' * stands still while the floor rotates out from under it.
' */
Public Function SpinDelta(ByVal idx As Long) As Single
    If idx < 0 Or idx >= m_count Then Exit Function
    SpinDelta = mo_dAng(idx)
End Function

'/**
' * @brief Reads the two size numbers of an object.
' * @param idx The object id.
' * @param outA Half the local X extent of a turned box, the width of a billboard, the radius of a spinner.
' * @param outB The matching second number.
' */
Public Sub GetSize(ByVal idx As Long, ByRef outA As Single, ByRef outB As Single)
    If idx < 0 Or idx >= m_count Then Exit Sub
    outA = o_sizeA(idx)
    outB = o_sizeB(idx)
End Sub

'/**
' * @brief Reads the centre of an object.
' * @param idx The object id.
' * @param outX The centre X coordinate.
' * @param outY The centre Y coordinate.
' * @param outZ The centre Z coordinate.
' */
Public Sub GetCenter(ByVal idx As Long, ByRef outX As Single, ByRef outY As Single, ByRef outZ As Single)
    If idx < 0 Or idx >= m_count Then Exit Sub
    outX = o_cx(idx)
    outY = o_cy(idx)
    outZ = o_cz(idx)
End Sub

'/**
' * @brief Reads how far a slab hangs below its face.
' * @param idx The object id.
' * @return The thickness in world units.
' */
Public Function ThickOf(ByVal idx As Long) As Single
    If idx < 0 Or idx >= m_count Then Exit Function
    ThickOf = o_thick(idx)
End Function

'/**
' * @brief Reads the height an object was built at.
' * @param idx The object id.
' * @return The top face for a box, the low end for a ramp.
' */
Public Function TopOf(ByVal idx As Long) As Single
    If idx < 0 Or idx >= m_count Then Exit Function
    TopOf = o_zTop(idx)
End Function

'/**
' * @brief Reads the second height of an object.
' * @param idx The object id.
' * @return The high end for a ramp, and the same as TopOf for anything level.
' */
Public Function HighOf(ByVal idx As Long) As Single
    If idx < 0 Or idx >= m_count Then Exit Function
    HighOf = o_zHigh(idx)
End Function

'/**
' * @brief Reads which way a ramp climbs.
' * @param idx The object id.
' * @return paX or paY.
' */
Public Function AxisOf(ByVal idx As Long) As PAxis
    If idx < 0 Or idx >= m_count Then Exit Function
    AxisOf = o_axis(idx)
End Function

'/**
' * @brief Reads the motion an object was given.
' * @param idx The object id.
' * @param outAx The X part of the travel direction.
' * @param outAy The Y part.
' * @param outAz The Z part.
' * @param outAmp How far it travels, or nothing for a spin.
' * @param outSpeed Its rate, in cycles a second for travel and radians a second for a spin.
' * @param outPhase Where in the cycle it starts.
' * @return 0 when it is still, 1 when it travels, 2 when it spins.
' */
Public Function MotionOf(ByVal idx As Long, ByRef outAx As Single, ByRef outAy As Single, ByRef outAz As Single, _
                         ByRef outAmp As Single, ByRef outSpeed As Single, ByRef outPhase As Single) As Long
    If idx < 0 Or idx >= m_count Then Exit Function

    outAx = mo_ax(idx)
    outAy = mo_ay(idx)
    outAz = mo_az(idx)
    outAmp = mo_amp(idx)
    outSpeed = mo_speed(idx)
    outPhase = mo_phase(idx)
    MotionOf = mo_mode(idx)
End Function

'/**
' * @brief Attaches a number of your own to an object.
' * @param idx The object id.
' * @param value Whatever the number means to you; zero, the default, means nothing was said.
' * @description The scene deliberately does not know what a checkpoint, a door or a waypoint is. It
' * carries one number per object and lets the caller decide what it stands for, which is the difference
' * between a library you can build a game on and a library you can only build this game on.
' */
Public Sub SetTag(ByVal idx As Long, ByVal value As Long)
    If idx < 0 Or idx >= m_count Then Exit Sub
    o_tag(idx) = value
End Sub

'/**
' * @brief Reads the number attached to an object.
' * @param idx The object id.
' * @return The tag, or zero when none was set.
' */
Public Function TagOf(ByVal idx As Long) As Long
    If idx < 0 Or idx >= m_count Then Exit Function
    TagOf = o_tag(idx)
End Function

'/** @section Bounds and physics */

'/**
' * @brief Reads the horizontal bounds of an object.
' * @param idx The object id.
' * @param outX1 Receives the lower X bound.
' * @param outY1 Receives the lower Y bound.
' * @param outX2 Receives the upper X bound.
' * @param outY2 Receives the upper Y bound.
' */
Public Sub GetBoundsXY(ByVal idx As Long, ByRef outX1 As Single, ByRef outY1 As Single, ByRef outX2 As Single, ByRef outY2 As Single)
    If idx < 0 Or idx >= m_count Then Exit Sub

    outX1 = o_x1(idx)
    outY1 = o_y1(idx)
    outX2 = o_x2(idx)
    outY2 = o_y2(idx)
End Sub

'/**
' * @brief Reads the vertical span of an object.
' * @param idx The object id.
' * @param outBottom Receives the lowest point.
' * @param outTop Receives the highest point.
' */
Public Sub GetSpanZ(ByVal idx As Long, ByRef outBottom As Single, ByRef outTop As Single)
    If idx < 0 Or idx >= m_count Then Exit Sub

    outTop = PCore.MaxOf(o_zTop(idx), o_zHigh(idx))
    outBottom = PCore.MinOf(o_zTop(idx), o_zHigh(idx)) - o_thick(idx)
End Sub

'/**
' * @brief Measures the height of the walkable surface of an object under a point.
' * @param idx The object id.
' * @param X The world X coordinate to sample.
' * @param Y The world Y coordinate to sample.
' * @return The surface height, interpolated across the slope for ramps.
' */
Public Function TopZAt(ByVal idx As Long, ByVal X As Single, ByVal Y As Single) As Single
    Dim t As Single

    If idx < 0 Or idx >= m_count Then Exit Function

    If o_kind(idx) <> pkRamp Then
        TopZAt = o_zTop(idx)
        Exit Function
    End If

    If o_axis(idx) = paX Then
        t = PCore.InverseLerp(o_x1(idx), o_x2(idx), X)
    Else
        t = PCore.InverseLerp(o_y1(idx), o_y2(idx), Y)
    End If

    TopZAt = PCore.Lerp(o_zTop(idx), o_zHigh(idx), PCore.Clamp01(t))
End Function

'/**
' * @brief Answers whether an object stops a body that just touched it.
' * @param idx The object id.
' * @param landingFromAbove True when the body is descending onto the top surface.
' * @return True when the contact must be resolved.
' */
Public Function Blocks(ByVal idx As Long, Optional ByVal landingFromAbove As Boolean = False) As Boolean
    If idx < 0 Or idx >= m_count Then Exit Function
    If Not o_active(idx) Then Exit Function

    Blocks = PMaterials.ById(o_mat(idx)).BlocksMovement(landingFromAbove)
End Function

'/**
' * @brief Reads the oriented footprint of an object.
' * @param idx The object id.
' * @param outCx Receives the centre X coordinate.
' * @param outCy Receives the centre Y coordinate.
' * @param outHalfW Receives the half extent along the object's own X axis.
' * @param outHalfH Receives the half extent along its own Y axis.
' * @param outAngle Receives the rotation in radians.
' * @remarks For everything except a rotated box the angle is zero and the half extents come straight
' * from the bounds, so a caller can use this one shape of answer for every object it meets.
' */
Public Sub GetOrientedBox(ByVal idx As Long, _
                          ByRef outCx As Single, ByRef outCy As Single, _
                          ByRef outHalfW As Single, ByRef outHalfH As Single, _
                          ByRef outAngle As Single)
    If idx < 0 Or idx >= m_count Then Exit Sub

    outCx = o_cx(idx)
    outCy = o_cy(idx)

    If o_kind(idx) = pkRotatedBox Then
        outHalfW = o_sizeA(idx)
        outHalfH = o_sizeB(idx)
        outAngle = o_angle(idx)
    Else
        outHalfW = (o_x2(idx) - o_x1(idx)) * 0.5!
        outHalfH = (o_y2(idx) - o_y1(idx)) * 0.5!
        outAngle = 0!
    End If
End Sub

'/**
' * @brief Reports whether a point falls inside the real footprint of an object.
' * @param idx The object id.
' * @param X The world X coordinate to test.
' * @param Y The world Y coordinate to test.
' * @return True when the point is inside, taking rotation into account.
' * @remarks The bounds of a rotated box are the square that encloses it at any angle, so a point can
' * be inside those bounds and well outside the box itself. Anything that decides what a body is
' * standing on has to ask this rather than read the bounds.
' */
Public Function ContainsXY(ByVal idx As Long, ByVal X As Single, ByVal Y As Single) As Boolean
    Dim dx As Single
    Dim dy As Single
    Dim lx As Single
    Dim ly As Single
    Dim ca As Single
    Dim sa As Single

    If idx < 0 Or idx >= m_count Then Exit Function

    If o_kind(idx) <> pkRotatedBox Then
        ContainsXY = (X >= o_x1(idx) And X <= o_x2(idx) And Y >= o_y1(idx) And Y <= o_y2(idx))
        Exit Function
    End If

    dx = X - o_cx(idx)
    dy = Y - o_cy(idx)
    ca = Cos(o_angle(idx))
    sa = Sin(o_angle(idx))

    lx = dx * ca + dy * sa
    ly = -dx * sa + dy * ca

    ContainsXY = (Abs(lx) <= o_sizeA(idx) And Abs(ly) <= o_sizeB(idx))
End Function

'/** @section Spatial query */

'/**
' * @brief Collects every active object whose bounds overlap a rectangle of the world.
' * @param x1 The lower X bound of the query.
' * @param y1 The lower Y bound of the query.
' * @param x2 The upper X bound of the query.
' * @param y2 The upper Y bound of the query.
' * @return How many objects were found; read them back through ResultAt.
' * @remarks The results live in a reused array, so a second query overwrites the first.
' */
Public Function QueryBox(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Long
    Dim i As Long
    Dim k As Long
    Dim c As Long
    Dim cx As Long, cy As Long
    Dim cx1 As Long, cx2 As Long
    Dim cy1 As Long, cy2 As Long
    Dim base As Long
    Dim tick As Long

    q_count = 0

    If Not m_gridReady Then BuildGrid

    '/* Everything that moves was taken out of the grid when it moved, so it is swept here instead.
    ' * A world is nearly all scenery and a handful of platforms, and rebuilding a grid every frame
    ' * for those few would cost more than the grid saves. */
    For k = 0 To m_dynCount - 1
        i = dyn_idx(k)

        If o_active(i) Then
            If Not (o_x2(i) < x1 Or o_x1(i) > x2 Or o_y2(i) < y1 Or o_y1(i) > y2) Then
                q_idx(q_count) = i
                q_count = q_count + 1
            End If
        End If
    Next k

    If m_gw <= 0 Then
        QueryBox = q_count
        Exit Function
    End If

    '/* The stamp saves a clearing pass by moving the tick on instead of wiping the array. A Long has
    ' * an end, so on the day it is reached the grid is rebuilt, which zeroes every stamp and starts the
    ' * count over. */
    If m_stampTick >= 2100000000 Then
        BuildGrid

        If m_gw <= 0 Then
            QueryBox = q_count
            Exit Function
        End If
    End If

    m_stampTick = m_stampTick + 1
    tick = m_stampTick

    cx1 = GridX(x1)
    cx2 = GridX(x2)
    cy1 = GridY(y1)
    cy2 = GridY(y2)

    For cy = cy1 To cy2
        base = cy * m_gw

        For cx = cx1 To cx2
            c = base + cx

            For k = g_start(c) To g_start(c + 1) - 1
                i = g_items(k)

                '/* A box wide enough to cover several cells is listed in each of them. The stamp is
                ' * what stops it being answered several times, and it costs no clearing pass because
                ' * the tick moves on instead of the array being wiped. */
                If g_stamp(i) <> tick Then
                    g_stamp(i) = tick

                    If Not (o_x2(i) < x1 Or o_x1(i) > x2 Or o_y2(i) < y1 Or o_y1(i) > y2) Then
                        q_idx(q_count) = i
                        q_count = q_count + 1
                    End If
                End If
            Next k
        Next cx
    Next cy

    QueryBox = q_count
End Function

'/**
' * @brief Collects every active object within a radius of a point.
' * @param X The centre X coordinate.
' * @param Y The centre Y coordinate.
' * @param radius The search radius.
' * @return How many objects were found.
' */
Public Function QueryRadius(ByVal X As Single, ByVal Y As Single, ByVal radius As Single) As Long
    QueryRadius = QueryBox(X - radius, Y - radius, X + radius, Y + radius)
End Function

'/**
' * @brief Reads one entry of the last query result.
' * @param slot The position in the result list.
' * @return The object id, or P_INVALID_ID when the slot is out of range.
' */
Public Function ResultAt(ByVal slot As Long) As Long
    If slot < 0 Or slot >= q_count Then
        ResultAt = P_INVALID_ID
        Exit Function
    End If

    ResultAt = q_idx(slot)
End Function

'/**
' * @brief Reads how many objects the last query found.
' * @return The result count.
' */
Public Property Get ResultCount() As Long
    ResultCount = q_count
End Property

'/** @section Spatial index */

'/**
' * @brief Takes an object out of the static grid for good, because it has moved.
' * @param idx The object id.
' * @remarks Once is enough: something that moved will very likely move again, and putting it back into
' * a grid it would leave next frame is how a lookup structure ends up costing more than it saves.
' */
Private Sub MarkDynamic(ByVal idx As Long)
    If idx < 0 Or idx >= m_count Then Exit Sub
    If o_dyn(idx) Then Exit Sub

    o_dyn(idx) = True
    m_gridReady = False
End Sub

'/**
' * @brief Reads how many objects are swept rather than looked up.
' * @return The count of objects that moved at least once, plus those too big to be worth indexing.
' */
Public Property Get DynamicCount() As Long
    If Not m_gridReady Then BuildGrid
    DynamicCount = m_dynCount
End Property

'/**
' * @brief Reads how wide one cell of the lookup grid is.
' * @return The cell size in world units, or zero when nothing is indexed.
' */
Public Property Get CellSize() As Single
    If Not m_gridReady Then BuildGrid
    CellSize = m_cell
End Property

'/**
' * @brief Finds the grid column a world X coordinate falls in.
' * @param wx The world X coordinate.
' * @return The column, clamped to the grid.
' * @remarks Clamped rather than rejected: an object whose box reaches past the border is listed in the
' * border cell, so a query out there has to look at that cell rather than skip it. The exact overlap
' * test afterwards throws away whatever the clamp let in.
' */
Private Function GridX(ByVal wx As Single) As Long
    Dim c As Long

    c = Int((wx - m_gx0) / m_cell)

    If c < 0 Then c = 0
    If c > m_gw - 1 Then c = m_gw - 1

    GridX = c
End Function

'/**
' * @brief Finds the grid row a world Y coordinate falls in.
' * @param wy The world Y coordinate.
' * @return The row, clamped to the grid.
' */
Private Function GridY(ByVal wy As Single) As Long
    Dim c As Long

    c = Int((wy - m_gy0) / m_cell)

    If c < 0 Then c = 0
    If c > m_gh - 1 Then c = m_gh - 1

    GridY = c
End Function

'/**
' * @brief Works out how many cells fit along one axis of the world.
' * @param span The width of the world on that axis.
' * @param cell The cell size.
' * @return The cell count, at least one and never more than MAX_GRID_SIDE.
' * @remarks Counted in a Double and clamped before it is ever a Long. One object placed at a mistyped
' * coordinate asks for two billion cells along an axis, and the multiplication meant to catch that has
' * already overflowed by the time it runs. Clamping turns a broken world into a slow lookup instead of
' * a crash, and the exact overlap test afterwards keeps the answers right either way.
' */
Private Function SideCount(ByVal span As Single, ByVal cell As Single) As Long
    Dim n As Double

    n = Int(CDbl(span) / CDbl(cell)) + 1#

    If n < 1# Then n = 1#
    If n > CDbl(MAX_GRID_SIDE) Then n = CDbl(MAX_GRID_SIDE)

    SideCount = CLng(n)
End Function

'/**
' * @brief Counts how many cells an object's footprint covers.
' * @param i The object id.
' * @return The cell count.
' */
Private Function CellFootprint(ByVal i As Long) As Long
    CellFootprint = (GridX(o_x2(i)) - GridX(o_x1(i)) + 1) * (GridY(o_y2(i)) - GridY(o_y1(i)) + 1)
End Function

'/**
' * @brief Rebuilds the lookup grid over every object that has never moved.
' * @description Counted, prefix summed, then filled: two passes over the objects and one flat array
' * of ids, with no list per cell and no allocation while a frame is running. The cell is sized from the
' * average object, because a terrain of one metre tiles indexed in four metre cells would put sixteen
' * of them in every bucket and index nothing at all.
' */
Private Sub BuildGrid()
    Dim i As Long
    Dim c As Long
    Dim cx As Long, cy As Long
    Dim cx1 As Long, cx2 As Long
    Dim cy1 As Long, cy2 As Long
    Dim cells As Long
    Dim total As Long
    Dim statics As Long
    Dim sumSize As Single
    Dim minX As Single, minY As Single
    Dim maxX As Single, maxY As Single
    Dim spanX As Single, spanY As Single
    Dim w As Single, h As Single
    Dim cursor() As Long

    m_gridReady = True
    m_dynCount = 0
    m_gw = 0
    m_gh = 0
    m_stampTick = 0

    minX = 1E+18: minY = 1E+18
    maxX = -1E+18: maxY = -1E+18

    For i = 0 To m_count - 1
        If o_active(i) Then
            g_stamp(i) = 0

            If o_dyn(i) Then
                dyn_idx(m_dynCount) = i
                m_dynCount = m_dynCount + 1
            Else
                statics = statics + 1

                If o_x1(i) < minX Then minX = o_x1(i)
                If o_y1(i) < minY Then minY = o_y1(i)
                If o_x2(i) > maxX Then maxX = o_x2(i)
                If o_y2(i) > maxY Then maxY = o_y2(i)

                w = o_x2(i) - o_x1(i)
                h = o_y2(i) - o_y1(i)
                If h > w Then w = h
                sumSize = sumSize + w
            End If
        End If
    Next i

    If statics <= 0 Then Exit Sub

    m_gx0 = minX
    m_gy0 = minY
    spanX = maxX - minX
    spanY = maxY - minY

    m_cell = sumSize / statics
    If m_cell < 0.5 Then m_cell = 0.5

    m_gw = SideCount(spanX, m_cell)
    m_gh = SideCount(spanY, m_cell)

    '/* A long thin world would otherwise ask for a table with more cells than it has objects. */
    Do While m_gw * m_gh > MAX_GRID_CELLS
        m_cell = m_cell * 2!
        m_gw = SideCount(spanX, m_cell)
        m_gh = SideCount(spanY, m_cell)
    Loop

    cells = m_gw * m_gh
    ReDim g_start(0 To cells)

    '/* A ground plane covers every cell there is. Indexing it would fill each bucket with the one
    ' * object every query was going to reach anyway, so anything that big is swept with the movers
    ' * instead: the index stays about the small things, which is the only thing an index is for. */
    For i = 0 To m_count - 1
        If o_active(i) And Not o_dyn(i) Then
            If CellFootprint(i) > MAX_CELLS_PER_OBJECT Then
                dyn_idx(m_dynCount) = i
                m_dynCount = m_dynCount + 1
            End If
        End If
    Next i

    For i = 0 To m_count - 1
        If o_active(i) And Not o_dyn(i) And CellFootprint(i) <= MAX_CELLS_PER_OBJECT Then
            cx1 = GridX(o_x1(i)): cx2 = GridX(o_x2(i))
            cy1 = GridY(o_y1(i)): cy2 = GridY(o_y2(i))

            For cy = cy1 To cy2
                For cx = cx1 To cx2
                    c = cy * m_gw + cx + 1
                    g_start(c) = g_start(c) + 1
                Next cx
            Next cy
        End If
    Next i

    For c = 0 To cells - 1
        g_start(c + 1) = g_start(c + 1) + g_start(c)
    Next c

    total = g_start(cells)

    If total <= 0 Then
        m_gw = 0
        m_gh = 0
        Exit Sub
    End If

    ReDim g_items(0 To total - 1)
    ReDim cursor(0 To cells - 1)

    For c = 0 To cells - 1
        cursor(c) = g_start(c)
    Next c

    For i = 0 To m_count - 1
        If o_active(i) And Not o_dyn(i) And CellFootprint(i) <= MAX_CELLS_PER_OBJECT Then
            cx1 = GridX(o_x1(i)): cx2 = GridX(o_x2(i))
            cy1 = GridY(o_y1(i)): cy2 = GridY(o_y2(i))

            For cy = cy1 To cy2
                For cx = cx1 To cx2
                    c = cy * m_gw + cx
                    g_items(cursor(c)) = i
                    cursor(c) = cursor(c) + 1
                Next cx
            Next cy
        End If
    Next i
End Sub

'/** @section Rendering */

'/**
' * @brief Draws the whole scene through a renderer, frustum culling, budget spending and painter ordering included.
' * @param rd The renderer to draw with.
' * @return How many objects were submitted.
' * @description Runs in three stages. The budget is spent nearest first, so whatever gets dropped is always the far background and never the wall in front of the player. The survivors are seeded farthest centre first, and then reordered by a real occlusion sort: for every pair that can be separated by an axis-aligned plane, the side the eye sits on decides which one has to be painted second. Guessing from a single depth number cannot do that, which is what makes a small object shine through a wide wall or a ground plane paint over everything standing on it.
' */
Public Function Render(ByVal rd As PRenderer) As Long
    Dim cam As PCamera
    Dim cv As PCanvas
    Dim i As Long
    Dim src As Long
    Dim idx As Long
    Dim nearD As Single
    Dim midD As Single
    Dim spent As Long
    Dim budget As Long
    Dim selCount As Long
    Dim focal As Single
    Dim lodCut As Single
    Dim ex As Single, ey As Single, ez As Single
    Dim cy0 As Single, sy0 As Single, cp As Single, sp As Single
    Dim nearP As Single, farP As Single
    Dim tanH As Single, normH As Single
    Dim tanV As Single, normV As Single
    Dim dx As Single, dy As Single, dz As Single
    Dim flat As Single, vF As Single, vS As Single, vU As Single
    Dim rad As Single
    Dim visible As Boolean

    If rd Is Nothing Then Exit Function

    Set cam = rd.Camera
    Set cv = rd.Canvas
    If cam Is Nothing Or cv Is Nothing Then Exit Function

    '/* Every value the visibility test needs is read once here instead of once per object. They are
    ' * constant for the whole frame, and asking for them through a property is the most expensive way
    ' * to learn a number that is already known. */
    ex = cam.X: ey = cam.Y: ez = cam.Z
    cy0 = cam.CosYaw: sy0 = cam.SinYaw
    cp = cam.CosPitch: sp = cam.SinPitch
    nearP = cv.NearPlane: farP = cv.FarPlane
    tanH = cv.TanFovH: normH = cv.NormH
    tanV = cv.TanFovV: normV = cv.NormV

    d_count = 0
    focal = cv.FocalLength
    lodCut = 0!
    If m_lodOn Then lodCut = m_lodSize * 0.5!

    For i = 0 To m_count - 1
        If o_active(i) And Not o_invis(i) Then
            '/* The frustum test written out, so a frame costs no calls at all to decide what it sees. */
            rad = o_radius(i)
            dx = o_cx(i) - ex
            dy = o_cy(i) - ey
            dz = o_cz(i) - ez
            flat = dx * cy0 + dy * sy0
            vF = flat * cp + dz * sp

            visible = (vF + rad >= nearP) And (vF - rad <= farP)

            If visible Then
                vS = -dx * sy0 + dy * cy0
                If (vF * tanH - vS) * normH + rad < 0! Then visible = False
                If visible Then If (vF * tanH + vS) * normH + rad < 0! Then visible = False
            End If

            If visible Then
                vU = dz * cp - flat * sp
                If (vF * tanV - vU) * normV + rad < 0! Then visible = False
                If visible Then If (vF * tanV + vU) * normV + rad < 0! Then visible = False
            End If

            If visible Then
                nearD = vF - rad
                midD = vF
                d_idx(d_count) = i
                d_near(d_count) = nearD
                d_far(d_count) = midD + o_radius(i)
                d_mid(d_count) = midD
                d_ramp(d_count) = (o_kind(i) = pkRamp)

                '/* Projected radius in slide points: the same divide the projection itself does. */
                d_lod(d_count) = False
                If lodCut > 0! And midD > 0.001 Then
                    If o_radius(i) * focal / midD < lodCut Then
                        d_lod(d_count) = True
                    End If
                End If

                If o_zTop(i) > o_zHigh(i) Then
                    d_z2(d_count) = o_zTop(i)
                    d_z1(d_count) = o_zHigh(i) - o_thick(i)
                Else
                    d_z2(d_count) = o_zHigh(i)
                    d_z1(d_count) = o_zTop(i) - o_thick(i)
                End If

                d_count = d_count + 1
            End If
        End If
    Next i

    If d_count <= 0 Then Exit Function

    '/* Stage one: closest first, so the budget is spent on what the player is actually looking at. */
    If d_count > 1 Then SortNearFirst 0, d_count - 1

    budget = rd.PolyBudget
    selCount = 0

    For i = 0 To d_count - 1
        If d_lod(i) Then
            spent = spent + LodCost(d_idx(i))
        Else
            spent = spent + o_cost(d_idx(i))
        End If

        If spent > budget Then Exit For
        selCount = selCount + 1
    Next i

    If selCount <= 0 Then selCount = 1

    '/* Stage two: seed the painter order with the farthest centre first. */
    If selCount > 1 Then SortMidLast 0, selCount - 1

    '/* Stage three: work out what each survivor covers on the canvas. Only the survivors, and only
    ' * now, because this is the one consumer and the budget has already thrown the rest away. */
    For i = 0 To selCount - 1
        ScreenRect i, ex, ey, ez, cy0, sy0, cp, sp, nearP, focal, cv.CenterX, cv.CenterY
    Next i

    '/* Stage four: refine that seed into a real occlusion order. */
    BuildDrawOrder selCount, cam.X, cam.Y, cam.Z

    m_drawn = selCount

    For i = 0 To selCount - 1
        src = t_order(i)
        idx = d_idx(src)

        DrawObject rd, idx, d_lod(src)
        Render = Render + 1
    Next i
End Function

'/**
' * @brief Reads how many objects the last Render painted.
' * @return The count, which is what Render returned.
' */
Public Property Get DrawnCount() As Long
    DrawnCount = m_drawn
End Property

'/**
' * @brief Reads which object was painted at a given position of the last frame.
' * @param slot The position, zero being the first thing painted and therefore the furthest back.
' * @return The object id, or P_INVALID_ID when the slot is out of range.
' * @remarks The order the painter's algorithm settled on, which is the one thing about a frame that
' * cannot be checked by looking at the shapes afterwards. Reading it is how a test proves that a coin
' * resting on a platform is painted before the platform when you are underneath it.
' */
Public Function DrawnAt(ByVal slot As Long) As Long
    If slot < 0 Or slot >= m_drawn Then
        DrawnAt = P_INVALID_ID
        Exit Function
    End If

    DrawnAt = d_idx(t_order(slot))
End Function

'/**
' * @brief Draws a single object through a renderer.
' * @param rd The renderer to draw with.
' * @param idx The object id.
' */
Public Sub DrawObject(ByVal rd As PRenderer, ByVal idx As Long, Optional ByVal reduced As Boolean = False)
    If idx < 0 Or idx >= m_count Then Exit Sub

    Select Case o_kind(idx)
        Case pkBox
            If reduced Then
                rd.DrawBoxLod o_mat(idx), o_x1(idx), o_y1(idx), o_x2(idx), o_y2(idx), o_zTop(idx), o_thick(idx)
            Else
                rd.DrawBox o_mat(idx), o_x1(idx), o_y1(idx), o_x2(idx), o_y2(idx), o_zTop(idx), o_thick(idx)
            End If

        Case pkRotatedBox
            rd.DrawBoxRotated o_mat(idx), o_cx(idx), o_cy(idx), o_sizeA(idx), o_sizeB(idx), _
                                     o_zTop(idx), o_thick(idx), o_angle(idx)

        Case pkRamp
            rd.DrawRamp o_mat(idx), o_x1(idx), o_y1(idx), o_x2(idx), o_y2(idx), _
                               o_zTop(idx), o_zHigh(idx), o_thick(idx), o_axis(idx)

        Case pkBillboard
            rd.DrawBillboard o_mat(idx), o_cx(idx), o_cy(idx), o_cz(idx), o_sizeA(idx), o_sizeB(idx)

        Case Else
            If reduced Then
                '/* One plate instead of two: the crossed pair only reads as a coin up close. */
                rd.DrawBillboard o_mat(idx), o_cx(idx), o_cy(idx), o_cz(idx), o_sizeA(idx) * 2!, o_sizeA(idx) * 2!
            Else
                rd.DrawSpinner o_mat(idx), o_cx(idx), o_cy(idx), o_cz(idx), o_sizeA(idx), o_phase(idx)
            End If
    End Select
End Sub

'/**
' * @brief Reads what an object costs once reduced.
' * @param idx The object id.
' * @return The polygon count of the reduced form.
' */
Private Function LodCost(ByVal idx As Long) As Long
    Select Case o_kind(idx)
        Case pkBox:     LodCost = 1
        Case pkSpinner: LodCost = 1
        Case Else:      LodCost = o_cost(idx)
    End Select
End Function

'/** @section Private helpers */

'/**
' * @brief Claims the next object slot, growing the arrays when they fill.
' * @param kind The primitive the object is drawn as.
' * @param matId The material the object carries.
' * @param cost The polygon cost charged when the object is drawn.
' * @return The id of the claimed slot.
' */
Private Function NewSlot(ByVal kind As PObjectKind, ByVal matId As Long, ByVal cost As Long) As Long
    Dim i As Long

    '/* A slot freed by Remove is handed out again before the store grows. Without this a level that
    ' * creates and destroys, a projectile or an item that respawns, leaks a slot every time and the
    ' * arrays climb forever even though the object count never does. */
    If f_count > 0 Then
        f_count = f_count - 1
        i = f_free(f_count)
    Else
        If m_count >= m_capacity Then Reserve m_count + 1
        i = m_count
        m_count = m_count + 1
    End If

    f_gone(i) = False
    o_dyn(i) = False
    m_gridReady = False
    m_live = m_live + 1

    o_kind(i) = kind
    o_active(i) = True
    o_mat(i) = matId
    o_cost(i) = cost

    o_x1(i) = 0!
    o_y1(i) = 0!
    o_x2(i) = 0!
    o_y2(i) = 0!
    o_zTop(i) = 0!
    o_zHigh(i) = 0!
    o_thick(i) = 0.01!
    o_cx(i) = 0!
    o_cy(i) = 0!
    o_cz(i) = 0!
    o_radius(i) = 0!
    o_angle(i) = 0!
    o_phase(i) = 0!
    o_axis(i) = paX
    o_sizeA(i) = 0!
    o_sizeB(i) = 0!
    o_invis(i) = Not PMaterials.ById(matId).Visible
    mo_mode(i) = 0
    mo_t(i) = 0!
    mo_last(i) = 0!
    o_tag(i) = 0

    NewSlot = i
End Function

'/**
' * @brief Recomputes the centre and bounding sphere of an object after its shape changes.
' * @param i The object id.
' */
Private Sub RefreshBounds(ByVal i As Long)
    Dim halfX As Single
    Dim halfY As Single
    Dim halfZ As Single
    Dim topZ As Single
    Dim botZ As Single
    Dim reach As Single
    Dim ca As Single
    Dim sa As Single

    Select Case o_kind(i)
        Case pkRotatedBox
            '/* The box that actually encloses the slab at the angle it is at, not the circle it would
            ' * sweep if it spun. The sweep circle is angle free and therefore cheaper, but it claims
            ' * ground the slab does not occupy, and every separating plane test that could have settled
            ' * an ordering against a neighbour fails on ground that is not really there. */
            ca = Abs(Cos(o_angle(i)))
            sa = Abs(Sin(o_angle(i)))
            reach = o_sizeA(i) * ca + o_sizeB(i) * sa
            o_x1(i) = o_cx(i) - reach
            o_x2(i) = o_cx(i) + reach
            reach = o_sizeA(i) * sa + o_sizeB(i) * ca
            o_y1(i) = o_cy(i) - reach
            o_y2(i) = o_cy(i) + reach

        Case pkBillboard, pkSpinner
            reach = PCore.MaxOf(o_sizeA(i), o_sizeB(i)) * 0.5!
            If o_kind(i) = pkSpinner Then reach = o_sizeA(i)
            o_x1(i) = o_cx(i) - reach
            o_x2(i) = o_cx(i) + reach
            o_y1(i) = o_cy(i) - reach
            o_y2(i) = o_cy(i) + reach
            o_zTop(i) = o_cz(i) + reach
            o_zHigh(i) = o_zTop(i)
            o_thick(i) = reach * 2!
    End Select

    topZ = PCore.MaxOf(o_zTop(i), o_zHigh(i))
    botZ = PCore.MinOf(o_zTop(i), o_zHigh(i)) - o_thick(i)

    halfX = (o_x2(i) - o_x1(i)) * 0.5!
    halfY = (o_y2(i) - o_y1(i)) * 0.5!
    halfZ = (topZ - botZ) * 0.5!

    If o_kind(i) <> pkRotatedBox Then
        o_cx(i) = (o_x1(i) + o_x2(i)) * 0.5!
        o_cy(i) = (o_y1(i) + o_y2(i)) * 0.5!
    End If

    o_cz(i) = (topZ + botZ) * 0.5!
    o_radius(i) = Sqr(halfX * halfX + halfY * halfY + halfZ * halfZ)
End Sub

'/**
' * @brief Orders the draw list by the near edge of each bounding sphere, closest object first.
' * @param first The first slot of the range to sort.
' * @param last The last slot of the range to sort.
' * @remarks A plain quicksort over the parallel arrays, which beats moving object records around.
' */
Private Sub SortNearFirst(ByVal first As Long, ByVal last As Long)
    Dim i As Long
    Dim j As Long
    Dim pivot As Single

    i = first
    j = last
    pivot = d_near((first + last) \ 2)

    Do While i <= j
        Do While d_near(i) < pivot
            i = i + 1
        Loop

        Do While d_near(j) > pivot
            j = j - 1
        Loop

        If i <= j Then
            SwapEntries i, j
            i = i + 1
            j = j - 1
        End If
    Loop

    If first < j Then SortNearFirst first, j
    If i < last Then SortNearFirst i, last
End Sub

'/**
' * @brief Orders the draw list by centre depth, farthest object first.
' * @param first The first slot of the range to sort.
' * @param last The last slot of the range to sort.
' * @remarks This is only the seed handed to BuildDrawOrder; it decides the order of pairs that no separating plane can rank.
' */
Private Sub SortMidLast(ByVal first As Long, ByVal last As Long)
    Dim i As Long
    Dim j As Long
    Dim pivot As Single

    i = first
    j = last
    pivot = d_mid((first + last) \ 2)

    Do While i <= j
        Do While d_mid(i) > pivot
            i = i + 1
        Loop

        Do While d_mid(j) < pivot
            j = j - 1
        Loop

        If i <= j Then
            SwapEntries i, j
            i = i + 1
            j = j - 1
        End If
    Loop

    If first < j Then SortMidLast first, j
    If i < last Then SortMidLast i, last
End Sub

'/**
' * @brief Decides which of two objects has to be painted first.
' * @param pa The first slot of the draw list.
' * @param pb The second slot of the draw list.
' * @param eyeX The camera X coordinate.
' * @param eyeY The camera Y coordinate.
' * @param eyeZ The camera Z coordinate.
' * @return 1 when the first object must be painted before the second, -1 for the opposite, 0 when neither can be proven to be behind the other.
' * @description Looks for an axis-aligned plane that separates the two boxes. Once one exists, the answer is exact and free of depth guesswork: whichever box sits on the same side as the eye is the near one, so it is painted last. Boxes that overlap on all three axes cannot be ranked this way and fall back to the seed order.
' * @remarks The height test is taken over the patch of ground the two objects share, not over their full bounding boxes. A ramp is tall at one end and low at the other, so its box reaches far above a coin resting on its low end even though the ramp surface there is well below it; measuring inside the shared window is what separates the pair.
' */
Private Function ItemOrder(ByVal pa As Long, ByVal pb As Long, _
                           ByVal eyeX As Single, ByVal eyeY As Single, ByVal eyeZ As Single) As Long
    Dim a As Long
    Dim b As Long
    Dim s As Single
    Dim wx1 As Single, wy1 As Single, wx2 As Single, wy2 As Single
    Dim az1 As Single, az2 As Single
    Dim bz1 As Single, bz2 As Single

    a = d_idx(pa)
    b = d_idx(pb)

    If o_x2(a) <= o_x1(b) + SEPARATION_EPS Then
        s = (o_x2(a) + o_x1(b)) * 0.5!
        If eyeX > s Then ItemOrder = 1 Else ItemOrder = -1
        Exit Function
    End If

    If o_x2(b) <= o_x1(a) + SEPARATION_EPS Then
        s = (o_x2(b) + o_x1(a)) * 0.5!
        If eyeX > s Then ItemOrder = -1 Else ItemOrder = 1
        Exit Function
    End If

    If o_y2(a) <= o_y1(b) + SEPARATION_EPS Then
        s = (o_y2(a) + o_y1(b)) * 0.5!
        If eyeY > s Then ItemOrder = 1 Else ItemOrder = -1
        Exit Function
    End If

    If o_y2(b) <= o_y1(a) + SEPARATION_EPS Then
        s = (o_y2(b) + o_y1(a)) * 0.5!
        If eyeY > s Then ItemOrder = -1 Else ItemOrder = 1
        Exit Function
    End If

    '/* They overlap on both ground axes, so the shared window is the intersection of their footprints. */
    If d_ramp(pa) Or d_ramp(pb) Then
        If o_x1(a) > o_x1(b) Then wx1 = o_x1(a) Else wx1 = o_x1(b)
        If o_x2(a) < o_x2(b) Then wx2 = o_x2(a) Else wx2 = o_x2(b)
        If o_y1(a) > o_y1(b) Then wy1 = o_y1(a) Else wy1 = o_y1(b)
        If o_y2(a) < o_y2(b) Then wy2 = o_y2(a) Else wy2 = o_y2(b)

        SpanZWithin pa, wx1, wy1, wx2, wy2, az1, az2
        SpanZWithin pb, wx1, wy1, wx2, wy2, bz1, bz2
    Else
        az1 = d_z1(pa)
        az2 = d_z2(pa)
        bz1 = d_z1(pb)
        bz2 = d_z2(pb)
    End If

    If az2 <= bz1 + SEPARATION_EPS Then
        s = (az2 + bz1) * 0.5!
        If eyeZ > s Then ItemOrder = 1 Else ItemOrder = -1
        Exit Function
    End If

    If bz2 <= az1 + SEPARATION_EPS Then
        s = (bz2 + az1) * 0.5!
        If eyeZ > s Then ItemOrder = -1 Else ItemOrder = 1
        Exit Function
    End If

    '/* Still tied. If one bounding sphere ends before the other begins, depth alone settles it:
    ' * along any viewing ray that one is farther. This is checked after the planes, never before,
    ' * because the sphere is a coarse stand in for the box and its verdict contradicts the exact
    ' * plane test often enough to poison the constraint graph with cycles. */
    If d_near(pa) >= d_far(pb) Then
        ItemOrder = 1
        Exit Function
    End If

    If d_near(pb) >= d_far(pa) Then
        ItemOrder = -1
        Exit Function
    End If

    '/* Nothing separates them, which means the two shapes genuinely intersect: a coin sunk a few
    ' * centimetres into the ramp it rests on, a slab embedded in the ground. No plane answers this, so
    ' * the centres do: the one whose middle is further from the eye is painted first. Measured against
    ' * the alternative of painting the larger one first, depth is the better guess, and size stays on
    ' * as the tie break for the case where the two centres sit at the same distance. */
    If d_mid(pa) > d_mid(pb) + 0.001 Then
        ItemOrder = 1
    ElseIf d_mid(pb) > d_mid(pa) + 0.001 Then
        ItemOrder = -1
    ElseIf o_radius(a) > o_radius(b) + 0.01 Then
        ItemOrder = 1
    ElseIf o_radius(b) > o_radius(a) + 0.01 Then
        ItemOrder = -1
    End If
End Function

'/**
' * @brief Works out the canvas rectangle one draw list entry covers.
' * @param slot The position in the draw list.
' * @param ex The eye X coordinate.
' * @param ey The eye Y coordinate.
' * @param ez The eye Z coordinate.
' * @param cy0 The cosine of the camera yaw.
' * @param sy0 The sine of the camera yaw.
' * @param cp The cosine of the camera pitch.
' * @param sp The sine of the camera pitch.
' * @param nearP The near plane distance.
' * @param focal The focal length in slide points.
' * @param ctrX The canvas centre X coordinate.
' * @param ctrY The canvas centre Y coordinate.
' * @description The eight corners of the bounding box, projected, and the rectangle that holds them.
' * It is the bounding box rather than the true shape, so the rectangle is always a little generous,
' * which is the safe direction: an ordering that is kept needlessly costs a comparison, while one
' * dropped by mistake costs a visible artefact. A box with a corner behind the eye has no bounded
' * rectangle at all and is marked as covering everything.
' */
Private Sub ScreenRect(ByVal slot As Long, _
                       ByVal ex As Single, ByVal ey As Single, ByVal ez As Single, _
                       ByVal cy0 As Single, ByVal sy0 As Single, _
                       ByVal cp As Single, ByVal sp As Single, _
                       ByVal nearP As Single, ByVal focal As Single, _
                       ByVal ctrX As Single, ByVal ctrY As Single)
    Dim idx As Long
    Dim cx As Long, cy As Long, cz As Long
    Dim wx As Single, wy As Single, wz As Single
    Dim dx As Single, dy As Single, dz As Single
    Dim flat As Single, vF As Single, vS As Single, vU As Single
    Dim px As Single, py As Single
    Dim inv As Single
    Dim topZ As Single, botZ As Single
    Dim x1 As Single, y1 As Single, x2 As Single, y2 As Single

    idx = d_idx(slot)

    topZ = PCore.MaxOf(o_zTop(idx), o_zHigh(idx))
    botZ = PCore.MinOf(o_zTop(idx), o_zHigh(idx)) - o_thick(idx)

    x1 = 1E+18: y1 = 1E+18
    x2 = -1E+18: y2 = -1E+18

    For cx = 0 To 1
        If cx = 0 Then wx = o_x1(idx) Else wx = o_x2(idx)

        For cy = 0 To 1
            If cy = 0 Then wy = o_y1(idx) Else wy = o_y2(idx)

            For cz = 0 To 1
                If cz = 0 Then wz = botZ Else wz = topZ

                dx = wx - ex
                dy = wy - ey
                dz = wz - ez
                flat = dx * cy0 + dy * sy0
                vF = flat * cp + dz * sp

                If vF < nearP Then
                    d_rx1(slot) = -1E+18
                    d_ry1(slot) = -1E+18
                    d_rx2(slot) = 1E+18
                    d_ry2(slot) = 1E+18
                    Exit Sub
                End If

                vS = -dx * sy0 + dy * cy0
                vU = dz * cp - flat * sp

                inv = focal / vF
                px = ctrX + vS * inv
                py = ctrY - vU * inv

                If px < x1 Then x1 = px
                If px > x2 Then x2 = px
                If py < y1 Then y1 = py
                If py > y2 Then y2 = py
            Next cz
        Next cy
    Next cx

    d_rx1(slot) = x1
    d_ry1(slot) = y1
    d_rx2(slot) = x2
    d_ry2(slot) = y2
End Sub

'/**
' * @brief Measures the height an object occupies over one patch of ground.
' * @param slot The draw list slot holding the object.
' * @param wx1 The lower X bound of the window.
' * @param wy1 The lower Y bound of the window.
' * @param wx2 The upper X bound of the window.
' * @param wy2 The upper Y bound of the window.
' * @param outBottom Receives the lowest point inside the window.
' * @param outTop Receives the highest point inside the window.
' * @remarks Only a ramp answers differently from its bounding box; every other primitive has the same height everywhere, so it returns the span cached when the draw list was gathered.
' */
Private Sub SpanZWithin(ByVal slot As Long, _
                        ByVal wx1 As Single, ByVal wy1 As Single, ByVal wx2 As Single, ByVal wy2 As Single, _
                        ByRef outBottom As Single, ByRef outTop As Single)
    Dim idx As Long
    Dim za As Single
    Dim zb As Single

    idx = d_idx(slot)

    If o_kind(idx) <> pkRamp Then
        outBottom = d_z1(slot)
        outTop = d_z2(slot)
        Exit Sub
    End If

    za = TopZAt(idx, wx1, wy1)
    zb = TopZAt(idx, wx2, wy2)

    If za > zb Then
        outTop = za
        outBottom = zb - o_thick(idx)
    Else
        outTop = zb
        outBottom = za - o_thick(idx)
    End If
End Sub

'/**
' * @brief Turns the pairwise ordering constraints into one painting sequence.
' * @param n How many objects are in the selection.
' * @param eyeX The camera X coordinate.
' * @param eyeY The camera Y coordinate.
' * @param eyeZ The camera Z coordinate.
' * @description Every proven constraint becomes an edge, and a topological sort walks them into an order that satisfies all of them at once. Objects tangled in a cycle, or left over when the edge store fills up, keep their seed position instead of being dropped.
' */
Private Sub BuildDrawOrder(ByVal n As Long, ByVal eyeX As Single, ByVal eyeY As Single, ByVal eyeZ As Single)
    Dim a As Long
    Dim b As Long
    Dim i As Long
    Dim j As Long
    Dim rel As Long
    Dim qHead As Long
    Dim qTail As Long
    Dim pos As Long
    Dim pick As Long
    Dim e As Long

    For i = 0 To n - 1
        t_indeg(i) = 0
        t_adjHead(i) = -1
        t_placed(i) = False
        t_order(i) = i
    Next i

    If n <= 1 Then Exit Sub

    EnsureEdges n
    e_count = 0

    For a = 0 To n - 2
        For b = a + 1 To n - 1
            '/* Only objects that share canvas space need to be ordered against each other. Skipping
            ' * the rest is not just cheaper: an ordering between two objects that never touch on
            ' * screen can never be seen, yet it can still close a cycle, and the cycle then has to be
            ' * broken somewhere that can be seen. That is the whole bug: a coin resting on a platform
            ' * was being sacrificed to satisfy a constraint against a slab on the far side of the
            ' * level that shared not one pixel with either of them. */
            If d_rx2(a) < d_rx1(b) Or d_rx2(b) < d_rx1(a) _
               Or d_ry2(a) < d_ry1(b) Or d_ry2(b) < d_ry1(a) Then
                rel = 0
            Else
                rel = ItemOrder(a, b, eyeX, eyeY, eyeZ)
            End If

            If rel <> 0 And e_count < e_cap Then
                If rel > 0 Then
                    e_to(e_count) = b
                    e_next(e_count) = t_adjHead(a)
                    t_adjHead(a) = e_count
                    t_indeg(b) = t_indeg(b) + 1
                Else
                    e_to(e_count) = a
                    e_next(e_count) = t_adjHead(b)
                    t_adjHead(b) = e_count
                    t_indeg(a) = t_indeg(a) + 1
                End If

                e_count = e_count + 1
            End If
        Next b
    Next a

    qHead = 0
    qTail = 0

    For i = 0 To n - 1
        If t_indeg(i) = 0 Then
            t_queue(qTail) = i
            qTail = qTail + 1
            t_placed(i) = True
        End If
    Next i

    pos = 0

    Do While pos < n
        If qHead < qTail Then
            i = t_queue(qHead)
            qHead = qHead + 1
        Else
            '/* A cycle: three objects can each be proven to sit behind the next, and no order
            ' * satisfies all of it, so one constraint has to be dropped. Drop as few as possible,
            ' * which means force placing the object that still owes the fewest, and when several owe
            ' * the same number take the furthest one, because a tangle no rule can order exactly is
            ' * still best painted back to front. Scoring by the size of the blocker instead, which is
            ' * what this did before, reliably sacrificed whatever was small: a coin sitting on a
            ' * platform has exactly one small thing in front of it and was picked every time. */
            pick = -1

            For j = 0 To n - 1
                If Not t_placed(j) Then
                    If pick < 0 Then
                        pick = j
                    ElseIf t_indeg(j) < t_indeg(pick) Then
                        pick = j
                    ElseIf t_indeg(j) = t_indeg(pick) Then
                        If d_mid(j) > d_mid(pick) Then pick = j
                    End If
                End If
            Next j

            If pick < 0 Then Exit Do

            i = pick
            t_placed(i) = True
        End If

        t_queue(n + pos) = i
        pos = pos + 1

        e = t_adjHead(i)

        Do While e >= 0
            j = e_to(e)
            t_indeg(j) = t_indeg(j) - 1

            If t_indeg(j) <= 0 And Not t_placed(j) Then
                t_queue(qTail) = j
                qTail = qTail + 1
                t_placed(j) = True
            End If

            e = e_next(e)
        Loop
    Loop

    If pos < n Then
        For i = 0 To n - 1
            If Not t_placed(i) Then
                t_queue(n + pos) = i
                t_placed(i) = True
                pos = pos + 1
            End If
        Next i
    End If

    For i = 0 To n - 1
        t_order(i) = t_queue(n + i)
    Next i
End Sub

'/**
' * @brief Grows the constraint store to hold the pairs of one frame.
' * @param n How many objects are being ordered.
' */
Private Sub EnsureEdges(ByVal n As Long)
    Dim need As Long

    If n < 2 Then Exit Sub

    need = n * (n - 1) \ 2
    If need > MAX_ORDER_EDGES Then need = MAX_ORDER_EDGES
    If need <= e_cap Then Exit Sub

    ReDim Preserve e_to(0 To need - 1)
    ReDim Preserve e_next(0 To need - 1)
    e_cap = need
End Sub

'/**
' * @brief Swaps two entries of the draw list, keeping every parallel array in step.
' * @param a The first slot.
' * @param b The second slot.
' */
Private Sub SwapEntries(ByVal a As Long, ByVal b As Long)
    Dim tmpIdx As Long
    Dim tmpKey As Single

    tmpIdx = d_idx(a)
    d_idx(a) = d_idx(b)
    d_idx(b) = tmpIdx

    Dim tmpFlag As Boolean

    tmpKey = d_near(a)
    d_near(a) = d_near(b)
    d_near(b) = tmpKey

    tmpKey = d_far(a)
    d_far(a) = d_far(b)
    d_far(b) = tmpKey

    tmpFlag = d_ramp(a)
    d_ramp(a) = d_ramp(b)
    d_ramp(b) = tmpFlag

    tmpFlag = d_lod(a)
    d_lod(a) = d_lod(b)
    d_lod(b) = tmpFlag

    tmpKey = d_mid(a)
    d_mid(a) = d_mid(b)
    d_mid(b) = tmpKey

    tmpKey = d_z1(a)
    d_z1(a) = d_z1(b)
    d_z1(b) = tmpKey

    tmpKey = d_z2(a)
    d_z2(a) = d_z2(b)
    d_z2(b) = tmpKey
End Sub

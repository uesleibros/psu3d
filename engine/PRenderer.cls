VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/**
' * PRenderer - Psu3D Face Pipeline
' * @description Turns world space faces into PowerPoint polylines: backface rejection, view frustum clipping, perspective projection into a PCanvas, material shading and the pooled shape emission that keeps the slide from thrashing. One renderer drives exactly one canvas, so a slide can host several independent views at once.
' * @author UesleiDev
' * @version 1.0
' * @scope Private to this VBA project: class modules carry VB_Exposed = False, which is the class level equivalent of Option Private Module.
' * @remarks Shapes are named from a per renderer prefix and deleted as a single Range on the next frame, which is by far the cheapest way to clear a slide from VBA.
' */

Option Explicit

'/** @section Constants */

'/** @description Largest vertex count a polygon may carry through the clipping stages. */
Private Const MAX_VERTS As Long = 24

'/** @description Squared distance below which two consecutive vertices are merged. */
Private Const WELD_DIST_SQ As Single = 0.0025

'/** @section Instance state */

Private m_shapes As Shapes
Private m_canvas As PCanvas
Private m_camera As PCamera

Private m_prefix As String
Private m_names() As String
Private m_altNames() As String
Private m_frameNames() As Variant
Private m_poolSize As Long
Private m_count As Long

Private m_budget As Long
Private m_minBudget As Long
Private m_maxBudget As Long
Private m_fastFrame As Single
Private m_slowFrame As Single
Private m_budgetDown As Long
Private m_budgetUp As Long
Private m_autoBudget As Boolean
Private m_slowStreak As Long
Private m_fastStreak As Long

Private m_ptsBuf() As Single
Private m_ptsSize As Long

Private m_edgeInflate As Single
Private m_minArea As Single
Private m_farArea As Single
Private m_farDepth As Single
Private m_backEps As Single
Private m_backEpsSq As Single

Private m_eyeX As Single
Private m_eyeY As Single
Private m_eyeZ As Single
Private m_cosYaw As Single
Private m_sinYaw As Single
Private m_cosPitch As Single
Private m_sinPitch As Single
Private m_shiftY As Single

Private m_left As Single
Private m_top As Single
Private m_right As Single
Private m_bottom As Single
Private m_projCX As Single
Private m_projCY As Single
Private m_focal As Single
Private m_tanH As Single
Private m_tanV As Single
Private m_near As Single

Private m_overlay As Boolean
Private m_seamFill As Boolean
Private m_dryRun As Boolean
Private m_liveShapes As Long
Private m_doubleBuffer As Boolean
Private m_bank As Long
Private m_prevBank As Long
Private m_prevCount As Long
Private m_comAcc As Double
Private m_comOps As Long
Private m_profile As Boolean

'/** @description Cached copy of the material in use, refreshed only when the id changes. Reading a
' * dozen properties off a PMaterial for every face is the single most expensive thing this class can
' * do, and consecutive faces almost always share a material, so the cache nearly always hits. */
Private m_matId As Long
Private m_matVisible As Boolean
Private m_matTwoSided As Boolean
Private m_matUnlit As Boolean
Private m_matFogged As Boolean
Private m_matColor As Long
Private m_matFormat As Boolean
Private m_matTransp As Single
Private m_matEdge As Boolean
Private m_matEdgeCol As Long
Private m_matEdgeW As Single

'/** @description Fog settings copied once per frame so the band of a face is worked out inline. */
Private m_fogOn As Boolean
Private m_fogStart As Single
Private m_fogEnd As Single
Private m_fogScale As Single
Private m_fogInv As Single
Private m_fogSteps As Long
Private m_fogR As Long
Private m_fogG As Long
Private m_fogB As Long

'/** @section Lifetime */

'/**
' * @brief Applies the stock tuning so a renderer works before anything is configured.
' */
Private Sub Class_Initialize()
    m_prefix = "p3d_"

    m_minBudget = 56
    m_maxBudget = 160
    m_budget = m_maxBudget
    m_fastFrame = 0.018
    m_slowFrame = 0.03
    m_budgetDown = 6
    m_budgetUp = 4
    m_autoBudget = True

    m_edgeInflate = 0.45
    m_minArea = 1.1
    m_farArea = 2.6
    m_farDepth = 17!
    m_backEps = 0.004
    m_backEpsSq = m_backEps * m_backEps
    m_seamFill = False
    m_doubleBuffer = True
    m_profile = True

    SetPool m_maxBudget
End Sub

'/**
' * @brief Releases the slide and view references held by the renderer.
' */
Private Sub Class_Terminate()
    Set m_shapes = Nothing
    Set m_canvas = Nothing
    Set m_camera = Nothing
End Sub

'/**
' * @brief Binds the renderer to a slide, a canvas and a camera.
' * @param target The Shapes collection of the slide being drawn into.
' * @param cv The canvas describing where and how to project.
' * @param cam The camera describing where the view is taken from.
' * @return The renderer itself, so it can be created and bound in one expression.
' */
Public Function Attach(ByVal target As Shapes, ByVal cv As PCanvas, ByVal cam As PCamera) As PRenderer
    Set m_shapes = target
    Set m_canvas = cv
    Set m_camera = cam

    RefreshState
    Set Attach = Me
End Function

'/** @section Bindings */

'/**
' * @brief Reads the canvas this renderer draws into.
' * @return The bound canvas.
' */
Public Property Get Canvas() As PCanvas
    Set Canvas = m_canvas
End Property

'/**
' * @brief Binds a different canvas.
' * @param value The canvas to draw into.
' */
Public Property Set Canvas(ByVal value As PCanvas)
    Set m_canvas = value
    RefreshState
End Property

'/**
' * @brief Reads the camera driving this renderer.
' * @return The bound camera.
' */
Public Property Get Camera() As PCamera
    Set Camera = m_camera
End Property

'/**
' * @brief Binds a different camera.
' * @param value The camera to render from.
' */
Public Property Set Camera(ByVal value As PCamera)
    Set m_camera = value
    RefreshState
End Property

'/**
' * @brief Reads the shape collection being drawn into.
' * @return The bound Shapes collection.
' */
Public Property Get Target() As Shapes
    Set Target = m_shapes
End Property

'/**
' * @brief Binds a different shape collection, which is how a renderer follows the slide it lives on.
' * @param value The Shapes collection of the target slide.
' */
Public Property Set Target(ByVal value As Shapes)
    Set m_shapes = value
End Property

'/**
' * @brief Reads the prefix given to every shape this renderer creates.
' * @return The shape name prefix.
' */
Public Property Get Prefix() As String
    Prefix = m_prefix
End Property

'/**
' * @brief Sets the prefix given to every shape this renderer creates.
' * @param value The new prefix; give each renderer on a slide its own so cleanup never crosses over.
' */
Public Property Let Prefix(ByVal value As String)
    If Len(value) = 0 Then Exit Property

    m_prefix = value
    RebuildNames
End Property

'/**
' * @brief Reads the eye X coordinate cached for this frame.
' * @return The world X position of the camera.
' * @remarks Geometry helpers read the eye once per object to decide which faces face it. Going through
' * the camera object each time costs a reference fetch plus a property call for a number that was
' * already copied here when the frame opened.
' */
Friend Property Get EyeX() As Single
    EyeX = m_eyeX
End Property

'/**
' * @brief Reads the eye Y coordinate cached for this frame.
' * @return The world Y position of the camera.
' */
Friend Property Get EyeY() As Single
    EyeY = m_eyeY
End Property

'/**
' * @brief Reads the eye Z coordinate cached for this frame.
' * @return The world Z position of the camera.
' */
Friend Property Get EyeZ() As Single
    EyeZ = m_eyeZ
End Property

'/** @section Frame control */

'/**
' * @brief Opens a frame by wiping the shapes drawn last time and re-reading the canvas and camera.
' * @remarks Call once per frame before any draw call; the wipe is a single Range delete, not one delete per shape.
' */
Public Sub BeginFrame()
    If m_doubleBuffer Then
        '/* The frame already on the slide stays there while the next one is built over it, and only
        ' * then is retired. Clearing first is what makes the slide flicker: PowerPoint has no screen
        ' * updating switch, so it is free to repaint between the delete and the last AddPolyline, and
        ' * what it paints in that gap is an empty slide. Two alternating name banks cost nothing but
        ' * a second set of strings, and the slide is never seen half drawn. */
        m_comAcc = 0#
        m_comOps = 0
        m_prevBank = m_bank
        m_prevCount = m_liveShapes
        m_bank = 1 - m_bank
        m_liveShapes = 0
        m_count = 0
    Else
        m_comAcc = 0#
        m_comOps = 0
        ClearShapes
    End If

    RefreshState
End Sub

'/**
' * @brief Retires the frame that was on the slide before this one.
' * @remarks Only does anything while double buffered. Call it once per frame, after every draw call, or the retired frame stays on the slide and the shape count keeps climbing.
' */
Public Sub EndFrame()
    Dim i As Long
    Dim t0 As Double

    If Not m_doubleBuffer Then Exit Sub
    If m_prevCount <= 0 Then Exit Sub
    If m_shapes Is Nothing Then
        m_prevCount = 0
        Exit Sub
    End If

    If m_profile Then t0 = PCore.Seconds()

    If m_prevCount = 1 Then
        m_shapes(BankName(m_prevBank, 0)).Delete
    Else
        ReDim m_frameNames(0 To m_prevCount - 1)
        For i = 0 To m_prevCount - 1
            m_frameNames(i) = BankName(m_prevBank, i)
        Next i
        m_shapes.Range(m_frameNames).Delete
    End If

    If m_profile Then m_comAcc = m_comAcc + (PCore.Seconds() - t0)
    m_prevCount = 0
End Sub

'/**
' * @brief Reports whether the previous frame is kept on the slide while the next one is drawn.
' * @return True when double buffered.
' */
Public Property Get DoubleBuffer() As Boolean
    DoubleBuffer = m_doubleBuffer
End Property

'/**
' * @brief Chooses whether the previous frame is kept while the next is drawn.
' * @param value True to trade a moment of doubled shape count for a slide that is never seen blank.
' * @remarks Turning this off restores the clear-then-draw order and removes the need to call EndFrame.
' */
Public Property Let DoubleBuffer(ByVal value As Boolean)
    If m_doubleBuffer = value Then Exit Property

    m_doubleBuffer = value
    Purge
End Property

'/**
' * @brief Re-reads the cached canvas and camera state.
' * @remarks Only needed when the canvas or camera is moved in the middle of a frame, since BeginFrame already does it.
' */
Public Sub RefreshState()
    If Not m_camera Is Nothing Then
        m_eyeX = m_camera.X
        m_eyeY = m_camera.Y
        m_eyeZ = m_camera.Z
        m_cosYaw = m_camera.CosYaw
        m_sinYaw = m_camera.SinYaw
        m_cosPitch = m_camera.CosPitch
        m_sinPitch = m_camera.SinPitch
        m_shiftY = m_camera.ScreenShiftY
    End If

    If Not m_canvas Is Nothing Then
        m_left = m_canvas.X
        m_top = m_canvas.Y
        m_right = m_canvas.Right
        m_bottom = m_canvas.Bottom
        m_projCX = m_canvas.CenterX
        m_projCY = m_canvas.CenterY
        m_focal = m_canvas.FocalLength
        m_tanH = m_canvas.TanFovH
        m_tanV = m_canvas.TanFovV
        m_near = m_canvas.NearPlane
    End If

    m_fogOn = PLighting.FogEnabled
    m_fogStart = PLighting.FogStart
    m_fogEnd = PLighting.FogEnd
    m_fogSteps = PLighting.FogSteps
    m_fogInv = 1! / (m_fogEnd - m_fogStart)
    m_fogScale = m_fogSteps * m_fogInv
    PCore.ColorUnpack PLighting.FogColor, m_fogR, m_fogG, m_fogB

    '/* Materials may have been edited between frames, so drop the cache and bake the table once here
    ' * rather than testing for it on every face. */
    m_matId = P_INVALID_ID - 1
    PMaterials.Sync
End Sub

'/**
' * @brief Deletes every shape emitted during the current frame.
' */
Public Sub ClearShapes()
    Dim i As Long

    m_count = 0

    If m_liveShapes <= 0 Then Exit Sub
    If m_shapes Is Nothing Then
        m_liveShapes = 0
        Exit Sub
    End If

    If m_liveShapes = 1 Then
        m_shapes(BankName(m_bank, 0)).Delete
    Else
        ReDim m_frameNames(0 To m_liveShapes - 1)
        For i = 0 To m_liveShapes - 1
            m_frameNames(i) = BankName(m_bank, i)
        Next i
        m_shapes.Range(m_frameNames).Delete
    End If

    m_liveShapes = 0
End Sub

'/**
' * @brief Deletes every shape on the slide that carries this renderer prefix, including leftovers from an interrupted run.
' * @remarks Worth calling once at start up, since a show stopped mid frame leaves its polygons behind.
' */
Public Sub Purge()
    Dim s As Shape
    Dim names() As Variant
    Dim total As Long
    Dim i As Long
    Dim plen As Long

    m_count = 0
    m_liveShapes = 0
    m_prevCount = 0
    If m_shapes Is Nothing Then Exit Sub

    plen = Len(m_prefix)

    For Each s In m_shapes
        If Left$(s.Name, plen) = m_prefix Then total = total + 1
    Next s

    If total <= 0 Then Exit Sub

    ReDim names(0 To total - 1)
    i = 0

    For Each s In m_shapes
        If Left$(s.Name, plen) = m_prefix Then
            names(i) = s.Name
            i = i + 1
        End If
    Next s

    If total = 1 Then
        m_shapes(names(0)).Delete
    Else
        m_shapes.Range(names).Delete
    End If
End Sub

'/** @section Budget */

'/**
' * @brief Reserves shape name slots up front so no growth happens mid frame.
' * @param maxPolys How many polygon slots to reserve.
' * @remarks The pool is not a ceiling: it grows on demand. Reserving simply moves the allocation out of the render loop.
' */
Public Sub SetPool(ByVal maxPolys As Long)
    If maxPolys < 8 Then maxPolys = 8
    EnsureNames maxPolys
End Sub

'/**
' * @brief Sets the range the adaptive budget may travel inside.
' * @param minPolys The floor the budget never drops under, protecting the silhouette of the scene.
' * @param maxPolys The ceiling the budget never climbs over.
' */
Public Sub SetBudgetRange(ByVal minPolys As Long, ByVal maxPolys As Long)
    If maxPolys < 1 Then maxPolys = 1

    EnsureNames maxPolys

    m_minBudget = PCore.ClampLong(minPolys, 1, maxPolys)
    m_maxBudget = maxPolys
    m_budget = PCore.ClampLong(m_budget, m_minBudget, m_maxBudget)
End Sub

'/**
' * @brief Sets the frame times the adaptive budget aims between.
' * @param fastSeconds Frames quicker than this earn more polygons.
' * @param slowSeconds Frames slower than this give polygons back.
' */
Public Sub SetFrameTargets(ByVal fastSeconds As Single, ByVal slowSeconds As Single)
    m_fastFrame = fastSeconds
    m_slowFrame = slowSeconds
End Sub

'/**
' * @brief Nudges the polygon budget towards the frame time targets.
' * @param dt The duration of the frame just finished, in seconds.
' * @description Moves only after several frames agree, because the budget decides where the draw list is
' * cut, and a budget that changes every frame moves that cut every frame: objects sitting on the boundary
' * then blink in and out while the camera is barely moving. A single slow frame is noise, not a trend.
' * @remarks Does nothing while AutoBudget is False, which is the right setting for a hand authored scene whose worst case fits the budget.
' */
Public Sub AdaptBudget(ByVal dt As Single)
    If Not m_autoBudget Then Exit Sub

    If dt > m_slowFrame Then
        m_slowStreak = m_slowStreak + 1
        m_fastStreak = 0
    ElseIf dt < m_fastFrame Then
        m_fastStreak = m_fastStreak + 1
        m_slowStreak = 0
    Else
        m_slowStreak = 0
        m_fastStreak = 0
    End If

    If m_slowStreak >= 3 Then
        m_budget = m_budget - m_budgetDown
        m_slowStreak = 0
    ElseIf m_fastStreak >= 10 Then
        m_budget = m_budget + m_budgetUp
        m_fastStreak = 0
    End If

    m_budget = PCore.ClampLong(m_budget, m_minBudget, m_maxBudget)
End Sub

'/**
' * @brief Reports whether the budget follows the frame time.
' * @return True when AdaptBudget is allowed to move it.
' */
Public Property Get AutoBudget() As Boolean
    AutoBudget = m_autoBudget
End Property

'/**
' * @brief Chooses whether the budget follows the frame time.
' * @param value False to pin the budget where it is, which removes the last source of objects popping in and out.
' */
Public Property Let AutoBudget(ByVal value As Boolean)
    m_autoBudget = value
    m_slowStreak = 0
    m_fastStreak = 0
End Property

'/**
' * @brief Reads how many polygons may still be drawn this frame.
' * @return The remaining budget.
' */
Public Property Get Remaining() As Long
    Remaining = m_budget - m_count
End Property

'/**
' * @brief Reports whether the frame budget is spent.
' * @return True when further draw calls will be ignored.
' */
Public Property Get IsFull() As Boolean
    IsFull = (m_count >= m_budget)
End Property

'/**
' * @brief Reads how many polygons have been emitted this frame.
' * @return The current shape count.
' */
Public Property Get PolyCount() As Long
    PolyCount = m_count
End Property

'/**
' * @brief Reads the current polygon budget.
' * @return The number of polygons allowed this frame.
' */
Public Property Get PolyBudget() As Long
    PolyBudget = m_budget
End Property

'/**
' * @brief Overrides the polygon budget, pinning it until AdaptBudget moves it again.
' * @param value The new budget.
' */
Public Property Let PolyBudget(ByVal value As Long)
    If value < 1 Then value = 1

    m_budget = value
    If m_budget > m_maxBudget Then m_maxBudget = m_budget
    EnsureNames m_budget
End Property

'/**
' * @brief Reads how many shape name slots are currently reserved.
' * @return The reserved slot count, which grows on demand.
' */
Public Property Get PoolSize() As Long
    PoolSize = m_poolSize
End Property

'/** @section Quality tuning */

'/**
' * @brief Reads how long this frame spent inside PowerPoint.
' * @return Seconds spent creating, formatting and deleting shapes.
' * @remarks Compare against the whole frame time and the answer to where the frame went stops being a guess. Everything not counted here is arithmetic the engine itself did.
' */
Public Property Get ComSeconds() As Double
    ComSeconds = m_comAcc
End Property

'/**
' * @brief Reads how many shapes this frame created.
' * @return The emission count.
' */
Public Property Get ComOps() As Long
    ComOps = m_comOps
End Property

'/**
' * @brief Reports whether COM timing is being collected.
' * @return True when profiling.
' */
Public Property Get Profiling() As Boolean
    Profiling = m_profile
End Property

'/**
' * @brief Chooses whether COM timing is collected.
' * @param value False to drop the two clock reads per polygon, which cost well under a percent of a frame.
' */
Public Property Let Profiling(ByVal value As Boolean)
    m_profile = value
End Property

'/**
' * @brief Reports whether the pipeline is running without touching the slide.
' * @return True while shape creation is suppressed.
' */
Public Property Get DryRun() As Boolean
    DryRun = m_dryRun
End Property

'/**
' * @brief Runs the whole pipeline but creates no shapes, so the cost of the slide can be measured.
' * @param value True to suppress shape creation.
' * @description Culling, clipping, projection, shading and ordering all still run and PolyCount still
' * reports what would have been drawn; only the COM calls are skipped. Time a run with this on and a
' * run with it off, and the difference is what PowerPoint itself costs, which is the number worth
' * knowing before optimising anything else. Toggle it between frames, never inside one.
' */
Public Property Let DryRun(ByVal value As Boolean)
    m_dryRun = value
End Property

'/**
' * @brief Reports how the gap between neighbouring faces is closed.
' * @return True when the outline is painted in the fill colour, False when polygons are inflated instead.
' */
Public Property Get SeamFill() As Boolean
    SeamFill = m_seamFill
End Property

'/**
' * @brief Chooses how the gap between neighbouring faces is closed.
' * @param value True to stroke each polygon in its own fill colour, False to inflate its outline geometrically.
' * @description Two ways to hide the hairline seam PowerPoint leaves between adjacent polygons. Inflating
' * pushes every vertex away from the centre, which distorts small polygons badly because the push is a
' * fixed number of points regardless of how big the polygon is. Stroking the outline in the fill colour
' * expands the shape uniformly by half a line weight, costs exactly the same single COM call, and leaves
' * the geometry untouched. Set EdgeInflate to zero when turning this on, or the two stack.
' */
Public Property Let SeamFill(ByVal value As Boolean)
    m_seamFill = value
End Property

'/**
' * @brief Reads how far polygon borders are pushed outwards.
' * @return The inflation in points.
' */
Public Property Get EdgeInflate() As Single
    EdgeInflate = m_edgeInflate
End Property

'/**
' * @brief Sets how far polygon borders are pushed outwards, which hides the hairline gaps between neighbouring faces.
' * @param value The inflation in points; around half a point is usually enough.
' */
Public Property Let EdgeInflate(ByVal value As Single)
    m_edgeInflate = value
End Property

'/**
' * @brief Reads the smallest screen area a polygon may cover.
' * @return The area threshold in square points.
' */
Public Property Get MinPolyArea() As Single
    MinPolyArea = m_minArea
End Property

'/**
' * @brief Sets the smallest screen area a polygon may cover before it is dropped.
' * @param value The area threshold in square points.
' */
Public Property Let MinPolyArea(ByVal value As Single)
    m_minArea = value
End Property

'/**
' * @brief Sets the stricter area threshold applied to distant polygons.
' * @param area The area threshold in square points.
' * @param depth The view depth beyond which the stricter threshold applies.
' */
Public Sub SetFarCulling(ByVal area As Single, ByVal depth As Single)
    m_farArea = area
    m_farDepth = depth
End Sub

'/** @section Drawing */

'/**
' * @brief Draws a single world space face.
' * @param f The face to draw, wound counter-clockwise when seen from its visible side.
' * @param offX An offset added to every vertex X, which lets one face definition be reused across instances.
' * @param offY An offset added to every vertex Y.
' * @param offZ An offset added to every vertex Z.
' * @remarks The face is rejected early when it faces away, falls outside the frustum, or the frame budget is already spent.
' * @remarks Declared Friend, not Public: VBA refuses a user defined type from a standard module in the public signature of a class. Friend keeps it callable from every module of this project, which is the only scope the engine wants anyway. DrawQuad and DrawTriangle are the Public doors into the same pipeline.
' */
Friend Sub DrawFace(ByRef f As PFace, _
                    Optional ByVal offX As Single = 0!, _
                    Optional ByVal offY As Single = 0!, _
                    Optional ByVal offZ As Single = 0!)
    Dim n As Long
    Dim i As Long
    Dim plane As Long

    Dim wx(0 To 3) As Single
    Dim wy(0 To 3) As Single
    Dim wz(0 To 3) As Single

    Dim cx As Single, cy As Single, cz As Single
    Dim e1x As Single, e1y As Single, e1z As Single
    Dim e2x As Single, e2y As Single, e2z As Single
    Dim nx As Single, ny As Single, nz As Single
    Dim nLen As Single
    Dim nSq As Single
    Dim facing As Single

    Dim viewDX As Single, viewDY As Single, relZ As Single, flat As Single

    Dim pF(0 To MAX_VERTS) As Single
    Dim pS(0 To MAX_VERTS) As Single
    Dim pU(0 To MAX_VERTS) As Single
    Dim tF(0 To MAX_VERTS) As Single
    Dim tS(0 To MAX_VERTS) As Single
    Dim tU(0 To MAX_VERTS) As Single
    Dim dist(0 To MAX_VERTS) As Single
    Dim pCnt As Long
    Dim tCnt As Long

    Dim cF As Single, cS As Single, cU As Single, cOff As Single
    Dim anyOut As Boolean, allOut As Boolean
    Dim prevIdx As Long
    Dim d0 As Single, d1 As Single, t As Single

    Dim dirKey As Long
    Dim dynCol As Long

    If m_count >= m_budget Then Exit Sub
    If m_canvas Is Nothing Or m_camera Is Nothing Then Exit Sub

    '/* A dry run has no slide to draw onto and does not need one, which is what lets the whole
    ' * pipeline be exercised with nothing but objects in memory. */
    If m_shapes Is Nothing And Not m_dryRun Then Exit Sub

    n = f.nVerts
    If n < 3 Or n > 4 Then Exit Sub

    If f.matId <> m_matId Then RefreshMaterial f.matId
    If Not m_matVisible Then Exit Sub

    For i = 0 To n - 1
        wx(i) = f.vx(i) + offX
        wy(i) = f.vy(i) + offY
        wz(i) = f.vz(i) + offZ
        cx = cx + wx(i)
        cy = cy + wy(i)
        cz = cz + wz(i)
    Next i

    cx = cx / n
    cy = cy / n
    cz = cz / n

    e1x = wx(1) - wx(0)
    e1y = wy(1) - wy(0)
    e1z = wz(1) - wz(0)
    e2x = wx(2) - wx(0)
    e2y = wy(2) - wy(0)
    e2z = wz(2) - wz(0)

    nx = e1y * e2z - e1z * e2y
    ny = e1z * e2x - e1x * e2z
    nz = e1x * e2y - e1y * e2x

    '/* The normal is left unnormalised here on purpose. Dividing it by its own length costs a square
    ' * root and three divides on every face, and the only thing the backface test needs is the sign of
    ' * the dot product, which scaling by a positive number cannot change. Comparing the squares against
    ' * the squared tolerance keeps the threshold meaning what it meant, still without the root. The
    ' * unit normal is built only for a face whose shade has to be computed from its real orientation,
    ' * which is the minority of them: everything axis aligned reads its colour from the table instead. */
    nSq = nx * nx + ny * ny + nz * nz
    If nSq < 0.0000000001 Then Exit Sub

    facing = nx * (m_eyeX - cx) + ny * (m_eyeY - cy) + nz * (m_eyeZ - cz)

    dirKey = f.dirKey

    If m_matTwoSided Then
        If facing * facing <= m_backEpsSq * nSq Then Exit Sub
        If facing < 0! Then
            nx = -nx
            ny = -ny
            nz = -nz
            dirKey = pdNone
        End If
    Else
        If facing <= 0! Then Exit Sub
        If facing * facing <= m_backEpsSq * nSq Then Exit Sub
    End If

    For i = 0 To n - 1
        viewDX = wx(i) - m_eyeX
        viewDY = wy(i) - m_eyeY
        relZ = wz(i) - m_eyeZ
        flat = viewDX * m_cosYaw + viewDY * m_sinYaw
        pF(i) = flat * m_cosPitch + relZ * m_sinPitch
        pS(i) = -viewDX * m_sinYaw + viewDY * m_cosYaw
        pU(i) = relZ * m_cosPitch - flat * m_sinPitch
    Next i

    pCnt = n

    '/* The five frustum planes, clipped inline. Every plane is the same linear expression with
    ' * different coefficients, so the distance of a vertex is three multiplies rather than a call
    ' * into a helper: at five planes times five vertices per face, that call was costing more than
    ' * the whole projection. The pass over the distances also catches the two common cases early,
    ' * a face wholly inside a plane (nothing to do) and one wholly outside it (drop the face). */
    For plane = 0 To 4
        Select Case plane
            Case 0
                cF = 1!: cS = 0!: cU = 0!: cOff = -m_near
            Case 1
                cF = m_tanH: cS = -1!: cU = 0!: cOff = 0!
            Case 2
                cF = m_tanH: cS = 1!: cU = 0!: cOff = 0!
            Case 3
                cF = m_tanV: cS = 0!: cU = -1!: cOff = 0!
            Case Else
                cF = m_tanV: cS = 0!: cU = 1!: cOff = 0!
        End Select

        anyOut = False
        allOut = True

        For i = 0 To pCnt - 1
            dist(i) = pF(i) * cF + pS(i) * cS + pU(i) * cU + cOff
            If dist(i) < 0! Then anyOut = True Else allOut = False
        Next i

        If allOut Then Exit Sub

        If anyOut Then
            tCnt = 0
            prevIdx = pCnt - 1
            d0 = dist(prevIdx)

            For i = 0 To pCnt - 1
                d1 = dist(i)

                If (d0 >= 0!) Xor (d1 >= 0!) Then
                    t = d0 / (d0 - d1)
                    tF(tCnt) = pF(prevIdx) + (pF(i) - pF(prevIdx)) * t
                    tS(tCnt) = pS(prevIdx) + (pS(i) - pS(prevIdx)) * t
                    tU(tCnt) = pU(prevIdx) + (pU(i) - pU(prevIdx)) * t
                    tCnt = tCnt + 1
                End If

                If d1 >= 0! Then
                    tF(tCnt) = pF(i)
                    tS(tCnt) = pS(i)
                    tU(tCnt) = pU(i)
                    tCnt = tCnt + 1
                End If

                prevIdx = i
                d0 = d1
            Next i

            If tCnt < 3 Then Exit Sub

            pCnt = tCnt
            For i = 0 To pCnt - 1
                pF(i) = tF(i)
                pS(i) = tS(i)
                pU(i) = tU(i)
            Next i
        End If
    Next plane

    If dirKey < 0 Then
        If m_matUnlit Then
            dynCol = m_matColor
        Else
            '/* Only now is the unit normal worth building. */
            nLen = Sqr(nSq)
            dynCol = PLighting.ShadeColor(m_matColor, nx / nLen, ny / nLen, nz / nLen)
        End If
    End If

    EmitProjected pF, pS, pU, pCnt, dirKey, dynCol
End Sub

'/**
' * @brief Draws a quad straight from its four corners.
' * @param x0 The first corner X coordinate.
' * @param y0 The first corner Y coordinate.
' * @param z0 The first corner Z coordinate.
' * @param x1 The second corner X coordinate.
' * @param y1 The second corner Y coordinate.
' * @param z1 The second corner Z coordinate.
' * @param x2 The third corner X coordinate.
' * @param y2 The third corner Y coordinate.
' * @param z2 The third corner Z coordinate.
' * @param x3 The fourth corner X coordinate.
' * @param y3 The fourth corner Y coordinate.
' * @param z3 The fourth corner Z coordinate.
' * @param matId The material id to shade the quad with.
' * @param dirKey The axis-aligned normal key, or pdNone when the face is tilted.
' */
Public Sub DrawQuad(ByVal x0 As Single, ByVal y0 As Single, ByVal z0 As Single, _
                    ByVal x1 As Single, ByVal y1 As Single, ByVal z1 As Single, _
                    ByVal x2 As Single, ByVal y2 As Single, ByVal z2 As Single, _
                    ByVal x3 As Single, ByVal y3 As Single, ByVal z3 As Single, _
                    ByVal matId As Long, Optional ByVal dirKey As PDirection = pdNone)
    Dim f As PFace

    FaceQuad f, x0, y0, z0, x1, y1, z1, x2, y2, z2, x3, y3, z3, matId, dirKey
    DrawFace f
End Sub

'/**
' * @brief Draws a triangle straight from its three corners.
' * @param x0 The first corner X coordinate.
' * @param y0 The first corner Y coordinate.
' * @param z0 The first corner Z coordinate.
' * @param x1 The second corner X coordinate.
' * @param y1 The second corner Y coordinate.
' * @param z1 The second corner Z coordinate.
' * @param x2 The third corner X coordinate.
' * @param y2 The third corner Y coordinate.
' * @param z2 The third corner Z coordinate.
' * @param matId The material id to shade the triangle with.
' * @param dirKey The axis-aligned normal key, or pdNone when the face is tilted.
' */
Public Sub DrawTriangle(ByVal x0 As Single, ByVal y0 As Single, ByVal z0 As Single, _
                        ByVal x1 As Single, ByVal y1 As Single, ByVal z1 As Single, _
                        ByVal x2 As Single, ByVal y2 As Single, ByVal z2 As Single, _
                        ByVal matId As Long, Optional ByVal dirKey As PDirection = pdNone)
    Dim f As PFace

    FaceTri f, x0, y0, z0, x1, y1, z1, x2, y2, z2, matId, dirKey
    DrawFace f
End Sub

'/**
' * @brief Draws a polygon whose points are already in slide coordinates, bypassing the 3D pipeline.
' * @param ax The X coordinates of the polygon.
' * @param ay The Y coordinates of the polygon.
' * @param count How many points the arrays hold.
' * @param col The fill colour.
' * @param depth The view depth used only for the distance based area culling.
' * @param matId The material whose transparency and outline settings are applied, or P_INVALID_ID for a plain fill.
' * @remarks This is the entry point for crosshairs, HUD panels and any 2D overlay that should share the shape pool.
' * @remarks Overlay polygons ignore the frame budget on purpose: a crosshair that vanishes exactly when the scene gets busy is worse than one extra shape, and the caller controls how many of these it asks for.
' */
Public Sub DrawPolygon2D(ByRef ax() As Single, ByRef ay() As Single, ByVal count As Long, _
                         ByVal col As Long, Optional ByVal depth As Single = 0!, _
                         Optional ByVal matId As Long = P_INVALID_ID)
    Dim i As Long
    Dim bx(0 To MAX_VERTS + 8) As Single
    Dim by(0 To MAX_VERTS + 8) As Single

    If count < 3 Or count > MAX_VERTS Then Exit Sub

    If matId = P_INVALID_ID Then
        m_matId = P_INVALID_ID
        m_matFormat = False
    ElseIf matId <> m_matId Then
        RefreshMaterial matId
    End If

    For i = 0 To count - 1
        bx(i) = ax(i)
        by(i) = ay(i)
    Next i

    m_overlay = True
    EmitPolygon bx, by, count, col, depth
    m_overlay = False
End Sub

'/** @section Private material cache */

'/**
' * @brief Copies one material into the local cache.
' * @param matId The material id to load.
' * @remarks Called only when the id changes, which in practice is once per object rather than once per face.
' */
Private Sub RefreshMaterial(ByVal matId As Long)
    Dim mat As PMaterial

    Set mat = PMaterials.ById(matId)

    m_matId = matId
    m_matVisible = mat.Visible
    m_matTwoSided = mat.TwoSided
    m_matUnlit = mat.Unlit
    m_matFogged = mat.Fogged
    m_matColor = mat.Color
    m_matFormat = mat.NeedsShapeFormat
    m_matTransp = mat.Transparency
    m_matEdge = mat.EdgeVisible
    m_matEdgeCol = mat.EdgeColor
    m_matEdgeW = mat.EdgeWeight
End Sub

'/** @section Private clipping */

'/**
' * @brief Clips a slide space polygon against one canvas border.
' * @param inX The X coordinates of the incoming polygon.
' * @param inY The Y coordinates of the incoming polygon.
' * @param inCount How many vertices the incoming polygon has.
' * @param edge 0 for left, 1 for right, 2 for top, 3 for bottom.
' * @param outX Receives the X coordinates of the result.
' * @param outY Receives the Y coordinates of the result.
' * @return The vertex count of the clipped polygon.
' * @remarks The inside test is written out rather than called, because it runs once per vertex per border.
' */
Private Function ClipCanvasEdge(ByRef inX() As Single, ByRef inY() As Single, ByVal inCount As Long, ByVal edge As Long, _
                                ByRef outX() As Single, ByRef outY() As Single) As Long
    Dim i As Long
    Dim prevIdx As Long
    Dim outCount As Long
    Dim x0 As Single, y0 As Single
    Dim x1 As Single, y1 As Single
    Dim inside0 As Boolean
    Dim inside1 As Boolean
    Dim t As Single
    Dim limit As Single
    Dim useX As Boolean
    Dim keepAbove As Boolean

    If inCount <= 0 Then Exit Function

    Select Case edge
        Case 0
            limit = m_left: useX = True: keepAbove = True
        Case 1
            limit = m_right: useX = True: keepAbove = False
        Case 2
            limit = m_top: useX = False: keepAbove = True
        Case Else
            limit = m_bottom: useX = False: keepAbove = False
    End Select

    prevIdx = inCount - 1
    x0 = inX(prevIdx)
    y0 = inY(prevIdx)

    If useX Then
        If keepAbove Then inside0 = (x0 >= limit) Else inside0 = (x0 <= limit)
    Else
        If keepAbove Then inside0 = (y0 >= limit) Else inside0 = (y0 <= limit)
    End If

    For i = 0 To inCount - 1
        x1 = inX(i)
        y1 = inY(i)

        If useX Then
            If keepAbove Then inside1 = (x1 >= limit) Else inside1 = (x1 <= limit)
        Else
            If keepAbove Then inside1 = (y1 >= limit) Else inside1 = (y1 <= limit)
        End If

        If inside0 Xor inside1 Then
            If useX Then
                t = (limit - x0) / (x1 - x0)
                outX(outCount) = limit
                outY(outCount) = y0 + (y1 - y0) * t
            Else
                t = (limit - y0) / (y1 - y0)
                outX(outCount) = x0 + (x1 - x0) * t
                outY(outCount) = limit
            End If
            outCount = outCount + 1
        End If

        If inside1 Then
            outX(outCount) = x1
            outY(outCount) = y1
            outCount = outCount + 1
        End If

        x0 = x1
        y0 = y1
        inside0 = inside1
    Next i

    ClipCanvasEdge = outCount
End Function

'/** @section Private emission */

'/**
' * @brief Projects a clipped view space polygon onto the canvas and resolves its colour.
' * @param pF The forward components of the polygon.
' * @param pS The side components of the polygon.
' * @param pU The up components of the polygon.
' * @param pCnt How many vertices the polygon has.
' * @param dirKey The shading key, or pdNone to use the precomputed dynamic colour.
' * @param dynCol The already shaded colour used when dirKey is pdNone.
' * @remarks Fog is applied from the settings cached at the start of the frame. Asking PLighting per face
' * meant a chain of cross module calls for what is a subtraction and a multiply.
' */
Private Sub EmitProjected(ByRef pF() As Single, ByRef pS() As Single, ByRef pU() As Single, ByVal pCnt As Long, _
                          ByVal dirKey As Long, ByVal dynCol As Long)
    Dim i As Long
    Dim invD As Single
    Dim depth As Single
    Dim col As Long
    Dim band As Long
    Dim t As Single
    Dim cr As Long, cg As Long, cb As Long
    Dim ax(0 To MAX_VERTS + 8) As Single
    Dim ay(0 To MAX_VERTS + 8) As Single

    If pCnt < 3 Or pCnt > MAX_VERTS Then Exit Sub

    For i = 0 To pCnt - 1
        invD = m_focal / pF(i)
        ax(i) = m_projCX + pS(i) * invD
        ay(i) = m_projCY - pU(i) * invD + m_shiftY
        depth = depth + pF(i)
    Next i

    depth = depth / pCnt

    If dirKey >= 0 Then
        If m_fogOn Then
            If depth >= m_fogEnd Then
                band = m_fogSteps
            ElseIf depth > m_fogStart Then
                band = CLng((depth - m_fogStart) * m_fogScale)
            End If
        End If

        col = PMaterials.ShadeBand(m_matId, dirKey, band)
    Else
        col = dynCol

        If m_matFogged And m_fogOn And depth > m_fogStart Then
            If depth >= m_fogEnd Then
                col = m_fogR + m_fogG * 256& + m_fogB * 65536
            Else
                t = (depth - m_fogStart) * m_fogInv
                cr = col And &HFF
                cg = (col \ &H100) And &HFF
                cb = (col \ &H10000) And &HFF
                col = CLng(cr + (m_fogR - cr) * t) + _
                      CLng(cg + (m_fogG - cg) * t) * 256& + _
                      CLng(cb + (m_fogB - cb) * t) * 65536
            End If
        End If
    End If

    EmitPolygon ax, ay, pCnt, col, depth
End Sub

'/**
' * @brief Clips, welds, inflates and area tests a slide space polygon, then hands it to the shape pool.
' * @param ax The X coordinates, overwritten in place by the clipping stages.
' * @param ay The Y coordinates, overwritten in place.
' * @param inCount How many vertices the polygon has.
' * @param col The fill colour.
' * @param depth The view depth, used to pick the area threshold.
' */
Private Sub EmitPolygon(ByRef ax() As Single, ByRef ay() As Single, ByVal inCount As Long, ByVal col As Long, ByVal depth As Single)
    Dim i As Long
    Dim count As Long
    Dim kept As Long
    Dim area As Single
    Dim minArea As Single
    Dim dx As Single, dy As Single
    Dim seg As Single
    Dim minX As Single, minY As Single
    Dim maxX As Single, maxY As Single
    Dim ccx As Single, ccy As Single
    Dim bx(0 To MAX_VERTS + 8) As Single
    Dim by(0 To MAX_VERTS + 8) As Single

    count = inCount
    If count < 3 Then Exit Sub

    minX = ax(0)
    maxX = minX
    minY = ay(0)
    maxY = minY

    For i = 1 To count - 1
        If ax(i) < minX Then minX = ax(i)
        If ax(i) > maxX Then maxX = ax(i)
        If ay(i) < minY Then minY = ay(i)
        If ay(i) > maxY Then maxY = ay(i)
    Next i

    If maxX <= m_left Or minX >= m_right Or maxY <= m_top Or minY >= m_bottom Then Exit Sub

    If minX < m_left Then
        count = ClipCanvasEdge(ax, ay, count, 0, bx, by)
        If count < 3 Then Exit Sub
        CopyBack bx, by, ax, ay, count
    End If

    If maxX > m_right Then
        count = ClipCanvasEdge(ax, ay, count, 1, bx, by)
        If count < 3 Then Exit Sub
        CopyBack bx, by, ax, ay, count
    End If

    If minY < m_top Then
        count = ClipCanvasEdge(ax, ay, count, 2, bx, by)
        If count < 3 Then Exit Sub
        CopyBack bx, by, ax, ay, count
    End If

    If maxY > m_bottom Then
        count = ClipCanvasEdge(ax, ay, count, 3, bx, by)
        If count < 3 Then Exit Sub
        CopyBack bx, by, ax, ay, count
    End If

    kept = 0

    For i = 0 To count - 1
        If kept = 0 Then
            bx(0) = ax(i)
            by(0) = ay(i)
            kept = 1
        Else
            dx = ax(i) - bx(kept - 1)
            dy = ay(i) - by(kept - 1)
            If dx * dx + dy * dy > WELD_DIST_SQ Then
                bx(kept) = ax(i)
                by(kept) = ay(i)
                kept = kept + 1
            End If
        End If
    Next i

    If kept >= 2 Then
        dx = bx(kept - 1) - bx(0)
        dy = by(kept - 1) - by(0)
        If dx * dx + dy * dy <= WELD_DIST_SQ Then kept = kept - 1
    End If

    If kept < 3 Then Exit Sub
    If kept > MAX_VERTS Then Exit Sub

    If m_edgeInflate > 0! Then
        ccx = 0!
        ccy = 0!

        For i = 0 To kept - 1
            ccx = ccx + bx(i)
            ccy = ccy + by(i)
        Next i

        ccx = ccx / kept
        ccy = ccy / kept

        For i = 0 To kept - 1
            dx = bx(i) - ccx
            dy = by(i) - ccy
            seg = dx * dx + dy * dy
            If seg > 0.0001 Then
                seg = m_edgeInflate / Sqr(seg)
                bx(i) = bx(i) + dx * seg
                by(i) = by(i) + dy * seg
            End If
        Next i
    End If

    area = 0!
    For i = 0 To kept - 2
        area = area + bx(i) * by(i + 1) - bx(i + 1) * by(i)
    Next i
    area = area + bx(kept - 1) * by(0) - bx(0) * by(kept - 1)

    minArea = m_minArea
    If depth >= m_farDepth Then minArea = m_farArea
    If Abs(area) < minArea Then Exit Sub

    AddShapePoly bx, by, kept, col
End Sub

'/**
' * @brief Copies a clipped polygon back into the working arrays.
' * @param srcX The clipped X coordinates.
' * @param srcY The clipped Y coordinates.
' * @param dstX Receives the X coordinates.
' * @param dstY Receives the Y coordinates.
' * @param count How many vertices to copy.
' */
Private Sub CopyBack(ByRef srcX() As Single, ByRef srcY() As Single, ByRef dstX() As Single, ByRef dstY() As Single, ByVal count As Long)
    Dim i As Long

    For i = 0 To count - 1
        dstX(i) = srcX(i)
        dstY(i) = srcY(i)
    Next i
End Sub

'/**
' * @brief Creates the PowerPoint polyline for one finished polygon.
' * @param bx The X coordinates in slide points.
' * @param by The Y coordinates in slide points.
' * @param count How many vertices the polygon has.
' * @param col The fill colour.
' * @remarks Shapes are named from the pool so the next frame can delete them as one range; only materials that ask for it pay for extra formatting calls.
' */
Private Sub AddShapePoly(ByRef bx() As Single, ByRef by() As Single, ByVal count As Long, ByVal col As Long)
    Dim shp As Shape
    Dim k As Long
    Dim n As Long
    Dim t0 As Double

    If m_count >= m_budget And Not m_overlay Then Exit Sub
    If m_count >= m_poolSize Then EnsureNames m_count + 64

    n = count + 1

    If n <> m_ptsSize Then
        ReDim m_ptsBuf(1 To n, 1 To 2)
        m_ptsSize = n
    End If

    For k = 0 To count - 1
        m_ptsBuf(k + 1, 1) = bx(k)
        m_ptsBuf(k + 1, 2) = by(k)
    Next k

    m_ptsBuf(n, 1) = bx(0)
    m_ptsBuf(n, 2) = by(0)

    '/* Everything above this point is arithmetic. Everything below is COM, and COM is what a
    ' * PowerPoint renderer actually pays for. DryRun stops exactly here, so timing a run with it on
    ' * against a run with it off measures the split without a profiler. */
    If m_dryRun Then
        m_count = m_count + 1
        Exit Sub
    End If

    If m_profile Then t0 = PCore.Seconds()

    Set shp = m_shapes.AddPolyline(m_ptsBuf)

    With shp
        .Name = BankName(m_bank, m_liveShapes)
        .Fill.ForeColor.RGB = col

        If Not m_matFormat Then
            If m_seamFill Then
                '/* Outline painted in the fill colour instead of switched off. A polyline is stroked
                ' * centred on its path, so this fattens the shape by half a line weight on every side,
                ' * closing the hairline gap against its neighbour without moving a single vertex. It
                ' * costs one COM round trip more than hiding the line, which EdgeInflate does for free
                ' * in arithmetic; the trade is fidelity on small polygons against that round trip. */
                .Line.ForeColor.RGB = col
            Else
                .Line.Visible = msoFalse
            End If
        Else
            If m_matTransp > 0! Then .Fill.Transparency = m_matTransp

            If m_matEdge Then
                .Line.Visible = msoTrue
                .Line.ForeColor.RGB = m_matEdgeCol
                .Line.Weight = m_matEdgeW
            ElseIf m_seamFill Then
                .Line.ForeColor.RGB = col
            Else
                .Line.Visible = msoFalse
            End If
        End If
    End With

    If m_profile Then
        m_comAcc = m_comAcc + (PCore.Seconds() - t0)
        m_comOps = m_comOps + 1
    End If

    m_liveShapes = m_liveShapes + 1
    m_count = m_count + 1
End Sub

'/**
' * @brief Reads the pooled shape name for one slot of one bank.
' * @param bank The bank index, 0 or 1.
' * @param slot The slot inside that bank.
' * @return The shape name.
' */
Private Function BankName(ByVal bank As Long, ByVal slot As Long) As String
    If bank = 0 Then
        BankName = m_names(slot)
    Else
        BankName = m_altNames(slot)
    End If
End Function

'/**
' * @brief Grows the shape name table so at least the requested number of slots exists.
' * @param required How many slots are needed.
' * @remarks Names are precomputed strings because building them inside the draw loop would allocate on every polygon.
' */
Private Sub EnsureNames(ByVal required As Long)
    Dim i As Long
    Dim target As Long

    If required <= m_poolSize Then Exit Sub

    target = m_poolSize
    If target < 64 Then target = 64
    Do While target < required
        target = target * 2
    Loop

    ReDim Preserve m_names(0 To target - 1)
    ReDim Preserve m_altNames(0 To target - 1)

    For i = m_poolSize To target - 1
        m_names(i) = m_prefix & "a" & CStr(i)
        m_altNames(i) = m_prefix & "b" & CStr(i)
    Next i

    m_poolSize = target
End Sub

'/**
' * @brief Rebuilds every pooled shape name after the prefix changes.
' */
Private Sub RebuildNames()
    Dim i As Long

    If m_poolSize <= 0 Then Exit Sub

    For i = 0 To m_poolSize - 1
        m_names(i) = m_prefix & "a" & CStr(i)
        m_altNames(i) = m_prefix & "b" & CStr(i)
    Next i
End Sub

'/** @section Primitives */

'/**
' * @description Everything below builds a shape out of faces and hands them straight to DrawFace. There
' * is no mesh object anywhere in Psu3D: a box is not a thing that is kept, it is a call that happens, the
' * way a rectangle is on a 2D canvas. That is why these are commands on the renderer rather than a module
' * of their own, and why nothing has to be allocated to draw a thousand of them.
' */

'/** @subsection Face construction */

'/**
' * @brief Fills a face with four corners.
' * @param f The face to fill.
' * @param x0 The first corner X coordinate.
' * @param y0 The first corner Y coordinate.
' * @param z0 The first corner Z coordinate.
' * @param x1 The second corner X coordinate.
' * @param y1 The second corner Y coordinate.
' * @param z1 The second corner Z coordinate.
' * @param x2 The third corner X coordinate.
' * @param y2 The third corner Y coordinate.
' * @param z2 The third corner Z coordinate.
' * @param x3 The fourth corner X coordinate.
' * @param y3 The fourth corner Y coordinate.
' * @param z3 The fourth corner Z coordinate.
' * @param matId The material id the face carries.
' * @param dirKey The axis-aligned normal key, or pdNone for a tilted face.
' */
Private Sub FaceQuad(ByRef f As PFace, _
                    ByVal x0 As Single, ByVal y0 As Single, ByVal z0 As Single, _
                    ByVal x1 As Single, ByVal y1 As Single, ByVal z1 As Single, _
                    ByVal x2 As Single, ByVal y2 As Single, ByVal z2 As Single, _
                    ByVal x3 As Single, ByVal y3 As Single, ByVal z3 As Single, _
                    ByVal matId As Long, Optional ByVal dirKey As PDirection = pdNone)
    f.nVerts = 4
    f.vx(0) = x0: f.vy(0) = y0: f.vz(0) = z0
    f.vx(1) = x1: f.vy(1) = y1: f.vz(1) = z1
    f.vx(2) = x2: f.vy(2) = y2: f.vz(2) = z2
    f.vx(3) = x3: f.vy(3) = y3: f.vz(3) = z3
    f.matId = matId
    f.dirKey = dirKey
End Sub

'/**
' * @brief Fills a face with three corners.
' * @param f The face to fill.
' * @param x0 The first corner X coordinate.
' * @param y0 The first corner Y coordinate.
' * @param z0 The first corner Z coordinate.
' * @param x1 The second corner X coordinate.
' * @param y1 The second corner Y coordinate.
' * @param z1 The second corner Z coordinate.
' * @param x2 The third corner X coordinate.
' * @param y2 The third corner Y coordinate.
' * @param z2 The third corner Z coordinate.
' * @param matId The material id the face carries.
' * @param dirKey The axis-aligned normal key, or pdNone for a tilted face.
' */
Private Sub FaceTri(ByRef f As PFace, _
                   ByVal x0 As Single, ByVal y0 As Single, ByVal z0 As Single, _
                   ByVal x1 As Single, ByVal y1 As Single, ByVal z1 As Single, _
                   ByVal x2 As Single, ByVal y2 As Single, ByVal z2 As Single, _
                   ByVal matId As Long, Optional ByVal dirKey As PDirection = pdNone)
    f.nVerts = 3
    f.vx(0) = x0: f.vy(0) = y0: f.vz(0) = z0
    f.vx(1) = x1: f.vy(1) = y1: f.vz(1) = z1
    f.vx(2) = x2: f.vy(2) = y2: f.vz(2) = z2
    f.vx(3) = x2: f.vy(3) = y2: f.vz(3) = z2
    f.matId = matId
    f.dirKey = dirKey
End Sub

'/** @subsection Shapes */

'/**
' * @brief Draws an axis-aligned box, submitting only the faces the eye can see.
' * @param matId The material id of every face.
' * @param x1 The lower X bound.
' * @param y1 The lower Y bound.
' * @param x2 The upper X bound.
' * @param y2 The upper Y bound.
' * @param zTop The height of the top face.
' * @param thickness How far the box extends below its top face.
' */
Public Sub DrawBox(ByVal matId As Long, _
                   ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, _
                   ByVal zTop As Single, ByVal thickness As Single)
    Dim f As PFace
    Dim zBot As Single
    Dim ex As Single, ey As Single, ez As Single
    Dim t As Single

    If x2 < x1 Then
        t = x1: x1 = x2: x2 = t
    End If

    If y2 < y1 Then
        t = y1: y1 = y2: y2 = t
    End If

    zBot = zTop - thickness
    ex = m_eyeX
    ey = m_eyeY
    ez = m_eyeZ

    If ez > zTop Then
        FaceQuad f, x1, y1, zTop, x2, y1, zTop, x2, y2, zTop, x1, y2, zTop, matId, pdPosZ
        DrawFace f
    ElseIf ez < zBot Then
        FaceQuad f, x1, y2, zBot, x2, y2, zBot, x2, y1, zBot, x1, y1, zBot, matId, pdNegZ
        DrawFace f
    End If

    If ey < y1 Then
        FaceQuad f, x1, y1, zBot, x2, y1, zBot, x2, y1, zTop, x1, y1, zTop, matId, pdNegY
        DrawFace f
    ElseIf ey > y2 Then
        FaceQuad f, x2, y2, zBot, x1, y2, zBot, x1, y2, zTop, x2, y2, zTop, matId, pdPosY
        DrawFace f
    End If

    If ex > x2 Then
        FaceQuad f, x2, y1, zBot, x2, y2, zBot, x2, y2, zTop, x2, y1, zTop, matId, pdPosX
        DrawFace f
    ElseIf ex < x1 Then
        FaceQuad f, x1, y2, zBot, x1, y1, zBot, x1, y1, zTop, x1, y2, zTop, matId, pdNegX
        DrawFace f
    End If
End Sub

'/**
' * @brief Draws an axis-aligned box as the single face that covers the most of it from here.
' * @param matId The material id of the face.
' * @param x1 The lower X bound.
' * @param y1 The lower Y bound.
' * @param x2 The upper X bound.
' * @param y2 The upper Y bound.
' * @param zTop The height of the top face.
' * @param thickness How far the box extends below its top face.
' * @description The reduced form used once a box is small on screen. A box shows at most three faces,
' * and at distance the two minor ones are a few points wide and cost exactly as much to draw as the
' * dominant one. Which face dominates is decided by the eye offset measured in units of the box's own
' * half extents, so a wide flat slab keeps its top and a tall thin post keeps its side.
' */
Public Sub DrawBoxLod(ByVal matId As Long, _
                      ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, _
                      ByVal zTop As Single, ByVal thickness As Single)
    Dim f As PFace
    Dim zBot As Single
    Dim hx As Single, hy As Single, hz As Single
    Dim dx As Single, dy As Single, dz As Single
    Dim ax As Single, ay As Single, az As Single
    Dim t As Single

    If x2 < x1 Then
        t = x1: x1 = x2: x2 = t
    End If

    If y2 < y1 Then
        t = y1: y1 = y2: y2 = t
    End If

    zBot = zTop - thickness

    hx = (x2 - x1) * 0.5!
    hy = (y2 - y1) * 0.5!
    hz = thickness * 0.5!
    If hx < 0.001 Then hx = 0.001
    If hy < 0.001 Then hy = 0.001
    If hz < 0.001 Then hz = 0.001

    dx = m_eyeX - (x1 + x2) * 0.5!
    dy = m_eyeY - (y1 + y2) * 0.5!
    dz = m_eyeZ - (zTop + zBot) * 0.5!

    ax = Abs(dx) / hx
    ay = Abs(dy) / hy
    az = Abs(dz) / hz

    If az >= ax And az >= ay Then
        If dz >= 0! Then
            FaceQuad f, x1, y1, zTop, x2, y1, zTop, x2, y2, zTop, x1, y2, zTop, matId, pdPosZ
        Else
            FaceQuad f, x1, y2, zBot, x2, y2, zBot, x2, y1, zBot, x1, y1, zBot, matId, pdNegZ
        End If
    ElseIf ay >= ax Then
        If dy < 0! Then
            FaceQuad f, x1, y1, zBot, x2, y1, zBot, x2, y1, zTop, x1, y1, zTop, matId, pdNegY
        Else
            FaceQuad f, x2, y2, zBot, x1, y2, zBot, x1, y2, zTop, x2, y2, zTop, matId, pdPosY
        End If
    Else
        If dx >= 0! Then
            FaceQuad f, x2, y1, zBot, x2, y2, zBot, x2, y2, zTop, x2, y1, zTop, matId, pdPosX
        Else
            FaceQuad f, x1, y2, zBot, x1, y1, zBot, x1, y1, zTop, x1, y2, zTop, matId, pdNegX
        End If
    End If

    DrawFace f
End Sub

'/**
' * @brief Draws a box spun around its own vertical axis.
' * @param matId The material id of every face.
' * @param cx The centre X coordinate.
' * @param cy The centre Y coordinate.
' * @param halfW Half the extent along the local X axis.
' * @param halfH Half the extent along the local Y axis.
' * @param zTop The height of the top face.
' * @param thickness How far the box extends below its top face.
' * @param angleRad The rotation around the vertical axis, in radians.
' * @remarks The four side walls carry no direction key, so they are lit from their real normals as the box turns.
' */
Public Sub DrawBoxRotated(ByVal matId As Long, _
                          ByVal cx As Single, ByVal cy As Single, _
                          ByVal halfW As Single, ByVal halfH As Single, _
                          ByVal zTop As Single, ByVal thickness As Single, ByVal angleRad As Single)
    Dim f As PFace
    Dim ez As Single
    Dim ca As Single, sa As Single
    Dim x0 As Single, y0 As Single
    Dim x1 As Single, y1 As Single
    Dim x2 As Single, y2 As Single
    Dim x3 As Single, y3 As Single
    Dim zBot As Single

    ez = m_eyeZ
    zBot = zTop - thickness
    ca = Cos(angleRad)
    sa = Sin(angleRad)

    x0 = cx - halfW * ca + halfH * sa
    y0 = cy - halfW * sa - halfH * ca
    x1 = cx + halfW * ca + halfH * sa
    y1 = cy + halfW * sa - halfH * ca
    x2 = cx + halfW * ca - halfH * sa
    y2 = cy + halfW * sa + halfH * ca
    x3 = cx - halfW * ca - halfH * sa
    y3 = cy - halfW * sa + halfH * ca

    If ez > zTop Then
        FaceQuad f, x0, y0, zTop, x1, y1, zTop, x2, y2, zTop, x3, y3, zTop, matId, pdPosZ
        DrawFace f
    ElseIf ez < zBot Then
        FaceQuad f, x3, y3, zBot, x2, y2, zBot, x1, y1, zBot, x0, y0, zBot, matId, pdNegZ
        DrawFace f
    End If

    FaceQuad f, x0, y0, zBot, x1, y1, zBot, x1, y1, zTop, x0, y0, zTop, matId, pdNone
    DrawFace f

    FaceQuad f, x1, y1, zBot, x2, y2, zBot, x2, y2, zTop, x1, y1, zTop, matId, pdNone
    DrawFace f

    FaceQuad f, x2, y2, zBot, x3, y3, zBot, x3, y3, zTop, x2, y2, zTop, matId, pdNone
    DrawFace f

    FaceQuad f, x3, y3, zBot, x0, y0, zBot, x0, y0, zTop, x3, y3, zTop, matId, pdNone
    DrawFace f
End Sub

'/**
' * @brief Draws a sloped slab climbing along one axis.
' * @param matId The material id of every face.
' * @param x1 The lower X bound.
' * @param y1 The lower Y bound.
' * @param x2 The upper X bound.
' * @param y2 The upper Y bound.
' * @param zLow The height at the low end of the slope.
' * @param zHigh The height at the high end of the slope.
' * @param thickness How far the slab extends below its sloped face.
' * @param axis paX to climb along X, paY to climb along Y.
' */
Public Sub DrawRamp(ByVal matId As Long, _
                    ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, _
                    ByVal zLow As Single, ByVal zHigh As Single, ByVal thickness As Single, _
                    Optional ByVal axis As PAxis = paX)
    Dim f As PFace
    Dim ex As Single, ey As Single
    Dim vx0 As Single, vy0 As Single, vz0 As Single
    Dim vx1 As Single, vy1 As Single, vz1 As Single
    Dim vx2 As Single, vy2 As Single, vz2 As Single
    Dim vx3 As Single, vy3 As Single, vz3 As Single
    Dim b0 As Single, b1 As Single, b2 As Single, b3 As Single

    ex = m_eyeX
    ey = m_eyeY

    If axis = paX Then
        vx0 = x1: vy0 = y1: vz0 = zLow
        vx1 = x2: vy1 = y1: vz1 = zHigh
        vx2 = x2: vy2 = y2: vz2 = zHigh
        vx3 = x1: vy3 = y2: vz3 = zLow
    Else
        vx0 = x1: vy0 = y1: vz0 = zLow
        vx1 = x2: vy1 = y1: vz1 = zLow
        vx2 = x2: vy2 = y2: vz2 = zHigh
        vx3 = x1: vy3 = y2: vz3 = zHigh
    End If

    b0 = vz0 - thickness
    b1 = vz1 - thickness
    b2 = vz2 - thickness
    b3 = vz3 - thickness

    FaceQuad f, vx0, vy0, vz0, vx1, vy1, vz1, vx2, vy2, vz2, vx3, vy3, vz3, matId, pdNone
    DrawFace f

    FaceQuad f, vx3, vy3, b3, vx2, vy2, b2, vx1, vy1, b1, vx0, vy0, b0, matId, pdNone
    DrawFace f

    If ey < y1 Then
        FaceQuad f, vx1, vy1, vz1, vx0, vy0, vz0, vx0, vy0, b0, vx1, vy1, b1, matId, pdNegY
        DrawFace f
    ElseIf ey > y2 Then
        FaceQuad f, vx3, vy3, vz3, vx2, vy2, vz2, vx2, vy2, b2, vx3, vy3, b3, matId, pdPosY
        DrawFace f
    End If

    If ex > x2 Then
        FaceQuad f, vx2, vy2, vz2, vx1, vy1, vz1, vx1, vy1, b1, vx2, vy2, b2, matId, pdPosX
        DrawFace f
    ElseIf ex < x1 Then
        FaceQuad f, vx0, vy0, vz0, vx3, vy3, vz3, vx3, vy3, b3, vx0, vy0, b0, matId, pdNegX
        DrawFace f
    End If
End Sub

'/**
' * @brief Draws a single upright wall between two ground points.
' * @param matId The material id of the wall.
' * @param x1 The X coordinate of the first end.
' * @param y1 The Y coordinate of the first end.
' * @param x2 The X coordinate of the second end.
' * @param y2 The Y coordinate of the second end.
' * @param zBottom The height of the base.
' * @param zTop The height of the crown.
' * @remarks Pair this with a two sided material when the wall must be visible from both faces.
' */
Public Sub DrawWall(ByVal matId As Long, _
                    ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, _
                    ByVal zBottom As Single, ByVal zTop As Single)
    Dim f As PFace

    FaceQuad f, x1, y1, zBottom, x2, y2, zBottom, x2, y2, zTop, x1, y1, zTop, matId, pdNone
    DrawFace f
End Sub

'/**
' * @brief Draws a flat ground quad.
' * @param matId The material id of the quad.
' * @param x1 The lower X bound.
' * @param y1 The lower Y bound.
' * @param x2 The upper X bound.
' * @param y2 The upper Y bound.
' * @param Z The height of the quad.
' */
Public Sub DrawFloor(ByVal matId As Long, _
                     ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, _
                     ByVal Z As Single)
    Dim f As PFace

    FaceQuad f, x1, y1, Z, x2, y1, Z, x2, y2, Z, x1, y2, Z, matId, pdPosZ
    DrawFace f
End Sub

'/**
' * @brief Draws an upright quad that always turns its face to the camera.
' * @param matId The material id of the billboard.
' * @param X The centre X coordinate.
' * @param Y The centre Y coordinate.
' * @param Z The centre Z coordinate.
' * @param Width The width of the quad in world units.
' * @param Height The height of the quad in world units.
' * @remarks This is how coins, pickups and simple sprites are drawn without building real geometry.
' */
Public Sub DrawBillboard(ByVal matId As Long, _
                         ByVal X As Single, ByVal Y As Single, ByVal Z As Single, _
                         ByVal Width As Single, ByVal Height As Single)
    Dim f As PFace
    Dim cam As PCamera
    Dim rx As Single, ry As Single
    Dim hw As Single, hh As Single

    Set cam = m_camera
    If cam Is Nothing Then Exit Sub

    hw = Width * 0.5!
    hh = Height * 0.5!
    rx = cam.RightX * hw
    ry = cam.RightY * hw

    FaceQuad f, X - rx, Y - ry, Z - hh, _
                X + rx, Y + ry, Z - hh, _
                X + rx, Y + ry, Z + hh, _
                X - rx, Y - ry, Z + hh, matId, pdNone
    DrawFace f
End Sub

'/**
' * @brief Draws a spinning double sided plate, the classic collectible look.
' * @param matId The material id of the plate.
' * @param X The centre X coordinate.
' * @param Y The centre Y coordinate.
' * @param Z The centre Z coordinate.
' * @param radius The half size of the plate.
' * @param phase The spin angle in radians.
' */
Public Sub DrawSpinner(ByVal matId As Long, _
                       ByVal X As Single, ByVal Y As Single, ByVal Z As Single, _
                       ByVal radius As Single, ByVal phase As Single)
    Dim f As PFace
    Dim ca As Single
    Dim sa As Single

    ca = Cos(phase) * radius
    sa = Sin(phase) * radius

    FaceQuad f, X - ca, Y - sa, Z - radius, _
                X + ca, Y + sa, Z - radius, _
                X + ca, Y + sa, Z + radius, _
                X - ca, Y - sa, Z + radius, matId, pdNone
    DrawFace f

    FaceQuad f, X + sa, Y - ca, Z - radius, _
                X - sa, Y + ca, Z - radius, _
                X - sa, Y + ca, Z + radius, _
                X + sa, Y - ca, Z + radius, matId, pdNone
    DrawFace f
End Sub

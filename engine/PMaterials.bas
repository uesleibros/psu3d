Attribute VB_Name = "PMaterials"
'/**
' * PMaterials - Psu3D Material Registry
' * @description Owns every PMaterial in the project, hands out the small integer ids faces carry, and keeps a baked shading table so the renderer resolves a face colour with one array read instead of per-face lighting maths.
' * @author UesleiDev
' * @version 1.0
' * @remarks The table is rebuilt lazily whenever a material changes or PLighting reports a new revision, so callers never have to think about invalidation.
' */

Option Explicit
Option Private Module

'/** @section Constants */

'/** @description How many material slots are added whenever the registry runs out of room. */
Private Const REGISTRY_GROW As Long = 16

'/** @description Name of the material that always occupies id 0. */
Public Const P_DEFAULT_MATERIAL As String = "default"

'/** @section Module state */

Private m_mats() As PMaterial
Private m_keys() As String
Private m_count As Long
Private m_capacity As Long

Private m_lut() As Long
Private m_lutCount As Long
Private m_lutSteps As Long
Private m_lutRevision As Long
Private m_dirty As Boolean
Private m_ready As Boolean

'/** @section Registry */

'/**
' * @brief Creates a material, registers it and returns it for further tuning.
' * @param matName The lookup name, which must be unique.
' * @param col The base colour; white when omitted.
' * @param mode The collision mode; pcSolid when omitted.
' * @return The registered material, or the existing one when the name is already taken.
' */
Public Function Create(ByVal matName As String, _
                       Optional ByVal col As Long = P_WHITE, _
                       Optional ByVal mode As PCollision = pcSolid) As PMaterial
    Dim m As PMaterial

    EnsureReady

    Set m = ByName(matName)
    If Not m Is Nothing Then
        Set Create = m
        Exit Function
    End If

    Set m = New PMaterial
    m.Name = matName
    m.Color = col
    m.Collision = mode

    Add m
    Set Create = m
End Function

'/**
' * @brief Registers a material that was built by hand.
' * @param mat The material to store; its Name is used as the key.
' * @return The id assigned to the material, or the id it already holds.
' */
Public Function Add(ByVal mat As PMaterial) As Long
    EnsureReady

    If mat Is Nothing Then
        Add = P_INVALID_ID
        Exit Function
    End If

    If mat.Id >= 0 And mat.Id < m_count Then
        If m_mats(mat.Id) Is mat Then
            Add = mat.Id
            Exit Function
        End If
    End If

    If Len(mat.Name) = 0 Then mat.Name = "material_" & m_count
    If IdOf(mat.Name) >= 0 Then mat.Name = mat.Name & "_" & m_count

    If m_count >= m_capacity Then Grow

    mat.Id = m_count
    Set m_mats(m_count) = mat
    m_keys(m_count) = LCase$(mat.Name)

    m_count = m_count + 1
    m_dirty = True

    Add = mat.Id
End Function

'/**
' * @brief Looks a material up by name.
' * @param matName The registered name; the search is case insensitive.
' * @return The material, or Nothing when the name is unknown.
' */
Public Function ByName(ByVal matName As String) As PMaterial
    Dim id As Long

    EnsureReady
    id = IdOf(matName)
    If id < 0 Then Exit Function

    Set ByName = m_mats(id)
End Function

'/**
' * @brief Looks a material up by id.
' * @param id The registry id carried by a face.
' * @return The material, falling back to the default material when the id is out of range.
' */
Public Function ById(ByVal id As Long) As PMaterial
    EnsureReady
    If id < 0 Or id >= m_count Then id = 0

    Set ById = m_mats(id)
End Function

'/**
' * @brief Resolves a name into the id faces are tagged with.
' * @param matName The registered name.
' * @return The id, or P_INVALID_ID when the name is unknown.
' * @remarks A linear scan over a handful of materials, and only ever used at build time; the render loop works with ids.
' */
Public Function IdOf(ByVal matName As String) As Long
    Dim key As String
    Dim i As Long

    EnsureReady
    IdOf = P_INVALID_ID

    key = LCase$(matName)

    For i = 0 To m_count - 1
        If m_keys(i) = key Then
            IdOf = i
            Exit Function
        End If
    Next i
End Function

'/**
' * @brief Reports whether a name is already registered.
' * @param matName The name to test.
' * @return True when a material answers to that name.
' */
Public Function Exists(ByVal matName As String) As Boolean
    Exists = (IdOf(matName) >= 0)
End Function

'/**
' * @brief Moves a material to a new lookup name, keeping its id and every face that references it.
' * @param oldName The current name.
' * @param newName The name to move it to.
' * @return True when the rename succeeded.
' */
Public Function Rename(ByVal oldName As String, ByVal newName As String) As Boolean
    Dim id As Long

    id = IdOf(oldName)
    If id < 0 Then Exit Function
    If Len(newName) = 0 Then Exit Function
    If Exists(newName) Then Exit Function

    m_keys(id) = LCase$(newName)
    m_mats(id).Name = newName

    Rename = True
End Function

'/**
' * @brief Reads how many materials are registered.
' * @return The material count, always at least one.
' */
Public Property Get Count() As Long
    EnsureReady
    Count = m_count
End Property

'/**
' * @brief Drops every material and restores the lone default entry.
' * @remarks Any id held by existing geometry becomes meaningless, so rebuild the world after calling this.
' */
Public Sub Clear()
    m_ready = False
    m_count = 0
    m_capacity = 0
    Erase m_mats
    Erase m_keys
    EnsureReady
End Sub

'/** @section Stock materials */

'/**
' * @brief Registers a small palette of ready to use materials covering the common surface behaviours.
' * @description Creates ground, ice, mud, metal, glass, lava, water, one-way platform, decorative and pickup materials. Existing names are left untouched, so it is safe to call after your own setup.
' */
Public Sub CreateDefaults()
    Dim m As PMaterial

    EnsureReady

    Create "stone", PCore.ColorPack(210, 210, 210), pcSolid
    Create "grass", PCore.ColorPack(35, 155, 65), pcSolid
    Create "brick", PCore.ColorPack(235, 110, 45), pcSolid
    Create "metal", PCore.ColorPack(55, 145, 220), pcSolid

    Set m = Create("ice", PCore.ColorPack(170, 220, 245), pcSolid)
    m.Friction = 0.12
    m.SpeedMultiplier = 1.15

    Set m = Create("mud", PCore.ColorPack(96, 72, 48), pcSolid)
    m.Friction = 2.4
    m.SpeedMultiplier = 0.55

    Set m = Create("rubber", PCore.ColorPack(180, 55, 85), pcSolid)
    m.Bounce = 0.72

    Set m = Create("glass", PCore.ColorPack(200, 230, 255), pcSolid)
    m.Transparency = 0.55
    m.TwoSided = True
    m.EdgeVisible = True
    m.EdgeColor = PCore.ColorPack(120, 160, 190)

    Set m = Create("lava", PCore.ColorPack(240, 90, 25), pcSolid)
    m.Unlit = True
    m.DamagePerSecond = 100!

    Set m = Create("water", PCore.ColorPack(40, 110, 190), pcTrigger)
    m.Transparency = 0.4
    m.TwoSided = True
    m.SpeedMultiplier = 0.45
    m.Buoyancy = 1.06
    m.Drag = 3.4

    Set m = Create("platform", PCore.ColorPack(230, 190, 90), pcOneWay)

    Set m = Create("ladder", PCore.ColorPack(150, 110, 60), pcGhost)
    m.Climbable = True
    m.TwoSided = True

    Set m = Create("decor", P_WHITE, pcGhost)
    Set m = Create("pickup", PCore.ColorPack(250, 205, 60), pcTrigger)
    m.Unlit = True
    m.TwoSided = True

    Set m = Create("clip", P_WHITE, pcSolid)
    m.Visible = False
End Sub

'/** @section Shading */

'/**
' * @brief Resolves the final colour of an axis-aligned face straight from the baked table.
' * @param matId The material id carried by the face.
' * @param dirKey The PDirection entry naming the face orientation.
' * @param depth The view depth of the face, used to pick the fog band.
' * @return The shaded and fogged colour.
' */
Public Function ShadeColor(ByVal matId As Long, ByVal dirKey As PDirection, ByVal depth As Single) As Long
    EnsureTable

    If matId < 0 Or matId >= m_lutCount Then matId = 0
    If dirKey < pdPosX Or dirKey > pdNegZ Then dirKey = pdPosZ

    ShadeColor = m_lut(matId, dirKey, PCore.ClampLong(PLighting.FogStepOf(depth), 0, m_lutSteps))
End Function

'/**
' * @brief Reads a baked colour straight from its fog band.
' * @param matId The material id carried by the face.
' * @param dirKey The PDirection entry naming the face orientation.
' * @param band The fog band the caller already worked out.
' * @return The shaded and fogged colour.
' * @remarks The hot path of the renderer. ShadeColor has to ask PLighting which band a depth lands in, which is several cross module calls per face; a renderer that already caches the fog settings computes the band itself and lands here, where the work is two bounds checks and one array read.
' */
Public Function ShadeBand(ByVal matId As Long, ByVal dirKey As Long, ByVal band As Long) As Long
    If m_dirty Or m_lutCount = 0 Then EnsureTable

    If matId < 0 Or matId >= m_lutCount Then matId = 0
    If dirKey < 0 Or dirKey > 5 Then dirKey = pdPosZ
    If band < 0 Then band = 0
    If band > m_lutSteps Then band = m_lutSteps

    ShadeBand = m_lut(matId, dirKey, band)
End Function

'/**
' * @brief Rebuilds the shading table now, so a frame can be drawn without any table checks.
' */
Public Sub Sync()
    EnsureTable
End Sub

'/**
' * @brief Resolves the final colour of a face whose normal is not axis-aligned.
' * @param matId The material id carried by the face.
' * @param nx The X component of the unit normal.
' * @param ny The Y component of the unit normal.
' * @param nz The Z component of the unit normal.
' * @param depth The view depth of the face.
' * @return The shaded and fogged colour.
' * @remarks This path costs a dot product and a colour blend, which is why axis-aligned faces should always carry a real dirKey.
' */
Public Function ShadeColorDynamic(ByVal matId As Long, _
                                  ByVal nx As Single, ByVal ny As Single, ByVal nz As Single, _
                                  ByVal depth As Single) As Long
    Dim m As PMaterial
    Dim col As Long

    Set m = ById(matId)
    col = m.Color

    If Not m.Unlit Then col = PLighting.ShadeColor(col, nx, ny, nz)
    If m.Fogged Then col = PLighting.ApplyFog(col, depth)

    ShadeColorDynamic = col
End Function

'/**
' * @brief Reads the unlit base colour of a material by id.
' * @param matId The material id.
' * @return The packed colour.
' */
Public Function ColorOf(ByVal matId As Long) As Long
    ColorOf = ById(matId).Color
End Function

'/**
' * @brief Marks the baked shading table as stale.
' * @remarks PMaterial calls this for you whenever a shading property changes; call it by hand only after editing lighting state through a route the engine cannot see.
' */
Public Sub NotifyChanged()
    m_dirty = True
End Sub

'/**
' * @brief Forces the shading table to be rebuilt on the next draw.
' */
Public Sub Invalidate()
    m_dirty = True
End Sub

'/** @section Private helpers */

'/**
' * @brief Bakes one colour per material, face direction and fog band.
' * @description Rebuilds only when a material changed, a material was added, or PLighting moved to a new revision.
' */
Private Sub EnsureTable()
    Dim i As Long
    Dim d As Long
    Dim s As Long
    Dim steps As Long
    Dim m As PMaterial
    Dim shaded As Long

    '/* The dirty flag alone decides, because PLighting pokes it whenever the light or fog moves.
    ' * Asking PLighting for its revision here instead would put three cross module calls on the
    ' * shading path of every single face. */
    If Not m_dirty And m_lutCount > 0 And m_lutCount = m_count Then Exit Sub

    EnsureReady
    steps = PLighting.FogSteps

    ReDim m_lut(0 To m_count - 1, 0 To 5, 0 To steps)

    For i = 0 To m_count - 1
        Set m = m_mats(i)

        For d = 0 To 5
            If m.Unlit Then
                shaded = m.Color
            Else
                shaded = PLighting.ShadeColorByDir(m.Color, d)
            End If

            If m.Fogged Then
                For s = 0 To steps
                    m_lut(i, d, s) = PLighting.ApplyFogStep(shaded, s)
                Next s
            Else
                For s = 0 To steps
                    m_lut(i, d, s) = shaded
                Next s
            End If
        Next d
    Next i

    m_lutCount = m_count
    m_lutSteps = steps
    m_lutRevision = PLighting.Revision
    m_dirty = False
End Sub

'/**
' * @brief Widens the material array when the registry fills up.
' */
Private Sub Grow()
    m_capacity = m_capacity + REGISTRY_GROW
    ReDim Preserve m_mats(0 To m_capacity - 1)
    ReDim Preserve m_keys(0 To m_capacity - 1)
End Sub

'/**
' * @brief Builds the registry and guarantees that id 0 always resolves to a usable material.
' */
Private Sub EnsureReady()
    Dim m As PMaterial

    If m_ready Then Exit Sub
    m_ready = True

    m_capacity = REGISTRY_GROW
    ReDim m_mats(0 To m_capacity - 1)
    ReDim m_keys(0 To m_capacity - 1)
    m_count = 0

    Set m = New PMaterial
    m.Name = P_DEFAULT_MATERIAL
    m.Color = PCore.ColorPack(200, 200, 200)
    Add m
End Sub

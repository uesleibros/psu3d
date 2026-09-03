VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/**
' * PMaterial - Psu3D Surface Definition
' * @description A single surface description shared by the renderer and the physics layer: what the surface looks like, and whether a body can walk on it, fall through it or pass straight through. One material is meant to be authored once and referenced by many faces through its registry id.
' * @author UesleiDev
' * @version 1.0
' * @scope Private to this VBA project: class modules carry VB_Exposed = False, which is the class level equivalent of Option Private Module.
' * @remarks Every property setter that affects shading notifies PMaterials so cached shading tables are rebuilt on the next draw.
' */

Option Explicit

'/** @section Instance state */

Private m_id As Long
Private m_name As String
Private m_tag As String

Private m_color As Long
Private m_unlit As Boolean
Private m_fogged As Boolean
Private m_visible As Boolean
Private m_twoSided As Boolean
Private m_transparency As Single

Private m_edgeVisible As Boolean
Private m_edgeColor As Long
Private m_edgeWeight As Single

Private m_collision As PCollision
Private m_friction As Single
Private m_bounce As Single
Private m_speedMul As Single
Private m_damage As Single
Private m_climbable As Boolean
Private m_stepUp As Single
Private m_buoyancy As Single
Private m_drag As Single

'/** @section Lifetime */

'/**
' * @brief Applies neutral defaults so a freshly created material is already drawable and solid.
' */
Private Sub Class_Initialize()
    m_id = P_INVALID_ID
    m_name = vbNullString

    m_color = P_WHITE
    m_unlit = False
    m_fogged = True
    m_visible = True
    m_twoSided = False
    m_transparency = 0!

    m_edgeVisible = False
    m_edgeColor = P_BLACK
    m_edgeWeight = 0.75

    m_collision = pcSolid
    m_friction = 1!
    m_bounce = 0!
    m_speedMul = 1!
    m_damage = 0!
    m_climbable = False
    m_stepUp = 0!
    m_buoyancy = 0!
    m_drag = 0!
End Sub

'/** @section Identity */

'/**
' * @brief Reads the registry id assigned to this material.
' * @return The id, or P_INVALID_ID while the material is unregistered.
' */
Public Property Get Id() As Long
    Id = m_id
End Property

'/**
' * @brief Stamps the registry id onto the material.
' * @param value The id handed out by PMaterials.
' * @remarks Reserved for the registry; application code never calls this.
' */
Friend Property Let Id(ByVal value As Long)
    m_id = value
End Property

'/**
' * @brief Reads the lookup name of the material.
' * @return The name used as the registry key.
' */
Public Property Get Name() As String
    Name = m_name
End Property

'/**
' * @brief Sets the lookup name of the material.
' * @param value The new name.
' * @remarks Renaming a material already held by the registry does not move its key; use PMaterials.Rename instead.
' */
Public Property Let Name(ByVal value As String)
    m_name = value
End Property

'/**
' * @brief Reads the free form payload attached to this material.
' * @return The caller defined tag.
' */
Public Property Get Tag() As String
    Tag = m_tag
End Property

'/**
' * @brief Attaches a free form payload to this material, such as a sound name or a gameplay flag.
' * @param value The caller defined tag.
' */
Public Property Let Tag(ByVal value As String)
    m_tag = value
End Property

'/** @section Appearance */

'/**
' * @brief Reads the unlit base colour of the surface.
' * @return The packed colour.
' */
Public Property Get Color() As Long
    Color = m_color
End Property

'/**
' * @brief Sets the unlit base colour of the surface.
' * @param value The packed colour, as produced by PCore.ColorPack.
' */
Public Property Let Color(ByVal value As Long)
    m_color = value
    PMaterials.NotifyChanged
End Property

'/**
' * @brief Reports whether the surface ignores the directional light.
' * @return True when the material renders at full brightness.
' */
Public Property Get Unlit() As Boolean
    Unlit = m_unlit
End Property

'/**
' * @brief Chooses whether the surface ignores the directional light.
' * @param value True to skip shading, which suits coins, HUD panels and emissive props.
' */
Public Property Let Unlit(ByVal value As Boolean)
    m_unlit = value
    PMaterials.NotifyChanged
End Property

'/**
' * @brief Reports whether distance fog is applied to the surface.
' * @return True when the material fades into the fog colour.
' */
Public Property Get Fogged() As Boolean
    Fogged = m_fogged
End Property

'/**
' * @brief Chooses whether distance fog is applied to the surface.
' * @param value True to fade with depth, False to keep the colour constant at any distance.
' */
Public Property Let Fogged(ByVal value As Boolean)
    m_fogged = value
    PMaterials.NotifyChanged
End Property

'/**
' * @brief Reports whether the surface is drawn at all.
' * @return True when the material produces geometry.
' */
Public Property Get Visible() As Boolean
    Visible = m_visible
End Property

'/**
' * @brief Chooses whether the surface is drawn at all.
' * @param value False to build an invisible collider that still blocks movement.
' */
Public Property Let Visible(ByVal value As Boolean)
    m_visible = value
End Property

'/**
' * @brief Reports whether both sides of a face are drawn.
' * @return True when backface culling is skipped.
' */
Public Property Get TwoSided() As Boolean
    TwoSided = m_twoSided
End Property

'/**
' * @brief Chooses whether both sides of a face are drawn.
' * @param value True for flat props such as flags and billboards, which have no inside.
' */
Public Property Let TwoSided(ByVal value As Boolean)
    m_twoSided = value
End Property

'/**
' * @brief Reads how see-through the surface is.
' * @return The transparency, from 0 for opaque to 1 for invisible.
' */
Public Property Get Transparency() As Single
    Transparency = m_transparency
End Property

'/**
' * @brief Sets how see-through the surface is.
' * @param value The transparency, from 0 for opaque to 1 for invisible.
' * @remarks Any value above zero forces the renderer onto a slower path, since the fill has to be formatted per shape.
' */
Public Property Let Transparency(ByVal value As Single)
    m_transparency = PCore.Clamp01(value)
End Property

'/**
' * @brief Reports whether faces are outlined.
' * @return True when the shape outline is drawn.
' */
Public Property Get EdgeVisible() As Boolean
    EdgeVisible = m_edgeVisible
End Property

'/**
' * @brief Chooses whether faces are outlined, which is useful for wireframe and toon looks.
' * @param value True to draw the outline.
' */
Public Property Let EdgeVisible(ByVal value As Boolean)
    m_edgeVisible = value
End Property

'/**
' * @brief Reads the outline colour.
' * @return The packed colour.
' */
Public Property Get EdgeColor() As Long
    EdgeColor = m_edgeColor
End Property

'/**
' * @brief Sets the outline colour.
' * @param value The packed colour.
' */
Public Property Let EdgeColor(ByVal value As Long)
    m_edgeColor = value
End Property

'/**
' * @brief Reads the outline thickness.
' * @return The weight in points.
' */
Public Property Get EdgeWeight() As Single
    EdgeWeight = m_edgeWeight
End Property

'/**
' * @brief Sets the outline thickness.
' * @param value The weight in points.
' */
Public Property Let EdgeWeight(ByVal value As Single)
    If value < 0.25! Then value = 0.25!
    m_edgeWeight = value
End Property

'/**
' * @brief Reports whether the renderer must format each emitted shape individually.
' * @return True when transparency or outlines are in play.
' * @remarks Materials that answer False are drawn through the fast path, which only assigns a fill colour.
' */
Public Property Get NeedsShapeFormat() As Boolean
    NeedsShapeFormat = (m_transparency > 0! Or m_edgeVisible)
End Property

'/** @section Physics */

'/**
' * @brief Reads how the surface answers collision queries.
' * @return The PCollision mode.
' */
Public Property Get Collision() As PCollision
    Collision = m_collision
End Property

'/**
' * @brief Sets how the surface answers collision queries.
' * @param value pcGhost to let bodies pass, pcSolid to block from every side, pcOneWay to block only a landing from above, or pcTrigger to report the overlap without blocking.
' */
Public Property Let Collision(ByVal value As PCollision)
    m_collision = value
End Property

'/**
' * @brief Reports whether bodies are stopped from every direction.
' * @return True for pcSolid materials.
' */
Public Property Get IsSolid() As Boolean
    IsSolid = (m_collision = pcSolid)
End Property

'/**
' * @brief Reports whether bodies pass straight through with no notification.
' * @return True for pcGhost materials.
' */
Public Property Get IsGhost() As Boolean
    IsGhost = (m_collision = pcGhost)
End Property

'/**
' * @brief Reports whether the surface can only be landed on from above.
' * @return True for pcOneWay materials.
' */
Public Property Get IsOneWay() As Boolean
    IsOneWay = (m_collision = pcOneWay)
End Property

'/**
' * @brief Reports whether the surface only signals overlaps, as pickups and hazard volumes do.
' * @return True for pcTrigger materials.
' */
Public Property Get IsTrigger() As Boolean
    IsTrigger = (m_collision = pcTrigger)
End Property

'/**
' * @brief Answers the core physics question: does this surface stop the body that just touched it?
' * @param landingFromAbove True when the body is descending onto the top face of the surface.
' * @return True when the motion must be resolved instead of ignored.
' */
Public Function BlocksMovement(Optional ByVal landingFromAbove As Boolean = False) As Boolean
    Select Case m_collision
        Case pcSolid
            BlocksMovement = True
        Case pcOneWay
            BlocksMovement = landingFromAbove
        Case Else
            BlocksMovement = False
    End Select
End Function

'/**
' * @brief Reads how strongly the surface slows a body sliding across it.
' * @return The friction multiplier; 1 is the engine default, lower values feel like ice.
' */
Public Property Get Friction() As Single
    Friction = m_friction
End Property

'/**
' * @brief Sets how strongly the surface slows a body sliding across it.
' * @param value The friction multiplier, clamped to zero or above.
' */
Public Property Let Friction(ByVal value As Single)
    If value < 0! Then value = 0!
    m_friction = value
End Property

'/**
' * @brief Reads how much vertical speed a landing body keeps.
' * @return The restitution; 0 absorbs the impact, 1 bounces back at full speed.
' */
Public Property Get Bounce() As Single
    Bounce = m_bounce
End Property

'/**
' * @brief Sets how much vertical speed a landing body keeps, turning the surface into a trampoline.
' * @param value The restitution, clamped to zero or above.
' */
Public Property Let Bounce(ByVal value As Single)
    If value < 0! Then value = 0!
    m_bounce = value
End Property

'/**
' * @brief Reads the walking speed multiplier applied while standing on the surface.
' * @return The multiplier; 1 is normal speed.
' */
Public Property Get SpeedMultiplier() As Single
    SpeedMultiplier = m_speedMul
End Property

'/**
' * @brief Sets the walking speed multiplier applied while standing on the surface, for mud or conveyors.
' * @param value The multiplier, clamped to zero or above.
' */
Public Property Let SpeedMultiplier(ByVal value As Single)
    If value < 0! Then value = 0!
    m_speedMul = value
End Property

'/**
' * @brief Reads the damage dealt per second of contact.
' * @return The damage rate; zero for a harmless surface.
' */
Public Property Get DamagePerSecond() As Single
    DamagePerSecond = m_damage
End Property

'/**
' * @brief Sets the damage dealt per second of contact, which is how lava and spikes are described.
' * @param value The damage rate.
' */
Public Property Let DamagePerSecond(ByVal value As Single)
    m_damage = value
End Property

'/**
' * @brief Reports whether a body may climb the vertical faces of the surface.
' * @return True for ladders and similar geometry.
' */
Public Property Get Climbable() As Boolean
    Climbable = m_climbable
End Property

'/**
' * @brief Chooses whether a body may climb the vertical faces of the surface.
' * @param value True to allow climbing.
' */
Public Property Let Climbable(ByVal value As Boolean)
    m_climbable = value
End Property

'/**
' * @brief Reads the tallest lip a walking body steps over instead of colliding with.
' * @return The step height in world units, or zero to leave the decision to the body.
' */
Public Property Get StepHeight() As Single
    StepHeight = m_stepUp
End Property

'/**
' * @brief Sets the tallest lip a walking body steps over instead of colliding with.
' * @param value The step height in world units; zero, the default, means the body decides.
' * @description This is what separates a kerb from a cliff when both are the same shape. A material that
' * names a height overrides whatever the body would have allowed, so a low wall can be made unclimbable
' * and a tall stair can be made walkable without either of them changing size.
' */
Public Property Let StepHeight(ByVal value As Single)
    If value < 0! Then value = 0!
    m_stepUp = value
End Property

'/**
' * @brief Reads how strongly the surface holds a body up against gravity.
' * @return The fraction of gravity cancelled while submerged; zero for anything that is not a fluid.
' */
Public Property Get Buoyancy() As Single
    Buoyancy = m_buoyancy
End Property

'/**
' * @brief Turns the material into a fluid a body can be inside of.
' * @param value How much of gravity the fluid cancels: 0 leaves the body falling normally, 1 makes it
' * float in place, above 1 pushes it to the surface.
' * @remarks Read together with Collision. A fluid is a trigger, because it must be entered rather than
' * landed on, and it is the buoyancy that tells the physics this particular trigger is something to
' * swim in rather than a coin to pick up.
' */
Public Property Let Buoyancy(ByVal value As Single)
    If value < 0! Then value = 0!
    m_buoyancy = value
End Property

'/**
' * @brief Reads how quickly the material bleeds off the speed of a body inside it.
' * @return The damping in units per second.
' */
Public Property Get Drag() As Single
    Drag = m_drag
End Property

'/**
' * @brief Sets how quickly the material bleeds off the speed of a body inside it.
' * @param value The damping in units per second; this is what makes water feel thick and a jump land softly.
' */
Public Property Let Drag(ByVal value As Single)
    If value < 0! Then value = 0!
    m_drag = value
End Property

'/**
' * @brief Reports whether the material is something a body can be immersed in.
' * @return True for a trigger that carries buoyancy.
' */
Public Property Get IsFluid() As Boolean
    IsFluid = (m_collision = pcTrigger And m_buoyancy > 0!)
End Property

'/** @section Copying */

'/**
' * @brief Overwrites this material with the settings of another, leaving the registry id alone.
' * @param src The material to copy from.
' */
Public Sub CopyFrom(ByVal src As PMaterial)
    If src Is Nothing Then Exit Sub

    m_name = src.Name
    m_tag = src.Tag
    m_color = src.Color
    m_unlit = src.Unlit
    m_fogged = src.Fogged
    m_visible = src.Visible
    m_twoSided = src.TwoSided
    m_transparency = src.Transparency
    m_edgeVisible = src.EdgeVisible
    m_edgeColor = src.EdgeColor
    m_edgeWeight = src.EdgeWeight
    m_collision = src.Collision
    m_friction = src.Friction
    m_bounce = src.Bounce
    m_speedMul = src.SpeedMultiplier
    m_damage = src.DamagePerSecond
    m_climbable = src.Climbable
    m_stepUp = src.StepHeight
    m_buoyancy = src.Buoyancy
    m_drag = src.Drag

    PMaterials.NotifyChanged
End Sub

'/**
' * @brief Produces an unregistered copy of this material.
' * @param newName The name to give the copy; the source name is reused when omitted.
' * @return The cloned material, ready to be handed to PMaterials.Add.
' */
Public Function Clone(Optional ByVal newName As String = vbNullString) As PMaterial
    Dim r As PMaterial

    Set r = New PMaterial
    r.CopyFrom Me
    If Len(newName) > 0 Then r.Name = newName

    Set Clone = r
End Function

'/**
' * @brief Renders the material as readable text for logging and debugging.
' * @return A one line summary of the identity, colour and collision mode.
' */
Public Function Describe() As String
    Dim mode As String

    Select Case m_collision
        Case pcGhost:   mode = "ghost"
        Case pcSolid:   mode = "solid"
        Case pcOneWay:  mode = "one-way"
        Case Else:      mode = "trigger"
    End Select

    Describe = "PMaterial#" & m_id & " '" & m_name & "' " & PCore.ColorToHex(m_color) & _
               " [" & mode & " fric=" & Format$(m_friction, "0.00") & _
               " bounce=" & Format$(m_bounce, "0.00") & "]"
End Function

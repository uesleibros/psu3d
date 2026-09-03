VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PCamera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/**
' * PCamera - Psu3D Yaw And Pitch Observer
' * @description Holds the eye position and orientation of a view, caches the trigonometry the transform needs, and turns world points into the forward, side and up triple the canvas projects. Also answers the visibility question used to reject whole objects before any face work happens.
' * @author UesleiDev
' * @version 1.0
' * @scope Private to this VBA project: class modules carry VB_Exposed = False, which is the class level equivalent of Option Private Module.
' * @remarks Yaw turns around the Z axis with zero looking down positive X; pitch is positive when looking up. The cached sines and cosines are refreshed on assignment, so reading them in a tight loop costs nothing.
' */

Option Explicit

'/** @section Instance state */

Private m_x As Single
Private m_y As Single
Private m_z As Single

Private m_yaw As Single
Private m_pitch As Single

Private m_cosYaw As Single
Private m_sinYaw As Single
Private m_cosPitch As Single
Private m_sinPitch As Single

Private m_pitchMin As Single
Private m_pitchMax As Single
Private m_shiftY As Single

'/** @section Lifetime */

'/**
' * @brief Places the camera at the origin looking along positive X.
' */
Private Sub Class_Initialize()
    m_pitchMin = -1.4
    m_pitchMax = 1.4
    RefreshYaw
    RefreshPitch
End Sub

'/**
' * @brief Places and aims the camera in one call.
' * @param X The eye X coordinate.
' * @param Y The eye Y coordinate.
' * @param Z The eye Z coordinate.
' * @param yawRad The heading in radians.
' * @param pitchRad The elevation in radians.
' * @return The camera itself, so it can be created and aimed in one expression.
' */
Public Function Init(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, _
                     Optional ByVal yawRad As Single = 0!, Optional ByVal pitchRad As Single = 0!) As PCamera
    SetPosition X, Y, Z
    SetAngles yawRad, pitchRad
    Set Init = Me
End Function

'/** @section Position */

'/**
' * @brief Moves the eye to an absolute position.
' * @param X The eye X coordinate.
' * @param Y The eye Y coordinate.
' * @param Z The eye Z coordinate.
' */
Public Sub SetPosition(ByVal X As Single, ByVal Y As Single, ByVal Z As Single)
    m_x = X
    m_y = Y
    m_z = Z
End Sub

'/**
' * @brief Shifts the eye by an offset.
' * @param dx The change along X.
' * @param dy The change along Y.
' * @param dz The change along Z.
' */
Public Sub Translate(ByVal dx As Single, ByVal dy As Single, ByVal dz As Single)
    m_x = m_x + dx
    m_y = m_y + dy
    m_z = m_z + dz
End Sub

'/**
' * @brief Reads the eye X coordinate.
' * @return The world X position.
' */
Public Property Get X() As Single
    X = m_x
End Property

'/**
' * @brief Sets the eye X coordinate.
' * @param value The world X position.
' */
Public Property Let X(ByVal value As Single)
    m_x = value
End Property

'/**
' * @brief Reads the eye Y coordinate.
' * @return The world Y position.
' */
Public Property Get Y() As Single
    Y = m_y
End Property

'/**
' * @brief Sets the eye Y coordinate.
' * @param value The world Y position.
' */
Public Property Let Y(ByVal value As Single)
    m_y = value
End Property

'/**
' * @brief Reads the eye Z coordinate.
' * @return The world Z position.
' */
Public Property Get Z() As Single
    Z = m_z
End Property

'/**
' * @brief Sets the eye Z coordinate.
' * @param value The world Z position.
' */
Public Property Let Z(ByVal value As Single)
    m_z = value
End Property

'/** @section Orientation */

'/**
' * @brief Aims the camera at absolute angles.
' * @param yawRad The heading in radians.
' * @param pitchRad The elevation in radians, clamped to the configured limits.
' */
Public Sub SetAngles(ByVal yawRad As Single, ByVal pitchRad As Single)
    Yaw = yawRad
    Pitch = pitchRad
End Sub

'/**
' * @brief Turns the camera by a relative amount, the shape mouse look input arrives in.
' * @param dYaw The heading change in radians.
' * @param dPitch The elevation change in radians.
' */
Public Sub AddAngles(ByVal dYaw As Single, ByVal dPitch As Single)
    Yaw = m_yaw + dYaw
    Pitch = m_pitch + dPitch
End Sub

'/**
' * @brief Reads the heading.
' * @return The yaw in radians.
' */
Public Property Get Yaw() As Single
    Yaw = m_yaw
End Property

'/**
' * @brief Sets the heading and refreshes the cached trigonometry.
' * @param value The yaw in radians; values outside a single turn are wrapped.
' */
Public Property Let Yaw(ByVal value As Single)
    m_yaw = PCore.WrapAngle(value)
    RefreshYaw
End Property

'/**
' * @brief Reads the elevation.
' * @return The pitch in radians.
' */
Public Property Get Pitch() As Single
    Pitch = m_pitch
End Property

'/**
' * @brief Sets the elevation and refreshes the cached trigonometry.
' * @param value The pitch in radians, clamped to the configured limits.
' */
Public Property Let Pitch(ByVal value As Single)
    m_pitch = PCore.Clamp(value, m_pitchMin, m_pitchMax)
    RefreshPitch
End Property

'/**
' * @brief Restricts how far the camera may look up and down.
' * @param minRad The lowest pitch allowed, in radians.
' * @param maxRad The highest pitch allowed, in radians.
' */
Public Sub SetPitchLimits(ByVal minRad As Single, ByVal maxRad As Single)
    m_pitchMin = minRad
    m_pitchMax = maxRad
    Pitch = m_pitch
End Sub

'/**
' * @brief Aims the camera at a world point.
' * @param tx The target X coordinate.
' * @param ty The target Y coordinate.
' * @param tz The target Z coordinate.
' */
Public Sub LookAt(ByVal tx As Single, ByVal ty As Single, ByVal tz As Single)
    Dim dx As Single
    Dim dy As Single
    Dim dz As Single
    Dim flat As Single

    dx = tx - m_x
    dy = ty - m_y
    dz = tz - m_z
    flat = Sqr(dx * dx + dy * dy)

    If flat < P_EPSILON Then
        Pitch = PCore.SignOf(dz) * m_pitchMax
        Exit Sub
    End If

    Yaw = Atn2(dy, dx)
    Pitch = Atn(dz / flat)
End Sub

'/**
' * @brief Reads the cached cosine of the yaw.
' * @return The cosine, refreshed whenever the yaw changes.
' */
Public Property Get CosYaw() As Single
    CosYaw = m_cosYaw
End Property

'/**
' * @brief Reads the cached sine of the yaw.
' * @return The sine, refreshed whenever the yaw changes.
' */
Public Property Get SinYaw() As Single
    SinYaw = m_sinYaw
End Property

'/**
' * @brief Reads the cached cosine of the pitch.
' * @return The cosine, refreshed whenever the pitch changes.
' */
Public Property Get CosPitch() As Single
    CosPitch = m_cosPitch
End Property

'/**
' * @brief Reads the cached sine of the pitch.
' * @return The sine, refreshed whenever the pitch changes.
' */
Public Property Get SinPitch() As Single
    SinPitch = m_sinPitch
End Property

'/**
' * @brief Reads the X component of the flat forward vector, the direction a walking body advances in.
' * @return The forward X component.
' */
Public Property Get ForwardX() As Single
    ForwardX = m_cosYaw
End Property

'/**
' * @brief Reads the Y component of the flat forward vector.
' * @return The forward Y component.
' */
Public Property Get ForwardY() As Single
    ForwardY = m_sinYaw
End Property

'/**
' * @brief Reads the X component of the flat right vector, used for strafing.
' * @return The right X component.
' */
Public Property Get RightX() As Single
    RightX = -m_sinYaw
End Property

'/**
' * @brief Reads the Y component of the flat right vector.
' * @return The right Y component.
' */
Public Property Get RightY() As Single
    RightY = m_cosYaw
End Property

'/**
' * @brief Reads the vertical screen offset applied at projection time.
' * @return The offset in slide points.
' */
Public Property Get ScreenShiftY() As Single
    ScreenShiftY = m_shiftY
End Property

'/**
' * @brief Sets a vertical screen offset applied at projection time, which is how head bob and recoil are expressed.
' * @param value The offset in slide points.
' */
Public Property Let ScreenShiftY(ByVal value As Single)
    m_shiftY = value
End Property

'/** @section Transforms */

'/**
' * @brief Converts a world point into the view space triple the canvas projects.
' * @param wx The world X coordinate.
' * @param wy The world Y coordinate.
' * @param wz The world Z coordinate.
' * @param outFwd Receives the distance in front of the eye.
' * @param outSide Receives the distance to the right of the eye.
' * @param outUp Receives the distance above the eye.
' */
Public Sub WorldToView(ByVal wx As Single, ByVal wy As Single, ByVal wz As Single, _
                       ByRef outFwd As Single, ByRef outSide As Single, ByRef outUp As Single)
    Dim dx As Single
    Dim dy As Single
    Dim dz As Single
    Dim flat As Single

    dx = wx - m_x
    dy = wy - m_y
    dz = wz - m_z
    flat = dx * m_cosYaw + dy * m_sinYaw

    outFwd = flat * m_cosPitch + dz * m_sinPitch
    outSide = -dx * m_sinYaw + dy * m_cosYaw
    outUp = dz * m_cosPitch - flat * m_sinPitch
End Sub

'/**
' * @brief Measures how far in front of the eye a world point sits.
' * @param wx The world X coordinate.
' * @param wy The world Y coordinate.
' * @param wz The world Z coordinate.
' * @return The view depth, negative when the point is behind the camera.
' */
Public Function ViewDepth(ByVal wx As Single, ByVal wy As Single, ByVal wz As Single) As Single
    Dim dx As Single
    Dim dy As Single

    dx = wx - m_x
    dy = wy - m_y

    ViewDepth = (dx * m_cosYaw + dy * m_sinYaw) * m_cosPitch + (wz - m_z) * m_sinPitch
End Function

'/**
' * @brief Rejects an object against the view frustum using its bounding sphere.
' * @param cv The canvas whose frustum is tested against.
' * @param wx The sphere centre X coordinate.
' * @param wy The sphere centre Y coordinate.
' * @param wz The sphere centre Z coordinate.
' * @param radius The sphere radius in world units.
' * @param outNear Receives the depth of the closest point of the sphere, for sorting.
' * @param outMid Receives the depth of the sphere centre.
' * @return True when any part of the sphere can reach the canvas.
' * @remarks This is the cheap early exit that keeps whole objects out of the face pipeline.
' */
Public Function SphereVisible(ByVal cv As PCanvas, _
                              ByVal wx As Single, ByVal wy As Single, ByVal wz As Single, ByVal radius As Single, _
                              ByRef outNear As Single, ByRef outMid As Single) As Boolean
    Dim dx As Single
    Dim dy As Single
    Dim dz As Single
    Dim flat As Single
    Dim vF As Single
    Dim vS As Single
    Dim vU As Single

    If cv Is Nothing Then Exit Function

    dx = wx - m_x
    dy = wy - m_y
    dz = wz - m_z
    flat = dx * m_cosYaw + dy * m_sinYaw
    vF = flat * m_cosPitch + dz * m_sinPitch

    If vF + radius < cv.NearPlane Then Exit Function
    If vF - radius > cv.FarPlane Then Exit Function

    vS = -dx * m_sinYaw + dy * m_cosYaw
    If (vF * cv.TanFovH - vS) * cv.NormH + radius < 0! Then Exit Function
    If (vF * cv.TanFovH + vS) * cv.NormH + radius < 0! Then Exit Function

    vU = dz * m_cosPitch - flat * m_sinPitch
    If (vF * cv.TanFovV - vU) * cv.NormV + radius < 0! Then Exit Function
    If (vF * cv.TanFovV + vU) * cv.NormV + radius < 0! Then Exit Function

    outNear = vF - radius
    outMid = vF

    SphereVisible = True
End Function

'/**
' * @brief Renders the camera state as readable text for logging and debugging.
' * @return A one line summary of the position and orientation.
' */
Public Function Describe() As String
    Describe = "PCamera (" & Format$(m_x, "0.00") & ", " & Format$(m_y, "0.00") & ", " & Format$(m_z, "0.00") & _
               ") yaw=" & Format$(m_yaw * P_RAD2DEG, "0.0") & " pitch=" & Format$(m_pitch * P_RAD2DEG, "0.0")
End Function

'/** @section Private helpers */

'/**
' * @brief Rebuilds the cached yaw trigonometry.
' */
Private Sub RefreshYaw()
    m_cosYaw = Cos(m_yaw)
    m_sinYaw = Sin(m_yaw)
End Sub

'/**
' * @brief Rebuilds the cached pitch trigonometry.
' */
Private Sub RefreshPitch()
    m_cosPitch = Cos(m_pitch)
    m_sinPitch = Sin(m_pitch)
End Sub

'/**
' * @brief Four quadrant arc tangent, which VBA does not provide.
' * @param dy The vertical component.
' * @param dx The horizontal component.
' * @return The angle in radians inside -PI..PI.
' */
Private Function Atn2(ByVal dy As Single, ByVal dx As Single) As Single
    If dx > 0! Then
        Atn2 = Atn(dy / dx)
    ElseIf dx < 0! Then
        If dy >= 0! Then
            Atn2 = Atn(dy / dx) + P_PI
        Else
            Atn2 = Atn(dy / dx) - P_PI
        End If
    ElseIf dy > 0! Then
        Atn2 = P_HALF_PI
    ElseIf dy < 0! Then
        Atn2 = -P_HALF_PI
    End If
End Function

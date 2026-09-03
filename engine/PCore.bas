Attribute VB_Name = "PCore"
'/**
' * PCore - Psu3D Foundations
' * @description The layer everything else is written on: the shared types and enumerations, scalar and angle maths, a deterministic random source, colour arithmetic and a clock fine enough to time one frame. Nothing here knows about a slide, a shape or a scene, so any of it can be used on its own.
' * @author UesleiDev
' * @version 1.0
' * @remarks Types live in a standard module rather than a class because VBA forbids a class from exposing, in a public signature, a type it declared itself. Four small modules were merged into this one: they had no dependencies of their own, were imported together every time, and separating them bought nothing but four more names in the project tree.
' */

Option Explicit
Option Private Module

'/** @section Native declarations */

#If VBA7 Then
    '/** @description Reads the current value of the performance counter. */
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long

    '/** @description Reads how many counter ticks make one second. */
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
#Else
    '/** @description Reads the current value of the performance counter. */
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long

    '/** @description Reads how many counter ticks make one second. */
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
#End If

'/** @section Enumerations */

'/**
' * @enum PDirection
' * @brief Axis-aligned normal keys used to index precomputed shading tables.
' * @remarks pdNone marks a face whose normal is arbitrary, forcing the renderer onto the slower per-face lighting path.
' */
Public Enum PDirection
    pdNone = -1
    pdPosX = 0
    pdNegX = 1
    pdPosY = 2
    pdNegY = 3
    pdPosZ = 4
    pdNegZ = 5
End Enum

'/**
' * @enum PCollision
' * @brief How a material answers the physics query "can a body pass through me?".
' * @remarks pcGhost is purely decorative, pcSolid blocks from every side, pcOneWay blocks only a body landing on its top face, and pcTrigger reports the overlap without ever stopping the body.
' */
Public Enum PCollision
    pcGhost = 0
    pcSolid = 1
    pcOneWay = 2
    pcTrigger = 3
End Enum

'/**
' * @enum PAxis
' * @brief Which horizontal axis a ramp climbs along.
' */
Public Enum PAxis
    paX = 0
    paY = 1
End Enum

'/**
' * @enum PObjectKind
' * @brief The primitive a scene object is drawn as.
' */
Public Enum PObjectKind
    pkBox = 0
    pkRotatedBox = 1
    pkRamp = 2
    pkBillboard = 3
    pkSpinner = 4
End Enum

'/**
' * @enum PCanvasFit
' * @brief Strategy used by PCanvas.FitToSlide when reshaping a canvas against the slide bounds.
' */
Public Enum PCanvasFit
    pfStretch = 0
    pfContain = 1
    pfCover = 2
End Enum

'/** @section Public types */

'/**
' * @struct PFace
' * @brief A single triangle or quad in world space, tagged with the material and shading key used to colour it.
' * @remarks Vertices must be wound counter-clockwise when seen from the visible side; the renderer derives the normal from the first three of them.
' */
Public Type PFace
    vx(0 To 3) As Single
    vy(0 To 3) As Single
    vz(0 To 3) As Single
    nVerts As Long
    matId As Long
    dirKey As Long
End Type

'/** @section Shared limits */

'/** @description Highest number of fog bands a shading table may hold. */
Public Const P_MAX_FOG_STEPS As Long = 64

'/** @description Identifier returned whenever a material lookup fails. */
Public Const P_INVALID_ID As Long = -1

'/** @section Maths constants */

'/** @description Ratio of a circle's circumference to its diameter. */
Public Const P_PI As Single = 3.14159265

'/** @description A full turn expressed in radians. */
Public Const P_TWO_PI As Single = 6.28318531

'/** @description A quarter turn expressed in radians. */
Public Const P_HALF_PI As Single = 1.57079633

'/** @description Multiplier converting degrees into radians. */
Public Const P_DEG2RAD As Single = 0.0174532925

'/** @description Multiplier converting radians into degrees. */
Public Const P_RAD2DEG As Single = 57.2957795

'/** @description Tolerance below which a Single is treated as zero. */
Public Const P_EPSILON As Single = 0.000001

'/** @section Colour presets */

'/** @description Opaque white. */
Public Const P_WHITE As Long = 16777215

'/** @description Opaque black. */
Public Const P_BLACK As Long = 0

'/** @section Module state */

Private m_rngState As Double
Private m_rngSeed As Long
Private m_rngReady As Boolean
Private m_freq As Currency
Private m_ready As Boolean
Private m_ok As Boolean

'/** @section Clamping and interpolation */

'/**
' * @brief Constrains a value to an inclusive range.
' * @param value The value to constrain.
' * @param minVal The lowest accepted result.
' * @param maxVal The highest accepted result.
' * @return The value pulled inside the range.
' */
Public Function Clamp(ByVal value As Single, ByVal minVal As Single, ByVal maxVal As Single) As Single
    If value < minVal Then
        Clamp = minVal
    ElseIf value > maxVal Then
        Clamp = maxVal
    Else
        Clamp = value
    End If
End Function

'/**
' * @brief Constrains an integral value to an inclusive range.
' * @param value The value to constrain.
' * @param minVal The lowest accepted result.
' * @param maxVal The highest accepted result.
' * @return The value pulled inside the range.
' */
Public Function ClampLong(ByVal value As Long, ByVal minVal As Long, ByVal maxVal As Long) As Long
    If value < minVal Then
        ClampLong = minVal
    ElseIf value > maxVal Then
        ClampLong = maxVal
    Else
        ClampLong = value
    End If
End Function

'/**
' * @brief Constrains a value to the unit range.
' * @param value The value to constrain.
' * @return The value pulled inside 0..1.
' */
Public Function Clamp01(ByVal value As Single) As Single
    If value < 0! Then
        Clamp01 = 0!
    ElseIf value > 1! Then
        Clamp01 = 1!
    Else
        Clamp01 = value
    End If
End Function

'/**
' * @brief Linearly blends between two values.
' * @param a The value returned at t = 0.
' * @param b The value returned at t = 1.
' * @param t The unclamped blend factor.
' * @return The interpolated value.
' */
Public Function Lerp(ByVal a As Single, ByVal b As Single, ByVal t As Single) As Single
    Lerp = a + (b - a) * t
End Function

'/**
' * @brief Recovers the blend factor that maps a value back onto a range.
' * @param a The value standing for 0.
' * @param b The value standing for 1.
' * @param value The value to locate inside the range.
' * @return The blend factor, or zero when the range is degenerate.
' */
Public Function InverseLerp(ByVal a As Single, ByVal b As Single, ByVal value As Single) As Single
    If Abs(b - a) < P_EPSILON Then Exit Function
    InverseLerp = (value - a) / (b - a)
End Function

'/**
' * @brief Steps a value towards a target without ever overshooting it.
' * @param current The current value.
' * @param target The value being approached.
' * @param maxDelta The largest change allowed on this call.
' * @return The advanced value.
' */
Public Function MoveTowards(ByVal current As Single, ByVal target As Single, ByVal maxDelta As Single) As Single
    Dim d As Single

    d = target - current

    If Abs(d) <= maxDelta Then
        MoveTowards = target
    ElseIf d > 0! Then
        MoveTowards = current + maxDelta
    Else
        MoveTowards = current - maxDelta
    End If
End Function

'/**
' * @brief Smooths a value towards a target at a frame-rate independent rate.
' * @param current The current value.
' * @param target The value being approached.
' * @param rate How aggressively the gap is closed, in units per second.
' * @param dt The elapsed frame time in seconds.
' * @return The smoothed value.
' */
Public Function Approach(ByVal current As Single, ByVal target As Single, ByVal rate As Single, ByVal dt As Single) As Single
    Dim k As Single

    k = rate * dt

    If k <= 0! Then
        Approach = current
        Exit Function
    End If

    If k > 1! Then k = 1!

    Approach = current + (target - current) * k
End Function

'/** @section Comparison helpers */

'/**
' * @brief Reports the sign of a value.
' * @param value The value to inspect.
' * @return -1, 0 or 1.
' */
Public Function SignOf(ByVal value As Single) As Single
    If value > 0! Then
        SignOf = 1!
    ElseIf value < 0! Then
        SignOf = -1!
    End If
End Function

'/**
' * @brief Returns the smaller of two values.
' * @param a The first value.
' * @param b The second value.
' * @return The minimum of the pair.
' */
Public Function MinOf(ByVal a As Single, ByVal b As Single) As Single
    If a < b Then MinOf = a Else MinOf = b
End Function

'/**
' * @brief Returns the larger of two values.
' * @param a The first value.
' * @param b The second value.
' * @return The maximum of the pair.
' */
Public Function MaxOf(ByVal a As Single, ByVal b As Single) As Single
    If a > b Then MaxOf = a Else MaxOf = b
End Function

'/**
' * @brief Compares two values within a tolerance.
' * @param a The first value.
' * @param b The second value.
' * @param tol The largest difference still treated as equality.
' * @return True when the values match closely enough.
' */
Public Function NearlyEqual(ByVal a As Single, ByVal b As Single, Optional ByVal tol As Single = P_EPSILON) As Boolean
    NearlyEqual = (Abs(a - b) <= tol)
End Function

'/** @section Angles */

'/**
' * @brief Wraps an angle into the -PI..PI range.
' * @param angleRad The angle in radians.
' * @return The equivalent angle inside a single turn.
' */
Public Function WrapAngle(ByVal angleRad As Single) As Single
    Dim a As Single

    a = angleRad - Int(angleRad / P_TWO_PI) * P_TWO_PI

    If a > P_PI Then
        a = a - P_TWO_PI
    ElseIf a < -P_PI Then
        a = a + P_TWO_PI
    End If

    WrapAngle = a
End Function

'/**
' * @brief Measures the shortest signed rotation between two angles.
' * @param fromRad The starting angle in radians.
' * @param toRad The destination angle in radians.
' * @return The signed delta inside -PI..PI.
' */
Public Function AngleDelta(ByVal fromRad As Single, ByVal toRad As Single) As Single
    AngleDelta = WrapAngle(toRad - fromRad)
End Function

'/**
' * @brief Smooths an angle towards a target across the shortest arc.
' * @param current The current angle in radians.
' * @param target The destination angle in radians.
' * @param rate How aggressively the gap is closed, in units per second.
' * @param dt The elapsed frame time in seconds.
' * @return The smoothed angle.
' */
Public Function ApproachAngle(ByVal current As Single, ByVal target As Single, ByVal rate As Single, ByVal dt As Single) As Single
    ApproachAngle = current + AngleDelta(current, target) * Clamp01(rate * dt)
End Function

'/** @section Planar queries */

'/**
' * @brief Tests whether two one dimensional spans overlap.
' * @param a1 The start of the first span.
' * @param a2 The end of the first span.
' * @param b1 The start of the second span.
' * @param b2 The end of the second span.
' * @return True when the spans share any ground.
' */
Public Function SpansOverlap(ByVal a1 As Single, ByVal a2 As Single, ByVal b1 As Single, ByVal b2 As Single) As Boolean
    SpansOverlap = Not (a2 < b1 Or b2 < a1)
End Function

'/**
' * @brief Tests whether two axis-aligned rectangles overlap.
' * @param ax1 The left edge of the first rectangle.
' * @param ay1 The top edge of the first rectangle.
' * @param ax2 The right edge of the first rectangle.
' * @param ay2 The bottom edge of the first rectangle.
' * @param bx1 The left edge of the second rectangle.
' * @param by1 The top edge of the second rectangle.
' * @param bx2 The right edge of the second rectangle.
' * @param by2 The bottom edge of the second rectangle.
' * @return True when the rectangles intersect.
' */
Public Function BoxesOverlap(ByVal ax1 As Single, ByVal ay1 As Single, ByVal ax2 As Single, ByVal ay2 As Single, _
                             ByVal bx1 As Single, ByVal by1 As Single, ByVal bx2 As Single, ByVal by2 As Single) As Boolean
    If ax2 < bx1 Or bx2 < ax1 Then Exit Function
    If ay2 < by1 Or by2 < ay1 Then Exit Function
    BoxesOverlap = True
End Function

'/**
' * @brief Tests whether a point falls inside an axis-aligned rectangle.
' * @param px The point X coordinate.
' * @param py The point Y coordinate.
' * @param x1 The left edge.
' * @param y1 The top edge.
' * @param x2 The right edge.
' * @param y2 The bottom edge.
' * @return True when the point is inside the rectangle or on its border.
' */
Public Function PointInBox(ByVal px As Single, ByVal py As Single, _
                           ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Boolean
    PointInBox = (px >= x1 And px <= x2 And py >= y1 And py <= y2)
End Function

'/**
' * @brief Squared planar distance between two points.
' * @param ax The first point X coordinate.
' * @param ay The first point Y coordinate.
' * @param bx The second point X coordinate.
' * @param by The second point Y coordinate.
' * @return The squared distance, avoiding a square root.
' */
Public Function DistSq2D(ByVal ax As Single, ByVal ay As Single, ByVal bx As Single, ByVal by As Single) As Single
    Dim dx As Single
    Dim dy As Single

    dx = ax - bx
    dy = ay - by

    DistSq2D = dx * dx + dy * dy
End Function

'/** @section Deterministic random */

'/**
' * @brief Reseeds the engine random stream so a generated world can be reproduced exactly.
' * @param seed The seed value; any Long is accepted.
' */
Public Sub SeedRandom(ByVal seed As Long)
    m_rngSeed = seed
    m_rngState = seed Mod 2147483647
    If m_rngState <= 0 Then m_rngState = m_rngState + 2147483646
    m_rngReady = True
End Sub

'/**
' * @brief Reads the seed currently driving the random stream.
' * @return The last seed handed to SeedRandom.
' */
Public Property Get RandomSeed() As Long
    EnsureRng
    RandomSeed = m_rngSeed
End Property

'/**
' * @brief Draws the next value from the Lehmer generator backing the engine.
' * @return A double inside the 0..1 range.
' */
Public Function RandomNext() As Double
    Dim v As Double

    EnsureRng
    v = 16807# * m_rngState
    m_rngState = v - Int(v / 2147483647#) * 2147483647#
    RandomNext = m_rngState / 2147483647#
End Function

'/**
' * @brief Draws a value inside an arbitrary range.
' * @param lo The lowest value that can be produced.
' * @param hi The highest value that can be produced.
' * @return The random value.
' */
Public Function RandomRange(ByVal lo As Single, ByVal hi As Single) As Single
    RandomRange = lo + (hi - lo) * CSng(RandomNext())
End Function

'/**
' * @brief Draws an integer inside an inclusive range.
' * @param lo The lowest value that can be produced.
' * @param hi The highest value that can be produced.
' * @return The random integer.
' */
Public Function RandomInt(ByVal lo As Long, ByVal hi As Long) As Long
    If hi <= lo Then
        RandomInt = lo
        Exit Function
    End If

    RandomInt = lo + Int(RandomNext() * (hi - lo + 1))
End Function

'/**
' * @brief Draws a coin flip biased by a probability.
' * @param chance The probability of a True result, from 0 to 1.
' * @return True when the draw falls under the given chance.
' */
Public Function RandomChance(ByVal chance As Single) As Boolean
    RandomChance = (RandomNext() < chance)
End Function

'/** @section Random helpers */

'/**
' * @brief Seeds the random stream from the clock the first time it is touched.
' */
Private Sub EnsureRng()
    If m_rngReady Then Exit Sub
    SeedRandom CLng(Timer * 37) Mod 2147483000
End Sub

'/** @section Colour packing */

'/**
' * @brief Builds a PowerPoint colour long from three channels, clamping each one.
' * @param r The red channel, 0 to 255.
' * @param g The green channel, 0 to 255.
' * @param b The blue channel, 0 to 255.
' * @return The packed colour.
' */
Public Function ColorPack(ByVal r As Long, ByVal g As Long, ByVal b As Long) As Long
    If r < 0 Then r = 0
    If r > 255 Then r = 255
    If g < 0 Then g = 0
    If g > 255 Then g = 255
    If b < 0 Then b = 0
    If b > 255 Then b = 255

    ColorPack = r + g * 256& + b * 65536
End Function

'/**
' * @brief Extracts the red channel of a packed colour.
' * @param col The packed colour.
' * @return The channel value, 0 to 255.
' */
Public Function ColorRed(ByVal col As Long) As Long
    ColorRed = col And &HFF
End Function

'/**
' * @brief Extracts the green channel of a packed colour.
' * @param col The packed colour.
' * @return The channel value, 0 to 255.
' */
Public Function ColorGreen(ByVal col As Long) As Long
    ColorGreen = (col \ &H100) And &HFF
End Function

'/**
' * @brief Extracts the blue channel of a packed colour.
' * @param col The packed colour.
' * @return The channel value, 0 to 255.
' */
Public Function ColorBlue(ByVal col As Long) As Long
    ColorBlue = (col \ &H10000) And &HFF
End Function

'/**
' * @brief Splits a packed colour into its three channels in one pass.
' * @param col The packed colour.
' * @param outR Receives the red channel.
' * @param outG Receives the green channel.
' * @param outB Receives the blue channel.
' */
Public Sub ColorUnpack(ByVal col As Long, ByRef outR As Long, ByRef outG As Long, ByRef outB As Long)
    outR = col And &HFF
    outG = (col \ &H100) And &HFF
    outB = (col \ &H10000) And &HFF
End Sub

'/** @section Colour blending */

'/**
' * @brief Multiplies every channel of a colour by a factor.
' * @param col The packed colour.
' * @param factor The brightness multiplier; 1 leaves the colour untouched.
' * @return The scaled colour.
' */
Public Function ColorScale(ByVal col As Long, ByVal factor As Single) As Long
    ColorScale = ColorPack(CLng((col And &HFF) * factor), _
                           CLng(((col \ &H100) And &HFF) * factor), _
                           CLng(((col \ &H10000) And &HFF) * factor))
End Function

'/**
' * @brief Linearly blends two colours.
' * @param colA The colour returned at t = 0.
' * @param colB The colour returned at t = 1.
' * @param t The blend factor, clamped to 0..1.
' * @return The mixed colour.
' */
Public Function ColorMix(ByVal colA As Long, ByVal colB As Long, ByVal t As Single) As Long
    Dim ar As Long, ag As Long, ab As Long
    Dim br As Long, bg As Long, bb As Long

    If t <= 0! Then
        ColorMix = colA
        Exit Function
    End If

    If t >= 1! Then
        ColorMix = colB
        Exit Function
    End If

    ColorUnpack colA, ar, ag, ab
    ColorUnpack colB, br, bg, bb

    ColorMix = ColorPack(CLng(ar + (br - ar) * t), _
                         CLng(ag + (bg - ag) * t), _
                         CLng(ab + (bb - ab) * t))
End Function

'/**
' * @brief Pushes a colour towards white.
' * @param col The packed colour.
' * @param amount How far to travel, from 0 to 1.
' * @return The lightened colour.
' */
Public Function ColorLighten(ByVal col As Long, ByVal amount As Single) As Long
    ColorLighten = ColorMix(col, P_WHITE, amount)
End Function

'/**
' * @brief Pushes a colour towards black.
' * @param col The packed colour.
' * @param amount How far to travel, from 0 to 1.
' * @return The darkened colour.
' */
Public Function ColorDarken(ByVal col As Long, ByVal amount As Single) As Long
    ColorDarken = ColorMix(col, P_BLACK, amount)
End Function

'/**
' * @brief Computes the perceived brightness of a colour.
' * @param col The packed colour.
' * @return The luminance, from 0 to 1.
' */
Public Function ColorLuminance(ByVal col As Long) As Single
    Dim r As Long, g As Long, b As Long

    ColorUnpack col, r, g, b
    ColorLuminance = (0.2126! * r + 0.7152! * g + 0.0722! * b) / 255!
End Function

'/**
' * @brief Converts a colour to its grey equivalent.
' * @param col The packed colour.
' * @return The desaturated colour.
' */
Public Function ColorGrayscale(ByVal col As Long) As Long
    Dim v As Long

    v = CLng(ColorLuminance(col) * 255!)
    ColorGrayscale = ColorPack(v, v, v)
End Function

'/** @section Colour as text */

'/**
' * @brief Parses an HTML style hex string into a packed colour.
' * @param hexText The colour text, with or without a leading hash, in RGB or RRGGBB form.
' * @return The packed colour, or black when the text cannot be parsed.
' * @remarks Every character is validated by hand, so a malformed string returns black instead of raising.
' */
Public Function ColorFromHex(ByVal hexText As String) As Long
    Dim s As String
    Dim i As Long
    Dim v(0 To 5) As Long

    s = Replace$(Trim$(hexText), "#", vbNullString)

    If Len(s) = 3 Then
        s = Left$(s, 1) & Left$(s, 1) & Mid$(s, 2, 1) & Mid$(s, 2, 1) & Right$(s, 1) & Right$(s, 1)
    End If

    If Len(s) <> 6 Then Exit Function

    For i = 0 To 5
        v(i) = HexNibble(Mid$(s, i + 1, 1))
        If v(i) < 0 Then Exit Function
    Next i

    ColorFromHex = ColorPack(v(0) * 16 + v(1), v(2) * 16 + v(3), v(4) * 16 + v(5))
End Function

'/**
' * @brief Converts a single hexadecimal character into its value.
' * @param ch The character to convert.
' * @return The value from 0 to 15, or -1 when the character is not hexadecimal.
' */
Private Function HexNibble(ByVal ch As String) As Long
    Dim c As Long

    c = Asc(UCase$(ch))

    If c >= 48 And c <= 57 Then
        HexNibble = c - 48
    ElseIf c >= 65 And c <= 70 Then
        HexNibble = c - 55
    Else
        HexNibble = -1
    End If
End Function

'/**
' * @brief Renders a packed colour as an HTML style hex string.
' * @param col The packed colour.
' * @return The colour text in #RRGGBB form.
' */
Public Function ColorToHex(ByVal col As Long) As String
    Dim r As Long, g As Long, b As Long

    ColorUnpack col, r, g, b
    ColorToHex = "#" & Right$("0" & Hex$(r), 2) & Right$("0" & Hex$(g), 2) & Right$("0" & Hex$(b), 2)
End Function

'/** @section Clock */

'/**
' * @brief Reads the clock.
' * @return Seconds, with a resolution far finer than a frame.
' * @remarks Returns zero when the platform has no performance counter, which lets a caller subtract two readings and simply get zero rather than a wrong number.
' */
Public Function Seconds() As Double
    Dim c As Currency

    If Not m_ready Then
        m_ready = True
        m_ok = (QueryPerformanceFrequency(m_freq) <> 0)
        If m_freq = 0 Then m_ok = False
    End If

    If Not m_ok Then Exit Function

    QueryPerformanceCounter c
    Seconds = CDbl(c) / CDbl(m_freq)
End Function

'/**
' * @brief Reports whether the clock is usable.
' * @return True when the performance counter answered.
' */
Public Function Available() As Boolean
    If Not m_ready Then Seconds
    Available = m_ok
End Function

Attribute VB_Name = "Misc"
Sub clearDebugConsole()
For i = 0 To 100
    Debug.Print ""
Next i
End Sub

Sub BakCalculateRotationAngles(ByVal oMatrix As Inventor.Matrix, ByRef aRotAngles() As Double)

    Const PI = 3.14159265358979

    Const TODEGREES As Double = 180 / PI

    Dim dB As Double

    Dim dC As Double

    Dim dNumer As Double

    Dim dDenom As Double

    Dim dAcosValue As Double

       

    Dim oRotate As Inventor.Matrix

    Dim oAxis As Inventor.Vector

    Dim oCenter As Inventor.Point

   

    Set oRotate = ThisApplication.TransientGeometry.CreateMatrix

    Set oAxis = ThisApplication.TransientGeometry.CreateVector

    Set oCenter = ThisApplication.TransientGeometry.CreatePoint

 

    oCenter.x = 0

    oCenter.y = 0

    oCenter.Z = 0

 

    ' Choose aRotAngles[0] about x which transforms axes[2] onto the x-z plane

    '

    dB = oMatrix.Cell(2, 3)

    dC = oMatrix.Cell(3, 3)

 

    dNumer = dC

    dDenom = Sqr(dB * dB + dC * dC)

 

    ' Make sure we can do the division.  If not, then axes[2] is already in the x-z plane

    If (Abs(dDenom) <= 0.000001) Then

        aRotAngles(0) = 0#

    Else

        If (dNumer / dDenom >= 1#) Then

            dAcosValue = 0#

        Else

            If (dNumer / dDenom <= -1#) Then

                dAcosValue = PI

            Else

                dAcosValue = Acos(dNumer / dDenom)

            End If

        End If

   

        aRotAngles(0) = Sgn(dB) * dAcosValue

        oAxis.x = 1

        oAxis.y = 0

        oAxis.Z = 0

 

        Call oRotate.SetToRotation(aRotAngles(0), oAxis, oCenter)

        Call oMatrix.PreMultiplyBy(oRotate)

    End If

 

    '

    ' Choose aRotAngles[1] about y which transforms axes[3] onto the z axis

    '

    If (oMatrix.Cell(3, 3) >= 1#) Then

        dAcosValue = 0#

    Else

        If (oMatrix.Cell(3, 3) <= -1#) Then

            dAcosValue = PI

        Else

            dAcosValue = Acos(oMatrix.Cell(3, 3))

        End If

    End If

 

    aRotAngles(1) = Math.Sgn(-oMatrix.Cell(1, 3)) * dAcosValue

    oAxis.x = 0

    oAxis.y = 1

    oAxis.Z = 0

    Call oRotate.SetToRotation(aRotAngles(1), oAxis, oCenter)

    Call oMatrix.PreMultiplyBy(oRotate)

 

    '

    ' Choose aRotAngles[2] about z which transforms axes[0] onto the x axis

    '

    If (oMatrix.Cell(1, 1) >= 1#) Then

        dAcosValue = 0#

    Else

        If (oMatrix.Cell(1, 1) <= -1#) Then

            dAcosValue = PI

        Else

            dAcosValue = Acos(oMatrix.Cell(1, 1))

        End If

    End If

 

    aRotAngles(2) = Math.Sgn(-oMatrix.Cell(2, 1)) * dAcosValue

 

    'if you want to get the result in degrees

    aRotAngles(0) = aRotAngles(0) * TODEGREES

    aRotAngles(1) = aRotAngles(1) * TODEGREES

    aRotAngles(2) = aRotAngles(2) * TODEGREES

End Sub

Sub CalculateRotationAngles(ByVal oMatrix As Inventor.Matrix, ByRef aRotAngles() As Double)
    
    Const PI = 3.14159265358979

    Const TODEGREES As Double = 180 / PI
    
    Dim a11 As Double, a12 As Double, a13 As Double
    Dim a21 As Double, a22 As Double, a23 As Double
    Dim a31 As Double, a32 As Double, a33 As Double
    
    
    'WTF VB?
    a11 = oMatrix.Cell(1, 1)
    a12 = oMatrix.Cell(2, 1)
    a13 = oMatrix.Cell(3, 1)
    a21 = oMatrix.Cell(1, 2)
    a22 = oMatrix.Cell(2, 2)
    a23 = oMatrix.Cell(3, 2)
    a31 = oMatrix.Cell(1, 3)
    a32 = oMatrix.Cell(2, 3)
    a33 = oMatrix.Cell(3, 3)
    
    Debug.Print ("ROT MATRIX:")
    Debug.Print (a11 & " " & a12 & " " & a13)
    Debug.Print (a21 & " " & a22 & " " & a23)
    Debug.Print (a31 & " " & a32 & " " & a33)
    
     
    
    Dim a_list(1) As Double, b_list(1) As Double, c_list(1) As Double
    
    
    b_list(0) = Acos(a33)
    b_list(1) = Acos(a33)
    c_list(0) = Atan2(a31 / Math.Cos(b_list(0)), a32 / Math.Cos(b_list(0)))
    c_list(1) = Atan2(a31 / Math.Cos(b_list(1)), a32 / Math.Cos(b_list(1)))
    a_list(0) = Atan2(a13 / Math.Cos(b_list(0)), -a23 / Math.Cos(b_list(0)))
    a_list(1) = Atan2(a13 / Math.Cos(b_list(1)), -a23 / Math.Cos(b_list(1)))
    
    Dim min_a_index As Integer, min_b_index As Integer, min_c_index As Integer
     
    min_a_index = FindMinAbsIndex(a_list)
    min_b_index = FindMinAbsIndex(b_list)
    min_c_index = FindMinAbsIndex(c_list)
    
    aRotAngles(0) = a_list(1)
    aRotAngles(1) = b_list(1)
    aRotAngles(2) = c_list(1)

    'if you want to get the result in degrees

    aRotAngles(0) = aRotAngles(0) * TODEGREES

    aRotAngles(1) = aRotAngles(1) * TODEGREES

    aRotAngles(2) = aRotAngles(2) * TODEGREES

End Sub

Function Atan2(y As Double, x As Double) As Double
' returned value is in radians
    If x > 0 Then
        Atan2 = Math.Atn(y / x)
    ElseIf (x < 0 And y >= 0) Then
        Atan2 = VBA.Atn(y / x) + PI
    ElseIf (x < 0 And y < 0) Then
        Atan2 = VBA.Atn(y / x) - PI
    ElseIf (x = 0 And y > 0) Then
        Atan2 = PI / 2
    ElseIf (x = 0 And y < 0) Then
        Atan2 = -PI / 2
    ElseIf (x = 0 And y = 0) Then
        Atan2 = 0
    End If
End Function
Function Asin(x As Double) As Double
' returned value is in radians
    Asin = VBA.Atn(x / VBA.Sqr(-x * x + 1))
End Function

Private Function FindMinAbsIndex(v() As Double) As Integer
    Dim minValue As Double
    Dim minIndex As Integer
    
    minValue = v(0)
    minIndex = 0
    For i = 1 To UBound(v)
        If (Math.Abs(minValue) > Math.Abs(v(i))) Then
            minValue = v(i)
            minIndex = i
        End If
    Next
    FindMinAbsIndex = minIndex
End Function
Private Function Acos(value) As Double

    Acos = Math.Atn(-value / Math.Sqr(-value * value + 1)) + 2 * Math.Atn(1)

End Function

   Sub AppendFiles(destinyName As String, sourceName As String)

      Dim SourceNum As Integer
      Dim DestNum As Integer
      Dim Temp As String

      ' If an error occurs, close the files and end the macro.
      On Error GoTo ErrHandler

      ' Open the destination text file.
      DestNum = FreeFile()
      Open destinyName For Append As DestNum

      ' Open the source text file.
      SourceNum = FreeFile()
      Open sourceName For Input As SourceNum

      ' Include the following line if the first line of the source
      ' file is a header row that you do now want to append to the
      ' destination file:
      ' Line Input #SourceNum, Temp

      ' Read each line of the source file and append it to the
      ' destination file.
      Do While Not EOF(SourceNum)
         Line Input #SourceNum, Temp
         Print #DestNum, Temp
      Loop

CloseFiles:

      ' Close the destination file and the source file.
      Close #DestNum
      Close #SourceNum
      Exit Sub

ErrHandler:
      MsgBox "Error # " & Err & ": " & Error(Err)
      Resume CloseFiles

   End Sub


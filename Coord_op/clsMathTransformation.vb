Option Explicit On

Public Class clsMathTransformation
    ' /**
    ' * A 3D transformation.
    ' * @author
    ' *     Torge Riedel (TRi)
    ' * @date
    ' *     02.09.2003
    ' */


    ' /**
    ' * The transformation's matrix.
    ' */
    Public matrix As New clsMath3x3Matrix

    ' /**
    ' * The transformation's vector.
    ' */
    Public vector As New clsMathVector
    '

    ' /**
    ' *
    ' */
    Public Sub Init()
        Call matrix.Init()
        Call vector.Init()
    End Sub

    ' /**
    ' * Sets the transformation-coefs by an array of doubles.
    ' * @param coefs
    ' *     Array with double values.
    ' */
    Public Sub SetByArray(coefs() As Double)
        Call matrix.SetByArray(coefs)
        Dim vc(0 To 2) As Double
        Dim i As Integer
        For i = 9 To 11
            vc(i - 9) = coefs(i)
        Next
        Call vector.SetByArray(vc)
    End Sub

    ' /**
    ' * Sets the transformation-coefs by an array of variants.
    ' * @param coefs
    ' *     Array with variant values.
    ' */
    Public Sub SetByArrayV(coefs() As Object)
        Call matrix.SetByArrayV(coefs)
        Dim vc(0 To 2) As Object
        Dim i As Integer
        For i = 9 To 11
            vc(i - 9) = coefs(i)
        Next
        '最后一列3个数值变为法向
        Call vector.SetByArrayV(vc)
    End Sub

    ' /**
    ' * Returns the transformation-coefs in an array of doubles.
    ' * @param coefs
    ' *     Array with double values.
    ' */
    Public Sub GetArray(coefs() As Double)
        Call matrix.GetArray(coefs)
        Dim vc(0 To 2) As Double
        Call vector.GetArray(vc)
        Dim i As Integer
        For i = 9 To 11
            coefs(i) = vc(i - 9)
        Next
    End Sub

    ' /**
    ' * Returns the transformation-coefs in an array of variants.
    ' * @param coefs
    ' *     Array with variant values.
    ' */
    Public Sub GetArrayV(coefs() As Object)
        Call matrix.GetArrayV(coefs)
        Dim vc(0 To 2) As Object
        Call vector.GetArrayV(vc)
        Dim i As Integer
        For i = 9 To 11
            coefs(i) = vc(i - 9)
        Next
    End Sub

    ' /**
    ' * Copies the coefs from another transformation.
    ' * @param trans
    ' *     The transformation to copy the coefs from.
    ' */
    Public Sub Copy(ByVal trans As clsMathTransformation)
        Call matrix.Copy(trans.matrix)
        Call vector.Copy(trans.vector)
    End Sub

    ' /**
    ' * Initializes this transformation as a translation.
    ' * @param vec
    ' *     The vector defining the translation.
    ' */
    Public Sub SetTranslation(ByVal vec As clsMathVector)
        Dim coefs(0 To 8) As Double
        coefs(0) = 1
        coefs(4) = 1
        coefs(8) = 1
        Call matrix.SetByArray(coefs)
        Call vector.Copy(vec)
    End Sub

    ' /**
    ' * Multiplies a point with this transformation.
    ' * @param iPoint
    ' *     The point to multiply with.
    ' * @returns
    ' *     The resulting point.
    ' */
    Public Function MulVec(ByVal iPoint As clsMathVector) As clsMathVector
        MulVec = New clsMathVector
        Call MulVec.Copy(vector.Add(matrix.MulVec(iPoint)))
    End Function

    ' /**
    ' * Multiplies a transformation matrix with this transformation.
    ' * @param iTransfo
    ' *     The transformation matrix to multiply with.
    ' * @returns
    ' *     The resulting transformation matrix.
    ' */
    Public Function MulTrans(ByVal iTransfo As clsMathTransformation) As clsMathTransformation
        MulTrans = New clsMathTransformation

        MulTrans.matrix.col1.cx = matrix.col1.cx * iTransfo.matrix.col1.cx + _
                                  matrix.col2.cx * iTransfo.matrix.col1.cy + _
                                  matrix.col3.cx * iTransfo.matrix.col1.cz
        MulTrans.matrix.col1.cy = matrix.col1.cy * iTransfo.matrix.col1.cx + _
                                  matrix.col2.cy * iTransfo.matrix.col1.cy + _
                                  matrix.col3.cy * iTransfo.matrix.col1.cz
        MulTrans.matrix.col1.cz = matrix.col1.cz * iTransfo.matrix.col1.cx + _
                                  matrix.col2.cz * iTransfo.matrix.col1.cy + _
                                  matrix.col3.cz * iTransfo.matrix.col1.cz
        MulTrans.matrix.col2.cx = matrix.col1.cx * iTransfo.matrix.col2.cx + _
                                 matrix.col2.cx * iTransfo.matrix.col2.cy + _
                                 matrix.col3.cx * iTransfo.matrix.col2.cz
        MulTrans.matrix.col2.cy = matrix.col1.cy * iTransfo.matrix.col2.cx + _
                                  matrix.col2.cy * iTransfo.matrix.col2.cy + _
                                  matrix.col3.cy * iTransfo.matrix.col2.cz
        MulTrans.matrix.col2.cz = matrix.col1.cz * iTransfo.matrix.col2.cx + _
                                  matrix.col2.cz * iTransfo.matrix.col2.cy + _
                                  matrix.col3.cz * iTransfo.matrix.col2.cz
        MulTrans.matrix.col3.cx = matrix.col1.cx * iTransfo.matrix.col3.cx + _
                                  matrix.col2.cx * iTransfo.matrix.col3.cy + _
                                  matrix.col3.cx * iTransfo.matrix.col3.cz
        MulTrans.matrix.col3.cy = matrix.col1.cy * iTransfo.matrix.col3.cx + _
                                  matrix.col2.cy * iTransfo.matrix.col3.cy + _
                                  matrix.col3.cy * iTransfo.matrix.col3.cz
        MulTrans.matrix.col3.cz = matrix.col1.cz * iTransfo.matrix.col3.cx + _
                                  matrix.col2.cz * iTransfo.matrix.col3.cy + _
                                  matrix.col3.cz * iTransfo.matrix.col3.cz
        MulTrans.vector.cx = matrix.col1.cx * iTransfo.vector.cx + _
                             matrix.col2.cx * iTransfo.vector.cy + _
                             matrix.col3.cx * iTransfo.vector.cz + _
                             vector.cx * 1
        MulTrans.vector.cy = matrix.col1.cy * iTransfo.vector.cx + _
                             matrix.col2.cy * iTransfo.vector.cy + _
                             matrix.col3.cy * iTransfo.vector.cz + _
                             vector.cy * 1
        MulTrans.vector.cz = matrix.col1.cz * iTransfo.vector.cx + _
                             matrix.col2.cz * iTransfo.vector.cy + _
                             matrix.col3.cz * iTransfo.vector.cz + _
                             vector.cz * 1
    End Function



End Class


Option Explicit On

Public Class clsMath3x3Matrix
    ' /**
    ' * A 3x3 matrix.
    ' * @author
    ' *     Torge Riedel (TRi)
    ' * @date
    ' *     02.09.2003
    ' */


    ' /**
    ' * First column
    ' */
    Public col1 As New clsMathVector

    ' /**
    ' * Second column
    ' */
    Public col2 As New clsMathVector

    ' /**
    ' * Third column
    ' */
    Public col3 As New clsMathVector

    ' /**
    ' *
    ' */
    Public Sub Init()
        Call col1.SetVec(1, 0, 0)
        Call col2.SetVec(0, 1, 0)
        Call col3.SetVec(0, 0, 1)
    End Sub

    ' /**
    ' *
    ' */
    Public Sub SetMat( _
                ByVal a11 As Double, ByVal a12 As Double, ByVal a13 As Double, _
                ByVal a21 As Double, ByVal a22 As Double, ByVal a23 As Double, _
                ByVal a31 As Double, ByVal a32 As Double, ByVal a33 As Double)

        Call col1.SetVec(a11, a21, a31)
        Call col2.SetVec(a12, a22, a32)
        Call col3.SetVec(a13, a23, a33)
    End Sub

    ' /**
    ' * Sets the matrix-coefs by an array of doubles.
    ' * @param coefs
    ' *     Array with double values.
    ' */
    Public Sub SetByArray(coefs() As Double)
        Call col1.SetByArray(coefs)
        Dim vc(0 To 2) As Double
        Dim i As Integer
        For i = 3 To 5
            vc(i - 3) = coefs(i)
        Next
        Call col2.SetByArray(vc)
        For i = 6 To 8
            vc(i - 6) = coefs(i)
        Next
        Call col3.SetByArray(vc)
    End Sub

    ' /**
    ' * Sets the matrix-coefs by an array of variants.
    ' * @param coefs
    ' *     Array with variant values.
    ' */
    Public Sub SetByArrayV(coefs() As Object)
        Call col1.SetByArrayV(coefs)
        Dim vc(0 To 2) As Object
        Dim i As Integer
        For i = 3 To 5
            vc(i - 3) = coefs(i)
        Next
        Call col2.SetByArrayV(vc)
        For i = 6 To 8
            vc(i - 6) = coefs(i)
        Next
        Call col3.SetByArrayV(vc)
    End Sub


    ' /**
    ' * Initializes this matrix as a scaling matrix.
    ' * @param fac
    ' *     The scaling factor.
    ' */
    Public Sub SetScaling(ByVal fac As Double)
        Dim coefs(0 To 8) As Double
        coefs(0) = fac
        coefs(4) = fac
        coefs(8) = fac
        Call SetByArray(coefs)
    End Sub

    ' /**
    ' * @nodoc
    ' */
    Public Sub SetRotation(ByVal axis As clsMathVector, ByVal angle As Double)
    End Sub

    ' /**
    ' * Returns the matrix-coefs in an array of doubles.
    ' * @param coefs
    ' *     Array with double values.
    ' */
    Public Sub GetArray(coefs() As Double)
        Call col1.GetArray(coefs)
        Dim vc(0 To 2) As Double
        Call col2.GetArray(vc)
        Dim i As Integer
        For i = 0 To 2
            coefs(i + 3) = vc(i)
        Next
        Call col3.GetArray(vc)
        For i = 0 To 2
            coefs(i + 6) = vc(i)
        Next
    End Sub

    ' /**
    ' * Returns the matrix-coefs in an array of variants.
    ' * @param coefs
    ' *     Array with variant values.
    ' */
    Public Sub GetArrayV(coefs() As Object)
        Call col1.GetArrayV(coefs)
        Dim vc(0 To 2) As Object
        Call col2.GetArrayV(vc)
        Dim i As Integer
        For i = 0 To 2
            coefs(i + 3) = vc(i)
        Next
        Call col3.GetArrayV(vc)
        For i = 0 To 2
            coefs(i + 6) = vc(i)
        Next
    End Sub


    ' /**
    ' * Copies the coefs from another matrix.
    ' * @param mat
    ' *     The matrix to copy the coefs from.
    ' */
    Public Sub Copy(ByVal mat As clsMath3x3Matrix)
        Call col1.Copy(mat.col1)
        Call col2.Copy(mat.col2)
        Call col3.Copy(mat.col3)
    End Sub

    ' /**
    ' * Computes the determinant of <tt>this</tt> matrix.
    ' * @returns
    ' *     The determinant.
    ' */
    Public Function GetDeterminant() As Double
        GetDeterminant = col1.cx * (col2.cy * col3.cz - col3.cy * col2.cz) - _
                         col1.cy * (col2.cx * col3.cz - col3.cx * col2.cz) + _
                         col1.cz * (col2.cx * col3.cy - col3.cx * col2.cy)
    End Function

    ' /**
    ' * @nodoc
    ' */
    Public Function GetInversed() As clsMath3x3Matrix
        GetInversed = New clsMath3x3Matrix
    End Function

    ' /**
    ' * Computes the transposed of <tt>this</tt> matrix.
    ' * @returns
    ' *     The transposed matrix.
    ' */
    Public Function GetTransposed() As clsMath3x3Matrix
        GetTransposed = New clsMath3x3Matrix
        Dim srccoefs(0 To 8) As Double
        Call GetArray(srccoefs)
        Dim destcoefs(0 To 8) As Double
        Dim i As Integer
        For i = 0 To 2
            destcoefs(i) = srccoefs(i * 3)
            destcoefs(i + 3) = srccoefs(i * 3 + 1)
            destcoefs(i + 6) = srccoefs(i * 3 + 2)
        Next
        Call GetTransposed.SetByArray(destcoefs)
    End Function

    ' /**
    ' * Multiplies a vector with <tt>this</tt> matrix.
    ' * @param vec
    ' *     The vector to multiply <tt>this</tt> matrix with.
    ' * @returns
    ' *     The result vector.
    ' */
    Public Function MulVec(ByVal vec As clsMathVector) As clsMathVector
        MulVec = New clsMathVector
        MulVec.cx = vec.cx * col1.cx + vec.cy * col2.cx + vec.cz * col3.cx
        MulVec.cy = vec.cx * col1.cy + vec.cy * col2.cy + vec.cz * col3.cy
        MulVec.cz = vec.cx * col1.cz + vec.cy * col2.cz + vec.cz * col3.cz
    End Function

    ' /**
    ' * Multiplies a matrix with <tt>this</tt> matrix.
    ' * @param mat
    ' *     The matrix to multiply <tt>this</tt> matrix with.
    ' * @returns
    ' *     The result matrix.
    ' */
    Public Function MulMat(ByVal mat As clsMath3x3Matrix) As clsMath3x3Matrix
        MulMat = New clsMath3x3Matrix
        MulMat.col1 = MulVec(mat.col1)
        MulMat.col2 = MulVec(mat.col2)
        MulMat.col3 = MulVec(mat.col3)
    End Function





End Class

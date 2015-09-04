Option Explicit On
Public Class clsMathVector
    ' /**
    ' * A 3D vector.
    ' * @author
    ' *     Torge Riedel (TRi)
    ' * @date
    ' *     02.09.2003
    ' */


    ' /**
    ' * component in x-direction.
    ' */
    Public cx As Double

    ' /**
    ' * component in y-direction.
    ' */
    Public cy As Double

    ' /**
    ' * component in z-direction.
    ' */
    Public cz As Double
    '

    ' /**
    ' *
    ' */
    Public Sub Init()
        cx = 0
        cy = 0
        cz = 0
    End Sub

    ' /**
    ' * Sets the vector-components by an array of doubles.
    ' * @param comps
    ' *     Array with double values.
    ' */
    Public Sub SetByArray(comps() As Double)
        cx = comps(0)
        cy = comps(1)
        cz = comps(2)
    End Sub

    ' /**
    ' * Sets the vector-components by an array of variants.
    ' * @param comps
    ' *     Array with variant values.
    ' */
    Public Sub SetByArrayV(comps() As Object)
        cx = comps(0)
        cy = comps(1)
        cz = comps(2)
    End Sub

    ' /**
    ' * Returns the vector-components in an array of doubles.
    ' * @param comps
    ' *     Array with double values.
    ' */
    Public Sub GetArray(comps() As Double)
        comps(0) = cx
        comps(1) = cy
        comps(2) = cz
    End Sub

    ' /**
    ' * Returns the vector-components in an array of variants.
    ' * @param comps
    ' *     Array with variant values.
    ' */
    Public Sub GetArrayV(comps() As Object)
        comps(0) = cx
        comps(1) = cy
        comps(2) = cz
    End Sub

    ' /**
    ' * Sets the vector components.
    ' * @param acx
    ' *      Component in x-direction.
    ' * @param acy
    ' *      Component in y-direction.
    ' * @param acz
    ' *      Component in z-direction.
    ' */
    Public Sub SetVec(ByVal acx As Double, ByVal acy As Double, ByVal acz As Double)
        cx = acx
        cy = acy
        cz = acz
    End Sub

    ' /**
    ' * Copies the components from another vector.
    ' * @param vec
    ' *     The vector to copy the components from.
    ' */
    Public Sub Copy(ByVal vec As clsMathVector)
        cx = vec.cx
        cy = vec.cy
        cz = vec.cz
    End Sub

    ' /**
    ' * Creates a copy of <tt>this</tt> vector.
    ' * @returns
    ' *       A copy of this vector.
    ' */
    Public Function GetCopy() As clsMathVector
        GetCopy = New clsMathVector
        Call GetCopy.Copy(Me)
    End Function

    ' /**
    ' * Checks whether <tt>this</tt> is a Null-vector. That is, all components equal 0.
    ' * @returns
    ' *     <tt>True</tt>, if <tt>this</tt> is a Null-vector, otherwise <tt>False</tt>.
    ' */
    Public Function IsNull() As Boolean
        IsNull = ((0 = cx) And (0 = cy) And (0 = cz))
    End Function

    ' /**
    ' * Checks whether <tt>this</tt> vector is equal to another vector.
    ' * @param vec
    ' *     Vector to compare with.
    ' * @returns
    ' *     <tt>True</tt>, if <tt>this</tt> vector is equal to <tt>vec</tt>, otherwise <tt>False</tt>.
    ' */
    Public Function IsEqual(ByVal vec As clsMathVector) As Boolean
        IsEqual = ((cx = vec.cx) And (cy = vec.cy) And (cz = vec.cz))
    End Function

    ' /**
    ' * Checks whether <tt>this</tt> vector is not equal to another vector.
    ' * @param vec
    ' *     Vector to compare with.
    ' * @returns
    ' *     <tt>True</tt>, if <tt>this</tt> vector is not equal to <tt>vec</tt>, otherwise <tt>False</tt>.
    ' */
    Public Function IsNotEqual(ByVal vec As clsMathVector) As Boolean
        IsNotEqual = ((cx <> vec.cx) Or (cy <> vec.cy) Or (cz <> vec.cz))
    End Function

    ' /**
    ' * Computes the length of <tt>this</tt> vector.
    ' * @returns
    ' *     Length of vector.
    ' */
    Public Function GetLength() As Double
        GetLength = Math.Sqrt(cx ^ 2 + cy ^ 2 + cz ^ 2)
    End Function

    ' /**
    ' * Sets the length of the vector. The direction will not change.
    ' * @param alen
    ' *      The new length for <tt>this</tt> vector.
    ' */
    Public Sub SetLength(ByVal alen As Double)
        Call Normalize()
        Call Scale2(alen)
    End Sub

    ' /**
    ' * Scales this vector.
    ' * @param scalar
    ' *     Scaling factor.
    ' */
    Public Sub Scale2(ByVal scalar As Double)
        cx = scalar * cx
        cy = scalar * cy
        cz = scalar * cz
    End Sub

    ' /**
    ' * Normalizes this vector. That is setting it's length to 1.
    ' */
    Public Sub Normalize()
        Dim nLen As Double
        nLen = Math.Sqrt(cx ^ 2 + cy ^ 2 + cz ^ 2)
        cx = cx / nLen
        cy = cy / nLen
        cz = cz / nLen
    End Sub

    ' /**
    ' * Returns a normalized vector of <tt>this</tt> vector.
    ' * @returns
    ' *     Normalized vector.
    ' * @see Normalize
    ' */
    Public Function GetNormalized() As clsMathVector
        GetNormalized = GetCopy()
        Call GetNormalized.Normalize()
    End Function

    ' /**
    ' * Checks whether <tt>this</tt> vector and another are parallel.
    ' * @param vec
    ' *     Vector to check parallelism.
    ' * @returns
    ' *     <tt>True</tt>, if <tt>this</tt> and the other vector are parallel, otherwise <tt>False</tt>.
    ' */
    Public Function IsParallelTo(ByVal vec As clsMathVector) As Boolean
        IsParallelTo = MulVec(vec).IsNull()
    End Function

    ' /**
    ' * Checks whether <tt>this</tt> vector and another are orthogonal.
    ' * @param vec
    ' *     Vector to check orthogonal.
    ' * @returns
    ' *     <tt>True</tt>, if <tt>this</tt> and the other vector are orthogonal, otherwise <tt>False</tt>.
    ' */
    Public Function IsOrthogonalTo(ByVal vec As clsMathVector) As Boolean
        IsOrthogonalTo = (0 = Mul(vec))
    End Function


    ' /**
    ' * Multiplies this vector with a scalar.
    ' * @param scalar
    ' *     Scalar to multiply with.
    ' * @returns
    ' *     The result vector.
    ' */
    Public Function MulS(ByVal scalar As Double) As clsMathVector
        MulS = GetCopy()
        Call MulS.Scale2(scalar)
    End Function

    ' /**
    ' * Adds two vectors.
    ' * @param vec
    ' *     The vector to add.
    ' * @returns
    ' *     The result vector.
    ' */
    Public Function Add(ByVal vec As clsMathVector) As clsMathVector
        Add = New clsMathVector
        Add.cx = cx + vec.cx
        Add.cy = cy + vec.cy
        Add.cz = cz + vec.cz
    End Function

    ' /**
    ' * Subtracts two vectors.
    ' * @param vec
    ' *     The vector to subtract.
    ' * @returns
    ' *     The result vector.
    ' */
    Public Function Diff(ByVal vec As clsMathVector) As clsMathVector
        Diff = New clsMathVector
        Diff.cx = cx - vec.cx
        Diff.cy = cy - vec.cy
        Diff.cz = cz - vec.cz
    End Function

    ' /**
    ' * Scalar-Multiplication of two vectors.
    ' * @param vec
    ' *     The vector to multiply.
    ' * @returns
    ' *     The resulting scalar.
    ' */
    Public Function Mul(ByVal vec As clsMathVector) As Double
        Mul = cx * vec.cx + cy * vec.cy + cz * vec.cz
    End Function

    ' /**
    ' * Multiplication of two vectors.
    ' * @param vec
    ' *     The vector to multiply.
    ' * @returns
    ' *     The result vector.
    ' */
    Public Function MulVec(ByVal vec As clsMathVector) As clsMathVector
        MulVec = New clsMathVector
        MulVec.cx = cy * vec.cz - vec.cy * cz
        MulVec.cy = -(cx * vec.cz - vec.cx * cz)
        MulVec.cz = cx * vec.cy - vec.cx * cy
    End Function



End Class

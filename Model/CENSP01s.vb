Option Explicit On

Imports ProductStructureTypeLib
Imports mysqlsolution

Public Class CENSP01s


    Private SP01_ As New Dictionary(Of String, CENSP01)
    '  Public FastNames_ As List(Of String)
    'comment

    '
    '   Default constructor
    '
    Public Sub New()

    End Sub

    '
    '   Default destructor
    '
    Protected Overrides Sub Finalize()
        RemoveAll()
    End Sub

    '
    '   Adds a rivet to a collection
    '
    Public Sub Add(ByRef SP01 As CENSP01)
        If (InTheList(SP01.X & "_" & SP01.Y & "_" & SP01.Z) = False) Then
            SP01_.Add(SP01.X & "_" & SP01.Y & "_" & SP01.Z, SP01)
        Else
            'MsgBox ("There are two Fasteners for " & SP01.name & " at " & SP01.X & "_" & SP01.Y & "_" & SP01.Z)
        End If
    End Sub

    '
    '   Returns a rivet based on the Index that has to be withing the legal bounds
    '
    Public Function Item(ID As String) As CENSP01

        Item = SP01_.Item(ID)

    End Function

    Public Function Item2(Index As Integer) As CENSP01

        Item2 = SP01_.Values(Index - 1)

    End Function
    Public Sub New(ByRef SPProducts As IEnumerable(Of Product))
        ' FastNames_ = fstlist


        walk(SPProducts)

    End Sub

    Private Sub walk(ByRef SPProducts As IEnumerable(Of Product))


        Dim i
        For i = 0 To SPProducts.Count - 1

            'Go through each fastener geoset

            Dim MyProduct = SPProducts.ElementAt(i)
            Dim CATIA = MyProduct.Application
            Dim j
            '遍历带逗号的层
            For j = 1 To MyProduct.Products.Count

                Dim FastenerName As String
                Dim tmpfastname = MyProduct.Products.Item(j).Name
                If tmpfastname.Contains(",") Then
                    FastenerName = Mid(tmpfastname, 1, InStr(tmpfastname, ",") - 1)
                Else

                    FastenerName = tmpfastname
                End If

                If FastenerName.Contains(".") Then

                    FastenerName = Strings.Split(FastenerName, ".")(0)

                End If

                If FastenerName.Contains("D01") Then

                    FastenerName = Strings.Split(FastenerName, "D01")(0)

                End If


                'Make sure it is a legit fastener
                If FastenerName.Contains("-") And (FastenerName.Length = 13 Or FastenerName.Length = 14) Then

                    'Go through each component
                    Dim k
                    '遍历下一层，直接取铆钉组件所在位置
                    For k = 1 To MyProduct.Products.Item(j).Products.Count
                        'Store the fastener name and XYZ
                        Dim X As Double, Y As Double, Z As Double, MyName As String
                        MyName = FastenerName
                        'Get XYZ
                        Dim pos
                        pos = MyProduct.Products.Item(j).Products.Item(k).Position
                        Dim matrix(11)
                        pos.GetComponents(matrix)

                        'Put the XYZ in a 1x3 array since thats what we need to call the point transformation
                        X = matrix(9)
                        Y = matrix(10)
                        Z = matrix(11)

                        Dim cCoords(2)
                        cCoords(0) = X
                        cCoords(1) = Y
                        cCoords(2) = Z

                        'Transform wrt parent?
                        Dim ProductTransformation As clsMathTransformation
                        Dim PointTransformation As New clsMathVector
                        Dim aDoc
                        Dim TopProduct As Product
                        aDoc = CATIA.ActiveDocument
                        TopProduct = aDoc.Product

                        ProductTransformation = mdlProductTransformation.GetAbsPosition(MyProduct.Products.Item(j))
                        Call PointTransformation.SetByArrayV(cCoords)
                        PointTransformation = ProductTransformation.MulVec(PointTransformation)


                        Dim MyFast As CENSP01
                        'Add data to collection only if the rivet really exists
                        If (Len(MyName) > 0) Then


                            MyFast = New CENSP01

                            With MyFast
                                .name = MyName
                                .X = PointTransformation.cx
                                .Y = PointTransformation.cy
                                .Z = PointTransformation.cz
                            End With

                            'Add data to a collection
                            Add(MyFast)

                        End If

                        MyFast = Nothing





                    Next
                Else
                    'Skip it
                    'MsgBox ("Fastener " & FastenerName & " not in the lookup table")
                End If
            Next
        Next



    End Sub




    Public Function to_points() As CENPoints

        Dim tmppoints As New CENPoints

        For Each kk As CENSP01 In SP01_.Values
            tmppoints.Add(kk.to_point())

        Next
        Return tmppoints

    End Function
    '
    '   Returns the number of elements in a collection
    '
    Public Function count() As Integer
        count = SP01_.Count
    End Function

    '
    '   Removes an element from a collection
    '
    Public Sub Remove(Index As Integer)

        If (Index > 0 And Index <= SP01_.Count) Then


            SP01_.Remove(SP01_.Keys.ElementAt(Index - 1))
        End If

    End Sub

    '
    '   Removes all elements from a collection
    '
    Public Sub RemoveAll()

        SP01_.Clear()

    End Sub

    '
    '   Performs a check whether or not a rivet is in the list or not.
    '   Search is performed by the name of a rivet
    '
    Public Function InTheList(ByRef name As String) As Boolean

        Return SP01_.Keys.Contains(name)

    End Function

End Class

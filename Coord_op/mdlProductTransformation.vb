
Option Explicit On
Imports ProductStructureTypeLib
Imports MECMOD


Public Module mdlProductTransformation
    ' /**
    ' * Provides some useful function for the CATIA product-structure.<br>
    ' * Uses the following files:<br>
    ' * <ul><li>clsMathTransformation</li>
    ' * <li>clsObjectList</li></ul>
    ' * @author
    ' *     Torge Riedel (TRi)
    ' * @date
    ' *     03.09.2003
    ' */


    Private Const sFILTER_PRODUCT As String = "Product"
    Private Const sFILTER_PART As String = "Part"
    Private Const sFILTER_POINT As String = "Point"
    Private Const sFILTER_SHAPEPOINT As String = "ShapePoint"
    Private Const sFILTER_PROJECTION As String = "HybridShapeProject"
    Private Const sFILTER_HYBRIDSHAPE As String = "HybridShape"
    Private Const sFILTER_INTERSECTION As String = "HybridShapeIntersection"
    Private Const sFILTER_SYMMETRY As String = "HybridShapeSymmetry"
    '

    ' /**
    ' * Returns the transformation for a specified CATIA-product.
    ' * @param iProduct
    ' *     Product with transformation.
    ' * @return
    ' *     The product's transformation.
    ' */
    Public Function GetProductTransformation(ByVal iProduct As Product) As clsMathTransformation
        GetProductTransformation = New clsMathTransformation
        Dim dummy As Object
        dummy = iProduct
        Dim comps(0 To 11) As Object
        '获取最顶层组件的坐标系和法向
        Call dummy.Position.GetComponents(comps)
        Call GetProductTransformation.SetByArrayV(comps)
    End Function


    ' /**
    ' * Computes the absolute position (transformation) of a product in its context.
    ' * @param iProduct
    ' *     Product to compute the absolute position from.
    ' * @return
    ' *     Absolute position (transformation) of product in its context.
    ' */
    Public Function GetAbsPosition(ByVal iProduct As Product) As clsMathTransformation
        ' parent is the product-list of parent-product
        Dim ProdList As Products
        On Error Resume Next
        ProdList = iProduct.Parent
        On Error GoTo 0
        If (ProdList Is Nothing) Then
            ' -> top-most product reached, return std-trans as abs-position
            GetAbsPosition = New clsMathTransformation
            Call GetAbsPosition.Init()
            Exit Function
        End If

        ' get parent-product
        Dim parentprod As Product
        On Error Resume Next
        parentprod = ProdList.Parent
        On Error GoTo 0
        If (parentprod Is Nothing) Then
            ' -> top-most product reached, return std-trans as abs-position
            GetAbsPosition = New clsMathTransformation
            Call GetAbsPosition.Init()
            Exit Function
        End If
        '不断遍历 ，直到 iProduct变为最顶层，并且初始化matrix

        ' get abs-trans of parent-prod
        GetAbsPosition = GetAbsPosition(parentprod)

        ' get trans of this prod
        '储存了整个矩阵
        Dim thistrans As clsMathTransformation
        thistrans = GetProductTransformation(iProduct)
        '把每一层product的变化迭代施加变化 直到得到与最顶层一样的坐标变化
        ' compute absolute position
        '对thistrans实行GetAbsPosition的变化
        GetAbsPosition = GetAbsPosition.MulTrans(thistrans)
    End Function


    ' Public Function GetPartOf(ByVal oObject As Object) As Part
    ' -------------------------------------------------------------------------------
    '
    ' Transfer: oObject: "Joint Definition( xx)"-geometrical set
    '
    ' Retrieves the part of the transfered "Joint Definition( xx)"-geometrical set
    '
    ' -------------------------------------------------------------------------------
    Public Function GetPartOf(ByVal oObject As Object) As Part

        Dim count
        count = 0

        While TypeName(oObject) <> sFILTER_PART And count < 50
            oObject = oObject.Parent
            count = count + 1
        End While

        GetPartOf = oObject

    End Function
    Public Function GetAbsPosition(ByVal TopProd As Product, ByVal iProduct As Product) As clsMathTransformation
        Dim absPositionTopProduct As clsMathTransformation
        Dim parent As Products = Nothing
        Try
            parent = DirectCast(iProduct.Parent, Products)
        Catch exception1 As Exception

            Dim exception As Exception = exception1
            absPositionTopProduct = New clsMathTransformation
            absPositionTopProduct.Init()

            Return absPositionTopProduct
        End Try
        Dim product As Product = Nothing
        Try
            product = DirectCast(parent.Parent, Product)
        Catch exception3 As Exception

            Dim exception2 As Exception = exception3

        End Try
        If (product.Name = TopProd.Name) Then
            absPositionTopProduct = New clsMathTransformation
            absPositionTopProduct.Init()
        Else
            absPositionTopProduct = mdlProductTransformation.GetAbsPosition(TopProd, product)
        End If
        Dim productTransformation As clsMathTransformation = mdlProductTransformation.GetProductTransformation(iProduct)
        Return absPositionTopProduct.MulTrans(productTransformation)
    End Function
    Public Sub GetFatherProduct(ByVal oProduct As Product, ByVal oPartDoc As PartDocument, ByRef oFatherProduct As Product, Optional ByRef sPath As Object = Nothing)
        Dim ii As Integer
        Dim oSubproduct As Product

        On Error Resume Next

        If oFatherProduct Is Nothing Then

            ' browse all elements in transfered product
            For ii = 1 To oProduct.Products.Count
                oSubproduct = oProduct.Products.Item(ii)
                Err.Clear()
                ' found part is part whose father product is searched ?
                Dim tmpfullname As String
                tmpfullname = oSubproduct.ReferenceProduct.Parent.FullName
                If tmpfullname = oPartDoc.FullName Then

                    ' is found element actually a part ?
                    If Err.Number = 0 Then
                        oFatherProduct = oSubproduct
                        If Not sPath Is Nothing Then
                            sPath = oSubproduct.Name + "." + sPath
                        End If
                        Exit For
                    End If
                End If

                ' element is a product
                ' -> browse it
                If oSubproduct.Products.Count <> 0 And oFatherProduct Is Nothing Then

                    ' recursive call
                    Call GetFatherProduct(oSubproduct, oPartDoc, oFatherProduct, sPath)
                    If Not oFatherProduct Is Nothing And Not sPath Is Nothing Then
                        sPath = oSubproduct.Name + "." + sPath
                    End If

                End If
            Next
        Else
            If Not sPath Is Nothing Then
                sPath = oFatherProduct.Name + "." + sPath
            End If
        End If

        On Error GoTo 0

    End Sub

    Public Sub FindFatherProductInPPRDocument(ByVal oPPRProduct As PPR.PPRDocument, ByVal oPartDoc As PartDocument, ByRef oFatherProduct As Product, Optional ByRef sPath As Object = Nothing)
        Dim ii As Integer

        For ii = 1 To oPPRProduct.Products.Count
            Call GetFatherProduct(oPPRProduct.Products.Item(ii), oPartDoc, oFatherProduct, sPath)
            If Not (oFatherProduct Is Nothing) Then
                If Not sPath Is Nothing Then
                    sPath = oPPRProduct.Products.Item(ii).Name + "." + sPath
                End If
                Exit For
            End If
        Next
    End Sub

    ' Public Sub CheckHybridShapeItem(oItem asObject, oPart as PArt, sDiameter as string,
    '                                 sGrip as string, oParameter as parameter)
    ' ------------------------------------------------------------------------------------
    '
    ' Checks whether this Item belongs to the supported objects.
    ' The examination is carried out with the internal names of the points.
    '
    ' ------------------------------------------------------------------------------------

    'Public Function CheckHybridShapeItem(oItem As Object) As Boolean

    '    CheckHybridShapeItem = False
    '    If (TypeName(oItem) = "HybridShape") Or (TypeName(oItem) = "HybridShapePointCoord") Or (TypeName(oItem) = "HybridShapePointExplicit") Or (TypeName(oItem) = "HybridShapeProject") Or (TypeName(oItem) = "HybridShapeIntersection") Then
    '        CheckHybridShapeItem = True
    '    End If

    'End Function
    Public Function CheckHybridShapeItem(oItem As Object) As Boolean

        CheckHybridShapeItem = False
        If (TypeName(oItem) = "HybridShapePointCoord") Or (TypeName(oItem) = "HybridShapePointExplicit") Or (TypeName(oItem) = "HybridShapeProject") Or (TypeName(oItem) = "HybridShapeIntersection") Then
            CheckHybridShapeItem = True
        End If

    End Function


    Public Function CheckHybridShapeItem_fordoor(oItem As Object) As Boolean

        CheckHybridShapeItem_fordoor = False
        If (TypeName(oItem) = "HybridShapePointCoord") Or (TypeName(oItem) = "HybridShapePointExplicit") Then
            CheckHybridShapeItem_fordoor = True
        End If

    End Function





End Module

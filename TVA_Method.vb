

Option Explicit On
Imports ProductStructureTypeLib
Imports MECMOD
'Imports PARTITF
Imports HybridShapeTypeLib
Imports KnowledgewareTypeLib
Imports INFITF
Imports GoumangToolKit
Imports System.Text

Public Class TVA_Method
    Private CATIA As Application
    Public Part As Part
    Private PointsFatherProduct As Product
    Public TopProduct As Product
    Private points As CENPoints
    Public ifvec As Boolean = False
    Public coordswitch As Boolean = False
    Public PilotHoles As HybridBody = Nothing
    Private tree As processTreeList
    Dim hbtree As List(Of HybridBody)
    Public filename As String = ""
    Public pointGeo As String = "Pilot_Holes"
    Dim Fst_List As List(Of String) = Nothing
    Public infoBag As Dictionary(Of String, String)



    Public Property FstList As List(Of String)

        Get
            If Fst_List Is Nothing Then

                Fst_List = AutorivetDB.allfast_list

            End If
            Return Fst_List
        End Get
        Set(value As List(Of String))
            Fst_List = value
        End Set
    End Property

    Public Property bugList As Dictionary(Of String, HybridShape)

        Get
            Dim f As MainTools = localMethod.get_Form("MainTools")
            If f Is Nothing Then
                f = New MainTools()


            End If


            Return f.showPoints

        End Get
        Set(value As Dictionary(Of String, HybridShape))

            Dim f As MainTools = localMethod.get_Form("MainTools")
            If f Is Nothing Then
                f = New MainTools()


            End If

            Dim aa = f.showPoints
            For Each pp In value

                aa.Add(pp.Key, pp.Value)
            Next
            f.showPoints = aa



        End Set
    End Property



    'Dim Points_ As New CENPoints()

    Public Sub New(ByRef part1 As Part)
        'coordswitch = False
        CATIA = part1.Application
        Part = part1
        filename = part1.Parent.name
        ' PointsFatherProduct = fatherProduct



    End Sub

    Public Sub New(ByRef part1 As Part, fatherProduct As Product)
        coordswitch = True
        CATIA = part1.Application
        Part = part1

        PointsFatherProduct = fatherProduct
        filename = part1.Parent.name


    End Sub



    Public Sub New(filepath As String, Optional ProductPath As String = "")


        Dim partDocument1 = openPartDoc(filepath)
        Part = partDocument1.Part
        CATIA = Part.Application

        Dim documents1 As Documents
        documents1 = CATIA.Documents
        If ProductPath <> "" Then

            '如果存在组件，则一定需要转换坐标
            coordswitch = True
            Dim productDocument As ProductDocument = Nothing
            Dim newfile = True
            For Each dd As Document In documents1

                If dd.FullName.ToUpper() = ProductPath.ToUpper() Then
                    productDocument = dd
                    ' newfile = False
                    newfile = False

                    Exit For

                End If


            Next

            If newfile Then
                Try

                    productDocument = documents1.Open(ProductPath)
                    '  partDocument1 = documents1.Open(filepath)


                Catch ex As Exception
                    ' catcherr = 1
                    ' partDocument1 = documents1.GetItem(filepath)
                End Try
            End If
            productDocument.Activate()
            Dim oProduct = productDocument.Product
            GetFatherProduct(oProduct, partDocument1, PointsFatherProduct)




        End If






        filename = partDocument1.Name


    End Sub
#Region "Static Method"
    Public Shared Sub setVis(ifvis As Boolean)

        loadCATIA().Visible = ifvis

    End Sub

    Public Shared Function ifGeoVis(geo As Object) As Boolean



        Dim State
        Dim selection1 As Selection
        selection1 = geo.Application.ActiveDocument.Selection
        selection1.Clear()
        selection1.Add(geo)
        Dim oVisProps = selection1.VisProperties
        oVisProps.GetShow(State)
        If State <> 1 Then
            Return True
        Else
            Return False

        End If

    End Function




    Public Shared Function loadCATIA() As Application
        Dim CATIA = Nothing
        Try
            CATIA = GetObject(, "CATIA.Application")

        Catch ex As Exception

            Try
                CATIA = GetObject(, "DELMIA.Application")
            Catch ex2 As Exception


                Try
                    CATIA = CreateObject("CATIA.Application")
                    ' CATIA.Visible = True
                Catch ex3 As Exception
                    MsgBox("you have not installed CATIA/DELMIA,so you can't operate TVA by the tools! ")

                End Try

            End Try

            ' MsgBox("请打开catia，贱人")
        End Try

        Return CATIA
    End Function
    Public Shared Function openPartDoc(filepath As String) As PartDocument

        Dim CATIA = loadCATIA()


        Dim documents1 As Documents
        documents1 = CATIA.Documents
        Dim partDocument1 As PartDocument = Nothing
        'On Error Resume Next
        Dim catcherr As Integer
        catcherr = 0

        Dim newfile As Boolean = True
        For Each dd As Document In documents1

            If dd.FullName.ToUpper() = filepath.ToUpper() Then
                partDocument1 = dd
                newfile = False

                Exit For
            End If


        Next



        If newfile Then
            Try

                partDocument1 = documents1.Open(filepath)

                Threading.Thread.Sleep(1000)

            Catch ex As Exception

                MsgBox("打开TVA出错" + ex.Message)
                Return Nothing
                ' partDocument1 = documents1.GetItem(filepath)
            End Try
        End If

        Return partDocument1


    End Function

#End Region
#Region "Property"



    Public ReadOnly Property TVAPoints As CENPoints


        Get

            If points Is Nothing Or (Not ifvec) Then
                ifvec = True
                points = TreeListtoPoints(TVATreeList)
                Return points
            Else

                Return points

            End If

        End Get


    End Property
    Public ReadOnly Property TVAPointsnoVic As CENPoints


        Get

            If points Is Nothing Then
                ifvec = False
                points = TreeListtoPoints(TVATreeList)

                Return points
            Else

                Return points

            End If

        End Get

    End Property

    Public Property TVATreeList() As processTreeList


        Get

            If tree Is Nothing Then

                tree = New processTreeList(pilot_geoset())
                Return tree
            Else

                Return tree

            End If

        End Get
        Set(value As processTreeList)
            tree = value
        End Set
    End Property
    Public Property TVAVisTreeList As processTreeList


        Get




            Return treeList_obj(True)


        End Get
        Set(value As processTreeList)
            tree = value
        End Set
    End Property

#End Region


#Region "ProcessTree_to_Points"



    Public Function TreeListtoPoints(tmptreelist As processTreeList) As CENPoints

        Dim allpoints As New CENPoints
        allpoints.CATIA = CATIA
        For Each tmppt As processTree In tmptreelist.treeList.Values

            allpoints.Merge(TreetoPoints(tmppt))

        Next

        Return allpoints


    End Function
    Public Function TreeListtoPointsList(tmptreelist As processTreeList) As List(Of CENPoints)
        '每个points都绑定了相应的几何图形集
        Dim allpoints As New List(Of CENPoints)

        For Each tmppt As processTree In tmptreelist.treeList.Values

            allpoints.AddRange(TreetoPointsList(tmppt))

        Next

        Return allpoints
    End Function
    Public Function TreetoPoints(tmptree As processTree) As CENPoints

        Dim onetypeall As New CENPoints

        For Each kkk In tmptree.fastenertree


            onetypeall.Merge(BranchtoPoints(kkk, tmptree.fasternername, tmptree.framename))



        Next

        Return onetypeall

    End Function





    Public Function TreetoPointsList(tmptree As processTree) As List(Of CENPoints)

        Dim ptslist As New List(Of CENPoints)
        '返回带有几何图形集关系的points list
        For Each kkk In tmptree.fastenertree
            'Dim tmppoints As CENPoints
            Dim processinfo As String
            processinfo = kkk.Key
            Dim tmptreestr As String()
            tmptreestr = Strings.Split(processinfo, " - ")
            'processinfo.Split(" - ")
            Dim opgeoset As HybridBody
            opgeoset = tmptree.getnextGeo(tmptree.fastgeoset, tmptree.fasternername + " - " + tmptreestr(0))
            If (tmptreestr.Count > 1) Then



                opgeoset = tmptree.getnextGeo(opgeoset, tmptreestr(1) + " - " + tmptreestr(2))



            End If
            '  If Not opgeoset Is Nothing Then
            ptslist.Add(BranchtoPoints(kkk, tmptree.fasternername, tmptree.framename, opgeoset))
            '   Else

            '    End If


        Next
        Return ptslist

    End Function
    Public Function BranchtoPoints(ByRef kkk As KeyValuePair(Of String, List(Of HybridShape)), fasternername As String, framename As String, Optional ByRef opgeoset As HybridBody = Nothing) As CENPoints
        Dim onetypepoints As New CENPoints()
        onetypepoints.CATIA = CATIA
        Dim tmppoints = New List(Of HybridShape)()
        Dim tmplines = New List(Of HybridShape)()
        Dim oplist As List(Of HybridShape)
        Dim bugpoints As New Dictionary(Of String, HybridShape)
        ' Dim PointsFatherProduct As Product
        '找到product文档的根product
        'part与partDocument为父子关系，product之间为父子关系，所以只能通过文件路径匹配进行配对
        Dim part1 As Part
        part1 = Part
        'Dim aDoc
        'aDoc = CATIA.ActiveDocument.Product
        'TVA_Method.GetFatherProduct(aDoc, part1.Parent, PointsFatherProduct)
        ' Dim tvamodel As New TVA_Method(bindingpart, True)
        ' PointsFatherProduct = tvamodel.PointsFatherProduct

        oplist = kkk.Value
        For Each hs As HybridShape In oplist
            '判断点
            If hs.Name = "test5" Then

                Console.Write("test5")
            End If


            Dim shapesort = TVA_Method.CheckHybridShapeItem(hs)
            If shapesort Is Nothing Then
                'Just delete it  in fix process
                onetypepoints.wrongpoints.Add(hs)
            Else
                If shapesort Then
                    tmppoints.Add(hs)
                Else
                    tmplines.Add(hs)
                End If


            End If

        Next
        If (ifvec = True) Then
            Dim i As Integer = 0
            Do While (i < tmppoints.Count)
                Dim success As Boolean = False
                Dim pp = tmppoints.ElementAt(i)

                'Get the SPAWorkbench from the measurement
                Dim SPAWorkb As Workbench
                    Dim Measurement



                SPAWorkb = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")

                    'Get the measurement for the point
                    Measurement = SPAWorkb.GetMeasurable(pp)
                    Dim j As Integer = 0
                Do While (j < tmplines.Count)

                    Dim ppk = tmplines.ElementAt(j)


                    'Get the SPAWorkbench from the measurement
                    SPAWorkb = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")
                    Dim reference5 As Reference
                    reference5 = part1.CreateReferenceFromObject(ppk)
                    Dim MinimumDistance2 As Double
                    MinimumDistance2 = Measurement.GetMinimumDistance(reference5)

                    'Now get the XYZ of the point
                    '允许录入点线不重合的法向
                    If MinimumDistance2 < 1 Then
                        '当part存在时，将会录入Product坐标系的skin part坐标
                        If coordswitch Then
                            onetypepoints.Add(LeaftoPoint(pp, ppk, fasternername, kkk.Key, framename, pp.Name))
                        Else
                            onetypepoints.Add(LeaftoPoint(pp, fasternername, kkk.Key, framename, pp.Name, ppk))
                        End If
                        '找到后，把该项移除
                        tmplines.Remove(ppk)
                        'tmppoints.Remove(pp)
                        success = True
                        Exit Do
                    Else

                        j = j + 1
                    End If

                Loop
                If Not success Then
                    Dim tpt = LeaftoPoint(pp, "", "", "", "")
                    If onetypepoints.Contains(tpt) Then
                        onetypepoints.dupli.Add(tpt)
                    Else
                        bugpoints.Add(i.ToString() + "_singlepoints_" + pp.Name, pp)

                    End If

                End If


                i = i + 1

            Loop
            'Add single points and lines to specified lists
            onetypepoints.singleline.AddRange(tmplines)

        Else

            If coordswitch Then
                For Each pp As HybridShape In tmppoints
                    onetypepoints.Add(LeaftoPoint(pp, Nothing, fasternername, kkk.Key, framename, pp.Name))
                Next
            Else
                For Each pp As HybridShape In tmppoints
                    onetypepoints.Add(LeaftoPoint(pp, fasternername, kkk.Key, framename, pp.Name, Nothing))
                Next

            End If

            '   onetypepoints.Add(LeaftoPoint(pp, fasternername, kkk.Key, framename, pp.Name))




        End If

        onetypepoints.hb = opgeoset


        If (bugpoints.Count > 0) Then
            bugList = bugpoints
        End If


        Return onetypepoints


    End Function
    Public Function LeaftoPoint(ByRef PointObj As HybridShape, ByRef vector As HybridShape, fastenername As String, processtype As String, framename As String, pointname As String) As CENPoint
        '需要坐标转化的情况 法向可有可无
        ' Dim CATIA = PointObj.Application
        Dim MyPoint As CENPoint
        MyPoint = New CENPoint

        'Get Location of Point
        Dim SPAWorkb As Workbench
        Dim Measurement

        Dim Coords(2) As Object

        'Get the SPAWorkbench from the measurement
        SPAWorkb = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")
        '    MsgBox(CATIA.ActiveDocument.Name)
        '    Dim partDoc = CATIA.Documents.Item(filename)
        '  partDoc.Activate()
        '  MsgBox(CATIA.ActiveDocument.Name)
        '  SPAWorkb = partDoc.GetWorkbench("SPAWorkbench")
        'Get the measurement for the point
        Measurement = SPAWorkb.GetMeasurable(PointObj)

        'Get the coordinates (Part based) from this point
        Call Measurement.GetPoint(Coords)
        With MyPoint

            .uuidP = Math.Round(Coords(0), 2).ToString() + "_" + Math.Round(Coords(1), 2).ToString() + "_" + Math.Round(Coords(2), 2).ToString()


        End With
        'If coordswitch Then


        Dim ProductTransformation As clsMathTransformation
        Dim PointTransformation As New clsMathVector
        If Not TopProduct Is Nothing Then
            ProductTransformation = mdlProductTransformation.GetAbsPosition(TopProduct, PointsFatherProduct)
        Else
            ProductTransformation = mdlProductTransformation.GetAbsPosition(PointsFatherProduct)

        End If

        Call PointTransformation.SetByArrayV(Coords)
        PointTransformation = ProductTransformation.MulVec(PointTransformation)
        MyPoint.update("", framename, processtype, PointObj, vector, PointTransformation.cx, PointTransformation.cy, PointTransformation.cz, fastenername, pointname)

        Return MyPoint


    End Function

    Public Shared Function LeaftoPoint(X As String, Y As String, Z As String, fastenername As String, processtype As String, framename As String, pointname As String, Optional uuid As String = "", Optional pfname As String = "") As CENPoint
        '直接赋值
        Dim MyPoint As CENPoint
        MyPoint = New CENPoint

        'Get Location of Point


        ' MyPoint.update("", framename, processtype, Nothing, Nothing, PointTransformation.cx, PointTransformation.cy, PointTransformation.cz, fastenername)
        '  MyPoint.update("", framename, processtype, Nothing, Nothing, Coords(0), Coords(1), Coords(2), fastenername, pointname)
        'Add data to a collection
        With MyPoint

            .X = X
            .Y = Y
            .Z = Z
            .FastenerName = fastenername
            .Frame = framename
            .Diam = processtype
            .PointName = pointname
            .PFname = pfname
            .uuid = uuid
        End With
        Return MyPoint


    End Function
    Public Shared Function LeaftoPoint(uuid As String, Optional processtype As String = "", Optional framename As String = "", Optional fastenername As String = "", Optional pfname As String = "") As CENPoint
        '直接赋值
        Dim MyPoint As CENPoint
        MyPoint = New CENPoint

        'Get Location of Point


        ' MyPoint.update("", framename, processtype, Nothing, Nothing, PointTransformation.cx, PointTransformation.cy, PointTransformation.cz, fastenername)
        '  MyPoint.update("", framename, processtype, Nothing, Nothing, Coords(0), Coords(1), Coords(2), fastenername, pointname)
        'Add data to a collection
        With MyPoint

            .FastenerName = fastenername
            .Frame = framename
            .Diam = processtype

            .PFname = pfname
            .uuid = uuid
        End With
        Return MyPoint


    End Function
    Public Shared Function LeaftoPoint(ByRef PointObj As HybridShape, fastenername As String, processtype As String, framename As String, pointname As String, Optional ByRef vector As HybridShape = Nothing) As CENPoint
        '不考虑坐标转化的情况 法向可有可无
        Dim MyPoint As CENPoint
        MyPoint = New CENPoint

        'Get Location of Point
        Dim SPAWorkb As Workbench
        Dim Measurement
        '  Dim ProductTransformation As clsMathTransformation
        Dim PointTransformation As New clsMathVector
        Dim Coords(2) As Object
        Dim CATIA = PointObj.Application
        'Get the SPAWorkbench from the measurement
        SPAWorkb = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")

        'Get the measurement for the point
        Measurement = SPAWorkb.GetMeasurable(PointObj)

        'Get the coordinates (Part based) from this point
        Call Measurement.GetPoint(Coords)



        ' MyPoint.update("", framename, processtype, Nothing, Nothing, PointTransformation.cx, PointTransformation.cy, PointTransformation.cz, fastenername)
        MyPoint.update("", framename, processtype, PointObj, vector, Coords(0), Coords(1), Coords(2), fastenername, pointname)
        'Add data to a collection
        ' Points_.Add(MyPoint)

        Return MyPoint


    End Function














#End Region















    Public Function noTreePoints(Optional ByRef geoset As HybridBody = Nothing) As CENPoints
        If geoset Is Nothing Then

            geoset = pilot_geoset()
        End If
        '需要坐标转化

        points = New CENPoints()
        Dim tmppoints = New CENPoints()
        Dim tmplines = New List(Of HybridShape)
        Dim walkHB As Action(Of HybridBody)
        walkHB = Sub(hbb As HybridBody)
                     For Each bb In hbb.HybridBodies
                         walkHB(bb)
                     Next
                     'After lookthrough every sub-hybody,try to pair the non-paired points and lines
                     If hbb.HybridShapes.Count > 1 Then
                         Dim State
                         Dim myselection
                         Dim oVisProps
                         myselection = CATIA.ActiveDocument.Selection
                         Dim tmppoints2 = New CENPoints()
                         Dim tmplines2 = New List(Of HybridShape)
                         For Each hs As HybridShape In hbb.HybridShapes

                             myselection.Clear()
                             myselection.Add(hs)
                             oVisProps = CATIA.ActiveDocument.Selection.VisProperties
                             oVisProps.GetShow(State)
                             'If State = 1 Then
                             '    State1 = "Hidden"
                             'Else
                             '    State1 = "Shown"
                             'End If

                             If State <> 1 Then

                                 '判断点
                                 If CheckHybridShapeItem(hs) Then
                                     'For using cenpoints to filter duplicate points or lines
                                     tmppoints2.Add(LeaftoPoint(hs, "", "", "", ""))
                                 Else
                                     Dim hstype = TypeName(hs)
                                     If hstype.Contains("HybridShapeLine") Then
                                         tmplines2.Add(hs)
                                     End If
                                 End If
                             End If

                         Next

                         Dim todelpts As New List(Of HybridShape)()
                         Dim pairedDic = Pair_pt(tmppoints2, tmplines2, hbb.Name)
                         'Remove the points that have been paired
                         For Each pp In pairedDic

                             tmppoints2.Remove(pp.Key)
                             tmplines2.Remove(pp.Value)

                         Next

                         'For non-paired points and lines,add them to global var
                         tmppoints.Merge(tmppoints2)
                         tmplines.AddRange(tmplines2)


                         'MsgBox(hbb.Name)
                         'Dim hbc = hbb.HybridBodies.Count
                     End If








                 End Sub





        walkHB(geoset)
        'Check all non paired points and lines

        Pair_pt(tmppoints, tmplines, "SB_ENG")
                                                  Return points



    End Function

    Public Function Pair_pt(ptlist As CENPoints, linelist As List(Of HybridShape), frameName As String) As Dictionary(Of String, HybridShape)

        Dim paired As New Dictionary(Of String, HybridShape)()
        Dim SPAWorkb As Workbench

        Dim Measurement
        SPAWorkb = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")
        For Each ppd In ptlist.Points_
            'Get the SPAWorkbench from the measurement
            Dim pp = ppd.Value.MyPoint
            Dim Coords(2) As Object





            'Get the measurement for the point
            Measurement = SPAWorkb.GetMeasurable(pp)

            'Get the coordinates (Part based) from this point
            Call Measurement.GetPoint(Coords)
            Dim i As Integer = 0
            Do While i < linelist.Count

                Dim ppk = linelist.ElementAt(i)
                'Get the SPAWorkbench from the measurement
                SPAWorkb = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")
                Dim reference5 As INFITF.Reference
                reference5 = Part.CreateReferenceFromObject(ppk)
                Dim MinimumDistance2 As Double
                MinimumDistance2 = Measurement.GetMinimumDistance(reference5)

                'Now get the XYZ of the point
                If MinimumDistance2 < 0.1 Then
                    paired.Add(ppd.Key, ppk)
                    linelist.Remove(ppk)
                    points.Add(LeaftoPoint(pp, ppk, "", "TEMP", frameName, pp.name))
                    'points.Add(LeaftoPoint(pp, ppk, "", "TEMP", hbb.Name, pp.Name))
                    ''找到后，把该项移除
                    Exit Do

                    'todelpts.Add(pp)
                    'tmplines2.Remove(ppk)
                Else

                    i = i + 1

                End If



            Loop
            'For Each ppk In linelist


            '    'Get the SPAWorkbench from the measurement
            '    SPAWorkb = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")
            '    Dim reference5 As INFITF.Reference
            '    reference5 = Part.CreateReferenceFromObject(ppk)
            '    Dim MinimumDistance2 As Double
            '    MinimumDistance2 = Measurement.GetMinimumDistance(reference5)

            '    'Now get the XYZ of the point
            '    If MinimumDistance2 < 0.1 Then
            '        paired.Add(ppd.Key, ppk)
            '        points.Add(LeaftoPoint(pp, ppk, "", "TEMP", frameName, pp.name))
            '        'points.Add(LeaftoPoint(pp, ppk, "", "TEMP", hbb.Name, pp.Name))
            '        ''找到后，把该项移除


            '        'todelpts.Add(pp)
            '        'tmplines2.Remove(ppk)
            '        Exit For

            '    End If

            'Next

        Next

        Return paired
    End Function





    '保存文件

    Public Sub save()
        If filename <> "" Then
            Dim doc = CATIA.Documents.Item(filename)
            doc.Save()

            doc.Close()
        End If

        'CATIA.ActiveDocument.Save()
        'CATIA.ActiveDocument.Close()


    End Sub



    Public Function treeList_obj(Optional iffilvis As Boolean = False) As processTreeList

        Dim aa As New processTreeList
        aa.FastList = FstList
        aa.inputShape(pilot_geoset(), iffilvis)

        Return aa

    End Function


    Public Function pilot_geoset() As HybridBody


        If PilotHoles Is Nothing Then
            '   If (PilotHoles Is Nothing) Or (InStr(LCase(PilotHoles.Name), LCase(geoname)) = 0) Then
            Dim i
            For i = 1 To Part.HybridBodies.Count
                '找到pilot_holes目录
                If InStr(LCase(Part.HybridBodies.Item(i).Name), LCase(pointGeo)) <> 0 Then

                    PilotHoles = Part.HybridBodies.Item(i)
                    '  PilotHoles.Name = "Pilot_Holes"
                    Exit For
                End If

            Next

            If PilotHoles Is Nothing Then
                PilotHoles = Part.HybridBodies.Add()

                PilotHoles.Name = "Pilot_Holes"

            End If


        End If

        Return PilotHoles



    End Function
    Public Function out_surf(Optional geoname As String = "NC_Geometry") As List(Of HybridShape)

        Dim aa As New List(Of HybridShape)
        '     pointGeo = geoname
        ' If OuterSurface Is Nothing Then
        Dim NCGeo = Part.HybridBodies.Item(geoname)
        For Each kk In NCGeo.HybridShapes

            If TypeName(kk) = "HybridShapeExtract" Or TypeName(kk) = "HybridShapeAssemble" Then
                aa.Add(kk)

            End If
        Next



        ' End If

        Return aa

    End Function

    Public Function fix_all(Optional closeAfter As Boolean = True) As processStatic

        ifvec = True
        coordswitch = False

        Dim aa = New processStatic()
        aa.Add("Fix outlines/dupli", filename)
        aa.Add("--------------", "------------------")
        Dim aapt As List(Of CENPoints) = TreeListtoPointsList(TVATreeList)
        '删除重复点线不会解决有线无点和有点无线的情况，即是这些孤立的点线可能是重复的
        For Each ppp In aapt

            aa.Add(ppp.fix_outlines(filename, out_surf()))
            aa.Add(ppp.del_dupli(filename))
        Next
        If closeAfter Then
            save()
        End If

        Return aa
    End Function


    Public Sub rebuild(prodname As String)
        ' ifvec = True

        Dim dbpoints = New CENPoints
        dbpoints.importfromdb(prodname, True)

        For Each pp In TVAPoints.Points_.Values
            pp.Frame = dbpoints.Item(pp.uuid).Frame

        Next
        Dim tree = TVAPoints.toProcessTree
        tree.output_topart(filename, filename, pilot_geoset)
        tree.del(filename)

        save()

    End Sub

    'Public Sub fix_out(ss As HybridShapeExtract)

    '    ppp.fix_outlines(filename, out_surf())
    'End Sub

#Region "Check TVA"





    Public Function CheckTVA(Optional ByVal color As Boolean = True, Optional ByVal database As Boolean = True, Optional FastList As List(Of String) = Nothing, Optional ifreport As Boolean = True) As processStatic
        Dim wrongstatistic = New processStatic()
        Dim fastenerqty = New processStatic()
        Dim processtype = New processStatic()
        Dim checkTVApoint = New processStatic()
        Dim processtree = New processTreeBase()
        Dim bugcontainer = New processStatic()

        infoBag = New Dictionary(Of String, String)()
        'Use delegate to recurse

        'Dim MyGeoSet = pilot_geoset()
        Dim CheckRecursion As Action(Of HybridBody)

        CheckRecursion = Sub(MyGeoSet As HybridBody)

                             ' CATIA = MyGeoSet.Application


                             '2015.5.12进行重构，与界面元素相分离
                             Dim buglocation As String

                             Dim MySourceGeoSet As HybridBody
                             Dim shapecount As Integer
                             Dim tempstring As String

                             If TVA_Method.ifFastener(MyGeoSet.Name) Then

                                 '紧固件列表
                                 fastenerqty.Add(0, MyGeoSet.Name)



                                 Dim i
                                 '开始遍历紧固件的下层几何图形集
                                 For i = 1 To MyGeoSet.HybridBodies.Count


                                     Dim MyNewGeoSet As HybridBody

                                     MyNewGeoSet = MyGeoSet.HybridBodies.Item(i)

                                     MySourceGeoSet = MyNewGeoSet

                                     '计算该几何图形集中点的数量
                                     shapecount = Fix(MySourceGeoSet.HybridShapes.Count / 2)



                                     '判断几何图形集下是否有奇数个shape ,若有，则报错
                                     If (MySourceGeoSet.HybridShapes.Count / 2 - shapecount) <> 0 Then
                                         shapecount = shapecount + 1
                                         buglocation = MySourceGeoSet.Parent.Parent.name + " - " + MySourceGeoSet.Name

                                         wrongstatistic.Add(shapecount, "error in qty of point and line" + " - " + buglocation)

                                         '以下过程找出孤立的点，并加入到bug列表中
                                         bugcontainer.Add(Checkvector(MySourceGeoSet, buglocation))
                                     End If

                                     '  Checkdupli MySourceGeoSet
                                     If Strings.InStr(MyNewGeoSet.Name, "TEMP") And MyNewGeoSet.HybridShapes.Count <> 0 Then

                                         wrongstatistic.Add(shapecount, "Something in Temp GeoSet" + " - " + MyGeoSet.Parent.name + " - " + MyNewGeoSet.Name)

                                     End If




                                     If Strings.InStr(MyNewGeoSet.Name, "RESYNCING") Then

                                         If MyNewGeoSet.HybridShapes.Count <> 0 Then
                                             wrongstatistic.Add(shapecount, "Need select Target Type" + " - " + MyGeoSet.Parent.name + " - " + MyNewGeoSet.Name)

                                         End If


                                         Dim ii
                                         For ii = 0 To 2



                                             'Set MySourceGeoSet = MyNewGeoSet.HybridBodies.GetItem(UserForm1.ComboBox4.Items.Item(ii))
                                             If (ifGeoExist(MyNewGeoSet, processtree.sectree(ii)) = False) Then

                                                 MySourceGeoSet = MyNewGeoSet.HybridBodies.Add()
                                                 MySourceGeoSet.Name = processtree.sectree(ii)
                                             Else
                                                 MySourceGeoSet = MyNewGeoSet.HybridBodies.GetItem(processtree.sectree(ii))

                                             End If


                                             If color Then

                                                 Dim selection1 As Object


                                                 selection1 = CATIA.ActiveDocument.Selection
                                                 selection1.Clear()
                                                 Dim zz
                                                 For zz = 1 To MySourceGeoSet.HybridShapes.Count
                                                     selection1.Add(MySourceGeoSet.HybridShapes.Item(zz))

                                                 Next


                                                 If Strings.InStr(MyNewGeoSet.Name, "RESYNCING ONLY") Then

                                                     '红色
                                                     selection1.VisProperties.SetRealColor(255, 0, 0, 1)
                                                 Else
                                                     '绿色
                                                     selection1.VisProperties.SetRealColor(0, 128, 64, 1)

                                                 End If


                                                 If Strings.InStr(MySourceGeoSet.Name, "Final") Then



                                                     '  selection1.VisProperties.SetSymbolType(6)
                                                     If Strings.InStr(MyGeoSet.Name, "B020600") Then
                                                         '设为大圆
                                                         selection1.VisProperties.SetSymbolType(3)
                                                     Else

                                                         '设为十字
                                                         selection1.VisProperties.SetSymbolType(2)
                                                     End If
                                                 Else
                                                     If Strings.InStr(MySourceGeoSet.Name, "Fast Tack") Then
                                                         '星号
                                                         selection1.VisProperties.SetSymbolType(7)

                                                     Else
                                                         '同心圆

                                                         If Strings.InStr(MyNewGeoSet.Name, "DRILL") Then

                                                             selection1.VisProperties.SetSymbolType(5)
                                                         Else

                                                             selection1.VisProperties.SetSymbolType(4)

                                                         End If


                                                     End If

                                                 End If

                                             End If


                                             '检查校准的几何图形集



                                             'If database Then


                                             '    ' buglocation = MyGeoSet.Parent.name + " - " + MyNewGeoSet.Name + " - " + MySourceGeoSet.Name
                                             '    updatedatabase(MySourceGeoSet, MyGeoSet.Parent.name, MyGeoSet.Name, Strings.Split(MyNewGeoSet.Name, " - ")(1) + " - " + MySourceGeoSet.Name)
                                             'End If



                                             shapecount = Fix(MySourceGeoSet.HybridShapes.Count / 2)
                                             'MsgBox MySourceGeoSet.name



                                             If (MySourceGeoSet.HybridShapes.Count / 2 - shapecount) <> 0 Then
                                                 shapecount = shapecount + 1
                                                 buglocation = MyGeoSet.Parent.name + " - " + MyNewGeoSet.Name + " - " + MySourceGeoSet.Name


                                                 wrongstatistic.Add(shapecount, "error in qty of point and line" + " - " + buglocation)


                                                 bugcontainer.Add(Checkvector(MySourceGeoSet, buglocation))

                                             End If

                                             tempstring = Strings.Trim(Strings.Split(MyNewGeoSet.Name, " - ")(1)) + "-" + Strings.Trim(Strings.Split(MySourceGeoSet.Name, " - ")(1))

                                             checkTVApoint.Add(shapecount, MyNewGeoSet.Name + "-" + Strings.Trim(Strings.Split(MySourceGeoSet.Name, " - ")(1)))

                                             '统计各个紧固件数量Fastener_Qty
                                             fastenerqty.Add(shapecount, MyGeoSet.Name)
                                             processtype.Add(shapecount, tempstring)

                                         Next

                                     Else

                                         '不为resync
                                         '检查非校准的几何图形集
                                         If (MyNewGeoSet.Name.Contains("AUTOMATED FASTENING") Or MyNewGeoSet.Name.Contains("TEMP")) Then


                                             If color Then

                                                 Dim selection1 As Object

                                                 selection1 = CATIA.ActiveDocument.Selection
                                                 selection1.Clear()
                                                 Dim zzz
                                                 For zzz = 1 To MyNewGeoSet.HybridShapes.Count
                                                     selection1.Add(MyNewGeoSet.HybridShapes.Item(zzz))

                                                 Next



                                                 If Strings.InStr(MyNewGeoSet.Name, "BY") Then


                                                     If Strings.InStr(MyGeoSet.Name, "5-") Then

                                                         If Strings.InStr(MyGeoSet.Name, "5020AD") Or Strings.InStr(MyGeoSet.Name, "6002AG") Then
                                                             '蓝色
                                                             selection1.VisProperties.SetRealColor(0, 0, 255, 1)
                                                         Else
                                                             selection1.VisProperties.SetRealColor(0, 255, 255, 1)
                                                         End If

                                                     Else
                                                         '棕色

                                                         If Strings.InStr(MyGeoSet.Name, "5020AD") Or Strings.InStr(MyGeoSet.Name, "6002AG") Then

                                                             selection1.VisProperties.SetRealColor(128, 64, 64, 1)
                                                         Else
                                                             selection1.VisProperties.SetRealColor(255, 128, 0, 1)
                                                         End If
                                                     End If



                                                     If Strings.InStr(MyNewGeoSet.Name, "INSTALL") Then
                                                         If Strings.InStr(MyGeoSet.Name, "B020600") Then
                                                             '设为大圆
                                                             selection1.VisProperties.SetSymbolType(3)
                                                         Else

                                                             '设为十字
                                                             selection1.VisProperties.SetSymbolType(2)
                                                         End If


                                                         'selection1 = Nothing
                                                     Else

                                                         If Strings.InStr(MyNewGeoSet.Name, "DRILL") Then

                                                             '设为实心圆
                                                             selection1.VisProperties.SetSymbolType(5)

                                                         End If

                                                     End If
                                                 Else


                                                     If Strings.InStr(MyNewGeoSet.Name, "AFTER") Or Strings.InStr(MyNewGeoSet.Name, "TEMP") Then
                                                         '符号为X
                                                         'after 点为黑色
                                                         selection1.VisProperties.SetRealColor(0, 0, 0, 1)
                                                         selection1.VisProperties.SetSymbolType(1)
                                                     Else
                                                         selection1.VisProperties.SetRealColor(0, 0, 0, 1)
                                                         'If Strings.InStr(MyGeoSet.Name, "5-") Then
                                                         '    '都改为深色

                                                         '    If Strings.InStr(MyGeoSet.Name, "5020AD") Or Strings.InStr(MyGeoSet.Name, "6002AG") Then
                                                         '        '紫色
                                                         '        selection1.VisProperties.SetRealColor(255, 0, 255, 1)
                                                         '    Else
                                                         '        '深色
                                                         '        selection1.VisProperties.SetRealColor(255, 165, 0, 1)
                                                         '    End If

                                                         'Else
                                                         '    '棕色

                                                         '    If Strings.InStr(MyGeoSet.Name, "5020AD") Or Strings.InStr(MyGeoSet.Name, "6002AG") Then

                                                         '        selection1.VisProperties.SetRealColor(64, 32, 32, 1)
                                                         '    Else
                                                         '        'Make protruding head 6# darker
                                                         '        selection1.VisProperties.SetRealColor(64, 32, 0, 1)
                                                         '    End If
                                                         'End If

                                                         If Strings.InStr(MyNewGeoSet.Name, "DRILL") Then
                                                             selection1.VisProperties.SetSymbolType(4)
                                                         Else

                                                             If Strings.InStr(MyGeoSet.Name, "B020600") Then
                                                                 '设为大圆
                                                                 selection1.VisProperties.SetSymbolType(3)
                                                             Else

                                                                 '设为十字
                                                                 selection1.VisProperties.SetSymbolType(2)
                                                             End If

                                                         End If

                                                     End If




                                                 End If


                                             End If


                                             If Strings.InStr(MyNewGeoSet.Name, " - ") Then

                                                 tempstring = Strings.Trim(Strings.Split(MySourceGeoSet.Name, " - ")(1))
                                                 checkTVApoint.Add(shapecount, MySourceGeoSet.Name)

                                                 fastenerqty.Add(shapecount, MyGeoSet.Name)
                                                 processtype.Add(shapecount, tempstring)
                                             End If
                                         Else
                                             wrongstatistic.Add(shapecount, "error geoset" + " - " + MyNewGeoSet.Name + " - ")

                                         End If
                                     End If


                                 Next

                             Else
                                 If MyGeoSet.HybridShapes.Count <> 0 Then
                                     wrongstatistic.Add(shapecount, "keng brother bug" + " - " + MyGeoSet.Parent.name + " - " + MyGeoSet.Name)
                                 End If

                                 Dim k As Integer

                                 For k = 1 To MyGeoSet.HybridBodies.Count
                                     '开始递归
                                     CheckRecursion(MyGeoSet.HybridBodies.Item(k))
                                 Next
                             End If






                             '生产表格














                         End Sub
        CheckRecursion(pilot_geoset())

        Dim ReportRecursion As Action

        ReportRecursion = Sub()

                              Dim rivottypename As String
                              Dim sum As Integer
                              Dim tempcount As Integer
                              sum = 0
                              Dim autosum As Integer
                              autosum = 0
                              Dim resynsum As Integer
                              resynsum = 0


                              Dim frconlyfinal As Integer
                              frconlyfinal = 0

                              Dim frconlyfastack As Integer
                              frconlyfastack = 0

                              Dim highlitesum As Integer
                              highlitesum = 0
                              Dim instodrill As Integer
                              instodrill = 0
                              '原始TVA就是仅仅钻孔的
                              Dim orgdrill As Integer
                              orgdrill = 0
                              '高锁钻孔校准点数
                              Dim highliteresyn As Integer
                              highliteresyn = 0

                              '校准插钉
                              Dim highliterRC As Integer
                              highliterRC = 0


                              '仅校准导孔
                              Dim FRChole As Integer
                              FRChole = 0

                              '校准导孔并安装
                              Dim RChole As Integer
                              RChole = 0


                              '手铆和机器一共需用高锁数
                              Dim highliteall As Integer
                              highliteall = 0
                              '手铆和机器下架补铆数
                              Dim afterall As Integer
                              afterall = 0
                              '高锁下架补铆数
                              Dim highafterall As Integer
                              highafterall = 0

                              '5号铆钉钻铆
                              Dim fiverivet As Integer
                              fiverivet = 0


                              '5号高锁钻铆
                              Dim fivehilite As Integer
                              fivehilite = 0

                              '5号高锁钻孔
                              Dim fivedrill As Integer
                              fivedrill = 0

                              '6号高锁钻铆
                              Dim sixhilite As Integer
                              sixhilite = 0


                              '6号高锁钻孔
                              Dim sixdrill As Integer
                              sixdrill = 0


                              Dim excelobject As Object
                              Dim wb As Object




                              Dim m As Integer
                              Dim n As Integer
                              Dim p As Integer
                              Dim fastextract As String
                              Dim processextract As String
                              Dim fastcount As Integer
                              Dim proccesscount As Integer




                              excelobject = CreateObject("excel.application") '启动Excel程序


                              excelobject.Visible = True

                              wb = excelobject.Workbooks.Add()
                              excelobject.displayalerts = False
                              fastcount = fastenerqty.count
                              proccesscount = processtype.count

                              wb.Sheets(1).Cells(1, fastcount + 2).Value = "总计"

                              wb.Sheets(1).Cells(proccesscount + 2, 1).Value = "总计"
                              '制作表头，添加统计数据
                              For p = 1 To processtype.count
                                  Dim processtypetmp As String

                                  processtypetmp = processtype.Key(p)
                                  processtypetmp = Strings.Replace(processtypetmp, " AUTOMATED FASTENING", "")


                                  wb.Sheets(1).Cells(p + 1, 1).Value = processtypetmp

                                  wb.Sheets(1).Cells(p + 1, fastcount + 2).Value = processtype.Item2(p)
                              Next

                              For n = 1 To fastcount
                                  wb.Sheets(1).Cells(1, n + 1).Value = fastenerqty.Key(n)
                                  wb.Sheets(1).Cells(proccesscount + 2, n + 1).Value = fastenerqty.Item2(n)
                              Next


                              Dim ff As Integer
                              'For ff = 0 To buglistBox.Items.Count - 1
                              '    If fastenerqty.InTheList(buglistBox.Items.Item(ff)) Then

                              '        buglistBox.SetSelected(ff, True)
                              '    Else
                              '        buglistBox.SetSelected(ff, True)
                              '    End If

                              'Next

                              wb.Sheets(1).Cells(1, 1).Value = "加工类型/紧固件"



                              For m = 1 To checkTVApoint.count
                                  rivottypename = checkTVApoint.Key(m)
                                  tempcount = checkTVApoint.Item(rivottypename)
                                  fastextract = Strings.Trim(Strings.Split(rivottypename, " - ")(0))
                                  processextract = Strings.Trim(Strings.Split(rivottypename, " - ")(1))
                                  Dim rownum As Integer
                                  Dim colnum As Integer

                                  rownum = processtype.SearchIndex(processextract) + 1
                                  colnum = fastenerqty.SearchIndex(fastextract) + 1
                                  wb.Sheets(1).Cells(rownum, colnum).Value = tempcount
                                  'MsgBox (processextract)

                                  If (InStr(rivottypename, "AFTER AUTOMATED FASTENING") Or InStr(rivottypename, "BEFORE AUTOMATED FASTENING")) Then

                                      afterall = afterall + tempcount
                                  End If



                                  If (InStr(rivottypename, "RESYNCING")) Then 'aa
                                      resynsum = resynsum + tempcount
                                      '统计高锁校准点数
                                      If (InStr(rivottypename, "B020600")) Then 'bb
                                          highliteresyn = highliteresyn + tempcount
                                          If (InStr(rivottypename, "INSTALLED")) Then
                                              highliterRC = highliterRC + tempcount
                                          End If
                                      End If 'bb

                                      If (InStr(rivottypename, "Pilot Holes")) Then 'cc

                                          If (InStr(rivottypename, "ONLY")) Then

                                              FRChole = FRChole + tempcount

                                          Else
                                              RChole = RChole + tempcount

                                          End If

                                      Else 'cc

                                          If (InStr(rivottypename, "ONLY")) Then

                                              If (InStr(rivottypename, "Final")) Then
                                                  frconlyfinal = frconlyfinal + tempcount
                                              Else
                                                  frconlyfastack = frconlyfastack + tempcount
                                              End If
                                          End If


                                      End If 'cc

                                  Else 'aa


                                      If (InStr(rivottypename, "BY AUTOMATED FASTENING")) Then 'a
                                          If (InStr(rivottypename, "INSTALLED") Or InStr(rivottypename, "DRILL")) Then 'b
                                              autosum = autosum + tempcount

                                              If tempcount < 10 Then
                                                  wb.Sheets(1).Cells(rownum, colnum).Interior.ColorIndex = 3 ' 背景的颜色为3 红色
                                              End If
                                              '如果是高锁
                                              If (InStr(rivottypename, "B020600")) Then 'c
                                                  highlitesum = highlitesum + tempcount
                                                  '统计自动钻铆安装转为仅仅钻孔的高锁
                                                  If (InStr(rivottypename, "INSTALLED")) Then 'd
                                                      'instodrill = instodrill + tempcount
                                                  Else 'd
                                                      '刚开始就仅仅是自动钻铆钻孔的
                                                      If (InStr(rivottypename, "DRILL")) Then 'e
                                                          orgdrill = orgdrill + tempcount
                                                      End If 'e


                                                  End If 'd

                                              Else 'c
                                                  fiverivet = fiverivet + tempcount
                                              End If 'c

                                          End If 'b
                                      End If 'a


                                  End If 'aa









                                  If (InStr(rivottypename, "B020600")) Then 'a
                                      highliteall = highliteall + tempcount


                                      If (InStr(rivottypename, "AFTER AUTOMATED FASTENING")) Then 'b
                                          highafterall = highafterall + tempcount

                                      Else 'b

                                          If (InStr(rivottypename, "RESYNCING") = 0) Then 'c


                                              If (InStr(rivottypename, "FASTENER INSTALLED BY")) Then 'd

                                                  If (InStr(rivottypename, "AG5-")) Then
                                                      fivehilite = fivehilite + tempcount
                                                  Else
                                                      If (InStr(rivottypename, "AG6-")) Then
                                                          sixhilite = sixhilite + tempcount
                                                      End If
                                                  End If


                                              Else 'd

                                                  If (InStr(rivottypename, "DRILL ONLY BY")) Then 'e

                                                      If (InStr(rivottypename, "AG5-")) Then
                                                          fivedrill = fivedrill + tempcount
                                                      Else
                                                          If (InStr(rivottypename, "AG6-")) Then
                                                              sixdrill = sixdrill + tempcount
                                                          End If
                                                      End If


                                                  End If 'e


                                              End If 'd


                                          End If 'c


                                      End If 'b
                                  End If 'a

                                  sum = sum + tempcount

                              Next

                              wb.Sheets(1).Cells(processtype.count + 2, fastenerqty.count + 2).Value = sum




                              m = processtype.count + 4

                              wb.Sheets(1).Cells(m, 1).Value = "仅校准(终钉):"
                              wb.Sheets(1).Cells(m, 2).Value = frconlyfinal

                              m = m + 1


                              wb.Sheets(1).Cells(m, 1).Value = "仅校准(临时紧固件):"
                              wb.Sheets(1).Cells(m, 2).Value = frconlyfastack


                              m = m + 1

                              wb.Sheets(1).Cells(m, 1).Value = "仅校准(导孔):"
                              wb.Sheets(1).Cells(m, 2).Value = FRChole



                              m = m + 1

                              wb.Sheets(1).Cells(m, 1).Value = "校准(任意)插钉:"
                              wb.Sheets(1).Cells(m, 2).Value = highliterRC
                              highlitesum = highlitesum + highliterRC
                              autosum += highliterRC

                              m = m + 1

                              wb.Sheets(1).Cells(m, 1).Value = "校准导孔安装:"
                              wb.Sheets(1).Cells(m, 2).Value = RChole
                              autosum += RChole
                              m = m + 1

                              wb.Sheets(1).Cells(m, 1).Value = "5号铆钉钻铆:"
                              wb.Sheets(1).Cells(m, 2).Value = fiverivet

                              m = m + 1

                              wb.Sheets(1).Cells(m, 1).Value = "5号高锁钻铆:"
                              wb.Sheets(1).Cells(m, 2).Value = fivehilite

                              m = m + 1

                              wb.Sheets(1).Cells(m, 1).Value = "6号高锁钻铆:"
                              wb.Sheets(1).Cells(m, 2).Value = sixhilite




                              m = m + 1

                              wb.Sheets(1).Cells(m, 1).Value = "5号高锁仅钻孔:"
                              wb.Sheets(1).Cells(m, 2).Value = fivedrill

                              m = m + 1


                              wb.Sheets(1).Cells(m, 1).Value = "6号高锁仅钻孔:"
                              wb.Sheets(1).Cells(m, 2).Value = sixdrill



                              m = m + 1
                              wb.Sheets(1).Cells(m, 1).Value = "其他点位:"
                              wb.Sheets(1).Cells(m, 2).Value = afterall



                              m = m + 2
                              wb.Sheets(1).Cells(m, 1).Value = "Wrong:"
                              m = m + 1
                              wb.Sheets(1).Cells(m, 1).Value = "WrongType"
                              wb.Sheets(1).Cells(m, 2).Value = "FrameName"
                              wb.Sheets(1).Cells(m, 3).Value = "FastenerName"
                              wb.Sheets(1).Cells(m, 4).Value = "ProccessType"
                              wb.Sheets(1).Cells(m, 5).Value = "Qty"
                              Dim wrongnum As Integer



                              For wrongnum = 1 To wrongstatistic.count
                                  m = m + 1
                                  Dim wrongstr As String
                                  wrongstr = wrongstatistic.Key(wrongnum)

                                  wb.Sheets(1).Cells(m, 1).Value = Strings.Split(wrongstr, " - ")(0)
                                  wb.Sheets(1).Cells(m, 2).Value = Strings.Split(wrongstr, " - ")(1)
                                  wb.Sheets(1).Cells(m, 3).Value = Strings.Split(wrongstr, " - ")(2)
                                  wb.Sheets(1).Cells(m, 4).Value = Strings.Split(wrongstr, " - ")(3)
                                  wb.Sheets(1).Cells(m, 5).Value = wrongstatistic.Item2(wrongnum)

                              Next

                              m = m + 2
                              wb.Sheets(1).Cells(m, 1).Value = "Statistic:"
                              m = m + 1
                              wb.Sheets(1).Cells(m, 2).Value = "Qty"
                              wb.Sheets(1).Cells(m, 3).Value = "Percentage"

                              m = m + 1
                              wb.Sheets(1).Cells(m, 1).Value = "SUM"
                              wb.Sheets(1).Cells(m, 2).Value = sum
                              infoBag.Add("SUM", sum.ToString())
                              '输出需要自动钻铆安装总数

                              m = m + 1
                              wb.Sheets(1).Cells(m, 1).Value = "BY MACHINE"
                              wb.Sheets(1).Cells(m, 2).Value = autosum
                              wb.Sheets(1).Cells(m, 3).Value = autosum / sum
                              wb.Sheets(1).Cells(m, 3).NumberFormatLocal = "0.00%"
                              infoBag.Add("BY MACHINE", autosum.ToString())
                              infoBag.Add("PERCENTAGE", autosum / sum)
                              '输出需要手动安装总数

                              m = m + 1
                              wb.Sheets(1).Cells(m, 1).Value = "BY HAND"
                              wb.Sheets(1).Cells(m, 2).Value = sum - autosum
                              wb.Sheets(1).Cells(m, 3).Value = (sum - autosum) / sum
                              wb.Sheets(1).Cells(m, 3).NumberFormatLocal = "0.00%"

                              '架下补铆安装总共

                              m = m + 1
                              wb.Sheets(1).Cells(m, 1).Value = "AFTER"
                              wb.Sheets(1).Cells(m, 2).Value = afterall
                              wb.Sheets(1).Cells(m, 3).Value = afterall / sum
                              wb.Sheets(1).Cells(m, 3).NumberFormatLocal = "0.00%"


                              '需预定位的总数

                              m = m + 1
                              wb.Sheets(1).Cells(m, 1).Value = "RESYNC"
                              wb.Sheets(1).Cells(m, 2).Value = resynsum
                              wb.Sheets(1).Cells(m, 3).Value = resynsum / sum
                              wb.Sheets(1).Cells(m, 3).NumberFormatLocal = "0.00%"
                              m = m + 1

                              '共需高锁数
                              m = m + 1
                              wb.Sheets(1).Cells(m, 1).Value = "HI-LITE"
                              wb.Sheets(1).Cells(m, 2).Value = highliteall

                              m = m + 1
                              wb.Sheets(1).Cells(m, 1).Value = "BY MACHINE"
                              wb.Sheets(1).Cells(m, 2).Value = highlitesum
                              wb.Sheets(1).Cells(m, 3).Value = highlitesum / highliteall
                              wb.Sheets(1).Cells(m, 3).NumberFormatLocal = "0.00%"


                              '手动安装总数
                              m = m + 1
                              wb.Sheets(1).Cells(m, 1).Value = "BY HAND"
                              wb.Sheets(1).Cells(m, 2).Value = highliteall - highlitesum
                              wb.Sheets(1).Cells(m, 3).Value = (highliteall - highlitesum) / highliteall
                              wb.Sheets(1).Cells(m, 3).NumberFormatLocal = "0.00%"


                              '架下补铆安装高锁

                              m = m + 1
                              wb.Sheets(1).Cells(m, 1).Value = "AFTER"
                              wb.Sheets(1).Cells(m, 2).Value = highafterall
                              wb.Sheets(1).Cells(m, 3).Value = highafterall / highliteall
                              wb.Sheets(1).Cells(m, 3).NumberFormatLocal = "0.00%"

                              '作为校准点的高锁数
                              m = m + 1
                              wb.Sheets(1).Cells(m, 1).Value = "RESYNC"
                              wb.Sheets(1).Cells(m, 2).Value = highliteresyn
                              wb.Sheets(1).Cells(m, 3).Value = highliteresyn / highliteall
                              wb.Sheets(1).Cells(m, 3).NumberFormatLocal = "0.00%"


                              '高锁自动钻铆钻孔总数
                              m = m + 1
                              wb.Sheets(1).Cells(m, 1).Value = "DRILL ONLY"
                              wb.Sheets(1).Cells(m, 2).Value = orgdrill

                              m = m + 1

                              '共需铆钉数
                              m = m + 1
                              wb.Sheets(1).Cells(m, 1).Value = "RIVET"
                              wb.Sheets(1).Cells(m, 2).Value = sum - highliteall

                              m = m + 1
                              wb.Sheets(1).Cells(m, 1).Value = "BY MACHINE"
                              wb.Sheets(1).Cells(m, 2).Value = autosum - highlitesum

                              wb.Sheets(1).Cells(m, 3).Value = (autosum - highlitesum) / (sum - highliteall)
                              wb.Sheets(1).Cells(m, 3).NumberFormatLocal = "0.00%"

                              '手动安装总数
                              m = m + 1
                              wb.Sheets(1).Cells(m, 1).Value = "BY HAND"
                              wb.Sheets(1).Cells(m, 2).Value = (sum - highliteall) - (autosum - highlitesum)
                              wb.Sheets(1).Cells(m, 3).Value = ((sum - highliteall) - (autosum - highlitesum)) / (sum - highliteall)
                              wb.Sheets(1).Cells(m, 3).NumberFormatLocal = "0.00%"

                              '架下补铆安装铆钉

                              m = m + 1
                              wb.Sheets(1).Cells(m, 1).Value = "AFTER"
                              wb.Sheets(1).Cells(m, 2).Value = afterall - highafterall
                              wb.Sheets(1).Cells(m, 3).Value = (afterall - highafterall) / (sum - highliteall)
                              wb.Sheets(1).Cells(m, 3).NumberFormatLocal = "0.00%"
                              '作为校准点的铆钉数
                              m = m + 1
                              wb.Sheets(1).Cells(m, 1).Value = "RESYNC"
                              wb.Sheets(1).Cells(m, 2).Value = resynsum - highliteresyn
                              wb.Sheets(1).Cells(m, 3).Value = (resynsum - highliteresyn) / (sum - highliteall)
                              wb.Sheets(1).Cells(m, 3).NumberFormatLocal = "0.00%"
                              'wb.Sheets(1).Cells.EntireColumn.AutoFit

                              wb.Sheets(1).Columns("A:A").ColumnWidth = 30
                              wb.Sheets(1).Columns("B:Z").ColumnWidth = 15
                              wb.Sheets(1).Columns("B:Z").HorizontalAlignment = 3
                              wb.Sheets(1).Columns("A:Z").wraptext = True

                              wb.Sheets(1).Name = filename.Substring(0, 13)


                          End Sub


        If ifreport Then
            ReportRecursion()
        End If

                              If database Then
                                  updatedt()
                              End If
                              Return bugcontainer

    End Function


    'Private Sub creatCheckReport(Optional path = "")



    '    Dim rivottypecount As Integer
    '    Dim resultstrcount As Integer
    '    Dim rivottypename As String
    '    Dim sum As Integer
    '    Dim tempcount As Integer
    '    sum = 0
    '    Dim autosum As Integer
    '    autosum = 0
    '    Dim resynsum As Integer
    '    resynsum = 0


    '    Dim frconlyfinal As Integer
    '    frconlyfinal = 0

    '    Dim frconlyfastack As Integer
    '    frconlyfastack = 0

    '    Dim highlitesum As Integer
    '    highlitesum = 0
    '    Dim instodrill As Integer
    '    instodrill = 0
    '    '原始TVA就是仅仅钻孔的
    '    Dim orgdrill As Integer
    '    orgdrill = 0
    '    '高锁钻孔校准点数
    '    Dim highliteresyn As Integer
    '    highliteresyn = 0

    '    '校准插钉
    '    Dim highliterRC As Integer
    '    highliterRC = 0


    '    '仅校准导孔
    '    Dim FRChole As Integer
    '    FRChole = 0

    '    '校准导孔并安装
    '    Dim RChole As Integer
    '    RChole = 0


    '    '手铆和机器一共需用高锁数
    '    Dim highliteall As Integer
    '    highliteall = 0
    '    '手铆和机器下架补铆数
    '    Dim afterall As Integer
    '    afterall = 0
    '    '高锁下架补铆数
    '    Dim highafterall As Integer
    '    highafterall = 0

    '    '5号铆钉钻铆
    '    Dim fiverivet As Integer
    '    fiverivet = 0


    '    '5号高锁钻铆
    '    Dim fivehilite As Integer
    '    fivehilite = 0

    '    '5号高锁钻孔
    '    Dim fivedrill As Integer
    '    fivedrill = 0

    '    '6号高锁钻铆
    '    Dim sixhilite As Integer
    '    sixhilite = 0


    '    '6号高锁钻孔
    '    Dim sixdrill As Integer
    '    sixdrill = 0


    '    Dim excelobject As Object
    '    Dim wb As Object




    '    Dim m As Integer
    '    Dim n As Integer
    '    Dim p As Integer
    '    Dim fastextract As String
    '    Dim processextract As String
    '    Dim fastcount As Integer
    '    Dim proccesscount As Integer




    '    excelobject = CreateObject("excel.application") '启动Excel程序
    '    If path = "" Then

    '        excelobject.Visible = True
    '    Else
    '        excelobject.Visible = False
    '    End If
    '    wb = excelobject.Workbooks.Add()
    '    excelobject.displayalerts = False
    '    fastcount = fastenerqty.count
    '    proccesscount = processtype.count

    '    wb.Sheets(1).Cells(1, fastcount + 2).Value = "总计"

    '    wb.Sheets(1).Cells(proccesscount + 2, 1).Value = "总计"
    '    '制作表头，添加统计数据
    '    For p = 1 To processtype.count
    '        Dim processtypetmp As String

    '        processtypetmp = processtype.Key(p)
    '        processtypetmp = Strings.Replace(processtypetmp, " AUTOMATED FASTENING", "")


    '        wb.Sheets(1).Cells(p + 1, 1).Value = processtypetmp

    '        wb.Sheets(1).Cells(p + 1, fastcount + 2).Value = processtype.Item2(p)
    '    Next

    '    For n = 1 To fastcount
    '        wb.Sheets(1).Cells(1, n + 1).Value = fastenerqty.Key(n)
    '        wb.Sheets(1).Cells(proccesscount + 2, n + 1).Value = fastenerqty.Item2(n)
    '    Next


    '    Dim ff As Integer
    '    'For ff = 0 To buglistBox.Items.Count - 1
    '    '    If fastenerqty.InTheList(buglistBox.Items.Item(ff)) Then

    '    '        buglistBox.SetSelected(ff, True)
    '    '    Else
    '    '        buglistBox.SetSelected(ff, True)
    '    '    End If

    '    'Next

    '    wb.Sheets(1).Cells(1, 1).Value = "加工类型/紧固件"



    '    For m = 1 To checkTVApoint.count
    '        rivottypename = checkTVApoint.Key(m)
    '        tempcount = checkTVApoint.Item(rivottypename)
    '        fastextract = Strings.Trim(Strings.Split(rivottypename, " - ")(0))
    '        processextract = Strings.Trim(Strings.Split(rivottypename, " - ")(1))
    '        Dim rownum As Integer
    '        Dim colnum As Integer

    '        rownum = processtype.SearchIndex(processextract) + 1
    '        colnum = fastenerqty.SearchIndex(fastextract) + 1
    '        wb.Sheets(1).Cells(rownum, colnum).Value = tempcount
    '        'MsgBox (processextract)

    '        If (InStr(rivottypename, "AFTER AUTOMATED FASTENING") Or InStr(rivottypename, "BEFORE AUTOMATED FASTENING")) Then

    '            afterall = afterall + tempcount
    '        End If



    '        If (InStr(rivottypename, "RESYNCING")) Then 'aa
    '            resynsum = resynsum + tempcount
    '            '统计高锁校准点数
    '            If (InStr(rivottypename, "B020600")) Then 'bb
    '                highliteresyn = highliteresyn + tempcount
    '                If (InStr(rivottypename, "INSTALLED")) Then
    '                    highliterRC = highliterRC + tempcount
    '                End If
    '            End If 'bb

    '            If (InStr(rivottypename, "Pilot Holes")) Then 'cc

    '                If (InStr(rivottypename, "ONLY")) Then

    '                    FRChole = FRChole + tempcount

    '                Else
    '                    RChole = RChole + tempcount

    '                End If

    '            Else 'cc

    '                If (InStr(rivottypename, "ONLY")) Then

    '                    If (InStr(rivottypename, "Final")) Then
    '                        frconlyfinal = frconlyfinal + tempcount
    '                    Else
    '                        frconlyfastack = frconlyfastack + tempcount
    '                    End If
    '                End If


    '            End If 'cc

    '        Else 'aa


    '            If (InStr(rivottypename, "BY AUTOMATED FASTENING")) Then 'a
    '                If (InStr(rivottypename, "INSTALLED") Or InStr(rivottypename, "DRILL")) Then 'b
    '                    autosum = autosum + tempcount

    '                    If tempcount < 10 Then
    '                        wb.Sheets(1).Cells(rownum, colnum).Interior.ColorIndex = 3 ' 背景的颜色为3 红色
    '                    End If
    '                    '如果是高锁
    '                    If (InStr(rivottypename, "B020600")) Then 'c
    '                        highlitesum = highlitesum + tempcount
    '                        '统计自动钻铆安装转为仅仅钻孔的高锁
    '                        If (InStr(rivottypename, "INSTALLED")) Then 'd
    '                            'instodrill = instodrill + tempcount
    '                        Else 'd
    '                            '刚开始就仅仅是自动钻铆钻孔的
    '                            If (InStr(rivottypename, "DRILL")) Then 'e
    '                                orgdrill = orgdrill + tempcount
    '                            End If 'e


    '                        End If 'd

    '                    Else 'c
    '                        fiverivet = fiverivet + tempcount
    '                    End If 'c

    '                End If 'b
    '            End If 'a


    '        End If 'aa









    '        If (InStr(rivottypename, "B020600")) Then 'a
    '            highliteall = highliteall + tempcount


    '            If (InStr(rivottypename, "AFTER AUTOMATED FASTENING")) Then 'b
    '                highafterall = highafterall + tempcount

    '            Else 'b

    '                If (InStr(rivottypename, "RESYNCING") = 0) Then 'c


    '                    If (InStr(rivottypename, "FASTENER INSTALLED BY")) Then 'd

    '                        If (InStr(rivottypename, "AG5-")) Then
    '                            fivehilite = fivehilite + tempcount
    '                        Else
    '                            If (InStr(rivottypename, "AG6-")) Then
    '                                sixhilite = sixhilite + tempcount
    '                            End If
    '                        End If


    '                    Else 'd

    '                        If (InStr(rivottypename, "DRILL ONLY BY")) Then 'e

    '                            If (InStr(rivottypename, "AG5-")) Then
    '                                fivedrill = fivedrill + tempcount
    '                            Else
    '                                If (InStr(rivottypename, "AG6-")) Then
    '                                    sixdrill = sixdrill + tempcount
    '                                End If
    '                            End If


    '                        End If 'e


    '                    End If 'd


    '                End If 'c


    '            End If 'b
    '        End If 'a

    '        sum = sum + tempcount

    '    Next

    '    wb.Sheets(1).Cells(processtype.count + 2, fastenerqty.count + 2).Value = sum




    '    m = processtype.count + 4

    '    wb.Sheets(1).Cells(m, 1).Value = "仅校准(终钉):"
    '    wb.Sheets(1).Cells(m, 2).Value = frconlyfinal

    '    m = m + 1


    '    wb.Sheets(1).Cells(m, 1).Value = "仅校准(临时紧固件):"
    '    wb.Sheets(1).Cells(m, 2).Value = frconlyfastack


    '    m = m + 1

    '    wb.Sheets(1).Cells(m, 1).Value = "仅校准(导孔):"
    '    wb.Sheets(1).Cells(m, 2).Value = FRChole

    '    m = m + 1

    '    wb.Sheets(1).Cells(m, 1).Value = "校准(任意)插钉:"
    '    wb.Sheets(1).Cells(m, 2).Value = highliterRC



    '    m = m + 1

    '    wb.Sheets(1).Cells(m, 1).Value = "校准导孔安装:"
    '    wb.Sheets(1).Cells(m, 2).Value = RChole

    '    m = m + 1

    '    wb.Sheets(1).Cells(m, 1).Value = "5号铆钉钻铆:"
    '    wb.Sheets(1).Cells(m, 2).Value = fiverivet

    '    m = m + 1

    '    wb.Sheets(1).Cells(m, 1).Value = "5号高锁钻铆:"
    '    wb.Sheets(1).Cells(m, 2).Value = fivehilite

    '    m = m + 1

    '    wb.Sheets(1).Cells(m, 1).Value = "6号高锁钻铆:"
    '    wb.Sheets(1).Cells(m, 2).Value = sixhilite




    '    m = m + 1

    '    wb.Sheets(1).Cells(m, 1).Value = "5号高锁仅钻孔:"
    '    wb.Sheets(1).Cells(m, 2).Value = fivedrill

    '    m = m + 1


    '    wb.Sheets(1).Cells(m, 1).Value = "6号高锁仅钻孔:"
    '    wb.Sheets(1).Cells(m, 2).Value = sixdrill



    '    m = m + 1
    '    wb.Sheets(1).Cells(m, 1).Value = "其他点位:"
    '    wb.Sheets(1).Cells(m, 2).Value = afterall



    '    m = m + 2
    '    wb.Sheets(1).Cells(m, 1).Value = "Wrong:"
    '    m = m + 1
    '    wb.Sheets(1).Cells(m, 1).Value = "WrongType"
    '    wb.Sheets(1).Cells(m, 2).Value = "FrameName"
    '    wb.Sheets(1).Cells(m, 3).Value = "FastenerName"
    '    wb.Sheets(1).Cells(m, 4).Value = "ProccessType"
    '    wb.Sheets(1).Cells(m, 5).Value = "Qty"
    '    Dim wrongnum As Integer



    '    For wrongnum = 1 To wrongstatistic.count
    '        m = m + 1
    '        Dim wrongstr As String
    '        wrongstr = wrongstatistic.Key(wrongnum)

    '        wb.Sheets(1).Cells(m, 1).Value = Strings.Split(wrongstr, " - ")(0)
    '        wb.Sheets(1).Cells(m, 2).Value = Strings.Split(wrongstr, " - ")(1)
    '        wb.Sheets(1).Cells(m, 3).Value = Strings.Split(wrongstr, " - ")(2)
    '        wb.Sheets(1).Cells(m, 4).Value = Strings.Split(wrongstr, " - ")(3)
    '        wb.Sheets(1).Cells(m, 5).Value = wrongstatistic.Item2(wrongnum)

    '    Next

    '    m = m + 2
    '    wb.Sheets(1).Cells(m, 1).Value = "Statistic:"
    '    m = m + 1
    '    wb.Sheets(1).Cells(m, 2).Value = "Qty"
    '    wb.Sheets(1).Cells(m, 3).Value = "Percentage"

    '    m = m + 1
    '    wb.Sheets(1).Cells(m, 1).Value = "SUM"
    '    wb.Sheets(1).Cells(m, 2).Value = sum

    '    '输出需要自动钻铆安装总数
    '    On Error Resume Next
    '    m = m + 1
    '    wb.Sheets(1).Cells(m, 1).Value = "BY MACHINE"
    '    wb.Sheets(1).Cells(m, 2).Value = autosum
    '    wb.Sheets(1).Cells(m, 3).Value = autosum / sum
    '    wb.Sheets(1).Cells(m, 3).NumberFormatLocal = "0.00%"
    '    '输出需要手动安装总数

    '    m = m + 1
    '    wb.Sheets(1).Cells(m, 1).Value = "BY HAND"
    '    wb.Sheets(1).Cells(m, 2).Value = sum - autosum
    '    wb.Sheets(1).Cells(m, 3).Value = (sum - autosum) / sum
    '    wb.Sheets(1).Cells(m, 3).NumberFormatLocal = "0.00%"

    '    '架下补铆安装总共

    '    m = m + 1
    '    wb.Sheets(1).Cells(m, 1).Value = "AFTER"
    '    wb.Sheets(1).Cells(m, 2).Value = afterall
    '    wb.Sheets(1).Cells(m, 3).Value = afterall / sum
    '    wb.Sheets(1).Cells(m, 3).NumberFormatLocal = "0.00%"


    '    '需预定位的总数

    '    m = m + 1
    '    wb.Sheets(1).Cells(m, 1).Value = "RESYNC"
    '    wb.Sheets(1).Cells(m, 2).Value = resynsum
    '    wb.Sheets(1).Cells(m, 3).Value = resynsum / sum
    '    wb.Sheets(1).Cells(m, 3).NumberFormatLocal = "0.00%"
    '    m = m + 1

    '    '共需高锁数
    '    m = m + 1
    '    wb.Sheets(1).Cells(m, 1).Value = "HI-LITE"
    '    wb.Sheets(1).Cells(m, 2).Value = highliteall

    '    m = m + 1
    '    wb.Sheets(1).Cells(m, 1).Value = "BY MACHINE"
    '    wb.Sheets(1).Cells(m, 2).Value = highlitesum
    '    wb.Sheets(1).Cells(m, 3).Value = highlitesum / highliteall
    '    wb.Sheets(1).Cells(m, 3).NumberFormatLocal = "0.00%"


    '    '手动安装总数
    '    m = m + 1
    '    wb.Sheets(1).Cells(m, 1).Value = "BY HAND"
    '    wb.Sheets(1).Cells(m, 2).Value = highliteall - highlitesum
    '    wb.Sheets(1).Cells(m, 3).Value = (highliteall - highlitesum) / highliteall
    '    wb.Sheets(1).Cells(m, 3).NumberFormatLocal = "0.00%"


    '    '架下补铆安装高锁

    '    m = m + 1
    '    wb.Sheets(1).Cells(m, 1).Value = "AFTER"
    '    wb.Sheets(1).Cells(m, 2).Value = highafterall
    '    wb.Sheets(1).Cells(m, 3).Value = highafterall / highliteall
    '    wb.Sheets(1).Cells(m, 3).NumberFormatLocal = "0.00%"

    '    '作为校准点的高锁数
    '    m = m + 1
    '    wb.Sheets(1).Cells(m, 1).Value = "RESYNC"
    '    wb.Sheets(1).Cells(m, 2).Value = highliteresyn
    '    wb.Sheets(1).Cells(m, 3).Value = highliteresyn / highliteall
    '    wb.Sheets(1).Cells(m, 3).NumberFormatLocal = "0.00%"


    '    '高锁自动钻铆钻孔总数
    '    m = m + 1
    '    wb.Sheets(1).Cells(m, 1).Value = "DRILL ONLY"
    '    wb.Sheets(1).Cells(m, 2).Value = orgdrill

    '    m = m + 1

    '    '共需铆钉数
    '    m = m + 1
    '    wb.Sheets(1).Cells(m, 1).Value = "RIVET"
    '    wb.Sheets(1).Cells(m, 2).Value = sum - highliteall

    '    m = m + 1
    '    wb.Sheets(1).Cells(m, 1).Value = "BY MACHINE"
    '    wb.Sheets(1).Cells(m, 2).Value = autosum - highlitesum

    '    wb.Sheets(1).Cells(m, 3).Value = (autosum - highlitesum) / (sum - highliteall)
    '    wb.Sheets(1).Cells(m, 3).NumberFormatLocal = "0.00%"

    '    '手动安装总数
    '    m = m + 1
    '    wb.Sheets(1).Cells(m, 1).Value = "BY HAND"
    '    wb.Sheets(1).Cells(m, 2).Value = (sum - highliteall) - (autosum - highlitesum)
    '    wb.Sheets(1).Cells(m, 3).Value = ((sum - highliteall) - (autosum - highlitesum)) / (sum - highliteall)
    '    wb.Sheets(1).Cells(m, 3).NumberFormatLocal = "0.00%"

    '    '架下补铆安装铆钉

    '    m = m + 1
    '    wb.Sheets(1).Cells(m, 1).Value = "AFTER"
    '    wb.Sheets(1).Cells(m, 2).Value = afterall - highafterall
    '    wb.Sheets(1).Cells(m, 3).Value = (afterall - highafterall) / (sum - highliteall)
    '    wb.Sheets(1).Cells(m, 3).NumberFormatLocal = "0.00%"
    '    '作为校准点的铆钉数
    '    m = m + 1
    '    wb.Sheets(1).Cells(m, 1).Value = "RESYNC"
    '    wb.Sheets(1).Cells(m, 2).Value = resynsum - highliteresyn
    '    wb.Sheets(1).Cells(m, 3).Value = (resynsum - highliteresyn) / (sum - highliteall)
    '    wb.Sheets(1).Cells(m, 3).NumberFormatLocal = "0.00%"
    '    'wb.Sheets(1).Cells.EntireColumn.AutoFit

    '    wb.Sheets(1).Columns("A:A").ColumnWidth = 30
    '    wb.Sheets(1).Columns("B:Z").ColumnWidth = 15
    '    wb.Sheets(1).Columns("B:Z").HorizontalAlignment = 3
    '    wb.Sheets(1).Columns("A:Z").wraptext = True

    '    If path = "" Then

    '    Else
    '        wb.SaveAs(path)
    '        excelobject.Quit()
    '    End If




    'End Sub
    'Public Function CheckTVA(ByRef MyGeoSet As HybridBody, Optional ByVal color As Boolean = True)
    '    ' CATIA = MyGeoSet.Application


    '    '2015.5.12进行重构，与界面元素相分离
    '    Dim buglocation As String

    '    Dim MySourceGeoSet As HybridBody
    '    Dim shapecount As Integer
    '    Dim tempstring As String

    '    If TVA_Method.ifFastener(MyGeoSet.Name) Then

    '        '紧固件列表
    '        fastenerqty.Add(0, MyGeoSet.Name)



    '        Dim i
    '        '开始遍历紧固件的下层几何图形集
    '        For i = 1 To MyGeoSet.HybridBodies.Count


    '            Dim MyNewGeoSet As HybridBody

    '            MyNewGeoSet = MyGeoSet.HybridBodies.Item(i)

    '            MySourceGeoSet = MyNewGeoSet

    '            '计算该几何图形集中点的数量
    '            shapecount = Fix(MySourceGeoSet.HybridShapes.Count / 2)



    '            '判断几何图形集下是否有奇数个shape ,若有，则报错
    '            If (MySourceGeoSet.HybridShapes.Count / 2 - shapecount) <> 0 Then
    '                shapecount = shapecount + 1
    '                buglocation = MySourceGeoSet.Parent.Parent.name + " - " + MySourceGeoSet.Name

    '                wrongstatistic.Add(shapecount, "error in qty of point and line" + " - " + buglocation)

    '                '以下过程找出孤立的点，并加入到bug列表中
    '                bugcontainer.Add(Checkvector(MySourceGeoSet, buglocation))
    '            End If

    '            '  Checkdupli MySourceGeoSet
    '            If Strings.InStr(MyNewGeoSet.Name, "TEMP") And MyNewGeoSet.HybridShapes.Count <> 0 Then

    '                wrongstatistic.Add(shapecount, "Something in Temp GeoSet" + " - " + MyGeoSet.Parent.name + " - " + MyNewGeoSet.Name)

    '            End If




    '            If Strings.InStr(MyNewGeoSet.Name, "RESYNCING") Then

    '                If MyNewGeoSet.HybridShapes.Count <> 0 Then
    '                    wrongstatistic.Add(shapecount, "Need select Target Type" + " - " + MyGeoSet.Parent.name + " - " + MyNewGeoSet.Name)

    '                End If


    '                Dim ii
    '                For ii = 0 To 2



    '                    'Set MySourceGeoSet = MyNewGeoSet.HybridBodies.GetItem(UserForm1.ComboBox4.Items.Item(ii))
    '                    If (ifGeoExist(MyNewGeoSet, processtree.sectree(ii)) = False) Then

    '                        MySourceGeoSet = MyNewGeoSet.HybridBodies.Add()
    '                        MySourceGeoSet.Name = processtree.sectree(ii)
    '                    Else
    '                        MySourceGeoSet = MyNewGeoSet.HybridBodies.GetItem(processtree.sectree(ii))

    '                    End If


    '                    If color Then

    '                        Dim selection1 As Object


    '                        selection1 = CATIA.ActiveDocument.Selection
    '                        selection1.Clear()
    '                        Dim zz
    '                        For zz = 1 To MySourceGeoSet.HybridShapes.Count
    '                            selection1.Add(MySourceGeoSet.HybridShapes.Item(zz))

    '                        Next


    '                        If Strings.InStr(MyNewGeoSet.Name, "RESYNCING ONLY") Then

    '                            '红色
    '                            selection1.VisProperties.SetRealColor(255, 0, 0, 1)
    '                        Else
    '                            '绿色
    '                            selection1.VisProperties.SetRealColor(0, 128, 64, 1)

    '                        End If


    '                        If Strings.InStr(MySourceGeoSet.Name, "Final") Then



    '                            '  selection1.VisProperties.SetSymbolType(6)
    '                            If Strings.InStr(MyGeoSet.Name, "B020600") Then
    '                                '设为大圆
    '                                selection1.VisProperties.SetSymbolType(3)
    '                            Else

    '                                '设为十字
    '                                selection1.VisProperties.SetSymbolType(2)
    '                            End If
    '                        Else
    '                            If Strings.InStr(MySourceGeoSet.Name, "Fast Tack") Then
    '                                '星号
    '                                selection1.VisProperties.SetSymbolType(7)

    '                            Else
    '                                '同心圆

    '                                If Strings.InStr(MyNewGeoSet.Name, "DRILL") Then

    '                                    selection1.VisProperties.SetSymbolType(5)
    '                                Else

    '                                    selection1.VisProperties.SetSymbolType(4)

    '                                End If


    '                            End If

    '                        End If

    '                    End If


    '                    '检查校准的几何图形集



    '                    'If database Then


    '                    '    ' buglocation = MyGeoSet.Parent.name + " - " + MyNewGeoSet.Name + " - " + MySourceGeoSet.Name
    '                    '    updatedatabase(MySourceGeoSet, MyGeoSet.Parent.name, MyGeoSet.Name, Strings.Split(MyNewGeoSet.Name, " - ")(1) + " - " + MySourceGeoSet.Name)
    '                    'End If



    '                    shapecount = Fix(MySourceGeoSet.HybridShapes.Count / 2)
    '                    'MsgBox MySourceGeoSet.name



    '                    If (MySourceGeoSet.HybridShapes.Count / 2 - shapecount) <> 0 Then
    '                        shapecount = shapecount + 1
    '                        buglocation = MyGeoSet.Parent.name + " - " + MyNewGeoSet.Name + " - " + MySourceGeoSet.Name


    '                        wrongstatistic.Add(shapecount, "error in qty of point and line" + " - " + buglocation)


    '                        bugcontainer.Add(Checkvector(MySourceGeoSet, buglocation))

    '                    End If

    '                    tempstring = Strings.Trim(Strings.Split(MyNewGeoSet.Name, " - ")(1)) + "-" + Strings.Trim(Strings.Split(MySourceGeoSet.Name, " - ")(1))

    '                    checkTVApoint.Add(shapecount, MyNewGeoSet.Name + "-" + Strings.Trim(Strings.Split(MySourceGeoSet.Name, " - ")(1)))

    '                    '统计各个紧固件数量Fastener_Qty
    '                    fastenerqty.Add(shapecount, MyGeoSet.Name)
    '                    processtype.Add(shapecount, tempstring)

    '                Next

    '            Else

    '                '不为resync
    '                '检查非校准的几何图形集
    '                If (MyNewGeoSet.Name.Contains("AUTOMATED FASTENING") Or MyNewGeoSet.Name.Contains("TEMP")) Then



    '                    'If database Then

    '                    '    updatedatabase(MyNewGeoSet, MyGeoSet.Parent.name, MyGeoSet.Name, Strings.Split(MyNewGeoSet.Name, " - ")(1))
    '                    'End If

    '                    If color Then

    '                        Dim selection1 As Object

    '                        selection1 = CATIA.ActiveDocument.Selection
    '                        selection1.Clear()
    '                        Dim zzz
    '                        For zzz = 1 To MyNewGeoSet.HybridShapes.Count
    '                            selection1.Add(MyNewGeoSet.HybridShapes.Item(zzz))

    '                        Next



    '                        If Strings.InStr(MyNewGeoSet.Name, "BY") Then


    '                            If Strings.InStr(MyGeoSet.Name, "5-") Then

    '                                If Strings.InStr(MyGeoSet.Name, "5020AD") Or Strings.InStr(MyGeoSet.Name, "6002AG") Then
    '                                    '蓝色
    '                                    selection1.VisProperties.SetRealColor(0, 0, 255, 1)
    '                                Else
    '                                    selection1.VisProperties.SetRealColor(0, 255, 255, 1)
    '                                End If

    '                            Else
    '                                '棕色

    '                                If Strings.InStr(MyGeoSet.Name, "5020AD") Or Strings.InStr(MyGeoSet.Name, "6002AG") Then

    '                                    selection1.VisProperties.SetRealColor(128, 64, 64, 1)
    '                                Else
    '                                    selection1.VisProperties.SetRealColor(255, 128, 0, 1)
    '                                End If
    '                            End If



    '                            If Strings.InStr(MyNewGeoSet.Name, "INSTALL") Then
    '                                If Strings.InStr(MyGeoSet.Name, "B020600") Then
    '                                    '设为大圆
    '                                    selection1.VisProperties.SetSymbolType(3)
    '                                Else

    '                                    '设为十字
    '                                    selection1.VisProperties.SetSymbolType(2)
    '                                End If


    '                                'selection1 = Nothing
    '                            Else

    '                                If Strings.InStr(MyNewGeoSet.Name, "DRILL") Then

    '                                    '设为实心圆
    '                                    selection1.VisProperties.SetSymbolType(5)

    '                                End If

    '                            End If
    '                        Else


    '                            If Strings.InStr(MyNewGeoSet.Name, "AFTER") Then
    '                                '符号为X
    '                                'after 点为黑色
    '                                selection1.VisProperties.SetRealColor(0, 0, 0, 1)
    '                                selection1.VisProperties.SetSymbolType(1)
    '                            Else

    '                                If Strings.InStr(MyGeoSet.Name, "5-") Then
    '                                    '都改为深色

    '                                    If Strings.InStr(MyGeoSet.Name, "5020AD") Or Strings.InStr(MyGeoSet.Name, "6002AG") Then
    '                                        '紫色
    '                                        selection1.VisProperties.SetRealColor(255, 0, 255, 1)
    '                                    Else
    '                                        '深色
    '                                        selection1.VisProperties.SetRealColor(255, 165, 0, 1)
    '                                    End If

    '                                Else
    '                                    '棕色

    '                                    If Strings.InStr(MyGeoSet.Name, "5020AD") Or Strings.InStr(MyGeoSet.Name, "6002AG") Then

    '                                        selection1.VisProperties.SetRealColor(64, 32, 32, 1)
    '                                    Else
    '                                        'Make protruding head 6# darker
    '                                        selection1.VisProperties.SetRealColor(64, 32, 0, 1)
    '                                    End If
    '                                End If

    '                                If Strings.InStr(MyNewGeoSet.Name, "DRILL") Then
    '                                    selection1.VisProperties.SetSymbolType(4)
    '                                Else

    '                                    If Strings.InStr(MyGeoSet.Name, "B020600") Then
    '                                        '设为大圆
    '                                        selection1.VisProperties.SetSymbolType(3)
    '                                    Else

    '                                        '设为十字
    '                                        selection1.VisProperties.SetSymbolType(2)
    '                                    End If

    '                                End If

    '                            End If




    '                        End If


    '                    End If









    '                    If Strings.InStr(MyNewGeoSet.Name, " - ") Then

    '                        tempstring = Strings.Trim(Strings.Split(MySourceGeoSet.Name, " - ")(1))
    '                        checkTVApoint.Add(shapecount, MySourceGeoSet.Name)

    '                        fastenerqty.Add(shapecount, MyGeoSet.Name)
    '                        processtype.Add(shapecount, tempstring)
    '                    End If
    '                Else
    '                    wrongstatistic.Add(shapecount, "error geoset" + " - " + MyNewGeoSet.Name + " - ")

    '                End If
    '            End If


    '        Next

    '    Else
    '        If MyGeoSet.HybridShapes.Count <> 0 Then
    '            wrongstatistic.Add(shapecount, "keng brother bug" + " - " + MyGeoSet.Parent.name + " - " + MyGeoSet.Name)
    '        End If

    '        Dim k As Integer

    '        For k = 1 To MyGeoSet.HybridBodies.Count
    '            '开始递归
    '            CheckTVA(MyGeoSet.HybridBodies.Item(k))
    '        Next
    '    End If






    '    '生产表格













    '    Return Nothing


    'End Function


#End Region

#Region "Revise ProcType Batch"

    Public Function procBAT(FastList As List(Of String), sourceType As String, targetType As String, Optional pilotname As String = "") As processStatic
        Dim rpt = New processStatic()

        FstList = FastList
        hbtree = New List(Of HybridBody)()
        LoopThroughTVA(pilot_geoset())
        For Each pp In hbtree
            Dim dpp = New processTree(pp)
            rpt.Add(dpp.MoveType(filename, sourceType, targetType))

        Next

        Return rpt

    End Function

    Sub LoopThroughTVA(ByRef MyGeoSet As HybridBody)

        Dim k As Integer

        For k = 1 To MyGeoSet.HybridBodies.Count
            '开始递归
            If FstList.Contains(MyGeoSet.HybridBodies.Item(k).Name) Then
                hbtree.Add(MyGeoSet.HybridBodies.Item(k))
                '    tree.Add(New processTree(MyGeoSet.HybridBodies.Item(k)))
            Else

                LoopThroughTVA(MyGeoSet.HybridBodies.Item(k))

            End If

        Next


    End Sub

    'Sub LoopThroughTVA(ByRef MyGeoSet As HybridBody, proctype As String)

    '    Dim k As Integer

    '    For k = 1 To MyGeoSet.HybridBodies.Count
    '        '开始递归
    '        If MyGeoSet.HybridBodies.Item(k).Name.Contains(proctype) Then

    '            processtype.Add(MyGeoSet.HybridBodies.Item(k), processtype.count + 1)
    '        Else

    '            LoopThroughTVA(MyGeoSet.HybridBodies.Item(k), proctype)

    '        End If

    '    Next


    'End Sub

    Public Shared Function getprocHB(fsthb As HybridBody, processty As String) As HybridBody
        Dim splited = Strings.Split(processty, " - ")



        If processty.Contains("RESYNCING") Then

            Return fsthb.HybridBodies.GetItem(fsthb.Name + " - " + splited(0)).HybridBodies.GetItem(splited(1) + " - " + splited(2))
        Else

            Return fsthb.HybridBodies.GetItem(fsthb.Name + " - " + splited(0))

        End If






    End Function

    Public Shared Function ifFastener(name As String) As Boolean
        Return (name.Contains("-") And (name.Length = 13 Or name.Length = 14)) And (name.Contains("AD") Or name.Contains("AG"))
    End Function


    Sub setHide(ifshow As Boolean, proctype As String)
        hbtree = New List(Of HybridBody)()

        '  If hbtree Is Nothing Then
        LoopThroughTVA(pilot_geoset())
        ' End If

        Dim selection1 = CATIA.ActiveDocument.Selection
        selection1.Clear()



        For Each dpp In hbtree

            selection1.Add(getprocHB(dpp, proctype))



        Next
        If ifshow Then
            selection1.VisProperties.SetShow(0)
        Else

            selection1.VisProperties.SetShow(1)
        End If


    End Sub



#End Region

    Public Sub updatedt(Optional suffix As String = "")







        Dim productname As String = localMethod.skin_to_drawing(filename)
        If productname <> "" Then



            TVAPointsnoVic.outputdb(productname + suffix)
        Else

            Exit Sub


        End If








    End Sub

    Public Shared Sub update_TVAall(skinpathlist As List(Of String), Optional suffix As String = "")

        For Each ss As String In skinpathlist

            Dim aa = New TVA_Method(ss)

            aa.updatedt(suffix)



        Next
    End Sub


    Public Shared Sub GetFatherProduct(ByVal oProduct As Product, ByVal oPartDoc As PartDocument, ByRef oFatherProduct As Product, Optional ByRef sPath As Object = Nothing)
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

    Public Shared Function GetProductList(ByVal oProduct As Product) As List(Of String)
        Dim prodlist As New List(Of String)
        GetProductList(oProduct, prodlist, 0)
        Return prodlist
    End Function



    Public Shared Sub GetProductList(ByVal oProduct As Product, ll As List(Of String), cls As Integer)
        Dim ii As Integer
        Dim oSubproduct As Product


        Dim i As Integer = 0
        Dim mark As String = ""
        While i < cls

            mark = mark + "*"
            i = i + 1
        End While
        cls = cls + 1
        ' browse all elements in transfered product
        For ii = 1 To oProduct.Products.Count
            oSubproduct = oProduct.Products.Item(ii)
            '  Err.Clear()
            ' found part is part whose father product is searched ?
            ' Dim tmpfullname As String
            'tmpfullname = oSubproduct.ReferenceProduct.Parent.Name

            ll.Add(mark + " " + oSubproduct.PartNumber)
            ' element is a product
            ' -> browse it
            If oSubproduct.Products.Count <> 0 Then

                ' recursive call
                Call GetProductList(oSubproduct, ll, cls)


            End If
        Next




    End Sub


    Public Sub CheckFasteners(SPProducts As Dictionary(Of String, Product))


        Dim sourcepplist As CENPoints
        sourcepplist = TVAPoints


        Dim SP01s_ = New CENSP01s(SPProducts.Values.AsEnumerable())



        Dim comresult() As CENPoints = {New CENPoints, New CENPoints, New CENPoints}
        comresult = sourcepplist.compare(SP01s_)
        Dim changelist = comresult(0).toProcessTree()
        If changelist.count() > 0 Then
            Dim targetBody = PilotHoles.HybridBodies.Add
            targetBody.Name = "ChangeFstType" + Now.Date.ToShortDateString()
            changelist.output_topart(filename, filename, targetBody, 0)
            changelist.del(filename)
        End If



        Dim dellist = comresult(1).toProcessTree()
        If dellist.count() > 0 Then

            Dim targetBody2 = Part.HybridBodies.Add
            targetBody2.Name = "DoubtBeDeleted" + Now.Date.ToShortDateString()
            dellist.output_topart(filename, filename, targetBody2, 0)
            ' dellist.del(TVAfilename)
        End If
        Part.Update()

    End Sub



    Public Shared Function ifGeoExist(ByRef MyGeoSet As HybridBody, ii As String) As Boolean

        On Error GoTo Error_Handler

        Dim MySourceGeoSet As HybridBody
        MySourceGeoSet = MyGeoSet.HybridBodies.GetItem(ii)
        'Let the collection handle the search instead of VB

        If MySourceGeoSet Is Nothing Then
            ifGeoExist = False
        Else
            ifGeoExist = True
        End If
        Exit Function

Error_Handler:

        ifGeoExist = False

    End Function

    Public Function Checkvector(ByRef MyGeoSet As HybridBody, buglocation As String) As processStatic

        Dim bugfeedback As New processStatic()
        'On Error GoTo Here1

        Dim shapecount As Integer
        shapecount = MyGeoSet.HybridShapes.Count
        Dim foundvectall() As Boolean
        ReDim foundvectall(0 To shapecount)
        Dim i As Integer

        For i = 1 To shapecount
            foundvectall(i) = False
        Next




        Dim m
        For m = 1 To shapecount
            'if 1
            If CheckHybridShapeItem(MyGeoSet.HybridShapes.Item(m)) Then



                'Get XYZ of point
                Dim SPAWorkb As Workbench
                Dim Measurement
                Dim Coords(2) As Object
                '  Dim CATIA = MyGeoSet.Application
                ' Dim partDoc = CATIA.ActiveDocument

                ' Dim part1 As Part
                '   part1 = Part

                'Get the SPAWorkbench from the measurement
                SPAWorkb = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")

                'Get the measurement for the point
                Measurement = SPAWorkb.GetMeasurable(MyGeoSet.HybridShapes.Item(m))

                'Get the coordinates (Part based) from this point
                'Try

                Call Measurement.GetPoint(Coords)
                'Catch ex As Exception
                '    Console.Write(MyGeoSet.HybridShapes.Item(m).Name)
                'End Try


                Dim myVect As Object
                Dim s
                Dim foundvect
                foundvect = False
                For s = 1 To shapecount
                    'if 3
                    If (foundvectall(s) = False) Then
                        If TypeName(MyGeoSet.HybridShapes.Item(s)) = "HybridShapeLineNormal" Or TypeName(MyGeoSet.HybridShapes.Item(s)) = "HybridShapeLineExplicit" Then

                            myVect = MyGeoSet.HybridShapes.Item(s)

                            'Get the SPAWorkbench from the measurement
                            SPAWorkb = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")
                            Dim reference4 As Reference
                            reference4 = Part.CreateReferenceFromObject(MyGeoSet.HybridShapes.Item(s))
                            Dim MinimumDistance As Double
                            MinimumDistance = Measurement.GetMinimumDistance(reference4)

                            'Now get the XYZ of the point
                            If MinimumDistance <= 0.002 Then



                                foundvect = True
                                foundvectall(m) = True
                                foundvectall(s) = True
                                Exit For
                            End If



                        End If

                        'end if 3
                    End If
                Next





                If foundvect = False Then
                    bugfeedback.Add(MyGeoSet.HybridShapes.Item(m), "wrongpoint,NO.:" + CStr(m) + ":" + buglocation)

                    foundvectall(m) = True
                End If




                'end if 1
            End If

        Next
        For i = 1 To shapecount
            If foundvectall(i) = False Then
                bugfeedback.Add(MyGeoSet.HybridShapes.Item(i), "line,NO.:" + CStr(m) + ":" + buglocation)
            End If
        Next

        '  Exit Sub

        'Here1:

        ' MsgBox ("Error")

        Return bugfeedback




    End Function
    Public Sub setcolor()

        For Each dd As processTree In TVATreeList.treeList.Values
            dd.setcolor()



        Next


    End Sub

    Public Sub CopyPastePartBody(targetfile As String)
        'before copying,it need to activate the whole father product
        'Get a handle to the SkinPart
        Dim selection1

        Dim documents1 As INFITF.Documents
        documents1 = CATIA.Documents
        Dim partDocument1 = documents1.Item(filename)
        'Dim partDocument1
        'If PointsFatherProduct Is Nothing Then
        '    partDocument1 = documents1.Item(filename)
        'Else
        '    Dim skinproductstr As String
        '    skinproductstr = PointsFatherProduct.Parent.name

        '    partDocument1 = documents1.Item(skinproductstr)
        'End If
        '  partDocument1 = documents1.Item(filename)
        'Activate the name of the correct window
        partDocument1.Activate()

        selection1 = CATIA.ActiveDocument.Selection
        selection1.Clear()

        Dim part1 As Part
        part1 = Part

        'Get the title off of it
        Dim parameters1 As Parameters
        parameters1 = part1.Parameters

        Dim strParam1 As StrParam
        On Error Resume Next
        strParam1 = parameters1.Item("Properties\Title")

        Dim bodies1 As Bodies

        bodies1 = part1.Bodies

        Dim body1 As Body
        body1 = bodies1.Item("PartBody")

        If body1 Is Nothing Then
            body1 = bodies1.Item("Part Body")
        End If

        If body1 Is Nothing Then
            body1 = bodies1.Item("零件几何体")
        End If
        'If it is still blank, yell at the user
        If body1 Is Nothing Then
            MsgBox("Cannot fine 'Part Body' or 'PartBody'")
            Exit Sub
        End If
        Dim i

        'If skin part isnt open this errors out!!!
        selection1.Add(body1)
        selection1.Copy()

        'Create a 2nd part as a target for the copy operation
        Dim partDocument2 As PartDocument

        partDocument2 = documents1.Item(targetfile)
        partDocument2.Activate()

        Dim selection2
        selection2 = CATIA.ActiveDocument.Selection

        Dim part2 As Part
        part2 = partDocument2.Part

        'Copy title in both COS/TVA
        parameters1 = part2.Parameters

        Dim strParam2 As StrParam
        strParam2 = parameters1.Item("Properties\Title")

        strParam2.Value = strParam1.Value

        selection2.Clear()

        selection2.Add(part2)

        selection2.PasteSpecial("CATPrtResultWithOutLink")

        part2.Update()

        selection1.Clear()
        selection2.Clear()

        Dim body2 As Body

        bodies1 = part2.Bodies
        body2 = bodies1.Item(2)

        selection2.Add(body2.Shapes.Item(1))

        selection2.Copy()

        selection1 = CATIA.ActiveDocument.Selection

        body1 = bodies1.Item(1)

        selection1.Add(body1)

        selection1.Paste()

        selection1.Clear()

        selection2.Add(body2.Shapes.Item(1))
        selection2.Delete()

        part2.Update()

        selection1 = Nothing
        selection2 = Nothing

    End Sub


    Public Shared Function CheckHybridShapeItem(oItem As Object) As Boolean?


        Dim aa = TypeName(oItem)
        If aa.Contains("HybridShapePoint") Or (aa = "HybridShapeProject") Or (aa = "HybridShapeIntersection") Then
            Return True
        Else
            If aa.Contains("HybridShapeLine") Then
                Return False

            Else
                ' If the shape is neither point or line,return null
                Return Nothing
            End If




        End If



    End Function

End Class

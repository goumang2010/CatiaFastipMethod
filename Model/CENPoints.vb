Option Explicit On
Imports mysqlsolution
Imports MECMOD
Imports HybridShapeTypeLib
Imports INFITF



Public Class CENPoints

    Public hb As HybridBody
    Public Points_ As New Dictionary(Of String, CENPoint)
    Public dupli As New List(Of CENPoint)
    Dim wrongstatistic As New processStatic()
    Public CATIA
    'comment

    '
    '   Default constructor
    '
    Public Sub New()

    End Sub
    Public Sub New(ByRef hbb As HybridBody)
        hb = hbb
        walk(hbb)
    End Sub
    '通过几何图形集建立
    Public Function clone() As CENPoints

        Dim tt As New CENPoints
        With tt
            .CATIA = CATIA
            .wrongstatistic = wrongstatistic

            .hb = hb
        End With

        tt.Merge(Me)
        Return tt
    End Function
    Public Sub walk(ByRef hbb As HybridBody)
        CATIA = hbb.Application
        Dim part1 As Part
        part1 = CATIA.ActiveDocument.part
        Dim SPAWorkb As INFITF.Workbench

        Dim Measurement
        SPAWorkb = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")
        Dim tmppoints = New List(Of HybridShape)()
        Dim tmplines = New List(Of HybridShape)()
        For Each hs As HybridShape In hbb.HybridShapes
            '判断点
            If TVA_Method.CheckHybridShapeItem(hs) Then

                tmppoints.Add(hs)
            Else

                tmplines.Add(hs)
            End If

        Next

        For Each pp As HybridShape In tmppoints
            'Get the SPAWorkbench from the measurement

            Dim Coords(2) As Object





            'Get the measurement for the point
            Measurement = SPAWorkb.GetMeasurable(pp)

            'Get the coordinates (Part based) from this point
            Call Measurement.GetPoint(Coords)

            For Each ppk As HybridShape In tmplines


                'Get the SPAWorkbench from the measurement
                SPAWorkb = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")
                Dim reference5 As INFITF.Reference
                reference5 = part1.CreateReferenceFromObject(ppk)
                Dim MinimumDistance2 As Double
                MinimumDistance2 = Measurement.GetMinimumDistance(reference5)

                'Now get the XYZ of the point
                If MinimumDistance2 < 0.1 Then
                    Add(TVA_Method.LeaftoPoint(pp, "", "", "", pp.Name, ppk))
                    '找到后，把该项移除
                    tmplines.Remove(ppk)
                    Exit For

                End If

            Next

        Next

        For Each bb In hbb.HybridBodies
            walk(bb)

        Next

    End Sub






    Public Function fix_outlines(partfilename As String, surface As HybridShapeExtract) As processStatic

        '考虑只有一个面的情况


        Dim ppstatic As New processStatic
        Dim documents1 As Documents

        Dim partDocument1 As PartDocument
        documents1 = CATIA.Documents
        partDocument1 = documents1.Item(partfilename)
        Dim part1 As Part
        part1 = partDocument1.Part
        partDocument1.Activate()
        Dim SPAWorkb As INFITF.Workbench
        Dim Measurement
        Dim Coords(2) As Object
        Dim source, destination, sourceinter
        source = CATIA.ActiveDocument.Selection
        sourceinter = CATIA.ActiveDocument.Selection
        destination = CATIA.ActiveDocument.Selection






        Dim reference2 As Reference
        SPAWorkb = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")
        ' Dim tempgeo = hb.HybridBodies.Add()

        part1.InWorkObject = hb
        For Each pp As CENPoint In Points_.Values


            'Get the measurement for the point
            Measurement = SPAWorkb.GetMeasurable(pp.MyPoint)

            'Get the coordinates (Part based) from this point
            Call Measurement.GetPoint(Coords)
            ' reference1 = part1.CreateReferenceFromObject(pp.MyPoint)
            '量取线到法向的距离
            reference2 = part1.CreateReferenceFromObject(pp.MyVector)

            Dim MinimumDistance As Double
            MinimumDistance = Measurement.GetMinimumDistance(reference2)




            Dim reference1 As Reference
            reference1 = part1.CreateReferenceFromObject(surface)

            '点到面上的距离
            Dim MinimumDistance2 As Double
            MinimumDistance2 = Measurement.GetMinimumDistance(reference1)

            '线到面的距离 
            Dim TheMeasurable As SPATypeLib.Measurable
            TheMeasurable = SPAWorkb.GetMeasurable(reference2)
            Dim MinimumDistance3 As Double
            MinimumDistance3 = TheMeasurable.GetMinimumDistance(reference1)

            Dim noerr As Boolean = False

            If MinimumDistance3 > 0 Then

            Else

                noerr = True
            End If







            If noerr And (MinimumDistance > 0 Or MinimumDistance2 > 0) Then
                '检测到点不在法向上
                ppstatic.Add(1, pp.Frame + " - " + pp.Diam)
                Dim hybridShapeIntersection1 As HybridShapeIntersection
                hybridShapeIntersection1 = part1.HybridShapeFactory.AddNewIntersection(reference2, reference1)
                hybridShapeIntersection1.PointType = 0
                ' hybridShapeIntersection1.ExtendMode = 3
                hb.AppendHybridShape(hybridShapeIntersection1)
                part1.Update()
                ' hb.AppendHybridShape(hybridShapeIntersection1)
                sourceinter.Clear()
                sourceinter.Add(hybridShapeIntersection1)
                sourceinter.Copy()
                destination.Clear()
                destination.Add(hb)
                destination.PasteSpecial("CATPrtResultWithOutLink")

                sourceinter.Clear()
                sourceinter.Add(hybridShapeIntersection1)
                sourceinter.Delete()



                source.Clear()

                source.Add(pp.MyPoint)

                source.Delete()


                part1.Update()

                '重新绑定point
                '不做其他操作了，则无必要
                'pp.MyPoint = hb.HybridShapes.Item(hb.HybridShapes.Count)

            End If


        Next
        '整体删除临时几何图形集
        'sourceinter.Clear()
        'sourceinter.Add(hybridShapeIntersection1)
        'sourceinter.Delete()


        Return ppstatic

    End Function
    Public Function fix_outlines(partfilename As String, surflist As List(Of HybridShapeExtract)) As processStatic

        '2015.7.25同时修复点不在面上及点不在线上的问题


        Dim ppstatic As New processStatic
        Dim documents1 As Documents

        Dim partDocument1 As PartDocument
        documents1 = CATIA.Documents
        partDocument1 = documents1.Item(partfilename)
        Dim part1 As Part
        part1 = partDocument1.Part
        partDocument1.Activate()
        Dim SPAWorkb As INFITF.Workbench
        Dim Measurement
        Dim Coords(2) As Object
        Dim source, destination, sourceinter
        source = CATIA.ActiveDocument.Selection
        sourceinter = CATIA.ActiveDocument.Selection
        destination = CATIA.ActiveDocument.Selection

        Dim surfcount = surflist.Count()
   



        Dim reference2 As Reference
        SPAWorkb = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")
        ' Dim tempgeo = hb.HybridBodies.Add()

        part1.InWorkObject = hb
        For Each pp As CENPoint In Points_.Values


            'Get the measurement for the point
            Measurement = SPAWorkb.GetMeasurable(pp.MyPoint)

            'Get the coordinates (Part based) from this point
            Call Measurement.GetPoint(Coords)
            ' reference1 = part1.CreateReferenceFromObject(pp.MyPoint)
            '量取线到法向的距离
            reference2 = part1.CreateReferenceFromObject(pp.MyVector)

            Dim MinimumDistance As Double
            MinimumDistance = Measurement.GetMinimumDistance(reference2)



            Dim surface As HybridShapeExtract
            surface = surflist(0)
            Dim reference1 As Reference
            reference1 = part1.CreateReferenceFromObject(surface)

            '点到面上的距离
            Dim MinimumDistance2 As Double
            MinimumDistance2 = Measurement.GetMinimumDistance(reference1)

            '线到面的距离 
            Dim TheMeasurable As SPATypeLib.Measurable
            TheMeasurable = SPAWorkb.GetMeasurable(reference2)
            Dim MinimumDistance3 As Double
            MinimumDistance3 = TheMeasurable.GetMinimumDistance(reference1)

            Dim noerr As Boolean = False

            If MinimumDistance3 > 0 Then
                If surfcount > 1 Then


                    For si As Integer = 1 To surfcount - 1




                        surface = surflist(si)
                        reference1 = part1.CreateReferenceFromObject(surface)
                        MinimumDistance3 = TheMeasurable.GetMinimumDistance(reference1)
                        If MinimumDistance3 = 0 Then
                            noerr = True
                            Exit For


                        End If

                    Next

                Else

                    '线到面的距离有问题,不再往下进行
                    noerr = False

                End If
            Else

                noerr = True
            End If







            If noerr And (MinimumDistance > 0 Or MinimumDistance2 > 0) Then
                '检测到点不在法向上
                ppstatic.Add(1, "Fix_out_of_vec_surf:" + pp.Frame + " - " + pp.Diam)
                Dim hybridShapeIntersection1 As HybridShapeIntersection
                hybridShapeIntersection1 = part1.HybridShapeFactory.AddNewIntersection(reference2, reference1)
                hybridShapeIntersection1.PointType = 0
                ' hybridShapeIntersection1.ExtendMode = 3
                hb.AppendHybridShape(hybridShapeIntersection1)
                part1.Update()
                ' hb.AppendHybridShape(hybridShapeIntersection1)
                sourceinter.Clear()
                sourceinter.Add(hybridShapeIntersection1)
                sourceinter.Copy()
                destination.Clear()
                destination.Add(hb)
                destination.PasteSpecial("CATPrtResultWithOutLink")

                sourceinter.Clear()
                sourceinter.Add(hybridShapeIntersection1)
                sourceinter.Delete()



                source.Clear()

                source.Add(pp.MyPoint)

                source.Delete()


                part1.Update()

                '重新绑定point
                '不做其他操作了，则无必要
                'pp.MyPoint = hb.HybridShapes.Item(hb.HybridShapes.Count)

            End If


        Next
        '整体删除临时几何图形集
        'sourceinter.Clear()
        'sourceinter.Add(hybridShapeIntersection1)
        'sourceinter.Delete()


        Return ppstatic

    End Function
    Public Function del_dupli(partfilename As String) As processStatic
        Dim ppstatic As New processStatic

        If dupli.Count <> 0 Then


            Dim documents1 As Documents

            Dim partDocument1 As PartDocument
            documents1 = CATIA.Documents
            partDocument1 = documents1.Item(partfilename)
            Dim part1 As Part
            part1 = partDocument1.Part
            partDocument1.Activate()
            '  Dim SPAWorkb As INFITF.Workbench
            ' Dim Measurement
            Dim Coords(2) As Object
            Dim source
            source = CATIA.ActiveDocument.Selection

            source.Clear()

            For Each pp As CENPoint In dupli

                source.Add(pp.MyPoint)
                source.Add(pp.MyVector)
                ppstatic.Add(1, "Fix_dupli_points_vecs:" + pp.Frame + " - " + pp.Diam)
            Next



            source.Delete()


            part1.Update()

        End If
        Return ppstatic
    End Function


    '
    '   Default destructor
    '
    'Protected Overrides Sub Finalize()
    '    RemoveAll()
    'End Sub

    '
    '   Adds a rivet to a collection
    '
    Public Sub Add(ByRef tmppoint As CENPoint)
        If Points_.Keys.Contains(tmppoint.uuid) Then
            dupli.Add(tmppoint)
        Else

            Points_.Add(tmppoint.uuid, tmppoint)
        End If
    End Sub
    Public Sub Merge(tmp As CENPoints)
        For Each tmppoint As CENPoint In tmp.Points_.Values
            Add(tmppoint)


            'Catch
            '    dupli.Add(tmppoint)
            '    ' MsgBox("重复点：" + tmppoint.Frame + "\r\t" + tmppoint.Diam + "\r\t" + tmppoint.PointName)

            'End Try

        Next
        For Each tmppoint As CENPoint In tmp.dupli
            dupli.Add(tmppoint)

        Next



    End Sub
    Public Function toProcessTree() As processTreeList
        Dim ds As New processTreeList
        Dim i As Integer


        For i = 0 To Points_.Count - 1
            Dim tmppoint As CENPoint
            tmppoint = Points_.Values.ElementAt(i)

            ds.Add(tmppoint.Frame, tmppoint.FastenerName, tmppoint.Diam, tmppoint.MyPoint)
            If Not tmppoint.MyVector Is Nothing Then

                ds.Add(tmppoint.Frame, tmppoint.FastenerName, tmppoint.Diam, tmppoint.MyVector)

            End If

        Next

        Return ds

    End Function
    '
    '   Returns a rivet based on the Index that has to be withing the legal bounds
    '
    Public Function Item(ID As String) As CENPoint
        If Points_.Keys.Contains(ID) Then
            Return Points_.Item(ID)

        Else
            Dim idd = ID.Split("_")
            Dim xx = CDec(idd(0))
            Dim yy = CDec(idd(1))
            Dim zz = CDec(idd(2))

            For Each pp As CENPoint In Points_.Values

                If ((pp.xxr - xx) ^ 2 + (pp.yyr - yy) ^ 2 + (pp.zzr - zz) ^ 2) < 25 Then

                    Return pp
                End If

            Next


        End If

        Return Nothing
    End Function



    Public Function Item2(Index As Integer) As CENPoint

        Item2 = Points_.Values.ElementAt(Index - 1)

    End Function

    '
    '   Returns the number of elements in a collection
    '
    Public Function count() As Integer
        count = Points_.Count
    End Function

    '
    '   Removes an element from a collection
    '
    Public Sub Remove(Index As Integer)

        If (Index > 0 And Index <= Points_.Count) Then
            Points_.Remove(Points_.Keys(Index - 1))
        End If

    End Sub
    Public Sub Remove(uuid As String)


        Points_.Remove(uuid)


    End Sub
    Public Sub RemovebyPF(pfname As String)
        Dim kk = Points_.Values.Where(Function(p) p.PFname = pfname)
        If kk.Count > 0 Then
            Points_.Remove(kk.First.uuid)

        End If



    End Sub
    Public Function ItembyPF(pfname As String) As CENPoint
        Dim kk = Points_.Values.Where(Function(p) p.PFname = pfname)
        If kk.Count > 0 Then
            Return kk.First()
        Else
            Return Nothing
        End If




    End Function
    Public Sub outputdb(productname As String, Optional suffix As String = "")
        productname = productname + suffix



                       
        Dim pointlist As New List(Of String)

        pointlist.Add("Create table if not exists " + productname + "(PFname varchar(100) ,FastenerName varchar(100),X double,Y double,Z double,FrameName varchar(100),ProcessType varchar(100),XR INT(20),YR INT(20),ZR INT(20),UUID varchar(100) PRIMARY KEY,STRNO varchar(100),location varchar(100),pointname varchar(100),uuidP varchar(100));")


        Dim strSql22 As New System.Text.StringBuilder()

        strSql22.Append(String.Format("delete from {0};", productname))
        pointlist.Add(strSql22.ToString())


        Dim jj As Integer
        For jj = 1 To count()
            Dim xx As Double
            Dim yy As Double
            Dim zz As Double
            Dim xxr As Integer
            Dim yyr As Integer
            Dim zzr As Integer

            xx = Item2(jj).X
            yy = Item2(jj).Y
            zz = Item2(jj).Z

            xxr = Item2(jj).xxr
            yyr = Item2(jj).yyr
            zzr = Item2(jj).zzr



            Dim uuid As String

            uuid = Item2(jj).uuid()

            Dim strno As Integer
            strno = (yyr ^ 2 + zzr ^ 2)

            ' strno = 0

            strno = yy / Math.Abs(yy) * strno
            Dim framename As String
            Dim location2 As String
            location2 = ""
            framename = Item2(jj).Frame
            If (framename.ToLower.Contains("win")) Then
                location2 = "WIN"
            End If

            Dim strSqlname As New System.Text.StringBuilder()
            strSqlname.Append(String.Format("INSERT INTO {0} (FastenerName,X,Y,Z,FrameName,ProcessType,XR,YR,ZR,UUID,STRNO,location,pointname,PFname,uuidP", productname))
            strSqlname.Append(String.Format(") VALUES ('{0}',{1},{2},{3},'{4}','{5}',{6},{7},{8},'{9}','{10}','{11}','{12}','{13}','{14}')", Item2(jj).FastenerName, xx, yy, zz, framename, Item2(jj).Diam, xxr, yyr, zzr, uuid, strno.ToString, location2, Item2(jj).PointName, Item2(jj).PFname, Item2(jj).uuidP))
            ' Points_.Item2(jj).FastenerName = FastenerName And Points_.Item2(jj).Frame = Framename Then
            pointlist.Add(strSqlname.ToString())

        Next
        DbHelperSQL.ExecuteSqlTran(pointlist)


    End Sub


    Public Sub importfromdb(productname As String, Optional ifupdatefrm As Boolean = False)

        RemoveAll()


        '不带001
        Dim tmpdt As DataTable
        If ifupdatefrm Then
            tmpdt = DbHelperSQL.Query("select UUID,location,strno from " + productname).Tables(0)


            For Each kk As DataRow In tmpdt.Rows
                Add(TVA_Method.LeaftoPoint(framename:=(kk("location").ToString() + "_" + kk("strno").ToString()), uuid:=kk("UUID").ToString()))


            Next
        Else
            tmpdt = DbHelperSQL.Query("select X,Y,Z,FastenerName,ProcessType,FrameName,pointname,UUID,PFname from " + productname).Tables(0)


            For Each kk As DataRow In tmpdt.Rows
                Add(TVA_Method.LeaftoPoint(kk(0).ToString(), kk(1).ToString(), kk(2).ToString(), kk(3).ToString(), kk(4).ToString(), kk(5).ToString(), kk(6).ToString(), kk("UUID").ToString(), kk("PFname").ToString()))


            Next
        End If




    End Sub
    Public Function compare(targetpt As CENPoints) As CENPoints()



        Dim cmresult() As CENPoints = {New CENPoints, New CENPoints}
        ' ReDim cmresult(0 To 1)

        Dim i As Integer

        For i = 1 To Points_.Count
            Dim sspt As New CENPoint

            sspt = Points_.Item(i)
            Dim suuid As String
            suuid = sspt.uuid

            If targetpt.Points_.Keys.Contains(suuid) Then
                '目标中存在改点
                Dim ttpt As CENPoint

                ttpt = targetpt.Item(suuid)
                If sspt.Diam = ttpt.Diam Then

                Else
                    sspt.Diam = sspt.Diam + ";" + ttpt.Diam

                    cmresult(0).Add(sspt)


                End If
            Else
                sspt.Diam = "Source only;" + sspt.Diam
                cmresult(0).Add(sspt)

            End If
        Next

        '找出删除的点
        For i = 1 To targetpt.count
            Dim sspt As New CENPoint

            sspt = targetpt.Item2(i)
            Dim suuid As String
            suuid = sspt.uuid

            If Points_.Keys.Contains(suuid) Then


            Else
                sspt.Diam = "Target only;" + sspt.Diam
                cmresult(1).Add(sspt)

            End If
        Next


        Return cmresult







    End Function


    Public Function compare(targetsps As CENSP01s) As CENPoints()
        '重构 进行紧固件比较

        Dim targetpt = targetsps.to_points()
        Dim cmresult() As CENPoints = {New CENPoints, New CENPoints}
        ' ReDim cmresult(0 To 1)

        Dim i As Integer

        For i = 1 To Points_.Count


            Dim sspt = Item2(i)
            ' Dim suuid As String
            ' suuid = sspt.uuid



            '快速判断
            'If targetpt.InTheList(suuid) Then
            '    '目标中存在改点
            '    Dim ttpt As CENPoint

            '    ttpt = targetpt.Item(suuid)
            '    If sspt.FastenerName = ttpt.FastenerName Then

            '    Else
            '        wrongstatistic.Add(1, sspt.FastenerName + " - " + ttpt.FastenerName)
            '        ' cmresult(1).Add(sspt)
            '        ' sspt.Diam = sspt.FastenerName + ";" + ttpt.FastenerName
            '        sspt.FastenerName = ttpt.FastenerName
            '        sspt.Diam = sspt.Diam.ToString().Replace(sspt.FastenerName, ttpt.FastenerName)
            '        cmresult(0).Add(sspt)


            '    End If
            'Else


            '    'PVR紧固件组件中没有，代表该紧固件有可能已被删除

            '    cmresult(1).Add(sspt)

            'End If

            Dim findit As Boolean = False
            '对比紧固件采用全面的比对判断
            For j = 1 To targetpt.count
                Dim ttpt = targetpt.Item2(j)
                Dim chazhi = ((sspt.xxr - ttpt.xxr) ^ 2 + (sspt.yyr - ttpt.yyr) ^ 2 + (sspt.zzr - ttpt.zzr) ^ 2)
                If chazhi < 40 Then
                    findit = True
                    If sspt.FastenerName = ttpt.FastenerName Then

                    Else
                        wrongstatistic.Add(1, sspt.FastenerName + " - " + ttpt.FastenerName)
                        ' cmresult(1).Add(sspt)
                        ' sspt.Diam = sspt.FastenerName + ";" + ttpt.FastenerName
                        sspt.FastenerName = ttpt.FastenerName
                        sspt.Diam = sspt.Diam.ToString().Replace(sspt.FastenerName, ttpt.FastenerName)
                        cmresult(0).Add(sspt)


                    End If


                    Exit For
                End If

            Next

            'PVR紧固件组件中没有，代表该紧固件有可能已被删除
            If Not findit Then
                cmresult(1).Add(sspt)
            End If








        Next


        Return cmresult







    End Function
    Public Function compare(productname As String) As DataTable

        Dim targetpoints As New CENPoints

        targetpoints.importfromdb(productname)

        Dim resultpt As CENPoints
        Dim tttpt() As CENPoints = {New CENPoints, New CENPoints}
        tttpt = compare(targetpoints)
        tttpt(0).Merge(tttpt(1))
        resultpt = tttpt(0)


        Dim newtb As New DataTable
        newtb = resultpt.outputtable()
        Return newtb
    End Function

    Public Function outputtable() As DataTable
        Dim newdtb As New DataTable

        newdtb.Columns.Add("ID")
        newdtb.Columns.Add("UUID")
        newdtb.Columns.Add("PFName")
        newdtb.Columns.Add("FrameName")
        newdtb.Columns.Add("FastenerName")
        newdtb.Columns.Add("ProcessType")
        For i = 1 To Points_.Count
            Dim sspt As CENPoint

            sspt = Item2(i)

            Dim newRow = newdtb.NewRow()

            newRow("ID") = i
            newRow("UUID") = sspt.uuid
            newRow("PFName") = sspt.PFname
            newRow("FrameName") = sspt.Frame
            newRow("FastenerName") = sspt.FastenerName
            newRow("ProcessType") = sspt.Diam


            newdtb.Rows.Add(newRow)
        Next

        Return newdtb


    End Function

    '
    '   Removes all elements from a collection
    '
    Public Sub RemoveAll()

        Points_.Clear()

    End Sub

    '
    '   Performs a check whether or not a rivet is in the list or not.
    '   Search is performed by the name of a rivet
    '
    Public Function InTheList(ByRef name As String) As Boolean

        Return Points_.Keys.Contains(name)

    End Function


End Class

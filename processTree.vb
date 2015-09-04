

Option Explicit On
Imports ProductStructureTypeLib
Imports MECMOD
'Imports PARTITF
Imports HybridShapeTypeLib

Imports System.Collections.Generic
Imports KnowledgewareTypeLib
Imports INFITF
Imports mysqlsolution

Public Class processTree : Inherits processTreeBase
    Public fastenertree As New Dictionary(Of String, List(Of HybridShape))
    Public fastgeoset As HybridBody
    '  Private bindingpart As Part
    ' Public PointsFatherProduct As Product
    Private CATIA As Application






    Public Sub New(ByRef myGeoSet As HybridBody)
        'ppat不赋值的话，表示不需要录入组件坐标
        CATIA = myGeoSet.Application
        fastgeoset = myGeoSet
        framename = myGeoSet.Parent.Name
        '   bindingpart = ppat
        fasternername = myGeoSet.Name
        '  addFastername(myGeoSet.Name)
        '   PointsFatherProduct = Product1

        'sectree.Add("Target type - Final")
        'sectree.Add("Target type - Pilot Holes")
        'sectree.Add("Target type - Fast Tack")

        inifstree()



        Add(myGeoSet)
        ' fastenertree.Add()
        ' Dim partDocument1 As PartDocument
        ' partDocument1 = myGeoSet..Application.ActiveDocument


    End Sub
    Sub inifstree()
        For Each rr As String In firsttree

            If Strings.InStr(rr, "RESYNCING") Then
                For Each kk As String In sectree

                    fastenertree.Add(rr + " - " + kk, New List(Of HybridShape))

                Next
            Else

                fastenertree.Add(rr, New List(Of HybridShape))
            End If


        Next
    End Sub
    Public Sub New(frname As String, fstname As String)

        framename = frname

        fasternername = fstname

        '同样初始化树结构
        inifstree()


    End Sub
    Public Sub setcolor()
        Dim selection1 As Selection
        '  Dim selection_fstcolor As Selection

        selection1 = CATIA.ActiveDocument.Selection
        '  selection_fstcolor = CATIA.ActiveDocument.Selection
        selection1.Clear()
        '   selection_fstcolor.Clear()
        Dim treevalue As List(Of List(Of HybridShape)) = fastenertree.Values.ToList

        ' RESYNCING ONLY final
        For Each sh As HybridShape In treevalue(0)
            selection1.Add(sh)

        Next
        selection1.VisProperties.SetRealColor(255, 0, 0, 1)
        'selection1.VisProperties.SetSymbolType(6)
        set_fstsh(selection1)
        selection1.Clear()
        ' RESYNCING ONLY pilot hole
        For Each sh As HybridShape In treevalue(1)
            selection1.Add(sh)

        Next
        selection1.VisProperties.SetRealColor(255, 0, 0, 1)
        '同心圆
        selection1.VisProperties.SetSymbolType(4)
        selection1.Clear()

        ' RESYNCING ONLY fastack
        For Each sh As HybridShape In treevalue(2)
            selection1.Add(sh)

        Next
        selection1.VisProperties.SetRealColor(255, 0, 0, 1)
        '星号
        selection1.VisProperties.SetSymbolType(7)
        selection1.Clear()

        'FASTENER INSTALLED BY AUTOMATED FASTENING

        For Each sh As HybridShape In treevalue(3)
            selection1.Add(sh)
            '   selection_fstcolor.Add(sh)
        Next
        '设为十字
        set_fstcl(selection1)
        selection1.VisProperties.SetSymbolType(2)
        selection1.Clear()



        'RESYNCING AND DRILL/install ONLY BY AUTOMATED FASTENING-final


        For Each sh As HybridShape In treevalue(4)
            selection1.Add(sh)

        Next
        For Each sh As HybridShape In treevalue(7)
            selection1.Add(sh)

        Next
        '设为实心方块
        'selection1.VisProperties.SetSymbolType(6)
        set_fstsh(selection1)



        '绿色
        selection1.VisProperties.SetRealColor(0, 128, 64, 1)
        selection1.Clear()

        'RESYNCING AND DRILL/install ONLY BY AUTOMATED FASTENING-pilot hole

        For Each sh As HybridShape In treevalue(5)
            selection1.Add(sh)

        Next
        'For Each sh As HybridShape In treevalue(8)
        '    selection1.Add(sh)

        'Next
        '同心圆
        selection1.VisProperties.SetSymbolType(4)
        '绿色
        selection1.VisProperties.SetRealColor(0, 128, 64, 1)
        selection1.Clear()

        For Each sh As HybridShape In treevalue(8)
            selection1.Add(sh)

        Next
        '2015.8.28 RESYNCING AND DRILL ONLY BY AUTOMATED FASTENING-pilot hole更改为实心点
        selection1.VisProperties.SetSymbolType(5)
        '绿色
        selection1.VisProperties.SetRealColor(0, 128, 64, 1)
        selection1.Clear()






        'RESYNCING AND DRILL/install ONLY BY AUTOMATED FASTENING-fastack

        For Each sh As HybridShape In treevalue(6)
            selection1.Add(sh)

        Next
        For Each sh As HybridShape In treevalue(9)
            selection1.Add(sh)

        Next
        '星号
        selection1.VisProperties.SetSymbolType(7)
        '绿色
        selection1.VisProperties.SetRealColor(0, 128, 64, 1)
        selection1.Clear()


        'DRILL ONLY BY AUTOMATED FASTENING

        For Each sh As HybridShape In treevalue(10)

            selection1.Add(sh)
        Next

        set_fstcl(selection1)
        '设为实心圆
        selection1.VisProperties.SetSymbolType(5)
        selection1.Clear()


        'after/before/temp
        'DRILL BEFOR

        For Each sh As HybridShape In treevalue(11)

            selection1.Add(sh)
        Next
        'DRILL BEFOR

        set_HANDfstcl(selection1)
        'selection1.VisProperties.SetRealColor(0, 0, 0, 1)
        '符号为同心圆
        selection1.VisProperties.SetSymbolType(4)
        selection1.Clear()

        'INSTALL BEFORE
        For Each sh As HybridShape In treevalue(12)

            selection1.Add(sh)
        Next

        set_fstsh(selection1)
        set_HANDfstcl(selection1)
        selection1.Clear()

        'INSTALL AFTER/temp

        For Each sh As HybridShape In treevalue(13)

            selection1.Add(sh)
        Next

        For Each sh As HybridShape In treevalue(14)

            selection1.Add(sh)
        Next

        'after点为黑色
        selection1.VisProperties.SetRealColor(0, 0, 0, 1)
        '符号为X
        selection1.VisProperties.SetSymbolType(1)
        selection1.Clear()







    End Sub
    Sub set_HANDfstcl(ByRef selection_fstcolor As Selection)
        '改变铆接点位颜色

        If Strings.InStr(fasternername, "5-") Then

            If Strings.InStr(fasternername, "5020AD") Or Strings.InStr(fasternername, "6002AG") Then
                '紫色
                selection_fstcolor.VisProperties.SetRealColor(255, 0, 255, 1)
            Else
                '橙色
                selection_fstcolor.VisProperties.SetRealColor(255, 165, 0, 1)
            End If

        Else
            '深棕色

            If Strings.InStr(fasternername, "5020AD") Or Strings.InStr(fasternername, "6002AG") Then

                selection_fstcolor.VisProperties.SetRealColor(64, 32, 32, 1)
            Else


                selection_fstcolor.VisProperties.SetRealColor(128, 64, 0, 1)
            End If
        End If

        ' selection_fstcolor.Clear()
    End Sub
    Sub set_fstsh(ByRef selection_fstcolor As Selection)
        If Strings.InStr(fasternername, "B020600") Then
            '设为大圆
            selection_fstcolor.VisProperties.SetSymbolType(3)
        Else

            '设为十字
            selection_fstcolor.VisProperties.SetSymbolType(2)
        End If


    End Sub

    Sub set_fstcl(ByRef selection_fstcolor As Selection)
        '改变铆接点位颜色

        If Strings.InStr(fasternername, "5-") Then

            If Strings.InStr(fasternername, "5020AD") Or Strings.InStr(fasternername, "6002AG") Then
                '蓝色
                selection_fstcolor.VisProperties.SetRealColor(0, 0, 255, 1)
            Else
                selection_fstcolor.VisProperties.SetRealColor(0, 255, 255, 1)
            End If

        Else
            '棕色

            If Strings.InStr(fasternername, "5020AD") Or Strings.InStr(fasternername, "6002AG") Then

                selection_fstcolor.VisProperties.SetRealColor(128, 64, 64, 1)
            Else
                selection_fstcolor.VisProperties.SetRealColor(255, 128, 0, 1)
            End If
        End If
        If Strings.InStr(fasternername, "B020600") Then
            '设为大圆
            selection_fstcolor.VisProperties.SetSymbolType(3)
        Else

            '设为十字
            selection_fstcolor.VisProperties.SetSymbolType(2)
        End If
        selection_fstcolor.Clear()
    End Sub

    Public Sub Add(processtp As String, ByRef point As HybridShape)
        If fastenertree.Keys.Contains(processtp) Then

        Else

            fastenertree.Add(processtp, New List(Of HybridShape))

        End If

        fastenertree(processtp).Add(point)
    End Sub
    Public Sub Add(ByRef myGeoSet As HybridBody)

        'Dim framename As String
        'framename = myGeoSet.Parent.Name
        'Dim fasterner As String
        'fasterner = myGeoSet.Name

        For i = 1 To myGeoSet.HybridBodies.Count

            Dim MyNewGeoSet As HybridBody

            MyNewGeoSet = myGeoSet.HybridBodies.Item(i)
            Dim processty As String
            processty = MyNewGeoSet.Name.Replace(fasternername + " - ", "")
            If Strings.InStr(MyNewGeoSet.Name, "RESYNCING") Then

                For j = 1 To MyNewGeoSet.HybridBodies.Count
                    Dim MyNew2GeoSet As HybridBody
                    MyNew2GeoSet = MyNewGeoSet.HybridBodies.Item(j)

                    Dim tmplist2 As New List(Of HybridShape)

                    Dim processty2 As String
                    processty2 = MyNew2GeoSet.Name
                    For l = 1 To MyNew2GeoSet.HybridShapes.Count

                        fastenertree(processty + " - " + processty2).Add(MyNew2GeoSet.HybridShapes.Item(l))





                    Next

                    '  fastenertree(processty + ";" + processty2).AddRange(tmplist2)


                Next




            Else
                ' Dim tmplist As New List(Of HybridShape)
                For k = 1 To MyNewGeoSet.HybridShapes.Count

                    fastenertree(processty).Add(MyNewGeoSet.HybridShapes.Item(k))

                Next
                ' fastenertree(processty).AddRange(tmplist)
            End If
        Next





    End Sub
    Public Function getfast(processty As String) As List(Of HybridShape)



        Return fastenertree(processty)

    End Function

    Public Function output_shapes() As List(Of List(Of HybridShape))
        Return fastenertree.Values.ToList()
    End Function
    Public Function output_keys() As List(Of String)
        Return fastenertree.Keys.ToList()
    End Function
    Public Function prodic() As Dictionary(Of String, List(Of HybridShape))
        Return fastenertree
    End Function

    Public Sub merge(ttt As processTree)
        Dim targetdic As New Dictionary(Of String, List(Of HybridShape))
        targetdic = ttt.prodic

        For Each kkk In targetdic
            If fastenertree.Keys.Contains(kkk.Key) Then
                fastenertree(kkk.Key).AddRange(kkk.Value)
            End If

        Next


    End Sub


    Public Sub output(sourcepartname As String, targetpartname As String, ByRef myGeoSet As HybridBody, Optional para As Integer = 0)
        CATIA = myGeoSet.Application
        Dim documents1 As Documents

        Dim partDocument1 As PartDocument
        documents1 = CATIA.Documents
        partDocument1 = documents1.Item(targetpartname)

        'Dim part1 As Part
        'part1 = partDocument1.Part
        Dim fastgeo As HybridBody
        '到达目标几何图形集紧固件的目录
        fastgeo = getnextGeo(getnextGeo(myGeoSet, framename), fasternername)

        '建立要操作的几何图形集字典
        Dim targethb As New Dictionary(Of String, HybridBody)

        If para = 0 Then

            '生成整个树结构作为目标操作集
            For Each rr As String In firsttree
                Dim opgeo As HybridBody
                opgeo = getnextGeo(fastgeo, fasternername + " - " + rr, targetpartname)
                Dim processty As String
                processty = rr
                If Strings.InStr(rr, "RESYNCING") Then
                    For Each kk As String In sectree
                        processty = rr + " - " + kk
                        targethb.Add(processty, getnextGeo(opgeo, kk))

                    Next
                Else

                    targethb.Add(processty, opgeo)
                End If


            Next
        Else

            '不生成
            For Each pptp In fastenertree.Keys
                Dim opgeo As HybridBody
                opgeo = getnextGeo(fastgeo, pptp)
                targethb.Add(pptp, opgeo)

            Next

        End If

        For Each outputls In targethb

            copypaste(sourcepartname, targetpartname, outputls.Value, outputls.Key)

        Next

        ' part1.Update()
    End Sub


    Public Function getnextGeo(ByRef myGeoSet As HybridBody, ii As String, ByRef targetpartname As String) As HybridBody

        ' part1.Activate(part1)

        '该重载是创建根目录的参数
        Dim documents1 As Documents

        Dim partDocument1 As PartDocument
        documents1 = CATIA.Documents
        partDocument1 = documents1.Item(targetpartname)

        Dim part1 As Part
        part1 = partDocument1.Part
        Dim framenamegeo As HybridBody
        Dim selection1, selection2
        selection1 = CATIA.ActiveDocument.Selection
        selection2 = CATIA.ActiveDocument.Selection
        If TVA_Method.ifGeoExist(myGeoSet, ii) Then
            framenamegeo = myGeoSet.HybridBodies.GetItem(ii)
        Else

            framenamegeo = myGeoSet.HybridBodies.Add()
            framenamegeo.Name = ii



            'Add parameters
            Dim parameters1 As Parameters
            parameters1 = part1.Parameters
            Dim strParam1 As StrParam
            strParam1 = parameters1.CreateString(framename & " \ " & fasternername & " \ " & framenamegeo.Name, "")
            strParam1.Rename("Fastener_NO.1")
            strParam1.Value = fasternername




            'Activate the name of the correct window
            partDocument1.Activate()




            'Cut and paste into the geoset
            selection1 = CATIA.ActiveDocument.Selection
            selection2 = CATIA.ActiveDocument.Selection
            selection2.Clear()

            selection1.Clear()
            selection1.Add(strParam1)
            selection1.Cut()
            '  selection1.Copy()
            selection2.Add(framenamegeo)
            selection2.Paste()

            part1.Update()
            part1.UpdateObject(framenamegeo)




        End If

        Return framenamegeo
    End Function


    Public Function MoveType(sourcepartname As String, sourceType As String, targetType As String) As processStatic
        Dim rpt = New processStatic
        Dim dd = Strings.Split(targetType, (" - "))
        Dim opgeo = getnextGeo(fastgeoset, fasternername + " - " + dd(0))
        If dd.Count > 1 Then

            opgeo = getnextGeo(opgeo, dd(1) + " - " + dd(2))
        End If


        rpt.Add(copypaste(sourcepartname, sourcepartname, opgeo, sourceType))

        'delete original points

        delprocTp(sourceType)
        Return rpt
    End Function





    Function copypaste(sourcepartname As String, targetpartname As String, ByRef opgeo As HybridBody, processty As String) As processStatic
        'copy all processty shapes into opgeo
        Dim aa = New processStatic
        Dim selection1, selection2
        selection1 = CATIA.ActiveDocument.Selection
        selection2 = CATIA.ActiveDocument.Selection
        selection1.Clear()
        selection2.Clear()

        'Now copy the appropiate points and vectors
        Dim documents1 As Documents
        documents1 = CATIA.Documents

        Dim partDocument1

        partDocument1 = documents1.Item(sourcepartname)

        Dim partDocument2

        partDocument2 = documents1.Item(targetpartname)
        'Activate the name of the correct window
        partDocument1.Activate()



        selection1.Clear()
        selection1 = CATIA.ActiveDocument.Selection
        Dim oplist As New List(Of HybridShape)()
        oplist = getfast(processty)
        Dim ptcount = oplist.Count / 2
        If ptcount > 0 Then

            aa.Add(ptcount, processty + "- TO -" + opgeo.Name)
            For Each hs As HybridShape In oplist

                selection1.Add(hs)

            Next
            '  partDocument1.Activate()
            selection1.Copy()
            partDocument2.Activate()


            selection2 = CATIA.ActiveDocument.Selection
            selection2.Clear()
            selection2.Add(opgeo)
            selection2.PasteSpecial("CATPrtResultWithOutLink")
            selection1.Clear()
        End If
        Return aa
    End Function
    Sub delprocTp(procseetype As String)
        Dim selection1 As Selection
        selection1 = CATIA.ActiveDocument.Selection
        selection1.Clear()
        Dim oplist = fastenertree(procseetype)

        For Each hs As HybridShape In oplist

            selection1.Add(hs)

        Next
        If oplist.Count <> 0 Then
            selection1.Delete()

        End If


        selection1.Clear()
    End Sub
    Sub del(partname As String)


        Dim selection1 As Selection

        'Now copy the appropiate points and vectors
        Dim documents1 As Documents
        documents1 = CATIA.Documents

        Dim partDocument1

        partDocument1 = documents1.Item(partname)




        'Activate the name of the correct window
        partDocument1.Activate()


        selection1 = CATIA.ActiveDocument.Selection
        selection1.Clear()

        ' Dim oplist As New List(Of HybridShape)()
        'oplist = fastenertree.Values

        For Each oplist As List(Of HybridShape) In fastenertree.Values



            If oplist.Count <> 0 Then


                For Each hs As HybridShape In oplist

                    selection1.Add(hs)

                Next

            End If
        Next


        selection1.Delete()

        selection1.Clear()



    End Sub

    Public Function getnextGeo(ByRef myGeoSet As HybridBody, ii As String) As HybridBody
        Dim framenamegeo As HybridBody
        If TVA_Method.ifGeoExist(myGeoSet, ii) Then
            framenamegeo = myGeoSet.HybridBodies.GetItem(ii)
        Else

            framenamegeo = myGeoSet.HybridBodies.Add()
            framenamegeo.Name = ii




        End If

        Return framenamegeo
    End Function

















End Class

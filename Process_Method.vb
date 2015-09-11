Imports INFITF
Imports PPR
Imports FWOlpProcFeat
Imports DNBIgpTagPath
Imports SPATypeLib
Imports ProductStructureTypeLib
Imports mysqlsolution
Imports CENFasTIP
Imports System.Text.RegularExpressions
Imports KnowledgewareTypeLib
Imports MECMOD
Imports FWFasTProduct
Imports FWOlpBase

Public Class Process_Method

    Private CATIA As Application
    Private filename As String
    Private pprdoc As PPRDocument
    Private PCFeatContainer As OlpProcFeatContainer
    Private objspa As SPAWorkbench
    Private objProd As Product
    Private TVA_ As TVA_Method
    Private pfPoints_ As CENPoints

    Public Sub New(filepath As String)


        Dim cc = System.Diagnostics.Process.GetProcessesByName("CATSTART")
        If (cc.Count() = 0) Then



            System.Diagnostics.Process.Start("C:\OPT\DS\DELMIA\2104_64\win_b64\code\bin\CATSTART.exe", " -run ""CNEXT.exe"" -env FASTTRIM_FASTTIP_FASTCURVE_FASTSURF_V3R13 -direnv ""C:\OPT\CENIT\FAST\R14\SP4\2104_64\CATEnv"" -nowindow -object """ + filepath + """")

            Threading.Thread.Sleep(80000)

        Else

            Dim username = localMethod.GetProcessUserName(cc.First().Id)
            If (username = System.Environment.UserName) Then

                ' CATIA = GetObject(, "CATIA.Application")



                '  System.Diagnostics.Process.Start(filepath)




            Else

                System.Diagnostics.Process.Start("C:\OPT\DS\DELMIA\2104_64\win_b64\code\bin\CATSTART.exe", " -run ""CNEXT.exe"" -env FASTTRIM_FASTTIP_FASTCURVE_FASTSURF_V3R13 -direnv ""C:\OPT\CENIT\FAST\R14\SP4\2104_64\CATEnv"" -nowindow -object """ + filepath + """")
                Threading.Thread.Sleep(80000)
            End If

        End If












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


        '新的类库永远开启新的CATIA进程
        filename = filepath.Split("\\").Last()

        '  CATIA = CreateObject("CATIA.Application")
        '  CATIA.Visible = True
        '    CATIA.DisplayFileAlerts = False


        Dim documents1 As Documents
        documents1 = CATIA.Documents
        '  Dim partDocument1 As PartDocument
        'On Error Resume Next
        Dim catcherr As Integer
        catcherr = 0

        Dim newfile As Boolean = True
        For Each dd As Document In documents1

            If dd.FullName.ToUpper() = filepath.ToUpper() Then
                dd.Activate()
                pprdoc = CATIA.ActiveDocument.PPRDocument

                newfile = False
                Exit For
            End If


        Next



        If newfile Then
            Try

                pprdoc = documents1.Open(filepath).PPRDocument


            Catch ex As Exception
                catcherr = 1
                ' partDocument1 = documents1.GetItem(filepath)
            End Try
        End If



        If (catcherr = 0) Then
            ' TVAPart = partDocument1.Part





        End If

    End Sub

    Public Sub iniProc()

        If Not PCFeatContainer Is Nothing Then

            Exit Sub
        End If


        Dim products As PPRProducts = pprdoc.Products
        '  CATIA = part1.Application
        '   Part = partDocument1.Part
        '    filename = Part.Name
        Dim str2
        pprdoc.Activate()
        Dim oRobotTask As RobotTask = Nothing
        Try
            str2 = "CENOlpProcFeatContainer"
            PCFeatContainer = DirectCast(CATIA.ActiveDocument.GetItem(str2), OlpProcFeatContainer)
        Catch exception7 As Exception
            '  ProjectData.SetProjectError(exception7)
            Dim exception3 As Exception = exception7
            ' Dim obj4 = 1
            Dim factory As RobotTaskFactory = pprdoc.Resources.Item(1).GetTechnologicalObject("RobotTaskFactory")

            str2 = "Placeholder"
            factory.CreateRobotTask(str2, oRobotTask)
            str2 = "CENOlpProcFeatContainer"
            PCFeatContainer = DirectCast(pprdoc.GetItem(str2), OlpProcFeatContainer)

        End Try

        'SPA工作台
        str2 = "SPAWorkbench"

        objspa = DirectCast(pprdoc.GetWorkbench(str2), SPAWorkbench)



        If (objProd Is Nothing) Then

            products.Item(1).ApplyWorkMode(CatWorkModeType.DESIGN_MODE)
            Try

                objProd = products.Item(1)
            Catch exception8 As Exception

                Dim exception4 As Exception = exception8
                Interaction.MsgBox("Please Open a Product in the Product list", MsgBoxStyle.Critical, "No Product Found")

            End Try
        Else
            '进入设计模式
            objProd.ApplyWorkMode(CatWorkModeType.DESIGN_MODE)
        End If
        ' 定义选择的产品()
        'If (objProdSel Is Nothing) Then
        '    objProdSel = objProd
        'Else

        'Dim procFeats As OlpProcFeats = PCFeatContainer.ProcFeats

        'For Each pp As OlpProcFeat In procFeats
        '    Dim ioTag As Object() = New Object(12 - 1) {}
        '    pp.GetTag(ioTag)

        '    Dim obj2 As EnumParam = pp.EventParameter("E_FASTENER_PROCFEAT", "Op_Type")

        '    '  obj2 = DirectCast(obj2, EnumParam)
        '    '  obj2 = DirectCast(obj2, StrParam)
        '    Dim cc = obj2.ValueEnum

        '    Dim objArray2 = obj2.Name





        '    'Dim comps As Object() = New Object(3 - 1) {}
        '    'Dim objArray6 As Object() = New Object(3 - 1) {}

        '    'objArray6(0) = ioTag(6)

        '    'objArray6(1) = ioTag(7)

        '    'objArray6(2) = ioTag(8)
        '    'comps(0) = ioTag(9)
        '    'comps(1) = ioTag(10)
        '    'comps(2) = ioTag(11)
        '    'Dim item As LayerAnalysis = CATIA.GetItem("CENLayerAnalysis")

        '    'item.UseShowOnly = False
        '    'item.Direction = objArray6
        '    'item.Point = comps

        '    'Dim List = item.Run

        '    'Dim count As Integer = List.Count
        '    'Dim num = 1
        '    'Do While (num <= count)

        '    '    Dim instance = List.Item(num)

        '    '    instance.GetItem("Thickness")

        '    'Loop

        'Next

        'Dim item As LayerAnalysis = CATIA.GetItem("CENLayerAnalysis")


        'Private Function GetStatus(ByVal mystatus As Short) As String
        '    If (mystatus <> 0) Then
        '        If (mystatus = 1) Then
        '            Return "Modified"
        '        End If
        '        If (mystatus = 4) Then
        '            Return "Moved"
        '        End If
        '        If (mystatus = 5) Then
        '            Return "New"
        '        End If
        '    End If
        '    Return "UnModified"
        'End Function
        'Dim aa = From pp As OlpProcFeat In procFeats
        '            Where pp.



    End Sub



    Public Sub GET_OPNAMES()

        'MsgBox ("Sub CENVB_GET_OPNAMES called ")

        'This subroutine gets called during each operation selection
        'It parses through all the robot tasks and sets 3 header variables:
        'Robot Task Name, Current Operation Name, and Previous Operation Name

        Dim aDoc As Object
        Dim PrevFastAct As FASTActivity
        Dim oRobotResource
        Dim RobotTaskFactory
        Dim RtName
        Dim OpNameCurrent
        Dim OpNamePrev
        Dim RobotTasks(300)
        Dim done As Boolean
        Dim PrevName
        PrevName = ""
        Dim FastAct
        Dim MyFastAct As FASTActivity
        Dim MyRobotTaskName

        Dim oProcess As PPRActivity

        'Get a handle to the current process file
        On Error Resume Next
        aDoc = CATIA.ActiveDocument.PPRDocument
        oProcess = aDoc.Processes.Item(1)

        If Err.Number <> 0 Then
            aDoc = CATIA.ActiveDocument
            oProcess = aDoc.Processes.Item(1)
        End If

        On Error GoTo 0

        '**GET HANDLE TO ROBOT
        For index = 1 To aDoc.Resources.Count
            'MsgBox (aDoc.Resources.Item(index).PartNumber)
            If InStr(aDoc.Resources.Item(index).PartNumber, "G2000_SAC") Then
                oRobotResource = aDoc.Resources.Item(index)
                RobotTaskFactory = oRobotResource.GetTechnologicalObject("RobotTaskFactory")
                Exit For
            End If
        Next

        'Get all the robot tasks
        RobotTaskFactory.GetAllRobotTasks(RobotTasks)

        Dim index2 As Integer
        'go through each robot task
        For index2 = 0 To UBound(RobotTasks)
            On Error Resume Next
            'Go into each operation
            For Each oOpActivity In RobotTasks(index2).ChildrenActivities
                If oOpActivity.Type = "Operation" Then
                    'go through each FASTip Activity
                    For Each FastAct In oOpActivity.ChildrenActivities
                        If FastAct.Type = "CENFasTIPIgpActivity" Then
                            'Cast as FASTActivity so you can get header variables
                            MyFastAct = FastAct
                            OpNamePrev = MyFastAct.GetHeaderParameter("H_OPNAME_PREV$")
                            OpNameCurrent = MyFastAct.GetHeaderParameter("H_OPNAME_CURR$")
                            RtName = MyFastAct.GetHeaderParameter("H_RTNAME$")
                            MsgBox(MyFastAct.Geosets.Item(1).GeoElements.Item(1).LinkedElement.Name)
                            'Robot Task Name (2 levels up)
                            MyRobotTaskName = FastAct.Parent.Parent.Name
                            'Set the header variables
                            'RtName.ValuateFromString(MyRobotTaskName)
                            'OpNameCurrent.ValuateFromString(FastAct.Name)
                            'OpNamePrev.ValuateFromString(PrevName)
                            'Store for the next operation
                            PrevName = FastAct.Name
                            PrevFastAct = FastAct
                        End If
                    Next
                End If
            Next
            'Clear out the previous name for the next robot task
            PrevName = ""
            'If there are no more robot tasks then exit loop
            If RobotTasks(index2 + 1) Is Nothing Then Exit For
        Next

    End Sub

    Public Function GET_TASKNAMES() As List(Of String)

        'MsgBox ("Sub CENVB_GET_OPNAMES called ")

        'This subroutine gets called during each operation selection
        'It parses through all the robot tasks and sets 3 header variables:
        'Robot Task Name, Current Operation Name, and Previous Operation Name

        Dim aDoc As Object
        Dim PrevFastAct As FASTActivity
        Dim oRobotResource
        Dim RobotTaskFactory
        Dim RtName
        Dim OpNameCurrent
        Dim OpNamePrev
        Dim RobotTasks(300)
        Dim done As Boolean
        Dim PrevName
        PrevName = ""
        Dim FastAct
        Dim MyFastAct As FASTActivity
        Dim MyRobotTaskName

        Dim oProcess As PPRActivity

        'Get a handle to the current process file
        On Error Resume Next

        oProcess = pprdoc.Processes.Item(1)

        If Err.Number <> 0 Then

            oProcess = pprdoc.Processes.Item(1)
        End If

        On Error GoTo 0

        '**GET HANDLE TO ROBOT
        For index = 1 To pprdoc.Resources.Count
            'MsgBox (aDoc.Resources.Item(index).PartNumber)
            If InStr(pprdoc.Resources.Item(index).PartNumber, "G2000_SAC") Then
                oRobotResource = pprdoc.Resources.Item(index)
                RobotTaskFactory = oRobotResource.GetTechnologicalObject("RobotTaskFactory")
                Exit For
            End If
        Next

        'Get all the robot tasks
        RobotTaskFactory.GetAllRobotTasks(RobotTasks)
        Dim list As New List(Of String)
        Dim index2 As Integer
        'go through each robot task
        For index2 = 0 To UBound(RobotTasks)
            Dim task = RobotTasks(index2)
            list.Add(task.name)
            If RobotTasks(index2 + 1) Is Nothing Then

                Exit For
            End If


        Next
        Return list
    End Function
    Public Function GET_PRODUCTSNAMES() As List(Of String)


        Return TVA_Method.GetProductList(objProd)


    End Function


    Public Sub iniTVA(TVApath As String)
        If Not TVA_ Is Nothing Then

            Exit Sub
        End If


        Dim documents1 As Documents
        documents1 = CATIA.Documents
        Dim partDocument1 As PartDocument = Nothing

        Dim part1 As Part
        'On Error Resume Next
        '  Dim catcherr As Integer
        ' catcherr = 0

        Dim newfile As Boolean = True
        For Each dd As Document In documents1

            If dd.FullName.ToUpper() = TVApath.ToUpper() Then
                partDocument1 = dd

                newfile = False

                Exit For
            End If


        Next



        If newfile Then

            Try

                partDocument1 = documents1.Open(TVApath)

                Threading.Thread.Sleep(1000)

            Catch ex As Exception

                MsgBox("打开TVA出错" + ex.Message)
                Exit Sub
                ' partDocument1 = documents1.GetItem(filepath)
            End Try



        End If
        part1 = partDocument1.Part


        Dim pointsfather As Product = Nothing
        TVA_Method.GetFatherProduct(objProd, partDocument1, pointsfather)

        TVA_ = New TVA_Method(part1, pointsfather)

        TVA_.TopProduct = objProd

    End Sub

    Public Function PFvsTVA() As CENPoints()

        Dim rt() As CENPoints

        ReDim rt(0 To 1)

        'Define the OP Type

        Dim opDic As New Dictionary(Of String, String)
        opDic.Add("Fastening", "FASTENER INSTALLED BY AUTOMATED FASTENING")

        opDic.Add("Resync_Only", "RESYNCING ONLY BY AUTOMATED FASTENING")
        opDic.Add("Resync_Drill_Only", "RESYNCING AND DRILL ONLY BY AUTOMATED FASTENING")

        opDic.Add("Resync_Fastening", "RESYNCING AND FASTENER INSTALLED BY AUTOMATED FASTENING")

        opDic.Add("Drill_Only", "DRILL ONLY BY AUTOMATED FASTENING")

        opDic.Add("Before", "FASTENER INSTALLED BEFORE AUTOMATED FASTENING")
        opDic.Add("After", "FASTENER INSTALLED AFTER AUTOMATED FASTENING")


        If Not TVA_ Is Nothing Then

            If PCFeatContainer Is Nothing Then
                iniProc()
            End If

            TVA_.coordswitch = True
            TVA_.ifvec = False


            Dim tmppt = TVA_.TVAPointsnoVic.clone

            Dim TVAonly As New CENPoints
            TVAonly = tmppt.clone

            Dim Proconly As New CENPoints



            Dim procFeats As OlpProcFeats = PCFeatContainer.ProcFeats

            For Each pp As OlpProcFeat In procFeats
                '    Dim ioTag As Object() = New Object(12 - 1) {}
                '  pp.GetTagInContext(ioTag)






                Dim obj2 As StrParam = pp.EventParameter("E_FASTENER_PROCFEAT", "Unique_Point_ID$")
                Dim cc = obj2.Value.Split("_")


                'Dim xxr As Integer = Math.Round(CDec(cc(0).Remove(0, 2)), 0)
                'Dim yyr As Integer = Math.Round(CDec(cc(1).Remove(0, 2)), 0)
                'Dim zzr As Integer = Math.Round(CDec(cc(2).Remove(0, 2)), 0)
                Dim uuid_ = cc(0).Remove(0, 2) + "_" + cc(1).Remove(0, 2) + "_" + cc(2).Remove(0, 2)


                tmppt.Item(uuid_).PFname = pp.Name

                Dim processinfo = pp.EventParameter("E_FASTENER_PROCFEAT", "Op_Type")
                Dim processstr = processinfo.ValueEnum
                '  Dim processarray = Strings.Split(processstr, "_")
                Dim procstr = tmppt.Item(uuid_).Diam.ToString()
                If Not procstr.Contains(opDic(processstr)) Then
                    Proconly.Add(tmppt.Item(uuid_))
                    Exit For

                End If


                TVAonly.Remove(uuid_)


            Next


            rt(0) = TVAonly

            rt(1) = Proconly

            pfPoints_ = tmppt
            Return rt

        Else

            Return Nothing


        End If





    End Function

    Public Function PFvsPath() As CENPoints()

        Dim aDoc As Object

        Dim oRobotResource
        Dim RobotTaskFactory

        Dim RobotTasks(300)

        Dim PrevName
        PrevName = ""
        Dim FastAct
        Dim MyFastAct As FASTActivity


        Dim oProcess As PPRActivity

        'Get a handle to the current process file

        aDoc = CATIA.ActiveDocument.PPRDocument
        oProcess = aDoc.Processes.Item(1)

        If Err.Number <> 0 Then
            aDoc = CATIA.ActiveDocument
            oProcess = aDoc.Processes.Item(1)
        End If


        '**GET HANDLE TO ROBOT
        For index = 1 To aDoc.Resources.Count
            'MsgBox (aDoc.Resources.Item(index).PartNumber)
            If InStr(aDoc.Resources.Item(index).PartNumber, "G2000_SAC") Then
                oRobotResource = aDoc.Resources.Item(index)
                RobotTaskFactory = oRobotResource.GetTechnologicalObject("RobotTaskFactory")
                Exit For
            End If
        Next
        Dim tmppoints = PFpoints.clone
        Dim progedafter = New CENPoints

        'Get all the robot tasks
        RobotTaskFactory.GetAllRobotTasks(RobotTasks)

        Dim index2 As Integer
        'go through each robot task
        For index2 = 0 To UBound(RobotTasks)

            'Go into each operation
            For Each oOpActivity In RobotTasks(index2).ChildrenActivities
                If oOpActivity.Type = "Operation" Then
                    'go through each FASTip Activity
                    For Each FastAct In oOpActivity.ChildrenActivities
                        If FastAct.Type = "CENFasTIPIgpActivity" Then
                            'Cast as FASTActivity so you can get header variables
                            MyFastAct = FastAct
                            For Each gs As OlpGeoset In MyFastAct.Geosets

                                For Each ge As OlpGeoElement In gs.GeoElements

                                    Try
                                        Dim pfname = ge.LinkedElement.Name

                                        Dim ddd = tmppoints.ItembyPF(ge.LinkedElement.Name)
                                        If Not ddd Is Nothing Then
                                            progedafter.Add(ddd)
                                            tmppoints.RemovebyPF(pfname)
                                        End If

                                    Catch ex As Exception
                                        MsgBox("检查" + MyFastAct.Name + ";" + gs.Name + ";" + ge.Name)
                                    End Try

                                Next

                            Next










                        End If
                    Next
                End If
            Next

            If RobotTasks(index2 + 1) Is Nothing Then Exit For
        Next
        'points that havn't been programed

        Dim i As Integer = 0
        Do While i < tmppoints.count

            Dim dd = tmppoints.Points_.Values.ElementAt(i).Diam.ToString()
            If dd.Contains("DRILL ONLY BY") Or dd.Contains("FASTENER INSTALLED BY") Then
                i = i + 1

            Else
                tmppoints.Remove(i + 1)
            End If
        Loop

        'points have been programed,filter the after points
        Dim j As Integer = 0
        Do While j < progedafter.count

            Dim dd = progedafter.Points_.Values.ElementAt(j).Diam.ToString()
            If dd.Contains("DRILL ONLY BY") Or dd.Contains("FASTENER INSTALLED BY") Or dd.Contains("RESYNCING") Then
                progedafter.Remove(j + 1)

            Else
                j = j + 1
            End If
        Loop



        Dim dddd() As CENPoints
        ReDim dddd(0 To 1)
        dddd(0) = tmppoints
        dddd(1) = progedafter

        Return dddd


    End Function
    Public Property PFpoints As CENPoints

        Get

            If pfPoints_ Is Nothing Then

                If Not TVA_ Is Nothing Then

                    If PCFeatContainer Is Nothing Then
                        iniProc()
                    End If

                    TVA_.coordswitch = True
                    TVA_.ifvec = False


                    Dim tmppt = TVA_.TVAPointsnoVic

                    Dim procFeats As OlpProcFeats = PCFeatContainer.ProcFeats

                    For Each pp As OlpProcFeat In procFeats
                        '    Dim ioTag As Object() = New Object(12 - 1) {}
                        '  pp.GetTagInContext(ioTag)






                        Dim obj2 As StrParam = pp.EventParameter("E_FASTENER_PROCFEAT", "Unique_Point_ID$")
                        Dim cc = obj2.Value.Split("_")


                        'Dim xxr As Integer = Math.Round(CDec(cc(0).Remove(0, 2)), 0)
                        'Dim yyr As Integer = Math.Round(CDec(cc(1).Remove(0, 2)), 0)
                        'Dim zzr As Integer = Math.Round(CDec(cc(2).Remove(0, 2)), 0)
                        Dim uuid_ = cc(0).Remove(0, 2) + "_" + cc(1).Remove(0, 2) + "_" + cc(2).Remove(0, 2)


                        tmppt.Item(uuid_).PFname = pp.Name

                        ''pp.Nam


                        ''  obj2 = DirectCast(obj2, EnumParam)
                        ''  obj2 = DirectCast(obj2, StrParam)
                        'Dim cc = obj2.ValueEnum

                        'Dim objArray2 = obj2.Name

                        'Dim comps As Object() = New Object(3 - 1) {}
                        'Dim objArray6 As Object() = New Object(3 - 1) {}

                        'objArray6(0) = ioTag(6)

                        'objArray6(1) = ioTag(7)

                        'objArray6(2) = ioTag(8)
                        'comps(0) = ioTag(9)
                        'comps(1) = ioTag(10)
                        'comps(2) = ioTag(11)

                    Next


                    pfPoints_ = tmppt

                Else


                    pfPoints_ = Nothing


                End If


            End If
            Return pfPoints_
        End Get
        Set(value As CENPoints)
            pfPoints_ = value


        End Set
    End Property










    Public Sub close()

        pprdoc.Close()
    End Sub




    Public ReadOnly Property TVA As TVA_Method
        Get
            Return TVA_
            ' Return PCFeatContainer.ProcFeats.Count


        End Get
    End Property

    Public ReadOnly Property pfCount As Integer
        Get

            Return PCFeatContainer.ProcFeats.Count


        End Get
    End Property




End Class

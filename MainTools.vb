
Imports ProductStructureTypeLib
Imports MECMOD
Imports INFITF

Imports mysqlsolution
Imports OFFICE_Method
Imports HybridShapeTypeLib
Imports System.Threading

Public Class MainTools

    Dim fso As Object

    '紧固件列表
    Dim FastNames_ As List(Of String)

    Dim CATIA As Application
    Dim TVAfather As Product = Nothing
    Dim OuterSurface As HybridShapeExtract
    Dim sourcetvainner As Part
    Dim targetTVA As Part
    Dim SPProducts As New Dictionary(Of String, Product)
    Dim searchDic As Dictionary(Of String, HybridShape)
    Dim defTVAModel As TVA_Method
    Dim opGeo As HybridBody = Nothing
    Dim CATIA_task As Threading.Thread




    Public Property sourceTVA As Part
        Get

            If sourcetvainner Is Nothing Then
                MsgBox("Make sure you have a Skin/Source Part Selected")
                Return Nothing
            Else
                Return sourcetvainner
            End If

        End Get
        Set(value As Part)
            sourcetvainner = value
        End Set
    End Property



    Public Property showPoints As Dictionary(Of String, HybridShape)
        '注意赋值前把CATIA文档激活
        Get


            Return searchDic


        End Get
        Set(value As Dictionary(Of String, HybridShape))

            searchDic = value


            CATIA.Visible = True

            buglistBox.DataSource = searchDic.Keys.ToList()


        End Set
    End Property

    Public Sub EnableButtons(Status As Boolean)

        Button2.Enabled = Status
        Button3.Enabled = Status
        Button4.Enabled = Status
        Button5.Enabled = Status

    End Sub







    Private Sub UserForm1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        CATIA = TVA_Method.loadCATIA()

        lookuptable_path.Text = My.Settings.templateTVAPath

        'initial the fasterner list,if it can not be fetched from database,then load it via exceltable
        Try

            FastNames_ = autorivet_op.allfast_list()
            lookuptable_path.Text = "From Datebase"

        Catch ex As Exception
            Dim fstExcelPath_ = My.Settings.lookuptablePath
            If System.IO.File.Exists(fstExcelPath_) Then
                lookuptable_path.Text = fstExcelPath_
                '  Dim kk = From s As DataRow In excelMethod.LoadDataFromExcel(fstExcelPath_).Rows
                '    Select Case CStr(s("Fasteners"))

                Dim fstdt = excelMethod.LoadDataFromExcel(fstExcelPath_)
                Dim kk = From s As DataRow In fstdt
                         Select CStr(s("Fasteners"))

                FastNames_ = kk.ToList()
            Else
                lookuptable_path.Text = "Please choose fasterners table or all work can not continue!"
                MsgBox("Please choose fasterners table or all work can not continue!")
                lookuptable_select.Enabled = True
            End If



        End Try

        fastener_list.DataSource = FastNames_


        'initial templateTVA path
        Dim templateTVApath = My.Settings.templateTVAPath
        If System.IO.File.Exists(templateTVApath) Then

            TargetTVABox.Text = Strings.Split(templateTVApath, "\").Last()
        End If


        'load Product Table

        Dim ProdList As List(Of String) = Nothing

        Try

            ProdList = autorivet_op.fullname_list("concat(图号,'_',名称,'_',站位号)")


        Catch ex As Exception
            Dim prodExcelPath_ = My.Settings.producttablePath
            If System.IO.File.Exists(prodExcelPath_) Then

                Dim kk = From s In excelMethod.LoadDataFromExcel(prodExcelPath_).AsEnumerable()
                         Select CStr(s("图号")) + "_" + CStr(s("名称")) + "_" + CStr(s("站位号"))
                ProdList = kk.ToList()
            Else

                MsgBox("Please choose product table!")
            End If



        End Try



        EnableButtons(True)

    End Sub



    Private Sub TVA_select_Click(sender As Object, e As EventArgs) Handles TVA_select.Click
        ' loadcatia()
        defTVAselect()
    End Sub

    Private Sub defTVAselect()



        Me.Hide()

        Dim selection1 As Object
        selection1 = CATIA.ActiveDocument.Selection

        Dim FilterData(0)
        FilterData(0) = "Part"

        If (selection1.SelectElement2(FilterData, "Choose the Skin Part", False) <> "Normal") Then
            MsgBox("Issue with part selected, please select again")
        Else
            'This part contains both skin and points
            sourceTVA = selection1.Item(1).Value
            'PointsPart = selection1.Item(1).Value
            'Dim aDoc
            'aDoc = CATIA.ActiveDocument.Product
            'If aDoc.Products.count <> 0 Then
            '    SkinProduct = aDoc
            '    PointsProduct = aDoc
            'End If
            SourceTVABox.Text = sourceTVA.Parent.Name
            TVASelection.Text = sourceTVA.Parent.Name
            defTVAModel = New TVA_Method(sourceTVA)
            defTVAModel.FstList = FastNames_

        End If

        Me.Show()
        Me.Activate()
    End Sub

    Public Property inibyTVAmodel As TVA_Method

        Get
            Return defTVAModel
        End Get
        Set(value As TVA_Method)

            SourceTVABox.Text = value.filename
            TVASelection.Text = SourceTVABox.Text
            defTVAModel = value
            sourceTVA = value.Part
            defTVAModel.FstList = FastNames_
            opGeo = value.pilot_geoset
        End Set
    End Property





    Private Sub AddSP01CockPit_Click(sender As Object, e As EventArgs) Handles AddSP01CockPit.Click
        'Hide user interface
        Me.Hide()
        'Let the user select an Sp01 part in the Tree

        Dim selection1
        selection1 = CATIA.ActiveDocument.Selection

        Dim FilterData(0)
        FilterData(0) = "Product"

        Dim SP01Product As Product

        If (selection1.SelectElement2(FilterData, "Choose the SP01 Product", False) <> "Normal") Then
            MsgBox("Issue with product selected, please select again")
        Else
            SP01Product = selection1.Item(1).Value
            SPProducts.Add(SP01Product.Name, SP01Product)
            'Then add it to the listbox
            SP01CP.Items.Add(SP01Product.Name)

        End If

        Me.Show()
        Me.Activate()
    End Sub

    Private Sub DelSP01Cockpit_Click(sender As Object, e As EventArgs) Handles DelSP01Cockpit.Click
        If SP01CP.SelectedIndex <> -1 Then
            SPProducts.Remove(SP01CP.SelectedItem.ToString)
            SP01CP.Items.Remove(SP01CP.SelectedItem.ToString)
        End If



    End Sub

    Private Sub SkinCockPit_Click(sender As Object, e As EventArgs) Handles SkinCockPit.Click
        defTVAselect()
    End Sub

    Private Sub CPLookup_Click(sender As Object, e As EventArgs) Handles CPLookup.Click
        Me.Hide()

        Dim selection1 As Object
        selection1 = CATIA.ActiveDocument.Selection
        selection1.Clear()

        Dim FilterData(0)
        FilterData(0) = "Part"

        If (selection1.SelectElement2(FilterData, "Choose the TVA Template", False) <> "Normal") Then
            MsgBox("Issue with part selected, please select again")
        Else
            targetTVA = selection1.Item(1).Value
            selection1.Clear()
            TargetTVABox.Text = targetTVA.Parent.Name
        End If

        Me.Show()
        Me.Activate()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        createTVA()
        'CATIA_task = New Thread(New ThreadStart(Sub() createTVA()))

    End Sub

    Private Sub createTVA()

        Dim prodDoc = CATIA.ActiveDocument
        '   MsgBox(prodDoc.Name)
        Dim oProduct = prodDoc.Product

        'check if necessary info has been filled
        If SourceTVABox.Text = "" Or TargetTVABox.Text = "" Or lookuptable_path.Text = "" Or SPProducts.Count = 0 Then
            MsgBox("Make sure you have a Skin Part, SP01 Product, and a TVA Template Selected")
        Else
            If targetTVA Is Nothing Then
                ' On Error Resume Next
                targetTVA = TVA_Method.openPartDoc(My.Settings.templateTVAPath).Part
            End If



            'Save template TVA path
            Dim newTmppath = targetTVA.Parent.FullName
            If My.Settings.templateTVAPath <> newTmppath Then
                My.Settings.templateTVAPath = newTmppath
                My.Settings.Save()

            End If

            'save new TVAfile
            Dim SavePath As String
            Dim objShell As System.Windows.Forms.FolderBrowserDialog = New System.Windows.Forms.FolderBrowserDialog()
            objShell.Description = "Select the folder to place the output"
            If objShell.ShowDialog() Then
                SavePath = objShell.SelectedPath
                TargetTVABox.Text = SourceTVABox.Text.Replace(".", "_temp_TVA.")

                Try

                    targetTVA = CATIA.Documents.Item(TargetTVABox.Text).Part
                Catch ex As Exception
                    targetTVA.Parent.SaveAs(SavePath + "\" + TargetTVABox.Text)
                End Try
                'If CATIA.Documents.Item(TargetTVABox.Text) Is Nothing Then
                '    targetTVA.Parent.SaveAs(SavePath + "\" + TargetTVABox.Text)
                'Else

                '    targetTVA = CATIA.Documents.Item(TargetTVABox.Text).Part
                'End If
                prodDoc.Activate()

                If TVAfather Is Nothing Then

                    TVA_Method.GetFatherProduct(oProduct, sourceTVA.Parent, TVAfather)
                End If
                '     MsgBox(TVAfather.Name)

                Dim sorpart = New TVA_Method(sourceTVA, TVAfather)
                sorpart.FstList = FastNames_
                sorpart.pointGeo = rootGeoSet.Text
                Dim tarpart = New TVA_Method(targetTVA)
                tarpart.FstList = FastNames_

                MsgLabel.Text = "Extracting the skin points"
                Dim allpts = sorpart.noTreePoints()

                MsgLabel.Text = "Extracting the SP models"
                Dim SP01 = New CENSP01s(SPProducts.Values.AsEnumerable(), FastNames_)
                MsgLabel.Text = "Matching the Fasterner Name"
                Dim rspts = allpts.compare(SP01)(0)
                MsgLabel.Text = "Converting to process tree"
                Dim tree = rspts.toProcessTree
                MsgLabel.Text = "Copying body"
                prodDoc.Activate()
                sorpart.CopyPastePartBody(TargetTVABox.Text)


                MsgLabel.Text = "Converting to new Part"


                prodDoc.Activate()





                tree.output_topart(prodDoc.Name, TargetTVABox.Text, tarpart.pilot_geoset)
                MsgLabel.Text = "Done"

                ' tarpart.save()

            End If
        End If

    End Sub



    Private Sub Button3_Click(sender As Object, e As EventArgs)

    End Sub





    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        Dim source = ComboBox2.Text
        Dim target = ComboBox3.Text


        If ComboBox4.Visible = True Then
            source = source + " - " + ComboBox4.Text

        End If
        If ComboBox5.Visible = True Then
            target = target + " - " + ComboBox5.Text

        End If
        Dim ssmodel = defTVAModel
        ssmodel.FstList = FastNames_
        ssmodel.pointGeo = rootGeoSet.Text


        Dim fstlist = New List(Of String)
        For Each dd In fastener_list.SelectedItems
            fstlist.Add(dd)

        Next

        Dim rpt = ssmodel.procBAT(fstlist, source, target)
        rpt.report()


        MsgLabel.Text = "Done"

        '   Call changeAction()
        '  Call generateChangeReport()
    End Sub








    Private Sub surface_select_Click(sender As Object, e As EventArgs) Handles surface_select.Click
        ' loadcatia()

        Me.Hide()

        Dim selection1 As Object
        selection1 = CATIA.ActiveDocument.Selection

        Dim FilterData(0)
        FilterData(0) = "HybridShapeExtract"

        If (selection1.SelectElement2(FilterData, "Choose Outer surface extract", False) <> "Normal") Then
            MsgBox("Issue with part selected, please select again")
        Else
            'This part contains both skin and points
            OuterSurface = selection1.Item(1).Value

            surfaceText.Text = OuterSurface.Name



        End If

        Me.Show()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If Strings.InStr(ComboBox2.Text, "RESYNCING") = 0 Then
            ComboBox4.Visible = False
        Else
            ComboBox4.Visible = True
        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        If Strings.InStr(ComboBox3.Text, "RESYNCING") = 0 Then
            ComboBox5.Visible = False
        Else
            ComboBox5.Visible = True
        End If
    End Sub

    Private Sub opreateTVA(Optional ifrpt As Boolean = True)

        Dim aa = defTVAModel
        aa.PilotHoles = opGeo
        aa.pointGeo = rootGeoSet.Text
        aa.FstList = FastNames_


        If iffix.Checked Then

            Dim cc = aa.fix_all()
            If ifrpt Then
                cc.report()
            End If


            If ifcolor.Checked Then
                aa.setcolor()

            End If

        End If

        If ifCheck.Checked Then

            Dim conversps As Dictionary(Of String, Object)
            searchDic = New Dictionary(Of String, HybridShape)

            conversps = aa.CheckTVA(database:=False, color:=ifcolor.Checked, ifreport:=ifrpt).Points_


            For Each item In conversps

                searchDic.Add(item.Key, item.Value)
            Next

            showPoints = searchDic

        Else
            If ifcolor.Checked Then


                aa.setcolor()
            End If

        End If



        If ifdt.Checked Then

            aa.updatedt()

        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        CATIA_task = New Thread(New ThreadStart(Sub() opreateTVA()))
        opreateTVA()


    End Sub


    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim folder1 As Object

        Dim SavePath As String
        Dim objShell As System.Windows.Forms.FolderBrowserDialog = New System.Windows.Forms.FolderBrowserDialog()
        ' Dim objFolder As Folder
        'If you use 16384 instead of 1 on the next line,'files are also displayed
        objShell.Description = "Select the folder to execute batch command"
        If objShell.ShowDialog() Then
            SavePath = objShell.SelectedPath
            If SavePath <> "" Then

                fso = CreateObject("Scripting.FileSystemObject")
                folder1 = fso.GetFolder(SavePath)  '获得文件夹

                CATIA.Visible = False
                CATIA.DisplayFileAlerts = False

                walktreework(folder1, SavePath)
                CATIA.Visible = True
            End If
        Else
            Me.EnableButtons(True)
            Exit Sub
        End If


    End Sub

    Private Sub walktreework(ByRef folder1 As Object, path As String)

        Dim fd As Object

        Dim fds, fs, f
        If (InStr(LCase(path), "old")) Or InStr(LCase(path), "backup") Then
        Else

            fds = folder1.SubFolders        '子文件夹集合
            For Each fd In fds                  '遍历子文件夹
                walktreework(fd, path + "\" + fd.name)

            Next

            fs = folder1.Files          '文件集合
            For Each f In fs                '遍历文件
                Dim tmpstr2 As String
                tmpstr2 = LCase(f.name)

                If (InStr(tmpstr2, "catpart")) Then
                    If (InStr(tmpstr2, "xls") Or InStr(tmpstr2, "cfg")) Then
                    Else

                        opreateTVA(False)
                    End If

                End If
            Next

        End If

    End Sub



    Private Sub OptionButton1_Click(sender As Object, e As EventArgs) Handles OptionButton1.Click
        Dim i As Integer
        For i = 0 To fastener_list.Items.Count - 1
            fastener_list.SetSelected(i, True)
        Next
    End Sub

    Private Sub OptionButton2_Click(sender As Object, e As EventArgs) Handles OptionButton2.Click
        Dim i As Integer
        For i = 0 To fastener_list.Items.Count - 1
            If Strings.InStr(fastener_list.Items.Item(i), "B020600") = 0 Then

                fastener_list.SetSelected(i, True)
            Else
                fastener_list.SetSelected(i, False)
            End If

        Next
    End Sub

    Private Sub OptionButton3_Click(sender As Object, e As EventArgs) Handles OptionButton3.Click
        Dim i As Integer
        For i = 0 To fastener_list.Items.Count - 1
            If Strings.InStr(fastener_list.Items.Item(i), "B020600") = 0 Then

                fastener_list.SetSelected(i, False)
            Else
                fastener_list.SetSelected(i, True)
            End If
        Next
    End Sub

    Private Sub OptionButton4_Click(sender As Object, e As EventArgs) Handles OptionButton4.Click
        Dim i As Integer
        For i = 0 To fastener_list.Items.Count - 1
            fastener_list.SetSelected(i, False)
        Next
    End Sub

    Private Sub OptionButton5_Click(sender As Object, e As EventArgs) Handles OptionButton5.Click
        Dim i As Integer
        For i = 0 To fastener_list.Items.Count - 1
            If Strings.InStr(fastener_list.Items.Item(i), "B020600") = 0 Or Strings.InStr(fastener_list.Items.Item(i), "AG5") = 0 Then

                fastener_list.SetSelected(i, False)
            Else
                fastener_list.SetSelected(i, True)
            End If
        Next
    End Sub

    Private Sub CommandButton40_Click(sender As Object, e As EventArgs) Handles CommandButton40.Click
        ComboBox2.Text = ComboBox2.Items.Item(3)
        ComboBox4.Text = ComboBox4.Items.Item(2)
        ComboBox4.Visible = True
        ComboBox3.Text = ComboBox3.Items.Item(0)
        ComboBox5.Text = ComboBox5.Items.Item(0)
        ComboBox5.Visible = True
    End Sub

    Private Sub CommandButton43_Click(sender As Object, e As EventArgs) Handles CommandButton43.Click
        ComboBox2.Text = ComboBox2.Items.Item(0)
        ComboBox4.Text = ComboBox4.Items.Item(2)
        ComboBox3.Text = ComboBox3.Items.Item(0)
        ComboBox5.Text = ComboBox5.Items.Item(0)
        ComboBox4.Visible = True
        ComboBox5.Visible = True
    End Sub

    Private Sub CommandButton47_Click(sender As Object, e As EventArgs) Handles CommandButton47.Click
        ComboBox2.Text = ComboBox2.Items.Item(7)
        ComboBox4.Visible = False

        ComboBox3.Text = ComboBox3.Items.Item(6)
        ComboBox5.Visible = False
    End Sub



    Private Sub NCToolsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NCToolsToolStripMenuItem.Click
        Dim f1 = New program_input()
        f1.inputValue = ""
        f1.Show()
    End Sub

    Private Sub DatabaseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DatabaseToolStripMenuItem.Click
        Dim f1 = New setting_database
        f1.Show()

    End Sub

    Private Sub SavingFoldersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SavingFoldersToolStripMenuItem.Click
        Dim f1 = New Save_Folder
        f1.Show()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs)

    End Sub




    Private Sub TVASelection_TextChanged(sender As Object, e As EventArgs) Handles TVASelection.TextChanged

    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        'check if necessary info has been filled
        If SourceTVABox.Text = "" Or TargetTVABox.Text = "" Or TargetTVABox.Text = Strings.Split(My.Settings.templateTVAPath, "\").Last() Or lookuptable_path.Text = "" Then
            MsgBox("Make sure you have a Source Part and a Target Part Selected")
        Else


            Dim aa = defTVAModel
            aa.FstList = FastNames_
            MsgLabel.Text = "Reading Point data"
            Dim sourcetree = aa.TVATreeList
            MsgLabel.Text = "copying to target TVA"
            Dim bb = New TVA_Method(targetTVA)
            sourcetree.output_topart(SourceTVABox.Text, TargetTVABox.Text, bb.pilot_geoset)
            MsgLabel.Text = "Done"

        End If


    End Sub


    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        'check if necessary info has been filled
        If SourceTVABox.Text = "" Or TargetTVABox.Text = "" Or TargetTVABox.Text = Strings.Split(My.Settings.templateTVAPath, "\").Last() Or lookuptable_path.Text = "" Then
            MsgBox("Make sure you have a Source Part and a Target Part Selected")
        Else


            Dim ssmodal = defTVAModel
            ssmodal.FstList = FastNames_
            MsgLabel.Text = "Reading Source Points data"

            Dim sourcepplist = ssmodal.TVAPointsnoVic


            MsgLabel.Text = "Reading Target Points data"

            Dim targetmodel As New TVA_Method(targetTVA)
            targetmodel.FstList = FastNames_
            Dim targetpplist As New CENPoints
            targetpplist = targetmodel.TVAPointsnoVic

            Dim comresult() As CENPoints = {New CENPoints, New CENPoints}
            MsgLabel.Text = "Comparing souce TVA and target TVA"
            comresult = sourcepplist.compare(targetpplist)
            MsgLabel.Text = "Outputing the result to target TVA"
            Dim targetBody = targetTVA.HybridBodies.Add
            targetBody.Name = "Comparison" + Now.Date.ToShortDateString()
            comresult(0).toProcessTree.output_topart(SourceTVABox.Text, TargetTVABox.Text, targetBody, 1)
            comresult(1).toProcessTree.output_topart(TargetTVABox.Text, TargetTVABox.Text, targetBody, 1)
            MsgLabel.Text = "Done"



        End If
    End Sub




    Private Sub Button16_Click_1(sender As Object, e As EventArgs) Handles Button16.Click


        If SourceTVABox.Text = "" Or lookuptable_path.Text = "" Or SPProducts.Count = 0 Then
            MsgBox("Make sure you have a Source Part, SP01 Product Selected")
        Else
            MsgLabel.Text = "Getting pointsFatherProduct"
            Dim TVAfather As Product = Nothing
            TVA_Method.GetFatherProduct(CATIA.ActiveDocument.Product, sourceTVA.Parent, TVAfather)

            Dim sorpart = New TVA_Method(sourceTVA, TVAfather)
            sorpart.FstList = FastNames_
            MsgLabel.Text = "Getting TVA points"


            Dim sourcepplist As CENPoints
            sourcepplist = sorpart.TVAPoints

            MsgLabel.Text = "Extracting the SP models"
            Dim SP01s_ = New CENSP01s(SPProducts.Values.AsEnumerable(), FastNames_)

            MsgLabel.Text = "Comparing fasterners"
            Dim filename = SourceTVABox.Text
            Dim comresult() As CENPoints = {New CENPoints, New CENPoints, New CENPoints}
            comresult = sourcepplist.compare(SP01s_)
            Dim changelist = comresult(0).toProcessTree()
            If changelist.count() > 0 Then
                Dim targetBody = sorpart.pilot_geoset.HybridBodies.Add
                targetBody.Name = "ChangedFstType" + Now.Date.ToShortDateString()
                changelist.output_topart(filename, filename, targetBody, 0)
                changelist.del(filename)
            End If



            Dim dellist = comresult(1).toProcessTree()
            If dellist.count() > 0 Then

                Dim targetBody2 = sourceTVA.HybridBodies.Add
                targetBody2.Name = "MayHaveBeenDeleted" + Now.Date.ToShortDateString()
                dellist.output_topart(filename, filename, targetBody2, 0)
                ' dellist.del(TVAfilename)
            End If
            sorpart.Part.Update()

        End If
        MsgLabel.Text = "Done"
    End Sub




    Private Sub lookuptable_select_Click(sender As Object, e As EventArgs) Handles lookuptable_select.Click
        Dim fileDialog As New System.Windows.Forms.OpenFileDialog()

        fileDialog.InitialDirectory = "D://"

        fileDialog.Filter = "xls files (*.xls)|*.xls|All files (*.*)|*.*"

        fileDialog.FilterIndex = 1

        fileDialog.RestoreDirectory = True

        If (fileDialog.ShowDialog() = DialogResult.OK) Then


            Dim newpath = fileDialog.FileName


            Try

                Dim kk = From s In excelMethod.LoadDataFromExcel(newpath).Rows
                         Select CStr(s("Fasteners"))
                FastNames_ = kk.ToList()
            Catch ex As Exception
                Exit Sub
            End Try

            My.Settings.lookuptablePath = newpath
            My.Settings.Save()


        End If





    End Sub

    Private Sub buglistBox_DoubleClick(sender As Object, e As EventArgs) Handles buglistBox.DoubleClick

        Dim bugselect
        bugselect = CATIA.ActiveDocument.Selection
        bugselect.Clear()
        bugselect.Add(searchDic(buglistBox.SelectedItem.ToString()))
        ' bugselect.Reframe

        If (RadioButton2.Checked) Then
            CATIA.StartCommand("居中")
            CATIA.StartCommand("将图居中")

        Else
            CATIA.StartCommand("Reframe on")
            CATIA.StartCommand("Center graph")


        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        showHideHelper(True)


    End Sub

    Private Sub showHideHelper(ifshow As Boolean)
        Dim source = ComboBox6.Text
        defTVAModel.PilotHoles = opGeo
        defTVAModel.pointGeo = rootGeoSet.Text
        defTVAModel.FstList = FastNames_
        If ComboBox1.Visible = True Then
            source = source + " - " + ComboBox1.Text

        End If



        defTVAModel.setHide(ifshow, source)

    End Sub
    Private Sub showHideAllHelper(ifshow As Boolean)
        Dim ss = New processTreeBase
        For Each rr As String In ss.firsttree

            If Strings.InStr(rr, "RESYNCING") Then
                For Each kk As String In ss.sectree

                    defTVAModel.setHide(ifshow, rr + " - " + kk)

                Next
            Else

                defTVAModel.setHide(ifshow, rr)
            End If


        Next





    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        If Strings.InStr(ComboBox6.Text, "RESYNCING") = 0 Then
            ComboBox1.Visible = False
        Else
            ComboBox1.Visible = True
        End If
    End Sub



    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        showHideHelper(False)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Me.Hide()
        'Let the user select an Sp01 part in the Tree

        Dim selection1
        selection1 = CATIA.ActiveDocument.Selection

        Dim FilterData(0)
        FilterData(0) = "Product"



        If (selection1.SelectElement2(FilterData, "Choose the TVA father Product", False) <> "Normal") Then
            MsgBox("Issue with product selected, please select again")
        Else
            TVAfather = selection1.Item(1).Value
            TextBox1.Text = TVAfather.Name

        End If

        Me.Show()
        Me.Activate()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Me.Hide()
        'Let the user select an Sp01 part in the Tree

        Dim selection1
        selection1 = CATIA.ActiveDocument.Selection

        Dim FilterData(0)
        FilterData(0) = "HybridBody"



        If (selection1.SelectElement2(FilterData, "Choose the points geometry", False) <> "Normal") Then
            MsgBox("Issue with HybridBody selected, please select again")
        Else
            rootGeoSet.Text = selection1.Item(1).Value.Name
            opGeo = selection1.Item(1).Value

        End If

        Me.Show()
        Me.Activate()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        showHideAllHelper(True)
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        showHideAllHelper(False)
    End Sub
End Class

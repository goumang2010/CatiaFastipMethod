Imports MECMOD

Public Class TVA_locate
    Dim CATIA As INFITF.Application
    Dim unipoints As CENPoints
    Dim searchDic As New Dictionary(Of String, HybridShape)
    Private Sub TVA_locate_Load(sender As Object, e As EventArgs) Handles MyBase.Load




    End Sub

    Public Property expl As String
        Get

            Return Label2.Text
        End Get
        Set(value As String)
            Label2.Text = value
        End Set
    End Property



    Public Property showPoints As CENPoints
        '注意赋值前把CATIA文档激活
        Get


            Return unipoints


        End Get
        Set(value As CENPoints)

            unipoints = value

            CATIA = TVA_Method.loadCATIA()
            CATIA.Visible = True

            For Each pp As CENPoint In unipoints.Points_.Values

                If pp.PFname <> "" Then

                    searchDic.Add(pp.PFname, pp.MyPoint)

                Else

                    searchDic.Add("No PF name;" + pp.Diam + ";" + pp.uuid, pp.MyPoint)
                End If


            Next

            ListBox1.DataSource = searchDic.Keys.ToList()


        End Set
    End Property













    Private Sub ListBox1_MouseDoubleClick(sender As Object, e As Windows.Forms.MouseEventArgs) Handles ListBox1.MouseDoubleClick


        Dim bugselect = CATIA.ActiveDocument.Selection
        bugselect.Clear()
        bugselect.Add(searchDic(ListBox1.SelectedItem.ToString()))

        ' bugselect.Reframe
        'Dim specsAndGeomWindow1 As INFITF.SpecsAndGeomWindow
        'specsAndGeomWindow1 = CATIA.ActiveWindow

        'Dim viewer3D1 As INFITF.Viewer3D
        'viewer3D1 = specsAndGeomWindow1.ActiveViewer
        'viewer3D1.Activate()
        'MsgBox(CATIA.ActiveDocument.Name)
        ' CATIA.HSOSynchronized = True
        If CATIA.GetWorkbenchId <> "PrtCfg" Then
            CATIA.StartWorkbench("PrtCfg")
        End If
        '
        If (RadioButton2.Checked) Then
            CATIA.StartCommand("居中")
            CATIA.StartCommand("将图居中")

        Else


            CATIA.StartCommand("Reframe on")
            CATIA.StartCommand("Center graph")


        End If
        'CATIA.ActiveWindow.ActiveViewer.Reframe()


    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text <> "" Then
            ListBox1.DataSource = searchDic.Keys.Where(Function(p) p.Contains(TextBox1.Text)).ToList()
        Else

            ListBox1.DataSource = searchDic.Keys.ToList()

        End If

    End Sub
End Class
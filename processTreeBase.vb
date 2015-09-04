Public Class processTreeBase
    Public firsttree As New List(Of String)
    Public sectree As New List(Of String)
    Public framename As String
    Public fasternername As String

    Public Sub New()




        sectree.Add("Target type - Final")
        sectree.Add("Target type - Pilot Holes")
        sectree.Add("Target type - Fast Tack")

        firsttree.Add("RESYNCING ONLY BY AUTOMATED FASTENING")
        firsttree.Add("FASTENER INSTALLED BY AUTOMATED FASTENING")
        firsttree.Add("RESYNCING AND DRILL ONLY BY AUTOMATED FASTENING")
        firsttree.Add("RESYNCING AND FASTENER INSTALLED BY AUTOMATED FASTENING")
        firsttree.Add("DRILL ONLY BY AUTOMATED FASTENING")
        firsttree.Add("FASTENER DRILL BEFORE AUTOMATED FASTENING")
        firsttree.Add("FASTENER INSTALLED BEFORE AUTOMATED FASTENING")
        firsttree.Add("FASTENER INSTALLED AFTER AUTOMATED FASTENING")
        firsttree.Add("TEMP")




    End Sub


    Sub addFastername(fstname As String)
        fasternername = fstname




    End Sub




End Class

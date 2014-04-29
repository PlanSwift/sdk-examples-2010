Imports System
Imports System.IO




Public Class Form1
    Dim ps As PlanSwift9.PlanSwift
    Dim tlst As List(Of String)

    Sub LoadTakeOffItems(ByVal itm As PlanSwift9.IItem)
        Dim isItm As Boolean


        For idx = 0 To itm.ChildCount - 1
            Dim citm As PlanSwift9.IItem = itm.ChildItem(idx)
            If citm.GetProperty("type").ResultAsString = "Folder" Then
                LoadTakeOffItems(citm)
            Else
                isItm = citm.GetPropertyResultAsBoolean("IsItem", False)

                If (isItm = True) Then
                    tlst.Add(citm.GUID)
                End If
                If citm.ChildCount > 0 Then
                    LoadTakeOffItems(citm)
                End If
            End If
        Next

    End Sub
    Sub itemtoword(ByVal aitm As PlanSwift9.IItem, ByVal tbl As Microsoft.Office.Interop.Word.Table, ByVal ItmType As String, ByVal rowidx As Integer)
        If rowidx > 20 Then
            tbl.Rows.Add()
        End If
        If ItmType = "Digitizer" Then
            tbl.Cell(rowidx, 1).Range.Font.Bold = 1
            tbl.Cell(rowidx, 1).Range.Font.Italic = 0
        Else
            tbl.Cell(rowidx, 1).Range.Font.Bold = 0
            tbl.Cell(rowidx, 1).Range.Font.Italic = 1
        End If
        'Item Name
        tbl.Cell(rowidx, 1).Range.Text = aitm.Name
        'Item Qty
        tbl.Cell(rowidx, 2).Range.Text = aitm.GetProperty("Qty").ResultAsString
        'item Units
        tbl.Cell(rowidx, 3).Range.Text = aitm.GetProperty("Qty").Units
        'item Price Each
        tbl.Cell(rowidx, 4).Range.Text = aitm.GetPropertyResultAsString("Price Each", "")
        'item Price Total
        tbl.Cell(rowidx, 5).Range.Text = aitm.GetPropertyResultAsString("Price Total", "")

    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Form2.ShowDialog()
        If Form2.DialogResult <> vbOK Then
            Exit Sub
        End If
        Dim Template As Object = Directory.GetCurrentDirectory() & "\Includes\SampleQuote.dotx"
        Dim newTemplate As Object = False
        Dim DocType As Object = 0
        Dim isVisible As Object = True
        Dim msword = New Microsoft.Office.Interop.Word.Application
        Dim msdoc As Microsoft.Office.Interop.Word.Document


        msword.Visible = False

        msdoc = msword.Documents.Add(Template, newTemplate, DocType, isVisible)
        Dim tbl As Microsoft.Office.Interop.Word.Table = msdoc.Tables.Item(2)

        Dim ReportType As String = Form2.ComboBox1.Text
        ps.BeginUpdate()

        Dim rowidx As Integer = 1
        If ReportType = "Digitizer Items Only" Then
            For idx = 0 To tlst.Count - 1
                Dim isDigitizer As Boolean = False
                Dim aitm As PlanSwift9.IItem = ps.GetItem(tlst.Item(idx))
                If aitm.GetPropertyResultAsBoolean("IsArea", False) Then
                    isDigitizer = True
                End If
                If aitm.GetPropertyResultAsBoolean("IsLinear", False) Then
                    isDigitizer = True
                End If
                If aitm.GetPropertyResultAsBoolean("IsSegment", False) Then
                    isDigitizer = True
                End If
                If aitm.GetPropertyResultAsBoolean("IsCount", False) Then
                    isDigitizer = True
                End If

                If isDigitizer Then
                    rowidx = rowidx + 1
                    itemtoword(aitm, tbl, "Digitizer", rowidx)
                End If
                ProgressBar1.Increment(idx)
            Next
        End If
        If ReportType = "Parts Only" Then
            For idx = 0 To tlst.Count - 1
                Dim aitm As PlanSwift9.IItem = ps.GetItem(tlst.Item(idx))
                If aitm.GetPropertyResultAsBoolean("IsPart", False) Then
                    rowidx = rowidx + 1
                    itemtoword(aitm, tbl, "Part", rowidx)
                End If
                ProgressBar1.Increment(idx)
            Next
        End If
        If ReportType = "Digitizer Items w/parts" Then
            
            For idx = 0 To tlst.Count - 1
                Dim aitm As PlanSwift9.IItem = ps.GetItem(tlst.Item(idx))
                rowidx = rowidx + 1
                If aitm.GetPropertyResultAsBoolean("isPart", False) Then
                    rowidx = rowidx + 1
                    itemtoword(aitm, tbl, "Part", rowidx)
                Else
                    rowidx = rowidx + 1
                    itemtoword(aitm, tbl, "Digitizer", rowidx)
                End If
                ProgressBar1.Increment(idx)
            Next
        End If
        ProgressBar1.Value = 0
        ps.EndUpdate()

        msword.Visible = True
        msdoc = Nothing
        msword = Nothing
    End Sub

    Private Sub Form1_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ps = Nothing
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ps = New PlanSwift9.PlanSwift
        Dim tpath As PlanSwift9.IItem
        tpath = ps.GetItem(ps.Root.FullPath & "\Job\Takeoff")
        tlst = New List(Of String)
        'load all "Takeoff Items" into a list
        LoadTakeOffItems(tpath)
        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = tlst.Count
    End Sub
    Sub ItemToExcel(ByVal aitm As PlanSwift9.IItem, ByVal xlsheet As Microsoft.Office.Interop.Excel.Worksheet, ByVal rowidx As Integer)
        Dim row As Object = rowidx
        Dim xlcolumn As Object = 1
        If rowidx > 36 Then
            xlsheet.Cells.Item(row, xlcolumn).EntireRow.Insert()
        End If
        'QtyCell
        xlsheet.Cells.Item(rowidx, 1).value = aitm.GetPropertyResultAsString("Qty", "")
        'QtyUnits
        xlsheet.Cells.Item(rowidx, 2).value = aitm.GetProperty("Qty").Units
        'Item Number
        xlsheet.Cells.Item(rowidx, 3).value = aitm.GetPropertyResultAsString("Item #", "")
        'Item Name
        xlsheet.Cells.Item(rowidx, 4).value = aitm.Name
        'Item Price Each
        xlsheet.Cells.Item(rowidx, 6).value = aitm.GetPropertyResultAsString("Price Each", "")
        'Item Price Total
        xlsheet.Cells.Item(rowidx, 7).value = aitm.GetPropertyResultAsString("Price Total", "")
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Form2.ShowDialog()
        If Form2.DialogResult <> vbOK Then
            Exit Sub
        End If
        Dim excel As New Microsoft.Office.Interop.Excel.Application
        Dim xlbook As Microsoft.Office.Interop.Excel.Workbook
        Dim Template As Object = Directory.GetCurrentDirectory & "\Includes\Estimate.xlt"
        Dim rowidx As Integer = 17
        xlbook = excel.Workbooks.Add(Template)
        Dim xlsheet As Microsoft.Office.Interop.Excel.Worksheet
        xlsheet = xlbook.Worksheets.Item(1)

        Dim ReportType = Form2.ComboBox1.Text

        If ReportType = "Digitizer Items Only" Then
            For idx = 0 To tlst.Count - 1
                Dim isDigitizer As Boolean = False
                Dim aitm As PlanSwift9.IItem = ps.GetItem(tlst.Item(idx))
                If aitm.GetPropertyResultAsBoolean("IsArea", False) Then
                    isDigitizer = True
                End If
                If aitm.GetPropertyResultAsBoolean("IsLinear", False) Then
                    isDigitizer = True
                End If
                If aitm.GetPropertyResultAsBoolean("IsSegment", False) Then
                    isDigitizer = True
                End If
                If aitm.GetPropertyResultAsBoolean("IsCount", False) Then
                    isDigitizer = True
                End If

                If isDigitizer Then
                    rowidx = rowidx + 1
                    ItemToExcel(aitm, xlsheet, rowidx)
                End If
                ProgressBar1.Increment(idx)
            Next
        End If
        If ReportType = "Parts Only" Then
            For idx = 0 To tlst.Count - 1
                Dim aitm As PlanSwift9.IItem = ps.GetItem(tlst.Item(idx))
                If aitm.GetPropertyResultAsBoolean("IsPart", False) Then
                    rowidx = rowidx + 1
                    ItemToExcel(aitm, xlsheet, rowidx)
                End If
                ProgressBar1.Increment(idx)
            Next
        End If
        If ReportType = "Digitizer Items w/parts" Then

            For idx = 0 To tlst.Count - 1
                Dim aitm As PlanSwift9.IItem = ps.GetItem(tlst.Item(idx))
                rowidx = rowidx + 1
                ItemToExcel(aitm, xlsheet, rowidx)
                ProgressBar1.Increment(idx)
            Next
        End If
        ProgressBar1.Value = 0
        excel.Visible = True
        excel = Nothing
        xlsheet = Nothing
        xlbook = Nothing
    End Sub
 
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
       
    End Sub
    Private Function createMailbody(ByVal RT As String) As String
        Dim body As String
        Dim itm As PlanSwift9.IItem
        body = "<body>" & vbCr
        body = body & "<h1>Report Type: " & RT & "</h1>" & vbCr
        body = body & "<table style=" & Chr(34) & "width:auto; font-weight:Bold; font-family:Tahoma; font-size:14px;" & Chr(34) & " > " & vbCr
        body = body & "<tr>" & vbCr
        body = body & "<td style=" & Chr(34) & "width:200px; font-weight:Bold; font-family:Tahoma; font-size:14px;" & Chr(34) & " > "
        body = body & "Name"
        body = body & "</td>" & vbCr
        body = body & "<td style=" & Chr(34) & "width:100px; font-weight:Bold; font-family:Tahoma; font-size:14px;" & Chr(34) & " > "
        body = body & "Qty"
        body = body & "</td>" & vbCr
        body = body & "<td style=" & Chr(34) & "width:100px; font-weight:Bold; font-family:Tahoma; font-size:14px;" & Chr(34) & " > "
        body = body & "Price Each"
        body = body & "</td>" & vbCr
        body = body & "<td style=" & Chr(34) & "width:100px; font-weight:Bold; font-family:Tahoma;font-size:14px;" & Chr(34) & " > "
        body = body & "Price Total"
        body = body & "</td>" & vbCr
        body = body & "</tr>" & vbCr
        If RT = "Digitizer Items Only" Then
            For idx = 0 To tlst.Count - 1
                itm = ps.GetItem(tlst(idx))
                If itm.GetPropertyResultAsBoolean("isPart", False) = False Then
                    body = body & "<tr><td style=" & Chr(34) & "font-size:12px; font-weight:normal;" & Chr(34) & ">" & itm.Name & "</td>"
                    body = body & "<td style=" & Chr(34) & "font-size:12px; font-weight:normal;" & Chr(34) & ">" & itm.GetPropertyResultAsString("Qty", "") & "</td>"
                    body = body & "<td style=" & Chr(34) & "font-size:12px; font-weight:normal;" & Chr(34) & ">" & itm.GetPropertyResultAsString("Price Each", "") & "</td>"
                    body = body & "<td style=" & Chr(34) & "font-size:12px; font-weight:normal;" & Chr(34) & ">" & itm.GetPropertyResultAsString("Price Total", "") & "</td>"
                    body = body & "</tr>"
                End If

            Next

        End If
        If RT = "Parts Only" Then
            For idx = 0 To tlst.Count - 1
                itm = ps.GetItem(tlst(idx))
                If itm.GetPropertyResultAsBoolean("isPart", False) = True Then
                    body = body & "<tr><td style=" & Chr(34) & "font-size:12px; font-weight:normal;" & Chr(34) & ">" & itm.Name & "</td>"
                    body = body & "<td style=" & Chr(34) & "font-size:12px; font-weight:normal;" & Chr(34) & ">" & itm.GetPropertyResultAsString("Qty", "") & "</td>"
                    body = body & "<td style=" & Chr(34) & "font-size:12px; font-weight:normal;" & Chr(34) & ">" & itm.GetPropertyResultAsString("Price Each", "") & "</td>"
                    body = body & "<td style=" & Chr(34) & "font-size:12px; font-weight:normal;" & Chr(34) & ">" & itm.GetPropertyResultAsString("Price Total", "") & "</td>"
                    body = body & "</tr>"
                End If

            Next

        End If
        If RT = "Digitizer Items w/parts" Then
            For idx = 0 To tlst.Count - 1
                itm = ps.GetItem(tlst(idx))
                If itm.GetPropertyResultAsBoolean("isPart", False) = False Then
                    body = body & "<tr><td style=" & Chr(34) & "font-size:12px; font-weight:normal; color:#007dc3;" & Chr(34) & ">" & itm.Name & "</td>"
                    body = body & "<td style=" & Chr(34) & "font-size:12px; font-weight:normal; color:#007dc3;" & Chr(34) & ">" & itm.GetPropertyResultAsString("Qty", "") & "</td>"
                    body = body & "<td style=" & Chr(34) & "font-size:12px; font-weight:normal; color:#007dc3;" & Chr(34) & ">" & itm.GetPropertyResultAsString("Price Each", "") & "</td>"
                    body = body & "<td style=" & Chr(34) & "font-size:12px; font-weight:normal; color:#007dc3;" & Chr(34) & ">" & itm.GetPropertyResultAsString("Price Total", "") & "</td>"
                    body = body & "</tr>"
                End If
                If itm.GetPropertyResultAsBoolean("isPart", False) = True Then
                    body = body & "<tr><td style=" & Chr(34) & "font-size:12px; font-weight:normal; color:#0000FF;" & Chr(34) & ">" & itm.Name & "</td>"
                    body = body & "<td style=" & Chr(34) & "font-size:12px; font-weight:normal; color:#0000FF;" & Chr(34) & ">" & itm.GetPropertyResultAsString("Qty", "") & "</td>"
                    body = body & "<td style=" & Chr(34) & "font-size:12px; font-weight:normal; color:#0000FF;" & Chr(34) & ">" & itm.GetPropertyResultAsString("Price Each", "") & "</td>"
                    body = body & "<td style=" & Chr(34) & "font-size:12px; font-weight:normal; color:#0000FF;" & Chr(34) & ">" & itm.GetPropertyResultAsString("Price Total", "") & "</td>"
                    body = body & "</tr>"
                End If

            Next
        End If


        createMailbody = body
    End Function
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Form2.ShowDialog()
        If Form2.DialogResult <> vbOK Then
            Exit Sub
        End If
        Dim olook As Microsoft.Office.Interop.Outlook._Application
        Dim display As Object
        display = True
        olook = New Microsoft.Office.Interop.Outlook.Application
        Dim omail As Microsoft.Office.Interop.Outlook._MailItem
        omail = olook.CreateItem(0)
        omail.HTMLBody = createMailbody(Form2.ComboBox1.Text)
        omail.Display(display)
    End Sub
End Class

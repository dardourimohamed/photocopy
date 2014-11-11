Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim scds As New DataTable
        scds.Columns.Add("dev", GetType(WIA.DeviceInfo))
        scds.Columns.Add("name", GetType(String))
        For Each d As WIA.DeviceInfo In Scanner.getDevices
            Dim name = d.DeviceID
            For Each p As WIA.Property In d.Properties
                If p.Name = "Name" Then name = p.Value.ToString
            Next
            scds.Rows.Add(d, name)
        Next
        ComboBox1.DataSource = scds
        ComboBox1.DisplayMember = "name"
        ComboBox1.ValueMember = "dev"

        ComboBox2.Items.Clear()
        For Each p As String In Printing.PrinterSettings.InstalledPrinters
            ComboBox2.Items.Add(p)
        Next
        ComboBox2.SelectedItem = My.Settings.printer
        TrackBar1.Value = My.Settings.brightness
        TrackBar2.Value = My.Settings.contrast
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        Button1.Enabled = ComboBox1.SelectedValue IsNot Nothing
        Button3.Enabled = ComboBox1.SelectedValue IsNot Nothing
    End Sub

    Structure img
        Dim img As Image
        Dim nb As Integer
    End Structure

    Dim imglist As New List(Of img)

    Dim btn_scan As Boolean = True

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        PrintDocument1.PrinterSettings.PrinterName = ComboBox2.SelectedItem
        imglist.Add(New img With {.img = Scanner.Scan(ComboBox1.SelectedValue, TrackBar1.Value, TrackBar2.Value), .nb = NumericUpDown1.Value})
        PrintDocument1.Print()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        imglist.Add(New img With {.img = PictureBox1.Image, .nb = NumericUpDown1.Value})
        PrintDocument1.PrinterSettings.PrinterName = ComboBox2.SelectedItem
        PrintDocument1.Print()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Button2.Enabled = True
        PrintDocument1.PrinterSettings.PrinterName = ComboBox2.SelectedItem
        PictureBox1.Image = Scanner.Scan(ComboBox1.SelectedValue, TrackBar1.Value, TrackBar2.Value)
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        If imglist.Count = 0 Then : e.Cancel = True : Exit Sub : End If

        Dim myimg As Bitmap = imglist(0).img

        If imglist(0).nb > 1 Then
            imglist(0) = New img With {.img = myimg, .nb = imglist(0).nb - 1}
            e.HasMorePages = True
        Else
            imglist.RemoveAt(0)
            e.HasMorePages = imglist.Count > 0
        End If
        e.Graphics.DrawImage(myimg, 0, 0, e.PageSettings.PrintableArea.Width, e.PageSettings.PrintableArea.Height)
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        My.Settings.printer = ComboBox2.SelectedItem
    End Sub

    Private Sub TrackBar1_Scroll(sender As Object, e As EventArgs) Handles TrackBar1.Scroll
        My.Settings.brightness = TrackBar1.Value
    End Sub

    Private Sub TrackBar2_Scroll(sender As Object, e As EventArgs) Handles TrackBar2.Scroll
        My.Settings.contrast = TrackBar2.Value
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            Button1.Visible = True
            Button2.Visible = True
            Button3.Visible = False
            Me.Size = New Size(800, 600)
            SplitContainer1.Panel2Collapsed = False
        Else
            Button1.Visible = False
            Button2.Visible = False
            Button3.Visible = True
            Me.Size = New Size(SplitContainer1.Panel1.Width, 300)
            SplitContainer1.Panel2Collapsed = True
        End If
    End Sub

End Class

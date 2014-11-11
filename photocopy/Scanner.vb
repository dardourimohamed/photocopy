Imports System.Drawing
Imports System.Drawing.Drawing2D

Module Scanner

    Dim devm As New WIA.DeviceManager

    Public Function getDevices() As List(Of WIA.DeviceInfo)
        Dim rslt As New List(Of WIA.DeviceInfo)
        For Each d As WIA.DeviceInfo In devm.DeviceInfos
            If d.Type = WIA.WiaDeviceType.ScannerDeviceType Then rslt.Add(d)
        Next
        Return rslt
    End Function

    Public Function ResizeImage(ByVal image As Image, ByVal size As Size, Optional ByVal preserveAspectRatio As Boolean = True) As Image
        Dim newWidth As Integer
        Dim newHeight As Integer
        If preserveAspectRatio Then
            Dim originalWidth As Integer = image.Width
            Dim originalHeight As Integer = image.Height
            Dim percentWidth As Single = CSng(size.Width) / CSng(originalWidth)
            Dim percentHeight As Single = CSng(size.Height) / CSng(originalHeight)
            Dim percent As Single = If(percentHeight < percentWidth, percentHeight, percentWidth)
            newWidth = CInt(originalWidth * percent)
            newHeight = CInt(originalHeight * percent)
        Else
            newWidth = size.Width
            newHeight = size.Height
        End If
        Dim newImage As Image = New Bitmap(newWidth, newHeight)
        Using graphicsHandle As Graphics = Graphics.FromImage(newImage)
            graphicsHandle.InterpolationMode = InterpolationMode.HighQualityBicubic
            graphicsHandle.DrawImage(image, 0, 0, newWidth, newHeight)
        End Using
        Return newImage
    End Function

    Public Function Scan(device As WIA.DeviceInfo, brightness As Integer, contrast As Integer) As Bitmap
        Dim CD As New WIA.CommonDialog
        Dim dev As WIA.Device = device.Connect
        'dev = CD.ShowSelectDevice(WIA.WiaDeviceType.ScannerDeviceType, False, True)
        Try
            For Each prp As WIA.Property In dev.Items(1).Properties
                Select Case prp.Name
                    'Case "Color Profile Name" : prp.Value = "C:\Windows\system32\spool\drivers\color\sRGB Color Space Profile.icm"
                    'Case "Preview" : prp.Value = 1
                    'Case "Format" : prp.Value = "{B96B3CAA-0728-11D3-9D7B-0000F81EF32E}"
                    'Case "Media Type" : prp.Value = 128
                    Case "Data Type" : prp.Value = 2
                        'Case "Bits Per Pixel" : prp.Value = 24
                    Case "Compression" : prp.Value = 0
                    Case "Horizontal Resolution" : prp.Value = 300
                    Case "Vertical Re'solution" : prp.Value = 300
                        'Case "Horizontal Extent" : prp.Value = 637
                        'Case "Vertical Extent" : prp.Value = 877
                        'Case "Horizontal Start Position" : prp.Value = 0
                        'Case "Vertical Start Position" : prp.Value = 0
                    Case "Brightness" : prp.Value = brightness
                    Case "Contrast" : prp.Value = contrast
                        'Case "Current Intent" : prp.Value = 0
                        'Case "Threshold" : prp.Value = 128
                        'Case "Photometric Interpretation" : prp.Value = 0
                        'Case "Planar" : prp.Value = 0
                End Select
            Next
        Catch ex As Exception
            CD.ShowItemProperties(dev.Items(1), True)
        End Try
        Return New Bitmap(Image.FromStream(New IO.MemoryStream(CType(CD.ShowTransfer(dev.Items(1)).FileData.BinaryData, Byte()))))
    End Function

End Module

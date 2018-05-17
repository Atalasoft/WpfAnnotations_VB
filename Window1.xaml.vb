Imports Microsoft.Win32
Imports Atalasoft.Annotate.Wpf
Imports Atalasoft.Annotate.Icons
Imports Atalasoft.Annotate
Imports System.IO
Imports Atalasoft.Imaging.Codec
Imports System.Reflection
Imports Atalasoft.Imaging


Class Window1
    Private _colors As Dictionary(Of String, Color) = New Dictionary(Of String, Color)()
    Private _printScaleMode As WpfVisualScaleMode = WpfVisualScaleMode.ScaleToFit
    Private _mirroredTextEditing As Boolean
    Private _rotatedTextEditing As Boolean
    Private _lastPdfMarkup As Atalasoft.Annotate.Pdf.PdfMarkupType = Atalasoft.Annotate.Pdf.PdfMarkupType.Highlight
    Private _pdfMarkupButton As Button
    Private _pdfReadSupport As Boolean

    Shared Sub New()
        AtalaDemos.HelperMethods.PopulateDecoders(RegisteredDecoders.Decoders)
    End Sub

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        FillColorCombos()

        AddHandler AnnotationViewer.Annotations.HotSpotClicked, AddressOf Annotations_HotSpotClicked
        AddHandler AnnotationViewer.Annotations.SelectedAnnotationsChanged, AddressOf Annotations_SelectedAnnotationsChanged
        Me.AnnotationViewer.ImageViewer.IsCentered = True

        Me.AnnotationViewer.Annotations.Factories.Add(New WpfAnnotationUIFactory(Of TriangleAnnotation, TriangleData)())

        PrepareToolbar()

        Try
            ' This requires a PDF Reader license.
            Dim decoder As New Atalasoft.Imaging.Codec.Pdf.PdfDecoder()
            decoder.RenderSettings.AnnotationSettings = Atalasoft.Imaging.Codec.Pdf.AnnotationRenderSettings.None
            Atalasoft.Imaging.Codec.RegisteredDecoders.Decoders.Add(decoder)
            _pdfReadSupport = True
        Catch
        End Try
    End Sub

#Region "Toolbar"

    Private Sub PrepareToolbar()
        Me.AnnotationToolbar.Items.Add(CreateToolbarButton("Callout", AnnotateIcon.Callout))
        Me.AnnotationToolbar.Items.Add(CreateToolbarButton("Ellipse", AnnotateIcon.Ellipse))
        Me.AnnotationToolbar.Items.Add(CreateToolbarButton("Embedded Image", AnnotateIcon.EmbeddedImage))
        Me.AnnotationToolbar.Items.Add(CreateToolbarButton("Freehand", AnnotateIcon.Freehand))
        Me.AnnotationToolbar.Items.Add(CreateToolbarButton("Hot Spot", AnnotateIcon.RectangleHotspot))
        Me.AnnotationToolbar.Items.Add(CreateToolbarButton("Hot Spot Freehand", AnnotateIcon.FreehandHotspot))
        Me.AnnotationToolbar.Items.Add(CreateToolbarButton("Line", AnnotateIcon.Line))
        Me.AnnotationToolbar.Items.Add(CreateToolbarButton("Lines", AnnotateIcon.Lines))
        Me.AnnotationToolbar.Items.Add(CreateToolbarButton("Polygon", AnnotateIcon.Polygon))
        Me.AnnotationToolbar.Items.Add(CreateToolbarButton("Rectangle", AnnotateIcon.Rectangle))
        Me.AnnotationToolbar.Items.Add(CreateToolbarButton("Referenced Image", AnnotateIcon.ReferencedImage))
        Me.AnnotationToolbar.Items.Add(CreateToolbarButton("Rubber Stamp", AnnotateIcon.RubberStamp))
        Me.AnnotationToolbar.Items.Add(CreateToolbarButton("Text", AnnotateIcon.Text))
        Me.AnnotationToolbar.Items.Add(CreateToolbarButton("PDF Line", AnnotateIcon.PdfLine))
        _pdfMarkupButton = CreateToolbarButton("PDF Markup (Highlight)", AnnotateIcon.PdfMarkup)
        Me.AnnotationToolbar.Items.Add(_pdfMarkupButton)
    End Sub

    Private Function CreateToolbarButton(ByVal tooltip As String, ByVal icon As AnnotateIcon) As Button
        Dim img As New Image()
        img.Source = ExtractAnnotationIcon(icon, AnnotateIconSize.Size24)

        Dim btn As New Button()
        btn.Content = img
        btn.ToolTip = tooltip
        AddHandler btn.Click, AddressOf AnnotationToolbar_Click

        Return btn
    End Function

    Private Sub AnnotationToolbar_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim btn As Button = TryCast(sender, Button)
        If btn Is Nothing Then
            Return
        End If

        Select Case CStr(btn.ToolTip)
            Case "Callout"
                AddCallout()
            Case "Ellipse"
                AddEllipse()
            Case "Embedded Image"
                AddEmbeddedImage()
            Case "Freehand"
                AddFreehand()
            Case "Hot Spot"
                AddHotSpot()
            Case "Hot Spot Freehand"
                AddHotSpotFreehand()
            Case "Line"
                AddLine()
            Case "Lines"
                AddLines()
            Case "Polygon"
                AddPolygon()
            Case "Rectangle"
                AddRectangle()
            Case "Referenced Image"
                AddReferencedImage()
            Case "Rubber Stamp"
                AddRubberStamp()
            Case "Text"
                AddTextAnnotation()
            Case "PDF Line"
                AddPdfLine()
            Case "PDF Markup (Highlight)", "PDF Markup (StrikeOut)", "PDF Markup (Underline)", "PDF Markup (Squiggly)"
                AddPdfMarkup(_lastPdfMarkup)
        End Select
    End Sub

    Private Function ExtractAnnotationIcon(ByVal icon As AnnotateIcon, ByVal size As AnnotateIconSize) As BitmapSource
        Dim returnSource As BitmapSource = WpfObjectConverter.ConvertBitmap(DirectCast(IconResource.ExtractAnnotationIcon(icon, size), System.Drawing.Bitmap))

        If returnSource Is Nothing Then
            Dim assm As Assembly = Assembly.LoadFrom("Atalasoft.dotImage.dll")
            If assm IsNot Nothing Then
                Dim stream As System.IO.Stream = assm.GetManifestResourceStream("Atalasoft.Imaging.Annotate.Icons._" + size.ToString().Substring(4) + "." + icon.ToString() + ".png")
                returnSource = WpfObjectConverter.ConvertBitmap(DirectCast(System.Drawing.Image.FromStream(stream), System.Drawing.Bitmap))
            End If
            ' if it's STILL null, then give up and make placeholders
            If returnSource Is Nothing Then
                Select Case size.ToString()
                    Case "size16"
                        returnSource = WpfObjectConverter.ConvertBitmap(New AtalaImage(16, 16, Atalasoft.Imaging.PixelFormat.Pixel24bppBgr, System.Drawing.Color.White).ToBitmap())
                        Exit Select
                    Case "size24"
                        returnSource = WpfObjectConverter.ConvertBitmap(New AtalaImage(24, 24, Atalasoft.Imaging.PixelFormat.Pixel24bppBgr, System.Drawing.Color.White).ToBitmap())
                        Exit Select
                    Case "size32"
                        returnSource = WpfObjectConverter.ConvertBitmap(New AtalaImage(32, 32, Atalasoft.Imaging.PixelFormat.Pixel24bppBgr, System.Drawing.Color.White).ToBitmap())
                        Exit Select
                End Select
            End If
        End If

        Return returnSource
    End Function

#End Region

#Region "Color Combobox"

    Private Sub FillColorCombos()
        _colors.Add("AliceBlue", Colors.AliceBlue)
        _colors.Add("Black", Colors.Black)
        _colors.Add("Blue", Colors.Blue)
        _colors.Add("DarkBlue", Colors.DarkBlue)
        _colors.Add("Green", Colors.Green)
        _colors.Add("Lime", Colors.Lime)
        _colors.Add("Maroon", Colors.Maroon)
        _colors.Add("Orange", Colors.Orange)
        _colors.Add("Red", Colors.Red)
        _colors.Add("Transparent", Colors.Transparent)
        _colors.Add("White", Colors.White)
        _colors.Add("Yellow", Colors.Yellow)

        Me.FillComboBox.ItemsSource = _colors
        Me.OutlineComboBox.ItemsSource = _colors
    End Sub

    Private Sub SelectAddComboBoxItem(ByVal comboBox As ComboBox, ByVal color As Color)
        If (Not _colors.ContainsValue(color)) Then
            _colors.Add(color.ToString(), color)
        End If

        Dim count As Integer = comboBox.Items.Count
        For i As Integer = 0 To count - 1
            Dim item As KeyValuePair(Of String, Color) = CType(comboBox.Items(i), KeyValuePair(Of String, Color))
            If item.Value = color Then
                comboBox.SelectedIndex = i
                Return
            End If
        Next i
    End Sub

#End Region

#Region "Annotation Events"

    Private Sub Annotations_SelectedAnnotationsChanged(ByVal sender As Object, ByVal e As EventArgs)
        Me.statusInfo.Content = "Selection Changed: {Count=" & Me.AnnotationViewer.Annotations.SelectedAnnotations.Count.ToString() & "}"

        ' Update the property fields.
        Dim ann As WpfAnnotationUI = Me.AnnotationViewer.Annotations.ActiveAnnotation

        Dim enabled As Boolean = (ann IsNot Nothing)
        Me.textLocation.IsEnabled = enabled
        Me.textSize.IsEnabled = enabled
        Me.FillComboBox.IsEnabled = enabled
        Me.OutlineComboBox.IsEnabled = enabled
        Me.PenSizeSlider.IsEnabled = enabled
        Me.ShadowCheckBox.IsEnabled = enabled
        Me.OffsetSlider.IsEnabled = enabled

        If enabled Then
            Me.textLocation.Text = ann.Location.X.ToString() & ", " & ann.Location.Y.ToString()
            Me.textSize.Text = ann.Size.Width.ToString() & ", " & ann.Size.Height.ToString()

            Dim b As AnnotationBrush = CType(ann.GetValue(WpfAnnotationUI.FillProperty), AnnotationBrush)
            If b IsNot Nothing Then
                SelectAddComboBoxItem(Me.FillComboBox, WpfObjectConverter.ConvertColor(b.Color))
            End If

            Dim p As AnnotationPen = CType(ann.GetValue(WpfAnnotationUI.OutlineProperty), AnnotationPen)
            If p IsNot Nothing Then
                SelectAddComboBoxItem(Me.OutlineComboBox, WpfObjectConverter.ConvertColor(p.Color))
                Me.PenSizeSlider.Value = p.Width
            End If

            b = CType(ann.GetValue(WpfAnnotationUI.ShadowProperty), AnnotationBrush)
            Me.ShadowCheckBox.IsChecked = (b IsNot Nothing)

            If b IsNot Nothing Then
                Dim pt As Point = CType(ann.GetValue(WpfAnnotationUI.ShadowOffsetProperty), Point)
                Me.OffsetSlider.Value = pt.X
            End If
        End If
    End Sub

    Private Sub Annotations_HotSpotClicked(ByVal sender As Object, ByVal e As EventArgs(Of WpfAnnotationUI))
        Me.statusInfo.Content = "Hot Spot clicked!"
    End Sub

#End Region

#Region "File Menu"

    Private Sub FileCmdExecuted(ByVal sender As Object, ByVal e As ExecutedRoutedEventArgs)
        Dim rc As RoutedCommand = TryCast(e.Command, RoutedCommand)
        If rc Is Nothing Then
            Return
        End If

        Select Case rc.Name
            Case "Open"
                Dim dlg As New OpenFileDialog()
                dlg.Filter = AtalaDemos.HelperMethods.CreateDialogFilter(True)
                If dlg.ShowDialog().Value Then
                    Me.AnnotationViewer.Open(dlg.FileName, 0, Nothing)
                End If
            Case "SaveAs"
                Dim save As New SaveFileDialog()
                save.Filter = "JPEG (*.jpg)|*.jpg|TIFF (*.tif)|*.tif|PDF (*.pdf)|*.pdf"
                If save.ShowDialog().Value Then
                    If save.FilterIndex = 1 Then
                        SaveImage(save.FileName, New Atalasoft.Imaging.Codec.JpegEncoder())
                    ElseIf save.FilterIndex = 2 Then
                        SaveImage(save.FileName, New Atalasoft.Imaging.Codec.TiffEncoder())
                    Else
                        If Atalasoft.Imaging.AtalaImage.Edition = Atalasoft.Imaging.LicenseEdition.Document Then
                            SaveImage(save.FileName, New Atalasoft.Imaging.Codec.Pdf.PdfEncoder())
                            SavePdfAnnotations(save.FileName)
                        Else
                            MessageBox.Show("A 'DotImage Document Imaging' license is required to save as PDF.", "License Required", MessageBoxButton.OK, MessageBoxImage.Information)
                        End If
                    End If
                End If
            Case "Print"
                Dim pdlg As New PrintDialog()
                If pdlg.ShowDialog().Value Then
                    Dim sz As System.Printing.PageMediaSize = pdlg.PrintTicket.PageMediaSize
                    pdlg.PrintVisual(Me.AnnotationViewer.CreateVisual(New Size(If(sz.Width.HasValue, sz.Width.Value, 0), If(sz.Height.HasValue, sz.Height.Value, 0)), _printScaleMode, New Thickness(10)), "WPF Annotation Printing")
                End If
        End Select
    End Sub

    Private Sub FileCmdCanExecute(ByVal sender As Object, ByVal e As CanExecuteRoutedEventArgs)
        Dim rc As RoutedCommand = TryCast(e.Command, RoutedCommand)
        If rc Is Nothing Then
            Return
        End If

        Select Case rc.Name
            Case "Open"
                e.CanExecute = True
            Case "SaveAs"
                e.CanExecute = (Me.AnnotationViewer.ImageViewer.Image IsNot Nothing)
            Case "Print"
                ' You can print the annotations without an image.
                e.CanExecute = (Me.AnnotationViewer.ImageViewer.Image IsNot Nothing OrElse Me.AnnotationViewer.Annotations.CountAnnotations() > 0)
        End Select
    End Sub

    Private Sub SaveImage(ByVal fileName As String, ByVal encoder As Atalasoft.Imaging.Codec.ImageEncoder)
        Me.AnnotationViewer.Save(fileName, encoder, Nothing)
    End Sub

    Private Sub SavePdfAnnotations(ByVal fileName As String)
        Using fs As New FileStream(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite)
            Dim exp As New Atalasoft.Annotate.Exporters.PdfAnnotationDataExporter()

            Dim sz As System.Drawing.Size = Me.AnnotationViewer.ImageViewer.Image.Size
            Dim pageSize As New System.Drawing.SizeF(sz.Width, sz.Height)
            Dim resolution As Atalasoft.Imaging.Dpi = Me.AnnotationViewer.ImageViewer.Image.Resolution
            Dim data As Atalasoft.Annotate.LayerData = CType(Me.AnnotationViewer.Annotations.CurrentLayer.CreateDataSnapshot(), Atalasoft.Annotate.LayerData)

            exp.ExportOver(fs, pageSize, AnnotationUnit.Pixel, resolution, data, 0)
        End Using
    End Sub

    Private Sub AutoLoadXmp_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Me.AnnotationViewer.AutoLoadXmp = (CType(sender, MenuItem)).IsChecked
    End Sub

    Private Sub AutoLoadWang_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Me.AnnotationViewer.AutoLoadWang = (CType(sender, MenuItem)).IsChecked
    End Sub

    Private Sub AutoSaveXmp_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Me.AnnotationViewer.AutoSaveXmp = (CType(sender, MenuItem)).IsChecked
    End Sub

    Private Sub AutoSaveWang_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Me.AnnotationViewer.AutoSaveWang = (CType(sender, MenuItem)).IsChecked
    End Sub

    Private Sub Burn_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Me.AnnotationViewer.Burn()
        Me.AnnotationViewer.Annotations.Layers.Clear()
    End Sub

    Private Sub PrintScaleMode_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Me.PrintScaleModeFill.IsChecked = False
        Me.PrintScaleModeFit.IsChecked = False
        Me.PrintScaleModeNone.IsChecked = False
        Me.PrintScaleModeStretch.IsChecked = False

        Dim item As MenuItem = TryCast(sender, MenuItem)
        item.IsChecked = True

        Dim header As String = CStr(item.Header)
        Select Case header
            Case "None"
                _printScaleMode = WpfVisualScaleMode.None
            Case "Scale To Fit"
                _printScaleMode = WpfVisualScaleMode.ScaleToFit
            Case "Scale To Fill"
                _printScaleMode = WpfVisualScaleMode.ScaleToFill
            Case "Stretch To Fill"
                _printScaleMode = WpfVisualScaleMode.StretchToFill
        End Select
    End Sub

    Private Sub Exit_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Me.Close()
    End Sub

#End Region

#Region "Edit Menu"

    Private Sub EditCmdExecuted(ByVal sender As Object, ByVal e As ExecutedRoutedEventArgs)
        Dim rc As RoutedCommand = TryCast(e.Command, RoutedCommand)
        If rc Is Nothing Then
            Return
        End If

        Select Case rc.Name
            Case "Cut"
                Me.AnnotationViewer.Annotations.Cut()
            Case "Copy"
                Me.AnnotationViewer.Annotations.Copy()
            Case "Paste"
                Me.AnnotationViewer.Annotations.Paste()
        End Select
    End Sub

    Private Sub EditCmdCanExecute(ByVal sender As Object, ByVal e As CanExecuteRoutedEventArgs)
        Dim rc As RoutedCommand = TryCast(e.Command, RoutedCommand)
        If rc Is Nothing Then
            Return
        End If

        Select Case rc.Name
            Case "Cut", "Copy"
                e.CanExecute = (Me.AnnotationViewer.Annotations.SelectedAnnotations.Count > 0)
            Case "Paste"
                e.CanExecute = Me.AnnotationViewer.Annotations.CanPaste
        End Select
    End Sub

    Private Sub RotateDocument_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Select Case CStr((CType(sender, MenuItem)).Header)
            Case "90"
                Me.AnnotationViewer.RotateDocument(DocumentRotation.Rotate90)
            Case "180"
                Me.AnnotationViewer.RotateDocument(DocumentRotation.Rotate180)
            Case "270"
                Me.AnnotationViewer.RotateDocument(DocumentRotation.Rotate270)
        End Select
    End Sub

#End Region

#Region "Viewer Menu"

    Private Sub CenterImage_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Me.AnnotationViewer.ImageViewer.IsCentered = (CType(sender, MenuItem)).IsChecked
    End Sub

    Private Sub SnapRotation_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        'Make it easy for end-users to hit a 45 degree angle when rotating.
        'Setting the RotationSnapInterval to 0 will disable this feature.
        Me.AnnotationViewer.Annotations.RotationSnapInterval = (If((CType(sender, MenuItem)).IsChecked, 45, 0))
        Me.AnnotationViewer.Annotations.RotationSnapThreshold = 6
    End Sub

    Private Sub Zoom_SelectionChanged(ByVal sender As Object, ByVal e As SelectionChangedEventArgs)
        If Me.zoomCombo.Text.Length = 0 Then
            Return
        End If

        Dim item As ComboBoxItem = CType(e.AddedItems(0), ComboBoxItem)
        Dim val As Double = Int32.Parse((CStr(item.Content)).Replace("%", "")) / 100.0
        Me.AnnotationViewer.ImageViewer.Zoom = val
    End Sub

#End Region

#Region "Annotations Menu"

    Private Sub LoadAnnotationData_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim dlg As New OpenFileDialog()
        dlg.Filter = "XMP Data (*.xml)|*.xml|WANG Data (*.wng)|*.wng"
        If dlg.ShowDialog().Value Then
            If dlg.FilterIndex = 1 Then
                Me.AnnotationViewer.Annotations.Load(dlg.FileName, New Atalasoft.Annotate.Formatters.XmpFormatter())
            Else
                Me.AnnotationViewer.Annotations.Load(dlg.FileName, New Atalasoft.Annotate.Formatters.WangFormatter())
            End If
        End If
    End Sub

    Private Sub SaveAnnotationData_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim dlg As New SaveFileDialog()
        dlg.Filter = "XMP Data (*.xml)|*.xml|WANG Data (*.wng)|*.wng"
        If dlg.ShowDialog().Value Then
            If dlg.FilterIndex = 1 Then
                Me.AnnotationViewer.Annotations.Save(dlg.FileName, New Atalasoft.Annotate.Formatters.XmpFormatter())
            Else
                Me.AnnotationViewer.Annotations.Save(dlg.FileName, New Atalasoft.Annotate.Formatters.WangFormatter())
            End If
        End If
    End Sub

#Region "Add Annotations"

    Private Sub AddAnnotation_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim item As MenuItem = TryCast(sender, MenuItem)
        If item Is Nothing Then
            Return
        End If

        Select Case CStr(item.Header)
            Case "Callout"
                AddCallout()
            Case "Ellipse"
                AddEllipse()
            Case "Embedded Image"
                AddEmbeddedImage()
            Case "Freehand"
                AddFreehand()
            Case "Hot Spot"
                AddHotSpot()
            Case "Hot Spot Freehand"
                AddHotSpotFreehand()
            Case "Line"
                AddLine()
            Case "Lines"
                AddLines()
            Case "Polygon"
                AddPolygon()
            Case "Rectangle"
                AddRectangle()
            Case "Referenced Image"
                AddReferencedImage()
            Case "Rubber Stamp"
                AddRubberStamp()
            Case "Text"
                AddTextAnnotation()
            Case "PDF Line"
                AddPdfLine()
            Case "PDF Highlight"
                AddPdfMarkup(Atalasoft.Annotate.Pdf.PdfMarkupType.Highlight)
            Case "PDF StrikeOut"
                AddPdfMarkup(Atalasoft.Annotate.Pdf.PdfMarkupType.StrikeOut)
            Case "PDF Underline"
                AddPdfMarkup(Atalasoft.Annotate.Pdf.PdfMarkupType.Underline)
            Case "PDF Squiggly"
                AddPdfMarkup(Atalasoft.Annotate.Pdf.PdfMarkupType.Squiggly)
            Case "Triangle (Custom)"
                Me.AnnotationViewer.Annotations.CreateAnnotation(New TriangleAnnotation(Nothing, New AnnotationBrush(System.Drawing.Color.Orange), New AnnotationPen(System.Drawing.Color.Gold, 1)))
        End Select
    End Sub

    Private Sub AddPdfMarkup(ByVal markupType As Atalasoft.Annotate.Pdf.PdfMarkupType)
        _lastPdfMarkup = markupType
        _pdfMarkupButton.ToolTip = "PDF Markup (" & markupType.ToString() & ")"

        If markupType = Atalasoft.Annotate.Pdf.PdfMarkupType.Highlight Then
            Me.AnnotationViewer.Annotations.CreateAnnotation(New Atalasoft.Annotate.Wpf.Pdf.WpfPdfMarkupAnnotation(Nothing, markupType, New AnnotationBrush(System.Drawing.Color.FromArgb(120, System.Drawing.Color.Yellow)), Nothing, "", System.Environment.UserName, DateTime.Now))
        Else
            Me.AnnotationViewer.Annotations.CreateAnnotation(New Atalasoft.Annotate.Wpf.Pdf.WpfPdfMarkupAnnotation(Nothing, markupType, Nothing, New AnnotationPen(System.Drawing.Color.Black, 1), "", System.Environment.UserName, DateTime.Now))
        End If
    End Sub

    Private Sub AddPdfLine()
        Me.AnnotationViewer.Annotations.CreateAnnotation(New Atalasoft.Annotate.Wpf.Pdf.WpfPdfLineAnnotation(New Point(), New Point(), True, Nothing, New AnnotationPen(System.Drawing.Color.Black), "PDF Line", System.Environment.UserName, DateTime.Now))
    End Sub

    Private Sub AddCallout()
        Dim ann As New WpfCalloutAnnotation("Callout", New AnnotationFont("Verdana", CSng(10)), New AnnotationBrush(System.Drawing.Color.White), New AnnotationPen(System.Drawing.Color.Black, 1), 20, New AnnotationBrush(System.Drawing.Color.Navy), New AnnotationPen(System.Drawing.Color.Black, 1))
        ann.CanEditMirrored = _mirroredTextEditing
        ann.CanEditRotated = _rotatedTextEditing
        Me.AnnotationViewer.Annotations.CreateAnnotation(ann)
    End Sub

    Private Sub AddEllipse()
        Dim ann As New WpfEllipseAnnotation(New System.Windows.Rect(0, 0, 0, 0), New Atalasoft.Annotate.AnnotationBrush(System.Drawing.Color.Green), New Atalasoft.Annotate.AnnotationPen(System.Drawing.Color.Gold, 1))
        Me.AnnotationViewer.Annotations.CreateAnnotation(ann)
    End Sub

    Private Sub AddEmbeddedImage()
        Dim dlg As New OpenFileDialog()
        dlg.Filter = "Images|*.jpg;*.png;*.tif"
        If dlg.ShowDialog().Value Then
            Dim image As New AnnotationImage(dlg.FileName)
            Dim ann As New WpfEmbeddedImageAnnotation(image, New Point(0, 0), Nothing, New Point(0, 0))
            Me.AnnotationViewer.Annotations.CreateAnnotation(ann)
        End If
    End Sub

    Private Sub AddFreehand()
        Dim ann As WpfFreehandAnnotation = New WpfFreehandAnnotation(New AnnotationPen(System.Drawing.Color.Red, 1), WpfFreehandLineType.Bezier, False)
        ann.GripMode = AnnotationGripMode.Rectangular
        Me.AnnotationViewer.Annotations.CreateAnnotation(ann)
    End Sub

    Private Sub AddHotSpot()
        Me.AnnotationViewer.Annotations.CreateAnnotation(New WpfHotSpotAnnotation(New AnnotationBrush(System.Drawing.Color.Firebrick), Nothing))
    End Sub

    Private Sub AddHotSpotFreehand()
        Me.AnnotationViewer.Annotations.CreateAnnotation(New WpfHotSpotFreehandAnnotation(New AnnotationBrush(System.Drawing.Color.Firebrick), WpfFreehandLineType.Bezier))
    End Sub

    Private Sub AddLine()
        Dim line As New WpfLineAnnotation(New Point(0, 0), New Point(0, 0), New AnnotationPen(System.Drawing.Color.Red, 2 * 1))
        line.Outline.StartCap = New AnnotationLineCap(AnnotationLineCapStyle.ReversedFilledArrow, New System.Drawing.SizeF(12 * 1, 12))
        line.Outline.EndCap = New AnnotationLineCap(AnnotationLineCapStyle.FilledDiamond, New System.Drawing.SizeF(12 * 1, 12))
        Me.AnnotationViewer.Annotations.CreateAnnotation(line)
    End Sub

    Private Sub AddLines()
        Me.AnnotationViewer.Annotations.CreateAnnotation(New WpfLinesAnnotation(Nothing, New AnnotationPen(System.Drawing.Color.Red, 2 * 1)))
    End Sub

    Private Sub AddPolygon()
        Me.AnnotationViewer.Annotations.CreateAnnotation(New WpfPolygonAnnotation(Nothing, New AnnotationBrush(System.Drawing.Color.Orange), New AnnotationPen(System.Drawing.Color.Black, 1), Nothing, New Point(0, 0)))
    End Sub

    Private Sub AddRectangle()
        Dim ann As New WpfRectangleAnnotation(New System.Windows.Rect(0, 0, 0, 0), New Atalasoft.Annotate.AnnotationBrush(System.Drawing.Color.Red), New Atalasoft.Annotate.AnnotationPen(System.Drawing.Color.Maroon, 1))
        Me.AnnotationViewer.Annotations.CreateAnnotation(ann)
    End Sub

    Private Sub AddReferencedImage()
        Dim dlg As New OpenFileDialog()
        dlg.Filter = "Images|*.jpg;*.png;*.tif"
        If dlg.ShowDialog().Value Then
            Dim ann As New WpfReferencedImageAnnotation(dlg.FileName, New Point(0, 0), Nothing, New Point(0, 0))
            Me.AnnotationViewer.Annotations.CreateAnnotation(ann)
        End If
    End Sub

    Private Sub AddRubberStamp()
        Me.AnnotationViewer.Annotations.CreateAnnotation(New WpfRubberStampAnnotation("TOP SECRET", New AnnotationFont("Verdana", 24 * 1), New AnnotationBrush(System.Drawing.Color.Firebrick), Nothing, New AnnotationPen(System.Drawing.Color.Firebrick, 10 * 1), 30, 4))
    End Sub

    Private Sub AddTextAnnotation()
        Dim txt As New WpfTextAnnotation("Testing", New Atalasoft.Annotate.AnnotationFont("Verdana", 12), New Atalasoft.Annotate.AnnotationBrush(System.Drawing.Color.Red), New Atalasoft.Annotate.AnnotationBrush(System.Drawing.Color.White), New Atalasoft.Annotate.AnnotationPen(System.Drawing.Color.Black, 1))
        txt.Shadow = New Atalasoft.Annotate.AnnotationBrush(System.Drawing.Color.FromArgb(120, System.Drawing.Color.Silver))
        txt.ShadowOffset = New Point(6, 6)
        txt.CanEditMirrored = _mirroredTextEditing
        txt.CanEditRotated = _rotatedTextEditing
        Me.AnnotationViewer.Annotations.CreateAnnotation(txt)
    End Sub

#End Region

    Private Sub Clear_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Me.AnnotationViewer.Annotations.Layers.Clear()
    End Sub

    Private Sub ClipToDocument_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Me.AnnotationViewer.Annotations.ClipToDocument = (CType(sender, MenuItem)).IsChecked
    End Sub

    Private Sub FlipHorizontal_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        FlipAnnotation(True, False)
    End Sub

    Private Sub FlipHorizontalPivot_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        FlipAnnotation(True, True)
    End Sub

    Private Sub FlipVertical_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        FlipAnnotation(False, False)
    End Sub

    Private Sub FlipVerticalPivot_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        FlipAnnotation(False, True)
    End Sub

    Private Sub FlipAnnotation(ByVal horizontal As Boolean, ByVal pivot As Boolean)
        Dim ann As WpfAnnotationUI = Me.AnnotationViewer.Annotations.ActiveAnnotation
        If ann Is Nothing Then
            Return
        End If

        ann.Mirror((If(horizontal, MirrorDirection.Horizontal, MirrorDirection.Vertical)), (Not pivot))
    End Sub

    Private Sub GripMode_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim ann As WpfPointBaseAnnotation = TryCast(Me.AnnotationViewer.Annotations.ActiveAnnotation, WpfPointBaseAnnotation)
        If ann IsNot Nothing Then
            Dim item As MenuItem = TryCast(sender, MenuItem)
            If CStr(item.Header) = "Rectangular" Then
                ann.GripMode = AnnotationGripMode.Rectangular
            Else
                ann.GripMode = AnnotationGripMode.Points
            End If
        End If
    End Sub

    Private Sub InteractMode_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim item As MenuItem = TryCast(sender, MenuItem)
        Select Case CStr(item.Header)
            Case "None"
                Me.AnnotationViewer.Annotations.InteractMode = AnnotateInteractMode.None
            Case "Author"
                Me.AnnotationViewer.Annotations.InteractMode = AnnotateInteractMode.Author
            Case "View"
                Me.AnnotationViewer.Annotations.InteractMode = AnnotateInteractMode.View
        End Select
    End Sub

    Private Sub Group_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Me.AnnotationViewer.Annotations.Group()
    End Sub

    Private Sub Ungroup_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Me.AnnotationViewer.Annotations.Ungroup()
    End Sub

    Private Sub RemoveSelected_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        If Me.AnnotationViewer.Annotations.ActiveAnnotation IsNot Nothing Then
            Me.AnnotationViewer.Annotations.ActiveAnnotation.Remove()
        End If
    End Sub

    Private Sub SelectAll_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Me.AnnotationViewer.Annotations.SelectAll(True)
    End Sub

    Private Sub MirroredTextEditing_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        _mirroredTextEditing = (CType(sender, MenuItem)).IsChecked
        UpdateTextAnnotationEditingOptions()
    End Sub

    Private Sub RotatedTextEditing_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        _rotatedTextEditing = (CType(sender, MenuItem)).IsChecked
        UpdateTextAnnotationEditingOptions()
    End Sub

    Private Sub UpdateTextAnnotationEditingOptions()
        For Each layer As WpfLayerAnnotation In Me.AnnotationViewer.Annotations.Layers
            For Each ann As WpfAnnotationUI In layer.Items
                Dim txt As WpfTextAnnotation = TryCast(ann, WpfTextAnnotation)
                If txt IsNot Nothing Then
                    txt.CanEditMirrored = _mirroredTextEditing
                    txt.CanEditRotated = _rotatedTextEditing
                    If txt.EditMode Then
                        txt.EditMode = False
                        txt.EditMode = True
                    End If
                    Continue For
                End If

                Dim [call] As WpfCalloutAnnotation = TryCast(ann, WpfCalloutAnnotation)
                If [call] IsNot Nothing Then
                    [call].CanEditMirrored = _mirroredTextEditing
                    [call].CanEditRotated = _rotatedTextEditing
                    If [call].EditMode Then
                        [call].EditMode = False
                        [call].EditMode = True
                    End If
                End If
            Next ann
        Next layer
    End Sub

#End Region

#Region "Annotation Property Box"

    Private Sub ProperyExpander_Expanded(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Me.ProperyExpander.Width = 185
        'Add clipping so the grips will not render over the expander.
        Me.AnnotationViewer.Annotations.GripClip = New RectangleGeometry(New System.Windows.Rect(160, 0, Me.AnnotationViewer.Width - 160, Me.AnnotationViewer.Height))
    End Sub

    Private Sub ProperyExpander_Collapsed(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Me.ProperyExpander.Width = 24
        Me.AnnotationViewer.Annotations.GripClip = New RectangleGeometry(New System.Windows.Rect(0, 0, Me.AnnotationViewer.Width, Me.AnnotationViewer.Height))
    End Sub

    Private Sub textLocation_TextChanged(ByVal sender As Object, ByVal e As TextChangedEventArgs)
        Dim ann As WpfAnnotationUI = Me.AnnotationViewer.Annotations.ActiveAnnotation
        If ann Is Nothing Then
            Return
        End If

        Dim txt As String = Me.textLocation.Text
        If txt.Length = 0 Then
            Return
        End If

        Dim items() As String = txt.Split(","c)
        If items.Length <> 2 Then
            Return
        End If

        Dim x As Double = 0
        Dim y As Double = 0
        If Double.TryParse(items(0).Trim(), x) AndAlso Double.TryParse(items(1).Trim(), y) Then
            ann.Location = New Point(x, y)
        End If
    End Sub

    Private Sub textSize_TextChanged(ByVal sender As Object, ByVal e As TextChangedEventArgs)
        Dim ann As WpfAnnotationUI = Me.AnnotationViewer.Annotations.ActiveAnnotation
        If ann Is Nothing Then
            Return
        End If

        Dim txt As String = Me.textSize.Text
        If txt.Length = 0 Then
            Return
        End If

        Dim items() As String = txt.Split(","c)
        If items.Length <> 2 Then
            Return
        End If

        Dim width As Double = 0
        Dim height As Double = 0
        If Double.TryParse(items(0).Trim(), width) AndAlso Double.TryParse(items(1).Trim(), height) Then
            ann.Size = New Size(width, height)
        End If
    End Sub

    Private Sub PenSizeSlider_ValueChanged(ByVal sender As Object, ByVal e As RoutedPropertyChangedEventArgs(Of Double))
        Dim ann As WpfAnnotationUI = Me.AnnotationViewer.Annotations.ActiveAnnotation
        If ann Is Nothing Then
            Return
        End If

        If Me.PenSizeSlider.Value = 0 Then
            ann.SetValue(WpfAnnotationUI.OutlineProperty, Nothing)
        Else
            Dim pen As AnnotationPen = CType(ann.GetValue(WpfAnnotationUI.OutlineProperty), AnnotationPen)
            If pen Is Nothing Then
                Dim clr As Color = Colors.Black
                If Me.OutlineComboBox.SelectedIndex > -1 Then
                    clr = (CType(Me.OutlineComboBox.SelectedItem, KeyValuePair(Of String, Color))).Value
                End If

                pen = New AnnotationPen(WpfObjectConverter.ConvertColor(clr), CSng(Me.PenSizeSlider.Value))
                ann.SetValue(WpfAnnotationUI.OutlineProperty, pen)
            Else
                pen.Width = CSng(Me.PenSizeSlider.Value)
            End If
        End If
    End Sub

    Private Sub OutlineComboBox_SelectionChanged(ByVal sender As Object, ByVal e As SelectionChangedEventArgs)
        Dim ann As WpfAnnotationUI = Me.AnnotationViewer.Annotations.ActiveAnnotation
        If ann Is Nothing Then
            Return
        End If

        Dim val As KeyValuePair(Of String, Color) = CType(Me.OutlineComboBox.SelectedItem, KeyValuePair(Of String, Color))

        Dim pen As AnnotationPen = CType(ann.GetValue(WpfAnnotationUI.OutlineProperty), AnnotationPen)
        If pen Is Nothing Then
            If Me.PenSizeSlider.Value > 0 Then
                pen = New AnnotationPen(WpfObjectConverter.ConvertColor(val.Value), CSng(Me.PenSizeSlider.Value))
                ann.SetValue(WpfAnnotationUI.OutlineProperty, pen)
            End If
        Else
            pen.Color = WpfObjectConverter.ConvertColor(val.Value)
        End If
    End Sub

    Private Sub FillComboBox_SelectionChanged(ByVal sender As Object, ByVal e As SelectionChangedEventArgs)
        Dim ann As WpfAnnotationUI = Me.AnnotationViewer.Annotations.ActiveAnnotation
        If ann Is Nothing Then
            Return
        End If

        Dim val As KeyValuePair(Of String, Color) = CType(Me.FillComboBox.SelectedItem, KeyValuePair(Of String, Color))

        Dim brush As AnnotationBrush = CType(ann.GetValue(WpfAnnotationUI.FillProperty), AnnotationBrush)
        If brush Is Nothing Then
            brush = New AnnotationBrush(WpfObjectConverter.ConvertColor(val.Value))
            ann.SetValue(WpfAnnotationUI.FillProperty, brush)
        Else
            brush.Color = WpfObjectConverter.ConvertColor(val.Value)
        End If
    End Sub

    Private Sub OffsetSlider_ValueChanged(ByVal sender As Object, ByVal e As RoutedPropertyChangedEventArgs(Of Double))
        Dim ann As WpfAnnotationUI = Me.AnnotationViewer.Annotations.ActiveAnnotation
        If ann IsNot Nothing Then
            ann.SetValue(WpfAnnotationUI.ShadowOffsetProperty, New Point(e.NewValue, e.NewValue))
        End If
    End Sub

    Private Sub ShadowCheckBox_Checked(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim ann As WpfAnnotationUI = Me.AnnotationViewer.Annotations.ActiveAnnotation
        If ann IsNot Nothing Then
            ann.SetValue(WpfAnnotationUI.ShadowProperty, New AnnotationBrush(System.Drawing.Color.FromArgb(120, System.Drawing.Color.Silver)))
        End If
    End Sub

    Private Sub ShadowCheckBox_Unchecked(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim ann As WpfAnnotationUI = Me.AnnotationViewer.Annotations.ActiveAnnotation
        If ann IsNot Nothing Then
            ann.SetValue(WpfAnnotationUI.ShadowProperty, Nothing)
        End If
    End Sub

#End Region

#Region "Multi-Select Annotations"

    ' Use the left CTRL key for multi-select.
    Protected Overrides Sub OnPreviewKeyDown(ByVal e As KeyEventArgs)
        If (Not e.IsRepeat) Then
            Me.AnnotationViewer.Annotations.SelectionMode = (If(e.Key = Key.LeftCtrl, WpfAnnotationSelectionMode.Multiple, WpfAnnotationSelectionMode.Single))
        End If

        MyBase.OnPreviewKeyDown(e)
    End Sub

    Protected Overrides Sub OnPreviewKeyUp(ByVal e As KeyEventArgs)
        Me.AnnotationViewer.Annotations.SelectionMode = WpfAnnotationSelectionMode.Single
        MyBase.OnPreviewKeyUp(e)
    End Sub

#End Region

    Private Sub Canvas_SizeChanged(ByVal sender As Object, ByVal e As SizeChangedEventArgs)
        Me.ProperyExpander.Height = e.NewSize.Height
        Me.AnnotationViewer.Height = e.NewSize.Height
        Me.AnnotationViewer.Width = e.NewSize.Width - 24
    End Sub

    Private Sub HelpAbout_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs)
        ' Instantiate the dialog box
        Dim dlg As About = New About("About Atalasoft WPF Annotations Demo", "WPF Annotations Demo")
        dlg.Description = "This is a WPF version of our very popular and powerful DotAnnotate Demo." & vbCrLf & vbCrLf & _
                              "The source code should provide a good working example of how yo use our AtalaAnnotationViewer (the WPF version of our AnnotateViewer Winforms control) and our annotations in a WPF application."
        dlg.Link = "www.atalasoft.com/Support/Sample-Applications"

        ' Configure the dialog box
        dlg.Owner = Me

        ' Open the dialog box modally 
        dlg.ShowDialog()
    End Sub
End Class
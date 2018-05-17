Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Reflection
Imports System.Windows
Imports Atalasoft.Imaging
Imports Atalasoft.Imaging.Codec
Imports Atalasoft.Imaging.Codec.CadCam
Imports Atalasoft.Imaging.Codec.Dicom
Imports Atalasoft.Imaging.Codec.Jbig2
Imports Atalasoft.Imaging.Codec.Jpeg2000
Imports Atalasoft.Imaging.Codec.Pdf
Imports Atalasoft.Imaging.Codec.Tiff
' In order to use the OfficeDecoder, you will need to
' 1) reference Atalasoft.dotImage.Office.dll
' 2) add the dlls from either
'       C:\Program Files (x86)\Atalasoft\DotImage 10.7\bin\PerceptiveDocumentFilters\intel-32
'       or
'       C:\Program Files (x86)\Atalasoft\DotImage 10.7\bin\PerceptiveDocumentFilters\intel-64
'   to the bin directory of this solution
' 3) uncomment the Imports statement for Atalasoft.Imaging.Codec.Office below
' 4) In the HelperMethods() method below, find the commented out try/catch entry related to OfficeDecoder and uncomment it
'Imports Atalasoft.Imaging.Codec.Office

Namespace AtalaDemos
    Public Structure ImageFormatInformation
        Public Filter As String
        Public Description As String
        Public Encoder As ImageEncoder
        Public Decoder As ImageDecoder

        Public Sub New(ByVal encoder_Renamed As ImageEncoder, ByVal description_Renamed As String, ByVal filter_Renamed As String)
            Me.Encoder = encoder_Renamed
            Me.Decoder = Nothing
            Me.Description = description_Renamed
            Me.Filter = filter_Renamed
        End Sub

        Public Sub New(ByVal decoder_Renamed As ImageDecoder, ByVal description_Renamed As String, ByVal filter_Renamed As String)
            Me.Encoder = Nothing
            Me.Decoder = decoder_Renamed
            Me.Description = description_Renamed
            Me.Filter = filter_Renamed
        End Sub
    End Structure

    ''' <summary>
    ''' A collection of static methods.
    ''' </summary>
    Public NotInheritable Class HelperMethods
        Private Shared _decoderImageFormats As System.Collections.ArrayList = New System.Collections.ArrayList
        Private Shared _encoderImageFormats As System.Collections.ArrayList = New System.Collections.ArrayList

        Shared Sub New()
            'Decoders
            _decoderImageFormats.Add(New ImageFormatInformation(New JpegDecoder, "Joint Photographic Experts Group (*.jpg)", "*.jpg"))
            _decoderImageFormats.Add(New ImageFormatInformation(New PngDecoder, "Portable Network Graphic (*.png)", "*.png"))
            _decoderImageFormats.Add(New ImageFormatInformation(New TiffDecoder, "Tagged Image File (*.tif, *.tiff)", "*.tif;*.tiff"))
            _decoderImageFormats.Add(New ImageFormatInformation(New PcxDecoder, "ZSoft PaintBrush (*.pcx)", "*.pcx"))
            _decoderImageFormats.Add(New ImageFormatInformation(New TgaDecoder, "Truevision Targa (*.tga)", "*.tga"))
            _decoderImageFormats.Add(New ImageFormatInformation(New BmpDecoder, "Windows Bitmap (*.bmp)", "*.bmp"))
            _decoderImageFormats.Add(New ImageFormatInformation(New WmfDecoder, "Windows Meta File (*.wmf)", "*.wmf"))
            _decoderImageFormats.Add(New ImageFormatInformation(New EmfDecoder, "Enhanced Windows Meta File (*.emf)", "*.emf"))
            _decoderImageFormats.Add(New ImageFormatInformation(New PsdDecoder, "Adobe (tm) Photoshop format (*.psd)", "*.psd"))
            _decoderImageFormats.Add(New ImageFormatInformation(New WbmpDecoder, "Wireless Bitmap (*.wbmp)", "*.wbmp"))
            _decoderImageFormats.Add(New ImageFormatInformation(New GifDecoder, "Graphics Interchange Format (*.gif)", "*.gif"))
            _decoderImageFormats.Add(New ImageFormatInformation(New TlaDecoder, "Smaller Animals TLA (*.tla)", "*.tla"))
            _decoderImageFormats.Add(New ImageFormatInformation(New PcdDecoder, "Kodak (tm) PhotoCD (*.pcd)", "*.pcd"))
            _decoderImageFormats.Add(New ImageFormatInformation(New RawDecoder, "RAW Images", "*.dcr;*.dng;*.eff;*.mrw;*.nef;*.orf;*.pef;*.raf;*.srf;*.x3f;*.crw;*.cr2;*.tif;*.ppm"))

            Try
                _decoderImageFormats.Add(New ImageFormatInformation(New DwgDecoder, "Cad/Cam (*.dwg, *.dxf)", "*.dwg;*.dxf"))
            Catch e1 As AtalasoftLicenseException
            End Try

            Try
                _decoderImageFormats.Add(New ImageFormatInformation(New DicomDecoder, "Dicom (*.dcm, *.dce)", "*.dcm;*.dce"))
            Catch e1 As AtalasoftLicenseException
            End Try

            Try
                _decoderImageFormats.Add(New ImageFormatInformation(New Jb2Decoder, "JBIG2 (*.jb2)", "*.jb2"))
            Catch e1 As AtalasoftLicenseException
            End Try

            Try
                _decoderImageFormats.Add(New ImageFormatInformation(New Jp2Decoder, "JPEG2000 (*.jpf, *.jp2, *.jpc, *.j2c, *.j2k)", "*.jpf;*.jp2;*.jpc;*.j2c;*.j2k"))
            Catch e1 As AtalasoftLicenseException
            End Try

            Try
                _decoderImageFormats.Add(New ImageFormatInformation(New PdfDecoder() With {.Resolution = 200, .RenderSettings = New RenderSettings() With {.AnnotationSettings = AnnotationRenderSettings.None}}, "PDF (*.pdf)", "*.pdf"))
            Catch e1 As AtalasoftLicenseException
            End Try

            '' OfficeDecoer only exists in 10.7 and newer. please see the instructions at the top of this file for enableing OfficeDecoder.. and uncomment the following:
            'Try
            '    _decoderImageFormats.Add(New ImageFormatInformation(New OfficeDecoder() With {.Resolution = 200}, "Office Doc (*.doc *.docx *.rtf *.xls *.xlsx *.ppt)", "*.doc;*.docx;*.rtf;*.xls;*.xlsx; *.ppt"))
            'Catch e1 As AtalasoftLicenseException
            'End Try


            'Encoders
            _encoderImageFormats.Add(New ImageFormatInformation(New JpegEncoder, "Joint Photographic Experts Group (*.jpg)", "*.jpg"))
            _encoderImageFormats.Add(New ImageFormatInformation(New PngEncoder, "Portable Network Graphic (*.png)", "*.png"))
            _encoderImageFormats.Add(New ImageFormatInformation(New TiffEncoder, "Tagged Image File (*.tif, *.tiff)", "*.tif;*.tiff"))
            _encoderImageFormats.Add(New ImageFormatInformation(New PcxEncoder, "ZSoft PaintBrush (*.pcx)", "*.pcx"))
            _encoderImageFormats.Add(New ImageFormatInformation(New TgaEncoder, "Truevision Targa (*.tga)", "*.tga"))
            _encoderImageFormats.Add(New ImageFormatInformation(New BmpEncoder, "Windows Bitmap (*.bmp)", "*.bmp"))
            _encoderImageFormats.Add(New ImageFormatInformation(New WmfEncoder, "Windows Meta File (*.wmf)", "*.wmf"))
            _encoderImageFormats.Add(New ImageFormatInformation(New EmfEncoder, "Enhanced Windows Meta File (*.emf)", "*.emf"))
            _encoderImageFormats.Add(New ImageFormatInformation(New PsdEncoder, "Adobe (tm) Photoshop format (*.psd)", "*.psd"))
            _encoderImageFormats.Add(New ImageFormatInformation(New WbmpEncoder, "Wireless Bitmap (*.wbmp)", "*.wbmp"))
            _encoderImageFormats.Add(New ImageFormatInformation(New GifEncoder, "Graphics Interchange Format (*.gif)", "*.gif"))
            _encoderImageFormats.Add(New ImageFormatInformation(New TlaEncoder, "Smaller Animals TLA (*.tla)", "*.tla"))

            Try
                _encoderImageFormats.Add(New ImageFormatInformation(New Jb2Encoder, "JBIG2 (*.jb2)", "*.jb2"))
            Catch e1 As AtalasoftLicenseException
            End Try

            Try
                _encoderImageFormats.Add(New ImageFormatInformation(New Jp2Encoder, "JPEG2000 (*.jpf, *.jp2, *.jpc, *.j2c, *.j2k)", "*.jpf;*.jp2;*.jpc;*.j2c;*.j2k"))
            Catch e1 As AtalasoftLicenseException
            End Try

            Try
                _encoderImageFormats.Add(New ImageFormatInformation(New PdfEncoder, "PDF (*.pdf)", "*.pdf"))
            Catch e1 As AtalasoftLicenseException
            End Try

        End Sub

        Public Shared Function HaveDecoder(ByVal dec As ImageDecoder) As Boolean
            For Each info As ImageFormatInformation In _decoderImageFormats
                If dec.GetType().Equals(info.Decoder.GetType()) Then
                    Return True
                End If
            Next

            Return False
        End Function

        Public Shared Function HaveEncoder(ByVal enc As ImageEncoder) As Boolean
            For Each info As ImageFormatInformation In _encoderImageFormats
                If enc.GetType().Equals(info.Encoder.GetType()) Then
                    Return True
                End If
            Next

            Return False
        End Function

        ''' <summary>
        ''' Use this when your demo needs a DotImage Document Imaging license
        ''' </summary>
        Public Shared Function HaveDotImage() As Boolean
            If Atalasoft.Imaging.AtalaImage.Edition <> LicenseEdition.Document Then
                LicenseCheckFailure("This demo requires a Document Imaging License.\r\nYour current license is for '" + AtalaImage.Edition.ToString() + "'.")
            End If
        End Function

        ''' <summary>
        ''' Use this when your demo needs a DotImage ADC add-on license
        ''' </summary>
        Public Shared Function HaveADC() As Boolean
            If Atalasoft.Licensing.AtalaLicenseProvider.GetLicenseFlag("Atalasoft.dotImage", "AdvancedDocClean") Is Nothing Then
                LicenseCheckFailure("This demo requires a DotImage Document Imaging license and an Advanced DocClean license.")
                Return False
            End If

            Return True
        End Function

        ''' <summary>
        ''' Creates the filter string for open and save dialogs.
        ''' </summary>
        ''' <param name="isOpenDialog">Set to true if this filter if for an open dialog.</param>
        ''' <returns>The filter string for an open or save dialog.</returns>
        Public Shared Function CreateDialogFilter(ByVal isOpenDialog As Boolean) As String
            Dim filter As System.Text.StringBuilder = New System.Text.StringBuilder

            If isOpenDialog Then
                'All supported formats
                filter.Append("All Supported Images|")
                For Each info As ImageFormatInformation In _decoderImageFormats
                    filter.Append(info.Filter & ";")
                Next info

                'Individual format filter
                For Each info As ImageFormatInformation In _decoderImageFormats
                    filter.Append("|" & info.Description & "|" & info.Filter)
                Next

                'Add all files filter, this will cover, e.g. Dicom files without extension
                filter.Append("|All files (*.*)|*.*")

            Else
                For Each info As ImageFormatInformation In _encoderImageFormats
                    filter.Append("|" & info.Description & "|" & info.Filter)
                Next

                filter.Append("|Animated GIF (*.gif)|*.gif")
                filter.Remove(0, 1) 'Remove leading "|"
            End If

            Return filter.ToString()
        End Function

        ''' <summary>
        ''' Based on licensed decoders, this will generate a search pattern string that may be used to search a directory for supported image formats
        ''' </summary>
        ''' <returns>search pattern string of supported extensions</returns>
        Public Shared Function GenerateDecoderSearchPattern() As String
            Dim pattern As New System.Text.StringBuilder
            For Each info As ImageFormatInformation In _decoderImageFormats
                pattern.Append(info.Filter & ";")
            Next

            pattern.Remove(pattern.Length - 1, 1) 'Remove trailing semicolon

            Return pattern.ToString
        End Function

        Public Shared Sub PopulateDecoders(ByVal col As DecoderCollection)
            For Each info As ImageFormatInformation In _decoderImageFormats
                If Not col.Contains(info.Decoder) Then
                    col.Add(info.Decoder)
                End If
            Next
        End Sub

        Public Shared Function GetImageEncoder(ByVal filterIndex As Integer) As ImageEncoder
        	' Since OpenFile/SaveFile dialogs are 1-indexed, we need to decrement this
        	filterIndex -= 1
            If filterIndex < 0 OrElse filterIndex >= _encoderImageFormats.Count Then
                Return Nothing
            End If

            Dim info As ImageFormatInformation = CType(_encoderImageFormats(filterIndex), ImageFormatInformation)
            Return info.Encoder
        End Function

        ''' <summary>
        ''' Fills the command arrays with method types.
        ''' </summary>
        ''' <param name="channelCommand"></param>
        ''' <param name="effectCommand"></param>
        ''' <param name="filterCommand"></param>
        ''' <param name="transformCommand"></param>
        Public Shared Sub GetReflectionMethods(<System.Runtime.InteropServices.Out()> ByRef channelCommand As Type(), <System.Runtime.InteropServices.Out()> ByRef effectCommand As Type(), <System.Runtime.InteropServices.Out()> ByRef filterCommand As Type(), <System.Runtime.InteropServices.Out()> ByRef transformCommand As Type())
            ' Just to make the compiler happy.
            channelCommand = Nothing
            effectCommand = Nothing
            filterCommand = Nothing
            transformCommand = Nothing

            ' Load the assumebly.
            Dim myAssembly As System.Reflection.Assembly = System.Reflection.Assembly.Load("Atalasoft.Imaging")
            If myAssembly Is Nothing Then
                Throw New ArgumentException("Unable to load the Atalasoft.Imaging assembly.")
            End If

            ' Get all of the assembly types.
            Dim myTypes As Type() = myAssembly.GetExportedTypes()

            ' Create temporary storage for the types.
            ' 100 elements each should be enough.
            Dim channels As Type() = New Type(99) {}
            Dim effects As Type() = New Type(99) {}
            Dim filters As Type() = New Type(99) {}
            Dim transforms As Type() = New Type(99) {}

            Dim channelCount As Integer = 0
            Dim effectCount As Integer = 0
            Dim filterCount As Integer = 0
            Dim transformCount As Integer = 0

            ' Loop through all of the types and fill out the arrays.
            For Each type As Type In myTypes
                If type.IsClass AndAlso type.IsPublic Then
                    Select Case type.Namespace
                        Case "Atalasoft.Imaging.Imaging.Channels"
                            channels(channelCount) = type
                            channelCount += 1
                        Case "Atalasoft.Imaging.Imaging.Effects"
                            effects(effectCount) = type
                            effectCount += 1
                        Case "Atalasoft.Imaging.Imaging.Filters"
                            filters(filterCount) = type
                            filterCount += 1
                        Case "Atalasoft.Imaging.Imaging.Transforms"
                            transforms(transformCount) = type
                            transformCount += 1
                    End Select
                End If
            Next type

            ' Copy the data to the arrays which were passed in.
            If channelCount > 0 Then
                channelCommand = New Type(channelCount - 1) {}
                Array.Copy(channels, 0, channelCommand, 0, channelCount)
            End If

            If effectCount > 0 Then
                effectCommand = New Type(effectCount - 1) {}
                Array.Copy(effects, 0, effectCommand, 0, effectCount)
            End If

            If filterCount > 0 Then
                filterCommand = New Type(filterCount - 1) {}
                Array.Copy(filters, 0, filterCommand, 0, filterCount)
            End If

            If transformCount > 0 Then
                transformCommand = New Type(transformCount - 1) {}
                Array.Copy(transforms, 0, transformCommand, 0, transformCount)
            End If

        End Sub

        ''' <summary>
        ''' Break apart the command name to make it more readable.
        ''' </summary>
        ''' <param name="commandName">Command name to separate.</param>
        ''' <returns>Formated command name.</returns>
        Public Shared Function SeparateCommandName(ByVal commandName As String) As String
            Dim letter As String = ""
            Dim nice As String = ""
            Dim lastPos As Integer = 0

            Dim i As Integer = 1
            'ORIGINAL LINE: for (int i = 1; i < commandName.Length; i += 1)
            'INSTANT VB NOTE: This 'for' loop was translated to a VB 'Do While' loop:
            Do While i < commandName.Length
                letter = commandName.Chars(i).ToString()
                If letter.ToUpper() = letter Then
                    nice &= (commandName.Substring(lastPos, i - lastPos) & " ")
                    lastPos = i
                End If
                i += 1
            Loop

            nice.Trim()
            Return nice
        End Function

        Public Shared Sub LicenseCheckFailure(ByVal message As String)
            If MessageBox.Show(Nothing, message + "\r\n\r\nWould you like to request an evaluation license?", "License Required", MessageBoxButton.YesNo, MessageBoxImage.Exclamation) = MessageBoxResult.Yes Then
                'Locate the activation utility
                Dim mPath As String = ""
                Dim key As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("Software\Atalasoft\dotImage\6.0")
                If Not key Is Nothing Then
                    mPath = Convert.ToString(key.GetValue("AssemblyBasePath"))
                    If Not mPath Is Nothing And mPath.Length > 5 Then
                        mPath = mPath.Substring(0, mPath.Length - 3) + "AtalasoftToolkitActivation.exe"
                    Else
                        mPath = Path.GetFullPath("C:\Program Files (x86)\Atalasoft\DotImage 10.7\AtalasoftToolkitActivation.exe")
                    End If
                    key.Close()
                End If
                If File.Exists(mPath) Then
                    System.Diagnostics.Process.Start(mPath)
                Else
                    MessageBox.Show(Nothing, "We were unable to location the DotImage activation utility.\r\nPlease run it from the Start menu shortcut.", "File Not Found")
                End If
            End If
        End Sub
    End Class
End Namespace



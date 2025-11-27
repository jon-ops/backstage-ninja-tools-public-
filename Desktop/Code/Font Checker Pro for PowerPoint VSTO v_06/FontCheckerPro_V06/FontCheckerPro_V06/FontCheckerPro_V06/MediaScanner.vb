' MediaScanner.vb
Option Strict On

Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Core
Imports System.IO
Imports System.Windows.Forms
Imports FontCheckerPro_V06
Imports System.IO.Compression
Imports System.Xml

Module MediaScanner
        ' Helper: get MediaType safely (returns Integer)
        Friend Function GetMediaTypeSafe(shape As Microsoft.Office.Interop.PowerPoint.Shape) As Integer
            Try
                Return CInt(shape.MediaType)
            Catch
                Return -1
            End Try
        End Function

        ' Helper: extract YouTube/Vimeo URL from alt text
        Friend Function ExtractOnlineVideoUrl(altText As String) As String
            If String.IsNullOrEmpty(altText) Then Return ""
            Dim urlPattern As String = "(https?://[^\u0000-\s"']+)"
            Dim m = System.Text.RegularExpressions.Regex.Match(altText, urlPattern)
            If m.Success AndAlso (m.Value.Contains("youtube.com") Or m.Value.Contains("youtu.be") Or m.Value.Contains("vimeo.com")) Then
                Return m.Value
            End If
            Return ""
        End Function
    Public Sub ShowMediaReport()
        Dim pres As Presentation = Nothing
        Try
            pres = Globals.ThisAddIn.Application.ActivePresentation
        Catch
        End Try

        If pres Is Nothing Then
            MessageBox.Show("No active presentation found.", "Media Report", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim items = GetMediaItems(pres)
        Dim frm As New MediaReportForm()
        frm.LoadData(items)
        frm.ShowDialog()
    End Sub


    Public Function GetMediaItems(pres As Presentation) As List(Of MediaScanResult)
        Dim embeddedSizes = GetEmbeddedMediaSizes(pres)
        Dim externalLinks = GetExternalMediaLinks(pres)

        Dim results As New List(Of MediaScanResult)()
        For Each slide As Microsoft.Office.Interop.PowerPoint.Slide In pres.Slides
            For Each shape As Microsoft.Office.Interop.PowerPoint.Shape In slide.Shapes
                If HasMedia(shape) Then
                    Try
                        Dim foundOnlineVideo As Boolean = False
                        Dim url As String = ""
                        ' Check for online video: msoMedia + ppMediaTypeOther or msoWebVideo
                        If (shape.Type = 16 AndAlso GetMediaTypeSafe(shape) = 3) Or shape.Type = 201 Then
                            url = ExtractOnlineVideoUrl(shape.AlternativeText)
                            If String.IsNullOrEmpty(url) Then
                                Try
                                    If shape.HasTextFrame = MsoTriState.msoTrue Then
                                        url = ExtractOnlineVideoUrl(shape.TextFrame.TextRange.Text)
                                    End If
                                Catch
                                End Try
                            End If
                            If Not String.IsNullOrEmpty(url) Then
                                Dim r As New MediaScanResult()
                                r.SlideIndex = slide.SlideIndex
                                r.ShapeName = shape.Name
                                r.ShapeId = shape.Id
                                r.MediaKind = "Video"
                                r.IsLinked = True
                                r.FileName = ""
                                r.FilePath = url
                                r.Duration = ""
                                r.Resolution = ""
                                r.FrameRate = ""
                                r.Compression = ""
                                r.EmbeddedFileSizeBytes = 0
                                results.Add(r)
                                foundOnlineVideo = True
                            End If
                        End If
                        If Not foundOnlineVideo Then
                            ' Try LinkFormat (some online videos may expose URL here)
                            Try
                                If shape.LinkFormat IsNot Nothing Then
                                    Dim srcTry = shape.LinkFormat.SourceFullName
                                    If Not String.IsNullOrEmpty(srcTry) AndAlso (srcTry.Contains("youtube.com") Or srcTry.Contains("youtu.be") Or srcTry.Contains("vimeo.com")) Then
                                        Dim r As New MediaScanResult()
                                        r.SlideIndex = slide.SlideIndex
                                        r.ShapeName = shape.Name
                                        r.ShapeId = shape.Id
                                        r.MediaKind = "Video"
                                        r.IsLinked = True
                                        r.FileName = ""
                                        r.FilePath = srcTry
                                        r.Duration = ""
                                        r.Resolution = ""
                                        r.FrameRate = ""
                                        r.Compression = ""
                                        r.EmbeddedFileSizeBytes = 0
                                        results.Add(r)
                                        Continue For
                                    End If
                                End If
                            Catch
                            End Try
                            ' If slide/shape has external link found in package, use that
                            Dim extUrl As String = Nothing
                            Dim slideMap As Dictionary(Of Integer, String) = Nothing
                            If externalLinks IsNot Nothing AndAlso externalLinks.TryGetValue(slide.SlideIndex, slideMap) Then
                                If slideMap.TryGetValue(shape.Id, extUrl) Then
                                    Dim r As New MediaScanResult()
                                    r.SlideIndex = slide.SlideIndex
                                    r.ShapeName = shape.Name
                                    r.ShapeId = shape.Id
                                    r.MediaKind = "Video"
                                    r.IsLinked = True
                                    r.FileName = ""
                                    r.FilePath = extUrl
                                    r.Duration = ""
                                    r.Resolution = ""
                                    r.FrameRate = ""
                                    r.Compression = ""
                                    r.EmbeddedFileSizeBytes = 0
                                    results.Add(r)
                                    Continue For
                                End If
                            End If
                            results.Add(BuildMediaResult(slide, shape, embeddedSizes))
                        End If
                    Catch
                        ' Skip problematic shapes
                    End Try
                    ' ...existing code...
                ' Detect online video (msoWebVideo or shape with YouTube/Vimeo hyperlink)
                ' msoWebVideo = 201
                ElseIf shape.Type = 201 Then
                    Dim r As New MediaScanResult()
                    r.SlideIndex = slide.SlideIndex
                    r.ShapeName = shape.Name
                    r.ShapeId = shape.Id
                    r.MediaKind = "Video"
                    r.IsLinked = True
                    r.FileName = ""
                    r.FilePath = ""
                    r.Duration = ""
                    r.Resolution = ""
                    r.FrameRate = ""
                    r.Compression = ""
                    r.EmbeddedFileSizeBytes = 0
                    results.Add(r)
                Else
                    ' Check for hyperlink or text/alt text containing YouTube/Vimeo on any shape
                    Try
                        Dim link As String = Nothing
                        If shape.ActionSettings(PpMouseActivation.ppMouseClick).Hyperlink IsNot Nothing Then
                            link = shape.ActionSettings(PpMouseActivation.ppMouseClick).Hyperlink.Address
                        End If
                        If String.IsNullOrEmpty(link) AndAlso Not String.IsNullOrEmpty(shape.AlternativeText) Then
                            link = shape.AlternativeText
                        End If
                        If String.IsNullOrEmpty(link) Then
                            Try
                                If shape.HasTextFrame = MsoTriState.msoTrue Then
                                    Dim txt = shape.TextFrame.TextRange.Text
                                    If Not String.IsNullOrEmpty(txt) Then
                                        link = txt
                                    End If
                                End If
                            Catch
                            End Try
                        End If
                        If Not String.IsNullOrEmpty(link) AndAlso (link.Contains("youtube.com") Or link.Contains("youtu.be") Or link.Contains("vimeo.com")) Then
                            Dim r As New MediaScanResult()
                            r.SlideIndex = slide.SlideIndex
                            r.ShapeName = shape.Name
                            r.ShapeId = shape.Id
                            r.MediaKind = "Video"
                            r.IsLinked = True
                            r.FileName = ""
                            r.FilePath = link
                            r.Duration = ""
                            r.Resolution = ""
                            r.FrameRate = ""
                            r.Compression = ""
                            r.EmbeddedFileSizeBytes = 0
                            results.Add(r)
                        End If
                    Catch
                    End Try
                End If
            Next
        Next
        Return results
    End Function

    Private Function HasMedia(shape As Microsoft.Office.Interop.PowerPoint.Shape) As Boolean
        Try
            If shape.Type = MsoShapeType.msoMedia Then
                Return True
            End If
            Dim mf = shape.MediaFormat
            Dim dummy = mf.Length ' Throws if not media
            Return True
        Catch
            Return False
        End Try
    End Function

    Private Function BuildMediaResult(slide As Microsoft.Office.Interop.PowerPoint.Slide, shape As Microsoft.Office.Interop.PowerPoint.Shape, embeddedSizes As Dictionary(Of Integer, Dictionary(Of Integer, Long))) As MediaScanResult
        Dim result As New MediaScanResult()
        result.SlideIndex = slide.SlideIndex
        result.ShapeName = shape.Name
        result.ShapeId = shape.Id

        Dim kind As String = ""
        Try
            Select Case shape.MediaType
                Case PpMediaType.ppMediaTypeMovie
                    kind = "Video"
                Case PpMediaType.ppMediaTypeSound
                    kind = "Audio"
                Case Else
                    kind = "Unknown"
            End Select
        Catch
            kind = "Unknown"
        End Try
        result.MediaKind = kind

        Dim mf = shape.MediaFormat
        result.IsLinked = False
        result.FileName = ""
        result.FilePath = ""

        Try
            result.IsLinked = mf.IsLinked
            If result.IsLinked Then
                Dim src As String = shape.LinkFormat.SourceFullName
                result.FilePath = src
                result.FileName = Path.GetFileName(src)
            End If
        Catch
            ' Embedded or error
        End Try

        ' Duration
        Try
            Dim ms = mf.Length
            Dim ts = TimeSpan.FromMilliseconds(ms)
            result.Duration = ts.ToString("hh\:mm\:ss")
        Catch
            result.Duration = ""
        End Try

        ' Video-specific
        If kind = "Video" Then
            Try
                result.Resolution = $"{mf.SampleWidth}x{mf.SampleHeight}"
            Catch
                result.Resolution = ""
            End Try
            Try
                result.FrameRate = $"{mf.VideoFrameRate} fps"
            Catch
                result.FrameRate = ""
            End Try
            Try
                result.Compression = mf.VideoCompressionType
            Catch
                result.Compression = ""
            End Try
        ElseIf kind = "Audio" Then
            result.Resolution = ""
            result.FrameRate = ""
            Try
                result.Compression = mf.AudioCompressionType
            Catch
                result.Compression = ""
            End Try
        Else
            result.Resolution = ""
            result.FrameRate = ""
            result.Compression = ""
        End If

        ' Embedded media size (if available)
        Try
            If Not result.IsLinked AndAlso embeddedSizes IsNot Nothing Then
                Dim slideMap As Dictionary(Of Integer, Long) = Nothing
                If embeddedSizes.TryGetValue(slide.SlideIndex, slideMap) Then
                    Dim size As Long = 0
                    If slideMap.TryGetValue(result.ShapeId, size) Then
                        result.EmbeddedFileSizeBytes = size
                    End If
                End If
            End If
        Catch
            ' best-effort; ignore errors
        End Try

        Return result
    End Function

    ' Map slide index + shape ID to embedded media file size in bytes
    Private Function GetEmbeddedMediaSizes(pres As Presentation) As Dictionary(Of Integer, Dictionary(Of Integer, Long))
        Dim sizes As New Dictionary(Of Integer, Dictionary(Of Integer, Long))()

        Dim presPath As String = ""
        Try
            presPath = pres.FullName
        Catch
        End Try
        If String.IsNullOrWhiteSpace(presPath) OrElse Not File.Exists(presPath) Then
            Return sizes
        End If

        Try
            Using archive = ZipFile.OpenRead(presPath)
                For slideIndex As Integer = 1 To pres.Slides.Count
                    Dim slideEntryName = $"ppt/slides/slide{slideIndex}.xml"
                    Dim slideRelsEntryName = $"ppt/slides/_rels/slide{slideIndex}.xml.rels"

                    Dim slideEntry = archive.GetEntry(slideEntryName)
                    If slideEntry Is Nothing Then
                        Continue For
                    End If

                    Dim relMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                    Dim relEntry = archive.GetEntry(slideRelsEntryName)
                    If relEntry IsNot Nothing Then
                        Dim relDoc As New XmlDocument()
                        Using relStream = relEntry.Open()
                            relDoc.Load(relStream)
                        End Using
                        Dim nsmgr As New XmlNamespaceManager(relDoc.NameTable)
                        nsmgr.AddNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships")

                        For Each relNode As XmlNode In relDoc.SelectNodes("//r:Relationship", nsmgr)
                            Dim idAttr = relNode.Attributes("Id")
                            Dim targetAttr = relNode.Attributes("Target")
                            If idAttr IsNot Nothing AndAlso targetAttr IsNot Nothing Then
                                relMap(idAttr.Value) = targetAttr.Value
                            End If
                        Next
                    End If

                    Dim slideDoc As New XmlDocument()
                    Using slideStream = slideEntry.Open()
                        slideDoc.Load(slideStream)
                    End Using

                    Dim mgr As New XmlNamespaceManager(slideDoc.NameTable)
                    mgr.AddNamespace("p", "http://schemas.openxmlformats.org/presentationml/2006/main")
                    mgr.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
                    mgr.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

                    Dim slideSizes As New Dictionary(Of Integer, Long)()

                    ' Find any element with an r:embed or r:link attribute
                    For Each attr As XmlAttribute In slideDoc.SelectNodes("//*/@r:embed | //*/@r:link", mgr)
                        Dim relId = attr.Value
                        If String.IsNullOrWhiteSpace(relId) OrElse Not relMap.ContainsKey(relId) Then
                            Continue For
                        End If

                        Dim shapeId = FindAncestorShapeId(attr.OwnerElement, mgr)
                        If shapeId <= 0 Then
                            Continue For
                        End If

                        Dim target = relMap(relId)
                        Dim partPath = NormalizePartPath(target)
                        Dim mediaEntry = archive.GetEntry(partPath)
                        If mediaEntry Is Nothing Then
                            Continue For
                        End If

                        If Not slideSizes.ContainsKey(shapeId) Then
                            slideSizes(shapeId) = mediaEntry.Length
                        End If
                    Next

                    If slideSizes.Count > 0 Then
                        sizes(slideIndex) = slideSizes
                    End If
                Next
            End Using
        Catch
            ' best-effort; ignore errors
        End Try

        Return sizes
    End Function

    ' Map slide index + shape ID to external media URL (e.g., YouTube/Vimeo) when stored as relationships with TargetMode="External"
    Private Function GetExternalMediaLinks(pres As Presentation) As Dictionary(Of Integer, Dictionary(Of Integer, String))
        Dim links As New Dictionary(Of Integer, Dictionary(Of Integer, String))()
        Dim presPath As String = ""
        Try
            presPath = pres.FullName
        Catch
        End Try
        If String.IsNullOrWhiteSpace(presPath) OrElse Not File.Exists(presPath) Then
            Return links
        End If

        Try
            Using archive = ZipFile.OpenRead(presPath)
                For slideIndex As Integer = 1 To pres.Slides.Count
                    Dim slideEntryName = $"ppt/slides/slide{slideIndex}.xml"
                    Dim slideRelsEntryName = $"ppt/slides/_rels/slide{slideIndex}.xml.rels"

                    Dim slideEntry = archive.GetEntry(slideEntryName)
                    If slideEntry Is Nothing Then
                        Continue For
                    End If

                    Dim relMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                    Dim relTargetMode As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                    Dim relEntry = archive.GetEntry(slideRelsEntryName)
                    If relEntry IsNot Nothing Then
                        Dim relDoc As New XmlDocument()
                        Using relStream = relEntry.Open()
                            relDoc.Load(relStream)
                        End Using
                        Dim nsmgr As New XmlNamespaceManager(relDoc.NameTable)
                        nsmgr.AddNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships")

                        For Each relNode As XmlNode In relDoc.SelectNodes("//r:Relationship", nsmgr)
                            Dim idAttr = relNode.Attributes("Id")
                            Dim targetAttr = relNode.Attributes("Target")
                            Dim modeAttr = relNode.Attributes("TargetMode")
                            If idAttr IsNot Nothing AndAlso targetAttr IsNot Nothing Then
                                relMap(idAttr.Value) = targetAttr.Value
                                If modeAttr IsNot Nothing Then
                                    relTargetMode(idAttr.Value) = modeAttr.Value
                                End If
                            End If
                        Next
                    End If

                    Dim slideDoc As New XmlDocument()
                    Using slideStream = slideEntry.Open()
                        slideDoc.Load(slideStream)
                    End Using

                    Dim mgr As New XmlNamespaceManager(slideDoc.NameTable)
                    mgr.AddNamespace("p", "http://schemas.openxmlformats.org/presentationml/2006/main")
                    mgr.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
                    mgr.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

                    Dim slideLinks As New Dictionary(Of Integer, String)()

                    ' Find any element with an r:embed or r:link attribute
                    For Each attr As XmlAttribute In slideDoc.SelectNodes("//*/@r:embed | //*/@r:link", mgr)
                        Dim relId = attr.Value
                        If String.IsNullOrWhiteSpace(relId) OrElse Not relMap.ContainsKey(relId) Then
                            Continue For
                        End If

                        Dim shapeId = FindAncestorShapeId(attr.OwnerElement, mgr)
                        If shapeId <= 0 Then
                            Continue For
                        End If

                        Dim target = relMap(relId)
                        Dim mode As String = Nothing
                        relTargetMode.TryGetValue(relId, mode)
                        ' If TargetMode is External, target is a URL
                        If Not String.IsNullOrEmpty(mode) AndAlso String.Equals(mode, "External", StringComparison.OrdinalIgnoreCase) Then
                            If target.StartsWith("http", StringComparison.OrdinalIgnoreCase) AndAlso (target.Contains("youtube.com") Or target.Contains("youtu.be") Or target.Contains("vimeo.com")) Then
                                If Not slideLinks.ContainsKey(shapeId) Then
                                    slideLinks(shapeId) = target
                                End If
                            End If
                        End If
                    Next

                    If slideLinks.Count > 0 Then
                        links(slideIndex) = slideLinks
                    End If
                Next
            End Using
        Catch
            ' best-effort; ignore errors
        End Try

        Return links
    End Function

    Private Function NormalizePartPath(target As String) As String
        Dim path = target.Replace("\"c, "/"c)
        While path.StartsWith("../", StringComparison.Ordinal)
            path = path.Substring(3)
        End While
        If path.StartsWith("./", StringComparison.Ordinal) Then
            path = path.Substring(2)
        End If
        If Not path.StartsWith("ppt/", StringComparison.OrdinalIgnoreCase) Then
            path = "ppt/" & path
        End If
        Return path
    End Function

    Private Function FindAncestorShapeId(node As XmlNode, mgr As XmlNamespaceManager) As Integer
        Dim current As XmlNode = node
        While current IsNot Nothing
            Dim cNvPr As XmlNode = Nothing
            cNvPr = TryFindChild(current, "p:nvPicPr/p:cNvPr", mgr)
            If cNvPr Is Nothing Then
                cNvPr = TryFindChild(current, "p:nvSpPr/p:cNvPr", mgr)
            End If
            If cNvPr Is Nothing Then
                cNvPr = TryFindChild(current, "p:nvGraphicFramePr/p:cNvPr", mgr)
            End If
            If cNvPr Is Nothing Then
                cNvPr = TryFindChild(current, "p:nvPr/p:cNvPr", mgr)
            End If

            If cNvPr IsNot Nothing Then
                Dim idAttr = cNvPr.Attributes("id")
                If idAttr IsNot Nothing Then
                    Dim idVal As Integer
                    If Integer.TryParse(idAttr.Value, idVal) Then
                        Return idVal
                    End If
                End If
            End If

            current = current.ParentNode
        End While
        Return 0
    End Function

    Private Function TryFindChild(node As XmlNode, xpath As String, mgr As XmlNamespaceManager) As XmlNode
        Try
            Return node.SelectSingleNode(xpath, mgr)
        Catch
            Return Nothing
        End Try
    End Function
End Module

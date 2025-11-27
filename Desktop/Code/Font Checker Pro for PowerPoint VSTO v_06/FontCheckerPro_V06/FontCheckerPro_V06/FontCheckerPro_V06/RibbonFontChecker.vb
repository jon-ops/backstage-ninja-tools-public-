
Imports System.Diagnostics
Imports System.Drawing
Imports Microsoft.Office.Core
Imports Microsoft.Office.Tools.Ribbon
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint
Imports System.Windows.Forms


Public Class RibbonFontChecker

    Private Sub RibbonFontChecker_Load(sender As Object, e As RibbonUIEventArgs) Handles MyBase.Load
        ' Debug message removed for production.
    End Sub

    Private Function IsFontInstalled(fontName As String) As Boolean
        Try
            Using testFont As New Drawing.Font(fontName, 8)
                Return testFont.Name.Equals(fontName, StringComparison.InvariantCultureIgnoreCase)
            End Using
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Resolves theme font tokens (e.g., +mn-lt, +mj-lt) to actual font names.
    ''' </summary>
    Private Function ResolveThemeFont(fontName As String, presentation As PowerPoint.Presentation) As String
        ' If it's not a theme font token, return as-is
        If String.IsNullOrEmpty(fontName) OrElse Not fontName.StartsWith("+") Then
            Return fontName
        End If

        Try
            ' Access the theme fonts from the presentation
            Dim themeColorScheme As Object = presentation.SlideMaster.Theme.ThemeFontScheme
            
            ' Map theme tokens to actual fonts
            Select Case fontName.ToLower()
                Case "+mn-lt" ' Minor Latin (body text)
                    Return themeColorScheme.MinorFont.Name
                Case "+mj-lt" ' Major Latin (headings)
                    Return themeColorScheme.MajorFont.Name
                Case "+mn-ea" ' Minor East Asian
                    Return themeColorScheme.MinorFont.Name
                Case "+mj-ea" ' Major East Asian
                    Return themeColorScheme.MajorFont.Name
                Case Else
                    ' For any other theme token, try to extract the font
                    ' Fall back to the original name if we can't resolve it
                    Return fontName
            End Select
        Catch ex As Exception
            ' If we can't resolve the theme font, return the original
            Debug.WriteLine("Could not resolve theme font: " & fontName & " - " & ex.Message)
            Return fontName
        End Try
    End Function

    Private Sub btnScanFonts_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        Dim allFonts As New HashSet(Of String)()
        Dim fontUsage As New Dictionary(Of String, List(Of String))()

        Try
            Dim presentation As PowerPoint.Presentation = Globals.ThisAddIn.Application.ActivePresentation

            ' Loop through all slides
            For Each slide As PowerPoint.Slide In presentation.Slides
                For Each shape As PowerPoint.Shape In slide.Shapes
                    If shape.HasTextFrame AndAlso shape.TextFrame.HasText Then
                        Dim fontName As String = Nothing
                        Try
                            Dim rawFontName As String = shape.TextFrame.TextRange.Font.Name
                            If String.IsNullOrWhiteSpace(rawFontName) Then Continue For
                            fontName = ResolveThemeFont(rawFontName, presentation)
                            If String.IsNullOrWhiteSpace(fontName) Then Continue For
                            fontName = fontName.Trim()
                        Catch ex As Exception
                            ' Skip problematic shape/font
                            Continue For
                        End Try

                        If String.IsNullOrWhiteSpace(fontName) Then Continue For

                        allFonts.Add(fontName)
                        If Not fontUsage.ContainsKey(fontName) Then
                            fontUsage(fontName) = New List(Of String)()
                        End If
                        fontUsage(fontName).Add(slide.SlideIndex.ToString())
                    End If
                Next
            Next

            ' Slide Masters
            For Each design As PowerPoint.Design In presentation.Designs
                Dim master As PowerPoint.Master = design.SlideMaster
                For Each shape As PowerPoint.Shape In master.Shapes
                    If shape.HasTextFrame AndAlso shape.TextFrame.HasText Then
                        Dim fontName As String = Nothing
                        Try
                            Dim rawFontName As String = shape.TextFrame.TextRange.Font.Name
                            If String.IsNullOrWhiteSpace(rawFontName) Then Continue For
                            fontName = ResolveThemeFont(rawFontName, presentation)
                            If String.IsNullOrWhiteSpace(fontName) Then Continue For
                            fontName = fontName.Trim()
                        Catch ex As Exception
                            ' Skip problematic shape/font
                            Continue For
                        End Try

                        If String.IsNullOrWhiteSpace(fontName) Then Continue For

                        allFonts.Add(fontName)
                        If Not fontUsage.ContainsKey(fontName) Then
                            fontUsage(fontName) = New List(Of String)()
                        End If
                        fontUsage(fontName).Add("Master Slide")
                    End If
                Next
            Next

            ' Display report
            Dim frm As New FormFontReport()
            frm.txtReport.Clear()

            ' === Fonts Used ===
            frm.txtReport.SelectionFont = New Font("Segoe UI", 11, FontStyle.Bold)
            frm.txtReport.SelectionColor = Color.DeepSkyBlue
            frm.txtReport.AppendText("=== FONTS USED IN SLIDES ===" & Environment.NewLine)

            frm.txtReport.SelectionFont = New Font("Segoe UI", 11, FontStyle.Regular)
            frm.txtReport.SelectionColor = Color.White
            frm.txtReport.AppendText(String.Join(", ", allFonts) & Environment.NewLine & Environment.NewLine)

            ' === Missing Fonts ===
            frm.txtReport.SelectionFont = New Font("Segoe UI", 11, FontStyle.Bold)
            frm.txtReport.SelectionColor = Color.DeepSkyBlue
            frm.txtReport.AppendText("=== MISSING FONTS ===" & Environment.NewLine)

            Dim missingFound As Boolean = False
            frm.txtReport.SelectionFont = New Font("Segoe UI", 11, FontStyle.Regular)

            For Each fontName In allFonts
                If String.IsNullOrWhiteSpace(fontName) Then Continue For
                Try
                    If Not IsFontInstalled(fontName) Then
                        missingFound = True
                        ' Red ❌, then white text
                        frm.txtReport.SelectionColor = Color.Red
                        frm.txtReport.AppendText("   ❌ ")
                        frm.txtReport.SelectionColor = Color.White
                        frm.txtReport.AppendText(fontName & " (Slides: " & FormatSlideList(fontUsage(fontName)) & ")" & Environment.NewLine)
                    End If
                Catch ex As Exception
                    ' Skip problematic font
                    Continue For
                End Try
            Next

            If Not missingFound Then
                frm.txtReport.SelectionColor = Color.Green
                frm.txtReport.AppendText("All fonts are installed ✅" & Environment.NewLine)
            End If

            frm.ShowDialog()

        Catch ex As Exception
            ' Debug message removed for production.
            Debug.WriteLine("Error scanning fonts: " & ex.Message)
        End Try
    End Sub

    Private Function FormatSlideList(slides As List(Of String)) As String
        ' Remove duplicates
        Dim uniqueSlides As New List(Of String)(slides.Distinct())

        ' Consolidate "Master Slide"
        If uniqueSlides.Contains("Master Slide") Then
            uniqueSlides.Remove("Master Slide")
            uniqueSlides.Add("Master Slides")
        End If

        ' Return as comma-separated string
        Return String.Join(", ", uniqueSlides)
    End Function


    Private Sub btnMediaReport_Click(sender As Object, e As RibbonControlEventArgs) Handles btnMediaReport.Click
        Try
            ' Call the MediaScanner module's public sub
            MediaScanner.ShowMediaReport()
        Catch ex As Exception
            MessageBox.Show("Error running Media Report: " & ex.Message, "Media Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnCheckUpdates_Click(sender As Object, e As RibbonControlEventArgs) Handles btnCheckUpdates.Click
        Try
            Process.Start(New ProcessStartInfo("https://backstageninjatools.lemonsqueezy.com/buy/a3675896-2153-4185-b3ee-f4cd32090f93") With {.UseShellExecute = True})
        Catch ex As Exception
            MessageBox.Show("Unable to open updates page: " & ex.Message, "Check for Updates", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class


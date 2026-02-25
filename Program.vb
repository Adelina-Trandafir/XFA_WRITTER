Imports System.IO
Imports System.Xml
Imports System.Windows.Forms

Module Program
    Private ReadOnly APP_VERSION As Single =
        CSng(Reflection.Assembly.GetExecutingAssembly().GetName().Version.Major)
    Private Const APP_TITLE As String = "XFA_WRITTER"

    Sub Main(args As String())
        Dim xmlPath As String
        Dim outputPdfPath As String
        Dim templatePath As String

#If DEBUG Then
        ' Debug mode: use hardcoded paths for convenience
        xmlPath = "C:\AVACONT\14\DDF_NR_1_REV_1.xml"
        outputPdfPath = "c:\avacont\output.pdf"
#Else
        ' Validate arguments: XFA_WRITTER.exe "data.xml" "output.pdf"
        If args.Length < 2 Then
            MessageBox.Show(
                $"{APP_TITLE} v{APP_VERSION} - Unelte pentru completare formulare XFA cu atașamente" & vbCrLf & vbCrLf &
                "Utilizare:" & vbCrLf &
                "  XFA_WRITTER.exe <fisier_date.xml> <fisier_output.pdf>" & vbCrLf & vbCrLf &
                "Parametri:" & vbCrLf &
                "  fisier_date.xml  - Fișierul XML cu datele de completat (poate conține și atașamente base64)" & vbCrLf &
                "  fisier_output.pdf - Calea fișierului PDF rezultat" & vbCrLf & vbCrLf &
                "Exemplu:" & vbCrLf &
                "  XFA_WRITTER.exe ""c:\avacont\data.xml"" ""c:\avacont\output.pdf""",
                $"{APP_TITLE} v{APP_VERSION}",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            )
            Environment.Exit(1)
            Return
        End If

        xmlPath = args(0)
        outputPdfPath = args(1)
#End If

        Try
            ' Validate input file exists
            If Not File.Exists(xmlPath) Then
                Throw New Exception($"Fișierul XML nu a fost găsit: {xmlPath}")
            End If

            ' Step 1: Parse XML - separate data from attachments
            Dim attachments As List(Of AttachmentModel) = Nothing
            Dim cleanXmlPath As String = Nothing
            ParseInputXml(xmlPath, attachments, cleanXmlPath)

            ' Step 2: Download/cache template from mfinante.gov.ro
#If DEBUG Then
            templatePath = TemplateDownloader.GetTemplatePath() '"C:\AVACONT\cache_ddf\DocumentFundamentare_cu_xml.pdf"
#Else
            templatePath  = TemplateDownloader.GetTemplatePath()

#End If
            ' Step 3: Process XFA and embed attachments
            Dim cPDF As New AdobeUtilsNS.AdobeUtils(attachments)
            cPDF.ProcessXfa(templatePath, outputPdfPath, cleanXmlPath)

            ' Cleanup temp XML if created
            If cleanXmlPath <> xmlPath AndAlso File.Exists(cleanXmlPath) Then
                Try
                    File.Delete(cleanXmlPath)
                Catch
                End Try
            End If

            Environment.Exit(0)

        Catch ex As Exception
            MessageBox.Show(
                $"Eroare: {ex.Message}",
                $"{APP_TITLE} v{APP_VERSION} - Eroare",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            )
            Environment.Exit(1)
        End Try
    End Sub

    ''' <summary>
    ''' Parses input XML: extracts attachments (base64) and produces a clean XML (without Attachments node) for XFA.
    ''' </summary>
    Private Sub ParseInputXml(xmlPath As String, ByRef attachments As List(Of AttachmentModel), ByRef cleanXmlPath As String)
        attachments = New List(Of AttachmentModel)

        Dim doc As New XmlDocument()
        doc.Load(xmlPath)

        ' Find and process Attachments node
        Dim attachmentsNode As XmlNode = doc.SelectSingleNode("//Attachments")

        If attachmentsNode IsNot Nothing Then
            For Each attNode As XmlNode In attachmentsNode.SelectNodes("Attachment")
                Dim fileNameNode = attNode.SelectSingleNode("FileName")
                Dim fileDataNode = attNode.SelectSingleNode("FileData")

                If fileNameNode IsNot Nothing AndAlso fileDataNode IsNot Nothing Then
                    Dim fileName = fileNameNode.InnerText.Trim()
                    Dim base64Data = fileDataNode.InnerText.Trim()

                    If Not String.IsNullOrEmpty(fileName) AndAlso Not String.IsNullOrEmpty(base64Data) Then
                        Try
                            attachments.Add(New AttachmentModel() With {
                                .fileName = fileName,
                                .FileData = Convert.FromBase64String(base64Data),
                                .IsDeleted = False
                            })
                        Catch ex As FormatException
                            ' Skip invalid base64
                        End Try
                    End If
                End If
            Next

            ' Remove Attachments node from XML for clean XFA data
            attachmentsNode.ParentNode.RemoveChild(attachmentsNode)

            ' Save clean XML to temp file
            cleanXmlPath = Path.Combine(Path.GetTempPath(), $"xfa_data_{Guid.NewGuid():N}.xml")
            doc.Save(cleanXmlPath)
        Else
            ' No attachments - use original XML as-is
            cleanXmlPath = xmlPath
        End If
    End Sub

End Module
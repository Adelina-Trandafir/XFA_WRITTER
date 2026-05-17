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
        Dim docType As String
        Dim templatePath As String

#If DEBUG Then
        xmlPath = "C:\AVACONT\ORD_TEST.xml"
        outputPdfPath = "c:\avacont\output.pdf"
        docType = "ORD"
        Logger.Init(docType)
        Logger.Log("INFO", $"args (DEBUG) — xmlPath: {xmlPath}, outputPdfPath: {outputPdfPath}, tip: {docType}")
#Else
        If args.Length < 3 Then
            MessageBox.Show(
                $"{APP_TITLE} v{APP_VERSION} - Unelte pentru completare formulare XFA cu atașamente" & vbCrLf & vbCrLf &
                "Utilizare:" & vbCrLf &
                "  XFA_WRITTER.exe <fisier_date.xml> <fisier_output.pdf> <tip_document>" & vbCrLf & vbCrLf &
                "Parametri:" & vbCrLf &
                "  fisier_date.xml   - Fișierul XML cu datele de completat (poate conține și atașamente base64)" & vbCrLf &
                "  fisier_output.pdf - Calea fișierului PDF rezultat" & vbCrLf &
                "  tip_document      - Tipul documentului: DDF, ORD" & vbCrLf & vbCrLf &
                "Exemplu:" & vbCrLf &
                "  XFA_WRITTER.exe ""c:\avacont\data.xml"" ""c:\avacont\output.pdf"" DDF",
                $"{APP_TITLE} v{APP_VERSION}",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            )
            Environment.Exit(1)
            Return
        End If

        xmlPath = args(0)
        outputPdfPath = args(1)
        docType = args(2).ToUpper()
        Logger.Init(docType)
        Logger.Log("INFO", $"args — xmlPath: {xmlPath}, outputPdfPath: {outputPdfPath}, tip: {docType}")

#End If

        Try
            If Not File.Exists(xmlPath) Then
                Throw New Exception($"Fișierul XML nu a fost găsit: {xmlPath}")
            End If

            ' Step 1: Separă datele de atașamente
            Logger.LogSection("ParseInputXml")
            Dim attachments As List(Of AttachmentModel) = Nothing
            Dim cleanXmlPath As String = Nothing
            ParseInputXml(xmlPath, attachments, cleanXmlPath)
            Logger.Log("INFO", $"ParseInputXml — attachments: {attachments.Count}, cleanXml: {cleanXmlPath}")

            ' Step 2: Descarcă/cache template din Flask pe baza tipului de document
            Logger.LogSection("GetTemplatePath")
            templatePath = TemplateDownloader.GetTemplatePath(docType)
            Logger.Log("INFO", $"GetTemplatePath — templatePath: {templatePath}")

            ' Step 3: Procesează XFA și embed atașamente
            Logger.LogSection("ProcessXfa")
            Dim cPDF As New AdobeUtilsNS.AdobeUtils(attachments)
            cPDF.ProcessXfa(templatePath, outputPdfPath, cleanXmlPath)

            ' Cleanup XML temporar
            If cleanXmlPath <> xmlPath AndAlso File.Exists(cleanXmlPath) Then
                Try
                    File.Delete(cleanXmlPath)
                Catch
                End Try
            End If

            Logger.Log("INFO", "Exit(0)")
            Environment.Exit(0)

        Catch ex As Exception
            Logger.Log("ERROR", $"Exception: {ex.Message}{vbCrLf}{ex.StackTrace}")
            MessageBox.Show(
                $"Eroare: {ex.Message}",
                $"{APP_TITLE} v{APP_VERSION} - Eroare",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            )
            Logger.Log("INFO", "Exit(1)")
            Environment.Exit(1)
        End Try
    End Sub

    Private Sub ParseInputXml(xmlPath As String, ByRef attachments As List(Of AttachmentModel), ByRef cleanXmlPath As String)
        attachments = New List(Of AttachmentModel)

        Dim doc As New XmlDocument()
        doc.Load(xmlPath)

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
                                .FileName = fileName,
                                .FileData = Convert.FromBase64String(base64Data),
                                .IsDeleted = False
                            })
                        Catch ex As FormatException
                            ' Skip invalid base64
                        End Try
                    End If
                End If
            Next

            ' Scoate nodul Attachments din XML înainte de a-l trimite la XFA
            attachmentsNode.ParentNode.RemoveChild(attachmentsNode)

            ' Salvează XML curat în temp
            cleanXmlPath = Path.Combine(Path.GetTempPath(), $"xfa_data_{Guid.NewGuid():N}.xml")
            doc.Save(cleanXmlPath)
        Else
            ' Fără atașamente - folosește XML-ul original ca atare
            cleanXmlPath = xmlPath
        End If
    End Sub

End Module
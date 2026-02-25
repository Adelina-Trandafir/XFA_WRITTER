Imports System.IO
Imports System.Threading
Imports System.Xml
Imports iTextSharp.text.pdf

Namespace AdobeUtilsNS
    <System.Runtime.Versioning.SupportedOSPlatform("windows")>
    Public Class AdobeUtils
        Private ReadOnly pAttachments As List(Of AttachmentModel)

        Public Sub New(Optional attList As List(Of AttachmentModel) = Nothing)
            If attList Is Nothing Then
                pAttachments = New List(Of AttachmentModel)
            Else
                pAttachments = attList
            End If
        End Sub

        ''' <summary>
        ''' Modifies XFA form fields/tables from XML data, then embeds attachments into the output PDF.
        ''' Does NOT open Adobe or wait for signature.
        ''' </summary>
        ''' <param name="inputPdfPath">Source PDF template path (must be XFA form)</param>
        ''' <param name="outputPdfPath">Output PDF path</param>
        ''' <param name="dataXMLPath">XML data file path (fields + tables, without Attachments node)</param>
        Public Sub ProcessXfa(inputPdfPath As String, outputPdfPath As String, dataXMLPath As String)
            ' Step 1: Modify XFA
            Dim rsp = ModifyXfaFromXml(inputPdfPath, outputPdfPath, dataXMLPath)
            If rsp <> "OK" Then
                Throw New Exception(rsp)
            End If

            ' Step 2: Embed attachments (in same file)
            If Not AddAttachmentsInPlace(outputPdfPath, pAttachments) Then
                Throw New Exception("Eroare la adăugarea atașamentelor în PDF.")
            End If
        End Sub

        ''' <summary>
        ''' Modifies XFA form, opens Adobe Reader, waits for signature, and verifies
        ''' </summary>
        ''' <param name="inputPdfPath">Source PDF path (must be XFA form)</param>
        ''' <param name="outputPdfPath">Output PDF path after XFA modification</param>
        ''' <param name="dataXMLPath">XML data file path</param>
        ''' <returns>True if PDF was signed and saved, False otherwise</returns>
        Public Function ProcessAndVerifySignature(inputPdfPath As String, outputPdfPath As String, dataXMLPath As String) As Boolean
            ' Step 1: Modify XFA
            Dim rsp = ModifyXfaFromXml(inputPdfPath, outputPdfPath, dataXMLPath)
            If rsp <> "OK" Then
                Throw New Exception(rsp)
            End If

            ' Step 2: Adaugă attach-urile (în același fișier)
            Dim rsp2 = AddAttachmentsInPlace(outputPdfPath, pAttachments)
            If Not rsp2 Then
                Throw New Exception("Eroare la adăugarea atașamentelor în PDF.")
            End If

            ' Step 3: Open Adobe and wait for signature
            If Not VerifyPdfSignature(outputPdfPath) Then
                Throw New Exception("PDF nu a fost semnat sau nu a fost salvat.")
            End If

            Return True
        End Function

        ''' <summary>
        ''' Function to modify XFA form fields and tables based on XML configuration
        ''' </summary>
        ''' <param name="inputPdfPath">Template PDF FILE</param>
        ''' <param name="outputPdfPath">Output file before signature</param>
        ''' <param name="configXmlPath"> MUST MATCH the XFA field names exactly!</param>
        ''' <returns></returns>
        Private Shared Function ModifyXfaFromXml(inputPdfPath As String, outputPdfPath As String, configXmlPath As String) As String
            Dim reader As New PdfReader(inputPdfPath)
            Dim stamper As New PdfStamper(reader, New FileStream(outputPdfPath, FileMode.Create), "\0", True)
            Try
                Dim xfaForm As XfaForm = stamper.AcroFields.Xfa
                Dim domDoc As System.Xml.XmlDocument = xfaForm.DomDocument

                ' Load config XML
                Dim configDoc As New System.Xml.XmlDocument()
                configDoc.Load(configXmlPath)

                ' Process all nodes recursively
                ProcessXmlNodes(configDoc.DocumentElement, domDoc)

                ' Write back XFA
                xfaForm.DomDocument = domDoc
                xfaForm.Changed = True

                Return "OK"

            Catch ex As Exception
                Return $"Eroare la modificarea XFA: {ex.Message}"
            Finally
                stamper?.Close()
                reader?.Close()
            End Try
        End Function

        ''' <summary>
        ''' Recursively processes XML nodes and applies to XFA PDF
        ''' </summary>
        Private Shared Sub ProcessXmlNodes(xmlNode As System.Xml.XmlNode, pdfDoc As System.Xml.XmlDocument)
            ' Check if this node is a table (has Row1 children)
            Dim isTable = xmlNode.SelectSingleNode("Row1") IsNot Nothing

            If isTable Then
                ' Handle table: add rows
                Dim tableName = xmlNode.Name
                Dim pdfTableNode = pdfDoc.SelectSingleNode($"//{tableName}")

                If pdfTableNode IsNot Nothing Then
                    ' Remove existing rows (keep header/footer)
                    Dim existingRows = pdfTableNode.SelectNodes("Row1")
                    For Each row In existingRows
                        pdfTableNode.RemoveChild(row)
                    Next

                    ' Add new rows from config
                    For Each row In xmlNode.SelectNodes("Row1")
                        Dim newRow As System.Xml.XmlElement = pdfDoc.CreateElement("Row1")
                        For Each cell In row.ChildNodes
                            Dim cellValue = CStr(cell.InnerText)
                            newRow.AppendChild(CreateElement(pdfDoc, cell.Name, cellValue))
                        Next
                        pdfTableNode.AppendChild(newRow)
                    Next
                End If

            Else
                ' Handle regular fields
                For Each childNode In xmlNode.ChildNodes
                    Dim nodeName = childNode.Name
                    Dim nodeValue = CStr(childNode.InnerText)

                    ' If node has children and is not empty, recurse
                    If childNode.HasChildNodes AndAlso childNode.SelectNodes("*").Count > 0 Then
                        ProcessXmlNodes(childNode, pdfDoc)
                    ElseIf Not String.IsNullOrWhiteSpace(nodeValue) Then
                        ' Set leaf node value in PDF
                        SetNodeValue(pdfDoc, $"//{nodeName}", nodeValue)
                    End If
                Next
            End If
        End Sub

        ''' <summary>
        ''' Adds attachments to existing PDF in place
        ''' </summary>
        Private Shared Function AddAttachmentsInPlace(pPdfPath As String, pAttList As List(Of AttachmentModel)) As Boolean
            If pAttList Is Nothing OrElse pAttList.Count = 0 Then
                Return True
            End If

            Dim pdfBytes As Byte() = File.ReadAllBytes(pPdfPath)
            Dim output As New MemoryStream()

            Dim reader As New PdfReader(pdfBytes)
            Dim stamper As New PdfStamper(reader, output, "\0", True)

            For Each att As AttachmentModel In pAttList
                If att Is Nothing Then Continue For
                If att.IsDeleted Then Continue For
                If att.FileData Is Nothing OrElse att.FileData.Length = 0 Then Continue For

                Dim pName As String = Path.GetFileName(att.FileName)

                Dim fileSpec As PdfFileSpecification =
                    PdfFileSpecification.FileEmbedded(
                        stamper.Writer,
                        att.FileName,
                        pName,
                        att.FileData
                    )

                stamper.AddFileAttachment(pName, fileSpec)
            Next

            stamper.Close()
            reader.Close()

            File.WriteAllBytes(pPdfPath, output.ToArray())

            Return True
        End Function

        ''' <summary>
        ''' Opens PDF in Adobe Reader, waits for user to sign and save, then verifies signature
        ''' </summary>
        ''' <param name="pdfPath">Full path to PDF (e.g., c:\avacont\output_itextsharp.pdf)</param>
        ''' <returns>True if signed, False if not signed or user didn't save</returns>
        Private Shared Function VerifyPdfSignature(pdfPath As String) As Boolean
            Dim adobePath = GetAdobeReaderPath()
            If String.IsNullOrEmpty(adobePath) Then
                Throw New Exception("Adobe Reader not found")
            End If

            Dim originalModifyTime = New FileInfo(pdfPath).LastWriteTime
            Dim isSigned = False
            Dim signedEvent = New System.Threading.ManualResetEvent(False)
            Dim closedEvent = New System.Threading.ManualResetEvent(False)

            ' FileSystemWatcher pentru semnare
            Dim watcher = New FileSystemWatcher(Path.GetDirectoryName(pdfPath)) With {
                                                                                    .Filter = Path.GetFileName(pdfPath),
                                                                                    .NotifyFilter = NotifyFilters.LastWrite
                                                                                    }

            AddHandler watcher.Changed, Sub(sender, e)
                                            System.Threading.Thread.Sleep(300)
                                            Try
                                                Dim newModifyTime = New FileInfo(pdfPath).LastWriteTime
                                                If newModifyTime > originalModifyTime Then
                                                    isSigned = CheckIfPdfHasSignature(pdfPath)
                                                End If
                                            Catch
                                            End Try
                                        End Sub

            ' Start Adobe process
            Dim process As New Process()
            process.StartInfo.FileName = adobePath
            process.StartInfo.Arguments = $"""{pdfPath}"""
            process.Start()

            ' Task pentru a detecta închidere
            Dim closeTask = System.Threading.Tasks.Task.Run(Sub()
                                                                process.WaitForExit()
                                                                closedEvent.Set()
                                                            End Sub)

            ' Aștept max 10 minute - fie semnare, fie închidere
            watcher.EnableRaisingEvents = True

            Dim h = {signedEvent, closedEvent}
            Dim waitIndex = WaitHandle.WaitAny(h, TimeSpan.FromMinutes(10))

            watcher.Dispose()
            process?.Dispose()

            Select Case waitIndex
                Case 0 ' Signed
                    Return True
                Case 1 ' Closed without signing
                    Return isSigned
                Case Else ' Timeout
                    Return False
            End Select
        End Function

        ''' <summary>
        ''' Checks if PDF contains valid signatures using iTextSharp
        ''' </summary>
        Private Shared Function CheckIfPdfHasSignature(pdfPath As String) As Boolean
            Dim reader As PdfReader = Nothing
            Try
                reader = New PdfReader(pdfPath)
                Dim af As AcroFields = reader.AcroFields

                ' Get all signature field names
                Dim signatureNames As List(Of String) = af.GetSignatureNames()

                ' If list has items, signatures exist
                Return signatureNames IsNot Nothing AndAlso signatureNames.Count > 0

            Catch ex As Exception
                Debug.WriteLine($"Error checking PDF signatures: {ex.Message}")
                Return False
            Finally
                reader?.Close()
            End Try
        End Function

        ''' <summary>
        ''' Finds Adobe Reader installation path
        ''' </summary>
        Private Shared Function GetAdobeReaderPath() As String

            Const REG_PATH As String = "Software\VB and VBA Program Settings\AVACONT\AdobeUtils"

            ' Hardcoded default paths
            Dim defaultPaths As String() = {
                                            "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
                                            "C:\Program Files (x86)\Adobe\Reader 11.0\Reader\AcroRd32.exe",
                                            "C:\Program Files (x86)\Adobe\Acrobat Reader\Reader\AcroRd32.exe",
                                            "C:\Program Files\Adobe\Acrobat Reader\Reader\AcroRd32.exe",
                                            "C:\Program Files (x86)\Adobe\Reader\Reader\AcroRd32.exe",
                                            "C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
                                            }

            Try
                Dim regKey As Microsoft.Win32.RegistryKey = If(Microsoft.Win32.Registry.CurrentUser.OpenSubKey(REG_PATH, True), Microsoft.Win32.Registry.CurrentUser.CreateSubKey(REG_PATH))

                ' Add missing default paths to registry
                For i = 0 To defaultPaths.Length - 1
                    Dim valueName = "AdobePath" & (i + 1)
                    Dim existingValue = regKey.GetValue(valueName)

                    If existingValue Is Nothing Then
                        regKey.SetValue(valueName, defaultPaths(i))
                    End If
                Next

                ' Search through all registry paths and return first existing one
                Dim valueNames = regKey.GetValueNames()
                For Each valueName In valueNames
                    Dim pathValue = regKey.GetValue(valueName)?.ToString()

                    If Not String.IsNullOrEmpty(pathValue) AndAlso File.Exists(pathValue) Then
                        regKey.Close()
                        Return pathValue
                    End If
                Next

                regKey.Close()

            Catch ex As Exception
                Debug.WriteLine($"Registry error: {ex.Message}")
            End Try

            Return Nothing
        End Function

        ''' <summary>
        ''' Attempts to find Adobe Reader path from Windows Registry
        ''' </summary>
        Private Shared Function GetAdobePathFromRegistry() As String
            Try
                Dim key As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\App Paths\AcroRd32.exe")

                If key IsNot Nothing Then
                    Dim path = key.GetValue("")?.ToString()
                    key.Close()
                    If Not String.IsNullOrEmpty(path) AndAlso File.Exists(path) Then
                        Return path
                    End If
                End If
            Catch
                ' Registry access failed, return nothing
            End Try

            Return Nothing
        End Function

        Private Shared Function CreateElement(doc As XmlDocument, name As String, value As String) As XmlElement
            Dim elem = doc.CreateElement(name)
            elem.InnerText = value
            Return elem
        End Function

        Private Shared Sub SetNodeValue(doc As XmlDocument, xpath As String, value As String)
            Dim node = doc.SelectSingleNode(xpath)
            If node IsNot Nothing Then
                node.InnerText = value
            End If
        End Sub
    End Class
End Namespace
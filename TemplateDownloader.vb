Imports System.IO
Imports System.Net.Http

<System.Runtime.Versioning.SupportedOSPlatform("windows")>
Public Class TemplateDownloader

    Public Shared Function GetTemplatePath(docType As String) As String
        Dim templateUrl As String = GetUrlForDocType(docType)

        If Not Directory.Exists(CACHE_DIR) Then
            Directory.CreateDirectory(CACHE_DIR)
        End If

        Using client As New HttpClient()
            client.Timeout = TimeSpan.FromSeconds(60)
            client.DefaultRequestHeaders.Add("X-API-KEY", API_KEY)

            Dim headRequest As New HttpRequestMessage(HttpMethod.Head, BASE_URL & templateUrl)
            Dim headResponse As HttpResponseMessage

            Try
                headResponse = client.SendAsync(headRequest).Result
                headResponse.EnsureSuccessStatusCode()
            Catch ex As Exception
                Return DownloadFull(client, templateUrl, docType)
            End Try

            Dim remoteFileName As String = GetFileNameFromResponse(headResponse)

            If String.IsNullOrEmpty(remoteFileName) Then
                Return DownloadFull(client, templateUrl, docType)
            End If

            Dim localPath As String = Path.Combine(CACHE_DIR, remoteFileName)
            If File.Exists(localPath) Then
                Return localPath
            End If

            Return DownloadFull(client, templateUrl, docType)
        End Using
    End Function

    ''' <summary>
    ''' Mapează docType → URL din Configs.
    ''' Adaugă aici orice tip nou de document.
    ''' </summary>
    Private Shared Function GetUrlForDocType(docType As String) As String
        Select Case docType.ToUpper()
            Case "DDF" : Return DDF_URL
            Case "ORD" : Return ORD_URL
            Case Else
                Throw New Exception($"Tip document necunoscut: '{docType}'. Tipuri acceptate: DDF, ORD.")
        End Select
    End Function

    Private Shared Function DownloadFull(client As HttpClient, templateUrl As String, docType As String) As String
        Dim response As HttpResponseMessage

        Try
            response = client.GetAsync(BASE_URL & templateUrl).Result
            response.EnsureSuccessStatusCode()
        Catch ex As Exception
            Throw New Exception($"Nu s-a putut descărca macheta {docType} de pe server: {ex.Message}")
        End Try

        Dim remoteFileName As String = GetFileNameFromResponse(response)
        If String.IsNullOrEmpty(remoteFileName) Then
            remoteFileName = $"{docType}_template_{DateTime.Now:yyyyMMdd}.pdf"
        End If

        Dim localPath As String = Path.Combine(CACHE_DIR, remoteFileName)
        Dim pdfBytes As Byte() = response.Content.ReadAsByteArrayAsync().Result
        File.WriteAllBytes(localPath, pdfBytes)

        Return localPath
    End Function

    Private Shared Function GetFileNameFromResponse(response As HttpResponseMessage) As String
        Dim cd = response.Content?.Headers?.ContentDisposition
        If cd IsNot Nothing AndAlso Not String.IsNullOrEmpty(cd.FileName) Then
            Return cd.FileName.Trim(""""c)
        End If
        Return Nothing
    End Function

End Class
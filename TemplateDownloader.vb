Imports System.IO
Imports System.Net.Http

''' <summary>
''' Downloads and caches the DocumentFundamentare PDF template from the local Flask server.
''' </summary>
<System.Runtime.Versioning.SupportedOSPlatform("windows")>
Public Class TemplateDownloader
    ''' <summary>
    ''' Returns local path to the cached template PDF.
    ''' Downloads from server if not already cached (or if filename changed).
    ''' </summary>
    Public Shared Function GetTemplatePath() As String
        If Not Directory.Exists(CACHE_DIR) Then
            Directory.CreateDirectory(CACHE_DIR)
        End If

        Using client As New HttpClient()
            client.Timeout = TimeSpan.FromSeconds(60)
            client.DefaultRequestHeaders.Add("X-API-KEY", API_KEY)

            ' HEAD request - ia doar numele fișierului fără a descărca
            Dim headRequest As New HttpRequestMessage(HttpMethod.Head, BASE_URL & DDF_URL)
            Dim headResponse As HttpResponseMessage

            Try
                headResponse = client.SendAsync(headRequest).Result
                headResponse.EnsureSuccessStatusCode()
            Catch ex As Exception
                ' Dacă HEAD eșuează, încearcă GET direct
                Return DownloadFull(client)
            End Try

            ' Extrage filename din Content-Disposition
            Dim remoteFileName As String = GetFileNameFromResponse(headResponse)

            If String.IsNullOrEmpty(remoteFileName) Then
                Return DownloadFull(client)
            End If

            ' Verifică dacă există deja în cache
            Dim localPath As String = Path.Combine(CACHE_DIR, remoteFileName)
            If File.Exists(localPath) Then
                Return localPath
            End If

            ' Nu există - descarcă
            Return DownloadFull(client)
        End Using
    End Function

    Private Shared Function DownloadFull(client As HttpClient) As String
        Dim response As HttpResponseMessage

        Try
            response = client.GetAsync(DDF_URL).Result
            response.EnsureSuccessStatusCode()
        Catch ex As Exception
            Throw New Exception($"Nu s-a putut descărca macheta de pe server: {ex.Message}")
        End Try

        Dim remoteFileName As String = GetFileNameFromResponse(response)
        If String.IsNullOrEmpty(remoteFileName) Then
            remoteFileName = $"DDF_template_{DateTime.Now:yyyyMMdd}.pdf"
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
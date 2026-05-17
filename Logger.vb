Imports System.IO

Module Logger
    Private _logPath As String = Nothing
    Private ReadOnly _lock As New Object()

    Sub Init(tipDocument As String)
        Try
            Dim dir = "C:\AVACONT\logs"
            Directory.CreateDirectory(dir)
            Dim ts = DateTime.Now.ToString("yyyyMMdd_HHmmss")
            Dim tip = If(String.IsNullOrWhiteSpace(tipDocument), "UNKNOWN", tipDocument.ToUpper())
            _logPath = Path.Combine(dir, $"xfa_{tip}_{ts}.txt")
        Catch
            _logPath = Path.Combine(Path.GetTempPath(), "xfa_fallback.txt")
        End Try
        Log("INFO", $"=== XFA_WRITTER start — tip: {tipDocument} ===")
    End Sub

    Sub Log(level As String, msg As String)
        Dim line = $"[{DateTime.Now:HH:mm:ss.fff}] [{level,-5}] {msg}"
        SyncLock _lock
            Try
                Dim filePath As String = If(_logPath, IO.Path.Combine(IO.Path.GetTempPath(), "xfa_fallback.txt"))
                Using sw = New StreamWriter(filePath, append:=True)
                    sw.WriteLine(line)
                    sw.Flush()
                End Using
            Catch
            End Try
        End SyncLock
    End Sub

    Sub LogSection(title As String)
        Log("INFO", $"====== {title} ======")
    End Sub
End Module

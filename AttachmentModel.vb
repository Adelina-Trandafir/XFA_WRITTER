''' <summary>
''' Model for PDF attachments parsed from XML (base64 encoded)
''' </summary>
Public Class AttachmentModel
    Public Property FileName As String
    Public Property FileData As Byte()
    Public Property IsDeleted As Boolean = False

    ''' <summary>
    ''' For compatibility with AdobeUtils reflection-based access
    ''' Maps FileName to FilePath
    ''' </summary>
    Public ReadOnly Property FilePath As String
        Get
            Return FileName
        End Get
    End Property
End Class

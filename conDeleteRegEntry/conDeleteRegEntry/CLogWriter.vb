Imports System
Imports System.IO
Imports Microsoft
Imports Microsoft.VisualBasic
Public Class CLogWriter
    Private m_fn As Integer = FreeFile()
    Private m_fileOpen As Boolean = False
    Private m_fileAndPath As String
    Private m_writeLog As Boolean = False
    Private m_dateTime As DateTime = DateTime.Now
    Private m_dateTimeString As String = m_dateTime.ToString("yyyy-d-MM : hh:m:ss t")
    Private Property FilePath() As String
        Get
            FilePath = m_fileAndPath
        End Get
        Set(ByVal value As String)
            If Not IsNothing(value) Then
                m_fileAndPath = value
            End If
        End Set
    End Property
    Private Property WriteLog() As Boolean
        Get
            WriteLog = m_writeLog
        End Get
        Set(ByVal value As Boolean)
            If Not IsNothing(value) Then
                m_writeLog = value
            End If
        End Set
    End Property
    Public Sub New(ByVal fileAndPath As String, ByVal writeMyLog As Boolean)
        FilePath = fileAndPath
        WriteLog = writeMyLog
        If WriteLog Then
            If File.Exists(FilePath) Then
                FileOpen(m_fn, FilePath, OpenMode.Append)
            Else
                Dim iEnd As Integer = InStrRev(FilePath, "\")
                Dim fpath As String = Left(FilePath, iEnd)
                Dim fname As String = Right(FilePath, Len(FilePath) - iEnd)
                If Not Directory.Exists(fpath) Then
                    Try
                        MkDir(fpath)
                    Catch ex As Exception
                    End Try
                End If
                Try
                    FileOpen(m_fn, FilePath, OpenMode.Output)
                    m_fileOpen = True
                Catch ex As Exception
                    m_fileOpen = False
                    Console.WriteLine("Cannot Open Log File : " & ex.Message.ToString)
                End Try
            End If
        End If
    End Sub
    Public Sub dispose()
        If m_fileOpen Then
            FileClose(m_fn)
        End If
        Me.dispose()
    End Sub
    Public Sub Print(ByVal varname As String, ByVal myObj As Object)
        WriteLine(m_fn, m_dateTimeString & FormatValues(varname, myObj))
    End Sub
    Public Overloads Shared Function FormatValues(ByVal varName As String, ByVal myObj As Integer) As String
        Return varName & " - " & myObj.ToString
    End Function
    Public Overloads Shared Function FormatValues(ByVal varName As String, ByVal myObj As DateTime) As String
        Return varName & " - " & myObj.ToString
    End Function
    Public Overloads Shared Function FormatValues(ByVal varName As String, ByVal myObj As String) As String
        Return varName & " - " & myObj
    End Function
    Public Overloads Shared Function FormatValues(ByVal varName As String, ByVal myObj As Boolean) As String
        Return varName & " - " & myObj.ToString
    End Function
    Public Overloads Shared Function FormatValues(ByVal varName As String, ByVal myObj() As String) As String
        Dim iMax As Integer = myObj.Length - 1
        Dim iIdx As Integer
        Dim sRtnVal As String = varName & " - "
        For iIdx = 0 To iMax
            If Not IsNothing(myObj(iIdx)) Then
                sRtnVal = sRtnVal & myObj(iIdx).ToString & vbCrLf
            End If
        Next
        Return sRtnVal
    End Function
End Class

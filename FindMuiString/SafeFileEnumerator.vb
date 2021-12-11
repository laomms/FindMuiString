Imports System.IO

Public Class SafeFileEnumerator
    Public Shared Function EnumerateDirectories(ByVal parentDirectory As String, ByVal searchPattern As String, ByVal searchOpt As SearchOption) As IEnumerable(Of String)
        Try
            Dim directories = Enumerable.Empty(Of String)()
            If searchOpt = SearchOption.AllDirectories Then
                directories = Directory.EnumerateDirectories(parentDirectory).SelectMany(Function(x) EnumerateDirectories(x, searchPattern, searchOpt))
            End If
            Return directories.Concat(Directory.EnumerateDirectories(parentDirectory, searchPattern))
        Catch ex As UnauthorizedAccessException
            Return Enumerable.Empty(Of String)()
        End Try
    End Function

    Public Shared Function EnumerateFiles(ByVal path As String, ByVal searchPattern As String, ByVal searchOpt As SearchOption) As IEnumerable(Of String)
        Try
            Dim dirFiles = Enumerable.Empty(Of String)()
            If searchOpt = SearchOption.AllDirectories Then
                dirFiles = Directory.EnumerateDirectories(path).SelectMany(Function(x) EnumerateFiles(x, searchPattern, searchOpt))
            End If
            Return dirFiles.Concat(Directory.EnumerateFiles(path, searchPattern))
        Catch ex As UnauthorizedAccessException
            Return Enumerable.Empty(Of String)()
        End Try
    End Function
End Class

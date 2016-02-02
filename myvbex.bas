Attribute VB_Name = "Module1"
Option Explicit

'²¹È«¼òÐ´
Public Function Cbln(ByVal a As Variant)
    Cbln = CBool(a)
End Function

Public Function Cbyt(ByVal a As Variant)
    Cbyt = CByte(a)
End Function

Public Function Cdtm(ByVal a As Variant)
    Cdtm = CDate(a)
End Function

Public Function Cdat(ByVal a As Variant)
    Cdat = CDate(a)
End Function

Public Function Cvnt(ByVal a As Variant)
    Cvnt = CVar(a)
End Function

Public Function Csgl(ByVal a As Variant)
    Csgl = CSng(a)
End Function

'²¹È«ËÄ×Ö´Ê
Public Function CLong(ByVal a As Variant)
    CLong = CLng(a)
End Function

'
Public Function Max(ByVal a As Variant, ByVal b As Variant) As Variant
    Max = IIf(a > b, a, b)
End Function

'fso
Public Sub CopyFileEx(Source As String, Destination As String, Optional OverWriteFiles As Boolean = True)
    Dim fso As FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim strArr() As String
    strArr = Split(Destination, "\")
    Dim i As Integer
    Dim strDir As String
    Dim indexOfFXG As Integer
    indexOfFXG = 1
    For i = 0 To UBound(strArr) - 1
        indexOfFXG = InStr(indexOfFXG + 1, Destination, "\")
        strDir = Left(Destination, indexOfFXG)
        If Not fso.FolderExists(strDir) Then
            fso.CreateFolder strDir
        End If
    Next
    fso.CopyFile Source, Destination, OverWriteFiles
End Sub

Public Sub CreateFolderEx(Destination As String)
    Dim fso As FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim strArr() As String
    strArr = Split(Destination, "\")
    Dim i As Integer
    Dim strDir As String
    Dim indexOfFXG As Integer
    indexOfFXG = 1
    For i = 0 To UBound(strArr) - 1
        indexOfFXG = InStr(indexOfFXG + 1, Destination, "\")
        strDir = Left(Destination, indexOfFXG)
        If Not fso.FolderExists(strDir) Then
            fso.CreateFolder strDir
        End If
    Next
    If Right(Destination, 1) <> "\" Then
        fso.CreateFolder Destination
    End If
End Sub



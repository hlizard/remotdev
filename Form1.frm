VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":0000
      Height          =   1575
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dtmLastCopyTime As Date
Dim strLastCopyTime As String
Dim fso As New FileSystemObject
Dim tempDir As Folder
Dim rootDir As Folder

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Function CopyModifiedFiles2(srcDir As Folder) As Boolean
    Dim fsobj As Object
    Dim fd As Folder
    Dim f As File
    Dim fqr As FsobjQueryResult
    For Each fsobj In srcDir.SubFolders
        If TypeOf fsobj Is Folder Then
            Set fd = fsobj
            If Not Left(fd.Name, Len(tempDir.Name)) = tempDir.Name _
                And Not LCase(Right(fd.Name, 3)) = "bak" Then
                fqr = QueryFolder(fd)
                If fqr = fqrChanged Or fqr = fqrNotFound Then
                    Dim strTargetDir As String
                    strTargetDir = tempDir.Path + Replace(fd.Path, rootDir.Path, "")
                    If Not fso.FolderExists(strTargetDir) Then
                        CreateFolderEx strTargetDir
                    End If
                    If fqr = fqrChanged Then DeleteFolderFromDB fd
                    InsertFolder fd
                End If
                CopyModifiedFiles2 fd
            End If
        End If
    Next
    For Each fsobj In srcDir.Files
        If TypeOf fsobj Is File Then
            Set f = fsobj
            If Not Left(f.Name, 12) = "remotdev.db3" Then
                fqr = QueryFile(f)
                If fqr = fqrChanged Or fqr = fqrNotFound Then
                    Dim strTargetFile As String
                    strTargetFile = tempDir.Path + Replace(f.Path, rootDir.Path, "")
                    If Not fso.FileExists(strTargetFile) Then
                        CopyFileEx f.Path, strTargetFile
                    End If
                    If fqr = fqrChanged Then DeleteFileFromDB f
                    InsertFile f
                End If
            End If
        End If
    Next
End Function

Function CopyModifiedFiles(srcDir As Folder) As Boolean
    Dim fsobj As Object
    Dim fd As Folder
    Dim f As File
    Dim dtmLastModified As Date
    For Each fsobj In srcDir.SubFolders
        If TypeOf fsobj Is Folder Then
            Set fd = fsobj
            If Not Left(fd.Name, Len(tempDir.Name)) = tempDir.Name _
                And Not LCase(Right(fd.Name, 3)) = "bak" Then
                dtmLastModified = Max(fd.DateCreated, fd.DateLastModified)
                If dtmLastModified > dtmLastCopyTime Then
                    Dim strTargetDir As String
                    strTargetDir = tempDir.Path + Replace(fd.Path, rootDir.Path, "")
                    If Not fso.FolderExists(strTargetDir) Then
                        fso.CreateFolder strTargetDir
                    End If
                End If
                CopyModifiedFiles fd
            End If
        End If
    Next
    For Each fsobj In srcDir.Files
        If TypeOf fsobj Is File Then
            Set f = fsobj
            dtmLastModified = Max(f.DateCreated, f.DateLastModified)
            If dtmLastModified > dtmLastCopyTime Then
                Dim strTargetFile As String
                strTargetFile = tempDir.Path + Replace(f.Path, rootDir.Path, "")
                If Not fso.FileExists(strTargetFile) Then
                    CopyFileEx f.Path, strTargetFile
                End If
            End If
        End If
    Next
End Function

Private Sub Command1_Click()
    Shell "explorer.exe """ + tempDir.Path + """"
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Load()
    'load lastCopyTime
    strLastCopyTime = GetSetting("remotdev", "Settings", "lastCopyTime")
    If strLastCopyTime = Empty Then
        dtmLastCopyTime = Now
        strLastCopyTime = CStr(dtmLastCopyTime)
        SaveSetting "remotdev", "Settings", "lastCopyTime", strLastCopyTime
    Else
        dtmLastCopyTime = Cdtm(strLastCopyTime)
    End If
    MsgBox strLastCopyTime
    'load db
    If Not EnsureDBConnection() Then
        MsgBox "Ensure db connection failure!"
        Exit Sub
    End If
    
    'Set Rs = QueryLastFilesTable()
    'create temp dir
    If fso.FolderExists(App.Path & "\temp") Then
        Set tempDir = fso.GetFolder(App.Path & "\temp")
        Dim strBackupDirName As String
        strBackupDirName = "temp" + Format(dtmLastCopyTime, " yyyymmdd-hhmmss")
        If Not fso.FolderExists(App.Path & "\" & strBackupDirName) Then
            tempDir.Name = strBackupDirName
        Else
            tempDir.Delete
        End If
    End If
    Set tempDir = fso.CreateFolder(App.Path & "\temp")
    'copy files
    Set rootDir = fso.GetFolder(App.Path)  'fso.GetFolder(".")
    CopyModifiedFiles2 rootDir
    CloseConnection
    
    'UpdateLastCopyTime
    dtmLastCopyTime = Now
    strLastCopyTime = CStr(dtmLastCopyTime)
    SaveSetting "remotdev", "Settings", "lastCopyTime", strLastCopyTime
End Sub

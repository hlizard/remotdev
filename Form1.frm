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
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
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

Function CopyModifiedFiles(srcDir As Folder) As Boolean
    Dim fsobj As Object
    Dim fd As Folder
    Dim f As File
    For Each fsobj In srcDir
        If TypeOf fsobj Is Folder Then
            Set fd = fsobj
            If fd.DateLastModified > dtmLastCopyTime Then
                Dim strTargetDir As String
                strTargetDir = tempDir.Path + "\" + Replace(fd.Path, rootDir.Path, "")
                If Not fso.FolderExists(strTargetDir) Then
                    fso.CreateFolder strTargetDir
                End If
                CopyModifiedFiles fd
            End If
        ElseIf TypeOf fsobj Is File Then
            Set f = fsobj
            If f.DateLastModified > dtmLastCopyTime Then
                Dim strTargetFile As String
                strTargetFile = tempDir.Path + "\" + Replace(f.Path, rootDir.Path, "")
                If Not fso.FileExists(strTargetFile) Then
                    fso.CopyFile f.Path, strTargetFile
                End If
            End If
        End If
    Next
End Function

Sub UpdateLastCopyTime()
    dtmLastCopyTime = Now
    strLastCopyTime = CStr(dtmLastCopyTime)
    SaveSetting "remotdev", "Settings", "lastCopyTime", strLastCopyTime
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Load()
    strLastCopyTime = GetSetting("remotdev", "Settings", "lastCopyTime")
    If strLastCopyTime = Empty Then
        UpdateLastCopyTime
    Else
        dtmLastCopyTime = Cdtm(strLastCopyTime)
    End If
    MsgBox strLastCopyTime
    'create temp dir
    If fso.FolderExists("temp") Then
        Set tempDir = fso.GetFolder("temp")
        Dim strBackupDirName As String
        strBackupDirName = "temp" + Format(dtmLastCopyTime, " yyyymmdd-hhmmss")
        If Not fso.FolderExists(strBackupDirName) Then
            tempDir.Name = strBackupDirName
        Else
            tempDir.Delete
        End If
    End If
    Set tempDir = fso.CreateFolder("temp")
    'copy files
    Set rootDir = fso.GetFolder(".")
    CopyModifiedFiles rootDir
    
    UpdateLastCopyTime
End Sub

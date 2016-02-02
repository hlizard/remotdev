Attribute VB_Name = "DB"
Option Explicit

Public Enum FsobjQueryResult
    fqrNotFound
    fqrNotChanged
    fqrChanged
End Enum

Private Const CMD_Insert$ = "InsertDescrAndPic"

Private conDB As cConnection

Public Property Get DBFileName() As String
    DBFileName = App.Path & "\remotdev.db3"
End Property

Public Function EnsureDBConnection() As Boolean
    On Error GoTo ExitFalse                                                     'Return False if operation fails.
    
    If New_c.fso.FileExists(DBFileName) Then                                    'normally this is the case
        Set conDB = New_c.Connection(DBFileName, DBOpenFromFile)
        
    Else                                                                        'create a new DB, a new Table + a persistent Insert-Command - and then populate the new table with Data
        Set conDB = New_c.Connection(DBFileName, DBCreateNewFileDB)
        conDB.Execute "Create Table LastFiles (Path Text PRIMARY KEY, IsFile Integer, CreateDate Text, LastModifiedDate Text)"
        conDB.CreateCommand("Insert Into LastFiles Values(?,?,?,?)").Save CMD_Insert
    End If
    
    conDB.BeginTrans
    EnsureDBConnection = True
ExitFalse:
End Function

Public Sub CloseConnection()
    conDB.CommitTrans
End Sub

Public Function QueryLastFilesTable() As cRecordset
    Set QueryLastFilesTable = conDB.OpenRecordset("Select * From LastFiles")
End Function

Public Function QueryFolder(fd As Folder) As FsobjQueryResult
    Dim strQueryDir As String
    strQueryDir = "select * from LastFiles where IsFile=0 and Path='" + fd.Path + "'"
    Dim records As cRecordset
    Set records = conDB.OpenRecordset(strQueryDir)
    If records.RecordCount <= 0 Then
        QueryFolder = fqrNotFound
    ElseIf records!createDate.Value = Format(fd.DateCreated, "yyyymmdd-hhmmss") And records!LastModifiedDate.Value = Format(fd.DateLastModified, "yyyymmdd-hhmmss") Then
        QueryFolder = fqrNotChanged
    Else
        QueryFolder = fqrChanged
    End If
End Function

Public Function QueryFile(fd As File) As FsobjQueryResult
    Dim strQueryDir As String
    strQueryDir = "select * from LastFiles where IsFile=-1 and Path='" + fd.Path + "'"
    Dim records As cRecordset
    Set records = conDB.OpenRecordset(strQueryDir)
    If records.RecordCount <= 0 Then
        QueryFile = fqrNotFound
    ElseIf records!createDate.Value = Format(fd.DateCreated, "yyyymmdd-hhmmss") And records!LastModifiedDate.Value = Format(fd.DateLastModified, "yyyymmdd-hhmmss") Then
        QueryFile = fqrNotChanged
    Else
        QueryFile = fqrChanged
    End If
End Function

Public Sub DeleteFolderFromDB(fd As Folder)
    Dim strQueryDir As String
    strQueryDir = "delete from LastFiles where IsFile=0 and Path='" + fd.Path + "'"
    conDB.Execute strQueryDir
End Sub

Public Sub DeleteFileFromDB(fd As File)
    Dim strQueryDir As String
    strQueryDir = "delete from LastFiles where IsFile=-1 and Path='" + fd.Path + "'"
    conDB.Execute strQueryDir
End Sub

Public Sub InsertFolder(fd As Folder)
    With conDB.CreateCommand(CMD_Insert) 'open the predefined (and persisted) Command-Object
      'update the 2 Parameters of the Cmd-Object in a typed and secure way
      .SetText 1, fd.Path
      .SetInt32 2, CInt(False)
      .SetText 3, Format(fd.DateCreated, "yyyymmdd-hhmmss")
      .SetText 4, Format(fd.DateLastModified, "yyyymmdd-hhmmss")
      .Execute 'insert the new Table-record into the DB-File per Cmd.Execute
  End With
End Sub

Public Sub InsertFile(fd As File)
    With conDB.CreateCommand(CMD_Insert) 'open the predefined (and persisted) Command-Object
      'update the 2 Parameters of the Cmd-Object in a typed and secure way
      .SetText 1, fd.Path
      .SetInt32 2, CInt(True)
      .SetText 3, Format(fd.DateCreated, "yyyymmdd-hhmmss")
      .SetText 4, Format(fd.DateLastModified, "yyyymmdd-hhmmss")
      .Execute 'insert the new Table-record into the DB-File per Cmd.Execute
  End With
End Sub

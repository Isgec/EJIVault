Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Namespace EJI
  Public Class ediAFile
    Implements ICloneable
    Public Function Clone() As Object Implements ICloneable.Clone
      Dim tmp As ediAFile = MyBase.MemberwiseClone()
      Return tmp
    End Function

    Public Property t_drid As String = "" 'PK
    Public Property t_dcid As String = ""
    Public Property t_hndl As String = ""
    Public Property t_indx As String = ""
    Public Property t_prcd As String = ""
    Private Property _t_fnam As String = ""
    Public Property t_lbcd As String = ""
    Public Property t_atby As String = ""
    Private Property _t_aton As String = ""
    Public Property t_Refcntd As Int32 = 0
    Public Property t_Refcntu As Int32 = 0
    Public Property t_fnam As String
      Get
        Return _t_fnam.Replace(",", "-").Replace("#", " ").Replace("'", "")
      End Get
      Set(value As String)
        _t_fnam = value
      End Set
    End Property

    Public Property t_aton As String
      Get
        If _t_aton <> "" Then
          Return Convert.ToDateTime(_t_aton).ToString("dd/MM/yyyy")
        End If
        Return ""
      End Get
      Set(value As String)
        _t_aton = value
      End Set
    End Property
    ''' <summary>
    ''' This is Main record of disk file
    ''' i.e. t_prcd = EJIMAIN
    ''' </summary>
    Public ReadOnly Property IsMainFile As Boolean
      Get
        If t_prcd = "EJIMAIN" Then
          Return True
        Else
          Return False
        End If
      End Get
    End Property
    ''' <summary>
    ''' Returns First Exact Matching Record
    ''' </summary>
    Public Shared Function GetFileByHandleIndex(ByVal t_hndl As String, ByVal t_indx As String) As EJI.ediAFile
      Dim Results As EJI.ediAFile = Nothing
      Using Con As SqlConnection = New SqlConnection(DBCommon.GetBaaNConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = "Select top 1 * from ttcisg132" & DBCommon.ERPCompany & " where t_hndl='" & t_hndl & "' and t_indx='" & t_indx & "'"
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          If Reader.Read() Then
            Results = New EJI.ediAFile(Reader)
          End If
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function
    ''' <summary>
    ''' Return All matching records
    ''' </summary>
    ''' <param name="t_hndl"></param>
    ''' <param name="t_indx"></param>
    ''' <returns>List of Files</returns>
    Public Shared Function GetFilesByHandleIndex(ByVal t_hndl As String, ByVal t_indx As String) As List(Of EJI.ediAFile)
      Dim Results As New List(Of EJI.ediAFile)
      Using Con As SqlConnection = New SqlConnection(DBCommon.GetBaaNConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = "Select * from ttcisg132" & DBCommon.ERPCompany & " where t_hndl='" & t_hndl & "' and t_indx='" & t_indx & "'"
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While Reader.Read()
            Results.Add(New EJI.ediAFile(Reader))
          End While
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function
    ''' <summary>
    ''' Returns All Exactly Matching Handle and Starts with Index
    ''' </summary>
    ''' <param name="t_hndl"></param>
    ''' <param name="t_indx"></param>
    ''' <returns>List of Files</returns>
    Public Shared Function GetFilesByHandleLikeIndex(ByVal t_hndl As String, ByVal t_indx As String) As List(Of EJI.ediAFile)
      Dim Results As New List(Of EJI.ediAFile)
      Using Con As SqlConnection = New SqlConnection(DBCommon.GetBaaNConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = "Select * from ttcisg132" & DBCommon.ERPCompany & " where t_hndl='" & t_hndl & "' and t_indx like '" & t_indx & "%'"
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While (Reader.Read())
            Results.Add(New EJI.ediAFile(Reader))
          End While
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function

    Public Shared Function GetFileByRecordID(ByVal t_drid As String) As EJI.ediAFile
      Dim Results As EJI.ediAFile = Nothing
      Using Con As SqlConnection = New SqlConnection(DBCommon.GetBaaNConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = "Select * from ttcisg132" & DBCommon.ERPCompany & " where t_drid='" & t_drid & "'"
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          If Reader.Read() Then
            Results = New EJI.ediAFile(Reader)
          End If
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function
    ''' <summary>
    ''' Returns all records of matching File ID
    ''' </summary>
    ''' <param name="t_dcid"></param>
    ''' <returns>Returns all records of matching File ID</returns>
    Public Shared Function GetFilesByFileID(ByVal t_dcid As String) As List(Of EJI.ediAFile)
      Dim Results As New List(Of EJI.ediAFile)
      Using Con As SqlConnection = New SqlConnection(DBCommon.GetBaaNConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = "Select * from ttcisg132" & DBCommon.ERPCompany & " where t_dcid='" & t_dcid & "'"
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While Reader.Read()
            Results.Add(New EJI.ediAFile(Reader))
          End While
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function
    ''' <summary>
    ''' Returns First Matching File ID
    ''' </summary>
    ''' <param name="t_dcid"></param>
    ''' <returns></returns>
    Public Shared Function GetFileByFileID(ByVal t_dcid As String) As EJI.ediAFile
      Dim Results As EJI.ediAFile = Nothing
      Using Con As SqlConnection = New SqlConnection(DBCommon.GetBaaNConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = "Select Top 1 * from ttcisg132" & DBCommon.ERPCompany & " where t_dcid='" & t_dcid & "'"
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While Reader.Read()
            Results = New EJI.ediAFile(Reader)
          End While
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function

    Public Shared Function InsertData(ByVal Record As EJI.ediAFile) As EJI.ediAFile
      Using Con As SqlConnection = New SqlConnection(DBCommon.GetBaaNConnectionString())
        Dim Sql As String = ""
        Sql &= "  INSERT [ttcisg132" & DBCommon.ERPCompany & "] ([t_drid],[t_dcid],[t_hndl],[t_indx],[t_prcd],[t_fnam],[t_lbcd],[t_atby],[t_aton],[t_Refcntd],[t_Refcntu]) "
        Sql &= "  VALUES ("
        Sql &= "   '" & Record.t_drid & "'"
        Sql &= "  ,'" & Record.t_dcid & "'"
        Sql &= "  ,'" & Record.t_hndl & "'"
        Sql &= "  ,'" & Record.t_indx & "'"
        Sql &= "  ,'" & Record.t_prcd & "'"
        Sql &= "  ,'" & Record.t_fnam & "'"
        Sql &= "  ,'" & Record.t_lbcd & "'"
        Sql &= "  ,'" & Record.t_atby & "'"
        Sql &= "  ,convert(datetime,'" & Convert.ToDateTime(Record.t_aton).ToString("dd/MM/yyyy") & "',103)"
        Sql &= "  ,0"
        Sql &= "  ,0"
        Sql &= "  )"
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = Sql
          Con.Open()
          Cmd.ExecuteNonQuery()
        End Using
      End Using
      Return Record
    End Function
    ''' <summary>
    ''' Simply Deletes the Record.
    ''' </summary>
    ''' <param name="Record"></param>
    ''' <returns></returns>
    Public Shared Function DeleteData(ByVal Record As EJI.ediAFile) As EJI.ediAFile
      Using Con As SqlConnection = New SqlConnection(DBCommon.GetBaaNConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = "Delete ttcisg132" & DBCommon.ERPCompany & " where t_drid = '" & Record.t_drid & "'"
          Con.Open()
          Cmd.ExecuteNonQuery()
        End Using
      End Using
      Return Record
    End Function
    ''' <summary>
    ''' Simply Delete the records of matching handle index
    ''' </summary>
    ''' <param name="t_hndl"></param>
    ''' <param name="t_indx"></param>
    ''' <returns>List of Deleted Records</returns>
    Public Shared Function DeleteDataByHandleIndex(ByVal t_hndl As String, ByVal t_indx As String) As List(Of EJI.ediAFile)
      Dim Results As New List(Of EJI.ediAFile)
      Using Con As SqlConnection = New SqlConnection(DBCommon.GetBaaNConnectionString())
        Con.Open()
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = "Select * from ttcisg132" & DBCommon.ERPCompany & " where t_hndl='" & t_hndl & "' and t_indx='" & t_indx & "'"
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While Reader.Read()
            Results.Add(New EJI.ediAFile(Reader))
          End While
          Reader.Close()
        End Using
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = "delete ttcisg132" & DBCommon.ERPCompany & " where t_hndl='" & t_hndl & "' and t_indx='" & t_indx & "'"
          Cmd.ExecuteNonQuery()
        End Using
      End Using
      Return Results
    End Function

    Public Shared Function UpdateData(ByVal Record As EJI.ediAFile) As EJI.ediAFile
      Using Con As SqlConnection = New SqlConnection(DBCommon.GetBaaNConnectionString())
        Dim Sql As String = ""
        Sql &= "  UPDATE [ttcisg132" & DBCommon.ERPCompany & "] "
        Sql &= "  SET "
        Sql &= "   [t_dcid] = '" & Record.t_dcid & "'"
        Sql &= "  ,[t_hndl] = '" & Record.t_hndl & "'"
        Sql &= "  ,[t_indx] = '" & Record.t_indx & "'"
        Sql &= "  ,[t_prcd] = '" & Record.t_prcd & "'"
        Sql &= "  ,[t_fnam] = '" & Record.t_fnam & "'"
        Sql &= "  ,[t_lbcd] = '" & Record.t_lbcd & "'"
        Sql &= "  ,[t_atby] = '" & Record.t_atby & "'"
        Sql &= "  ,convert(datetime,'" & Convert.ToDateTime(Record.t_aton).ToString("dd/MM/yyyy") & "',103)"
        Sql &= "  WHERE [t_drid] = '" & Record.t_drid & "'"
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = Sql
          Con.Open()
          Cmd.ExecuteNonQuery()
        End Using
      End Using
      Return Record
    End Function

    Public Shared Function GetDownloadPathFileName(ByVal Record As EJI.ediAFile) As String
      Dim ejiLib As EJI.ediALib = EJI.ediALib.GetLibraryByID(Record.t_lbcd)
      Dim TargetPathFileName As String = ejiLib.t_path & "\" & Record.t_dcid
      Dim TargetExists As Boolean = False
      Try
        TargetExists = My.Computer.FileSystem.FileExists(TargetPathFileName)
      Catch ex As Exception
        Throw New ediException(1008, "Error while checking Target File Exists.")
      End Try
      If Not TargetExists Then
        Throw New ediException(1010, "Error Disk File NOT Found in Vault.")
      End If
      Return TargetPathFileName
    End Function

    Public Shared Function InsertUploaded(ByVal Record As EJI.ediAFile, PathFileName As String) As EJI.ediAFile
      Dim FileExists As Boolean = False
      Try
        FileExists = My.Computer.FileSystem.FileExists(PathFileName)
      Catch ex As Exception
        Throw New ediException(1002, "Error while checking File Exists.")
      End Try
      If Not FileExists Then
        Throw New ediException(1006, "Uploaded File Not Found.")
      End If
      Dim ejiLib As EJI.ediALib = EJI.ediALib.GetActiveLibrary()
      With Record
        .t_drid = EJI.ediASeries.GetNextRecordID()
        .t_dcid = EJI.ediASeries.GetNextFileID()
        .t_prcd = "EJIMAIN"
        .t_aton = Now.ToString("dd/MM/yyyy")
        .t_lbcd = ejiLib.t_lbcd
      End With
      Dim TargetPathFileName As String = ejiLib.t_path & "\" & Record.t_dcid
      Dim mayInsertRecord As Boolean = False
      Dim TargetExists As Boolean = False
      Try
        TargetExists = My.Computer.FileSystem.FileExists(TargetPathFileName)
      Catch ex As Exception
        Throw New ediException(1008, "Error while checking Target File Exists.")
      End Try
      If Not TargetExists Then
        Try
          My.Computer.FileSystem.CopyFile(PathFileName, TargetPathFileName)
          mayInsertRecord = True
        Catch ex As Exception
          Throw New ediException(1009, "Error while copying disk file to Vault.")
        End Try
      Else
        mayInsertRecord = True
      End If
      If mayInsertRecord Then
        InsertData(Record)
      End If
      Return Record
    End Function
    ''' <summary>
    ''' If Copied Records Found, i.e. t_prcd of any record contains t_drid
    ''' </summary>
    ''' <param name="t_drid"></param>
    ''' <returns></returns>
    Public Shared Function IsFileCopied(ByVal t_drid As String) As Boolean
      Dim mRet As Boolean = False
      Using Con As SqlConnection = New SqlConnection(DBCommon.GetBaaNConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = "Select isnull(count(*),0) as cnt from ttcisg132" & DBCommon.ERPCompany & " where t_prcd='" & t_drid & "'"
          Con.Open()
          Dim cnt As Integer = Cmd.ExecuteScalar
          If cnt > 0 Then
            mRet = True
          End If
        End Using
      End Using
      Return mRet
    End Function
    ''' <summary>
    ''' Deletes only when record is not further copied
    ''' If Record is Copied => Deletes only meta data
    ''' If Record is Main   => Deletes also Disk file
    ''' </summary>
    ''' <param name="t_drid"></param>
    Public Shared Sub FileDelete(ByVal t_drid As String)
	'Use IO Functions
	'Map path if IsLocalIsgecVault
      Dim tmp As ediAFile = GetFileByRecordID(t_drid)
      If IsFileCopied(t_drid) Then
        Throw New ediException(1001, "File is copied CANNOT delete.")
      End If
      Dim mayDeleteRecord As Boolean = True
      If tmp.IsMainFile Then
        mayDeleteRecord = False
        Dim tmpLib As ediALib = ediALib.GetLibraryByID(tmp.t_lbcd)
        Dim PathFileName As String = tmpLib.LibraryPath & "\" & tmp.t_dcid
        If Not EJI.DBCommon.IsLocalISGECVault Then
          EJI.ediALib.ConnectISGECVault(tmpLib)
        End If
        Dim FileExists As Boolean = False
        Try
          FileExists = IO.File.Exists(PathFileName)
        Catch ex As Exception
          Throw New ediException(1002, "Error while checking File Exists.")
        End Try
        If FileExists Then
          Try
            IO.File.Delete(PathFileName)
            mayDeleteRecord = True
          Catch ex As Exception
            Throw New ediException(1003, "Error while deleting disk file.")
          End Try
        Else
          mayDeleteRecord = True
        End If
      End If
      If mayDeleteRecord Then
        Try
          Using Con As SqlConnection = New SqlConnection(DBCommon.GetBaaNConnectionString())
            Using Cmd As SqlCommand = Con.CreateCommand()
              Cmd.CommandType = CommandType.Text
              Cmd.CommandText = " delete ttcisg132" & DBCommon.ERPCompany & " where t_drid='" & t_drid & "'"
              Con.Open()
              Cmd.ExecuteNonQuery()
            End Using
          End Using
        Catch ex As Exception
          Throw New ediException(1004, "Error while deleting database record.")
        End Try
      End If
    End Sub
    ''' <summary>
    ''' Only for Main File
    ''' Deletes Disk File and all records which contains same File ID
    ''' </summary>
    ''' <param name="t_drid"></param>
    Public Shared Sub ForcedFileDelete(ByVal t_drid As String)
      Dim tmp As ediAFile = GetFileByRecordID(t_drid)
      If Not tmp.IsMainFile Then
        Throw New ediException(1005, "Forced File delete can be performed from EJIMAIN file.")
      End If
      Dim mayDeleteRecord As Boolean = True
      If tmp.IsMainFile Then
        mayDeleteRecord = False
        Dim tmpLib As ediALib = ediALib.GetLibraryByID(tmp.t_lbcd)
        Dim PathFileName As String = tmpLib.LibraryPath & "\" & tmp.t_dcid
        If Not EJI.DBCommon.IsLocalISGECVault Then
          EJI.ediALib.ConnectISGECVault(tmpLib)
        End If
        Dim FileExists As Boolean = False
        Try
          FileExists = IO.File.Exists(PathFileName)
        Catch ex As Exception
          Throw New ediException(1002, "Error while checking File Exists.")
        End Try
        If FileExists Then
          Try
            IO.File.Delete(PathFileName)
            mayDeleteRecord = True
          Catch ex As Exception
            Throw New ediException(1003, "Error while deleting disk file.")
          End Try
        Else
          mayDeleteRecord = True
        End If
      End If
      If mayDeleteRecord Then
        Try
          Using Con As SqlConnection = New SqlConnection(DBCommon.GetBaaNConnectionString())
            Using Cmd As SqlCommand = Con.CreateCommand()
              Cmd.CommandType = CommandType.Text
              Cmd.CommandText = " delete ttcisg132" & DBCommon.ERPCompany & " where t_dcid='" & tmp.t_dcid & "'"
              Con.Open()
              Cmd.ExecuteNonQuery()
            End Using
          End Using
        Catch ex As Exception
          Throw New ediException(1004, "Error while deleting database record.")
        End Try
      End If
    End Sub
    ''' <summary>
    ''' Deletes all copied records which contains same file ID
    ''' Does not deletes main file and disk file
    ''' </summary>
    ''' <param name="t_drid"></param>
    Public Shared Sub ForcedFileDeleteNotMain(ByVal t_drid As String)
      Dim tmp As ediAFile = GetFileByRecordID(t_drid)
      Try
        Using Con As SqlConnection = New SqlConnection(DBCommon.GetBaaNConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = " delete ttcisg132" & DBCommon.ERPCompany & " where t_dcid='" & tmp.t_dcid & "' and t_prcd <> 'EJIMAIN'"
            Con.Open()
            Cmd.ExecuteNonQuery()
          End Using
        End Using
      Catch ex As Exception
        Throw New ediException(1004, "Error while deleting database record.")
      End Try
    End Sub
    ''' <summary>
    ''' Deletes all records including main which contains same FileID
    ''' also deletes disk file
    ''' </summary>
    ''' <param name="t_drid"></param>
    Public Shared Sub ForcedFileDeleteIncludingMain(ByVal t_drid As String)
      Dim tmp As ediAFile = GetFileByRecordID(t_drid)
      Dim mayDeleteRecord As Boolean = False
      Dim tmpLib As ediALib = ediALib.GetLibraryByID(tmp.t_lbcd)
      Dim PathFileName As String = tmpLib.LibraryPath & "\" & tmp.t_dcid
      If Not EJI.DBCommon.IsLocalISGECVault Then
        EJI.ediALib.ConnectISGECVault(tmpLib)
      End If
      Dim FileExists As Boolean = False
      Try
        FileExists = IO.File.Exists(PathFileName)
      Catch ex As Exception
        Throw New ediException(1002, "Error while checking File Exists.")
      End Try
      If FileExists Then
        Try
          IO.File.Delete(PathFileName)
          mayDeleteRecord = True
        Catch ex As Exception
          Throw New ediException(1003, "Error while deleting disk file.")
        End Try
      Else
        mayDeleteRecord = True
      End If
      If mayDeleteRecord Then
        Try
          Using Con As SqlConnection = New SqlConnection(DBCommon.GetBaaNConnectionString())
            Using Cmd As SqlCommand = Con.CreateCommand()
              Cmd.CommandType = CommandType.Text
              Cmd.CommandText = " delete ttcisg132" & DBCommon.ERPCompany & " where t_dcid='" & tmp.t_dcid & "'"
              Con.Open()
              Cmd.ExecuteNonQuery()
            End Using
          End Using
        Catch ex As Exception
          Throw New ediException(1004, "Error while deleting database record.")
        End Try
      End If
    End Sub
    ''' <summary>
    ''' Returns the List of Immediate Child records which contains t_drid in t_prcd
    ''' </summary>
    ''' <param name="t_drid"></param>
    ''' <returns></returns>
    Public Shared Function GetChildFilesByRecordID(ByVal t_drid As String) As List(Of EJI.ediAFile)
      Dim Results As New List(Of EJI.ediAFile)
      Using Con As SqlConnection = New SqlConnection(DBCommon.GetBaaNConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = "Select * from ttcisg132" & DBCommon.ERPCompany & " where t_prcd='" & t_drid & "' "
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While Reader.Read()
            Results.Add(New EJI.ediAFile(Reader))
          End While
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function
    ''' <summary>
    ''' Hierarical Delete upto bottom from current branch
    ''' If current branch is main then deletes disk file too
    ''' </summary>
    ''' <param name="t_drid"></param>
    Public Shared Sub FileDeleteIncludingCopy(ByVal t_drid As String)
      Dim tmp As ediAFile = GetFileByRecordID(t_drid)
      If IsFileCopied(t_drid) Then
        Dim childs As List(Of ediAFile) = GetChildFilesByRecordID(t_drid)
        For Each chld As ediAFile In childs
          FileDeleteIncludingCopy(chld.t_drid)
        Next
      End If
      Dim mayDeleteRecord As Boolean = True
      If tmp.IsMainFile Then
        mayDeleteRecord = False
        Dim tmpLib As ediALib = ediALib.GetLibraryByID(tmp.t_lbcd)
        Dim PathFileName As String = tmpLib.LibraryPath & "\" & tmp.t_dcid

        If Not EJI.DBCommon.IsLocalISGECVault Then
          EJI.ediALib.ConnectISGECVault(tmpLib)
        End If
        Dim FileExists As Boolean = False
        Try
          FileExists = IO.File.Exists(PathFileName)
        Catch ex As Exception
          Throw New ediException(1002, "Error while checking File Exists.")
        End Try
        If FileExists Then
          Try
            IO.File.Delete(PathFileName)
            mayDeleteRecord = True
          Catch ex As Exception
            Throw New ediException(1003, "Error while deleting disk file.")
          End Try
        Else
          mayDeleteRecord = True
        End If
      End If
      If mayDeleteRecord Then
        Try
          Using Con As SqlConnection = New SqlConnection(DBCommon.GetBaaNConnectionString())
            Using Cmd As SqlCommand = Con.CreateCommand()
              Cmd.CommandType = CommandType.Text
              Cmd.CommandText = " delete ttcisg132" & DBCommon.ERPCompany & " where t_drid='" & t_drid & "'"
              Con.Open()
              Cmd.ExecuteNonQuery()
            End Using
          End Using
        Catch ex As Exception
          Throw New ediException(1004, "Error while deleting database record.")
        End Try
      End If
    End Sub
    ''' <summary>
    ''' Copies Exact Matching Records Only
    ''' </summary>
    ''' <param name="s_t_hndl"></param>
    ''' <param name="s_t_indx"></param>
    ''' <param name="t_t_hndl"></param>
    ''' <param name="t_t_indx"></param>
    ''' <param name="CreatedBy">May be empty string to ensure, this function is called</param>
    ''' <returns>CreatedBy must be supplied to insure to execute FileCopy by HandleIndex</returns>
    Public Shared Function FileCopy(ByVal s_t_hndl As String, ByVal s_t_indx As String, ByVal t_t_hndl As String, ByVal t_t_indx As String, CreatedBy As String) As List(Of EJI.ediAFile)
      Dim Results As List(Of EJI.ediAFile) = EJI.ediAFile.GetFilesByHandleIndex(s_t_hndl, s_t_indx)
      For Each tmp As EJI.ediAFile In Results
        With tmp
          .t_prcd = .t_drid 'Set Parent DRID
          .t_drid = EJI.ediASeries.GetNextRecordID()
          .t_hndl = t_t_hndl
          .t_indx = t_t_indx
          If CreatedBy <> "" Then
            .t_atby = CreatedBy
          End If
          .t_aton = Now.ToString("dd/MM/yyyy")
        End With
        tmp = EJI.ediAFile.InsertData(tmp)
      Next
      Return Results
    End Function
    Public Shared Function RecordCopy(ByVal t_drid As String, ByVal t_t_hndl As String, ByVal t_t_indx As String, CreatedBy As String) As EJI.ediAFile
      Dim Results As EJI.ediAFile = EJI.ediAFile.GetFileByRecordID(t_drid)
      With Results
        .t_prcd = .t_drid 'Set Parent DRID
        .t_drid = EJI.ediASeries.GetNextRecordID()
        .t_hndl = t_t_hndl
        .t_indx = t_t_indx
        If CreatedBy <> "" Then
          .t_atby = CreatedBy
        End If
        .t_aton = Now.ToString("dd/MM/yyyy")
      End With
      Results = EJI.ediAFile.InsertData(Results)
      Return Results
    End Function
    Public Sub New(ByVal Rd As SqlDataReader)
      DBCommon.NewObj(Me, Rd)
    End Sub
    Public Sub New()
    End Sub
  End Class
End Namespace

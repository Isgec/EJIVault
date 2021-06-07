Imports System.Data.SqlClient
Imports System.ComponentModel
Namespace EJI
  Public Class ediASeries
    Private Shared _RecordCount As Integer
    Public Property t_seri As String = ""
    Public Property t_rnum As Int32 = 0
    Public Property t_acti As String = ""
    Public Property t_Refcntd As Int32 = 0
    Public Property t_Refcntu As Int32 = 0
    Public Shared Function GetNextRecordID() As String
      Dim mRet As String = ""
      Using Con As SqlConnection = New SqlConnection(DBCommon.GetBaaNConnectionString())
        Con.Open()
        Dim mSql As String = ""
        mSql &= " BEGIN TRANSACTION "
        mSql &= " Update ttcisg131" & DBCommon.FixedCompany & " set t_rnum=t_rnum+1 where t_acti='Z' "
        mSql &= " Select top 1 isnull(t_seri + Ltrim(str(t_rnum)),'') as FileName from ttcisg131" & DBCommon.FixedCompany & " where t_acti = 'Z' "
        mSql &= " COMMIT TRANSACTION "
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = mSql
          mRet = Cmd.ExecuteScalar
        End Using
      End Using
      Return mRet
    End Function
    Public Shared Function GetNextFileID() As String
      Dim mRet As String = ""
      Using Con As SqlConnection = New SqlConnection(DBCommon.GetBaaNConnectionString())
        Con.Open()
        Dim mSql As String = ""
        mSql &= " BEGIN TRANSACTION "
        mSql &= " Update ttcisg131" & DBCommon.FixedCompany & " set t_rnum=t_rnum+1 where t_acti='T' "
        mSql &= " Select top 1 isnull(t_seri + '_' + Ltrim(str(t_rnum)),'') as FileName from ttcisg131" & DBCommon.FixedCompany & " where t_acti = 'T' "
        mSql &= " COMMIT TRANSACTION "
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = mSql
          mRet = Cmd.ExecuteScalar
        End Using
      End Using
      Return mRet
    End Function
    Public Shared Function GetActiveSeries() As EJI.ediASeries
      Dim Results As EJI.ediASeries = Nothing
      Using Con As SqlConnection = New SqlConnection(DBCommon.GetBaaNConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = "select top 1 * from ttcisg131" & DBCommon.FixedCompany & " where t_acti = 'T'"
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          If Reader.Read() Then
            Results = New EJI.ediASeries(Reader)
          End If
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function
    Public Sub New(ByVal Rd As SqlDataReader)
      DBCommon.NewObj(Me, Rd)
    End Sub
    Public Sub New()
    End Sub
  End Class
End Namespace

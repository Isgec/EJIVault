Imports System.Data.SqlClient
Imports System.ComponentModel
Imports System
Namespace EJI
  Public Class ediALib
    Public Property t_lbcd As String = ""
    Public Property t_desc As String = ""
    Public Property t_serv As String = ""
    Public Property t_path As String = ""
    Public Property t_acti As String = ""
    Public Property t_Refcntd As Int32 = 0
    Public Property t_Refcntu As Int32 = 0
    Public ReadOnly Property LibraryPath As String
      Get
        If EJI.DBCommon.IsLocalISGECVault Then
          Return t_desc.Trim & "\" & t_path.Trim
        Else
          Return "\\" & DBCommon.ISGECVaultIP & "\" & t_path.Trim
        End If
      End Get
    End Property
    Public Shared Function ConnectISGECVault(oLib As EJI.ediALib) As Boolean
      If Not EJI.DBCommon.IsLocalISGECVault Then
        Try
          ConnectToNetworkFunctions.disconnectFromNetwork("X:")
        Catch ex As Exception
        End Try
        Dim mRet As Boolean = False
        If Not EJI.DBCommon.IsLocalISGECVault Then
          mRet = ConnectToNetworkFunctions.connectToNetwork(oLib.LibraryPath, "X:", "administrator", "Indian@12345")
        End If
        Return mRet
      End If
      Return True
    End Function
    Public Shared Function DisconnectISGECVault() As Boolean
      If Not EJI.DBCommon.IsLocalISGECVault Then
        Try
          Return ConnectToNetworkFunctions.disconnectFromNetwork("X:")
        Catch ex As Exception
        End Try
        Return False
      Else
        Return True
      End If
    End Function
    Public Shared Function GetActiveLibrary() As EJI.ediALib
      Dim Results As EJI.ediALib = Nothing
      Using Con As SqlConnection = New SqlConnection(DBCommon.GetBaaNConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = "select top 1 * from ttcisg127" & DBCommon.FixedCompany & " where t_acti = 'T'"
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          If Reader.Read() Then
            Results = New EJI.ediALib(Reader)
          End If
          Reader.Close()
        End Using
      End Using
      If Results Is Nothing Then
        Throw New ediException(1007, "Active Library Record NOT Found. [t_acti=T]")
      End If
      Return Results
    End Function
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function GetLibraryByID(ByVal t_lbcd As String) As EJI.ediALib
      Dim Results As EJI.ediALib = Nothing
      Using Con As SqlConnection = New SqlConnection(DBCommon.GetBaaNConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = "select top 1 * from ttcisg127" & DBCommon.FixedCompany & " where t_lbcd = '" & t_lbcd & "'"
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          If Reader.Read() Then
            Results = New EJI.ediALib(Reader)
          End If
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function
    Public Sub New(ByVal rd As SqlDataReader)
      DBCommon.NewObj(Me, rd)
    End Sub
    Public Sub New()
    End Sub
  End Class
End Namespace

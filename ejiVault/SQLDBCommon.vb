Imports System.Data.SqlClient
Namespace EJI
  Public Class DBCommon
    Implements IDisposable
    Public Shared Property BaaNLive As Boolean = True
    Public Shared Property JoomlaLive As Boolean = True
    Public Shared Property ERPCompany As String = "200"
    Public Shared Property FixedCompany As String = "200"
    Public Shared Property IsLocalISGECVault As Boolean = True
    Public Shared Property ISGECVaultIP As String = ""
    Public Shared Function ConnectLibrary(LibraryID As String) As Boolean
      If Not EJI.DBCommon.IsLocalISGECVault Then
        ConnectToNetworkFunctions.disconnectFromNetwork("X:")
        Return ConnectToNetworkFunctions.connectToNetwork(ediALib.GetLibraryByID(LibraryID).LibraryPath, "X:", "administrator", "Indian@12345")
      Else
        Return True
      End If
    End Function
    Public Shared Function ConnectLibrary() As Boolean
      If Not EJI.DBCommon.IsLocalISGECVault Then
        ConnectToNetworkFunctions.disconnectFromNetwork("X:")
        Return ConnectToNetworkFunctions.connectToNetwork(ediALib.GetActiveLibrary.LibraryPath, "X:", "administrator", "Indian@12345")
      Else
        Return True
      End If
    End Function
    Public Shared Function DisconnectLibrary() As Boolean
      Try
        Return ConnectToNetworkFunctions.disconnectFromNetwork("X:")
      Catch ex As Exception
        Return False
      End Try
      Return True
    End Function

    ''' <summary>
    ''' To work with configuration setting of EJIVault
    ''' </summary>
    Public Shared Sub Initialize()
      modMain.Initialize()
    End Sub
    Public Shared Function GetBaaNConnectionString() As String
      If BaaNLive Then
        Return "Data Source=192.9.200.129;Initial Catalog=inforerpdb;Integrated Security=False;User Instance=False;Persist Security Info=True;User ID=lalit;Password=scorpions"
      Else
        Return "Data Source=192.9.200.45;Initial Catalog=inforerpdb;Integrated Security=False;User Instance=False;Persist Security Info=True;User ID=lalit;Password=scorpions"
      End If
    End Function
    Public Shared Function GetConnectionString() As String
      If JoomlaLive Then
        Return "Data Source=192.9.200.150;Initial Catalog=IJTPerks;Integrated Security=False;User Instance=False;Persist Security Info=True;User ID=sa;Password=isgec12345"
      Else
        Return "Data Source=.\LGSQL;Initial Catalog=IJTPerks;Integrated Security=False;User Instance=False;Persist Security Info=True;User ID=sa;Password=isgec12345"
      End If
    End Function
    Shared Sub New()
    End Sub
    Public Shared Sub AddDBParameter(ByRef Cmd As SqlCommand, ByVal name As String, ByVal type As SqlDbType, ByVal size As Integer, ByVal value As Object)
      Dim Parm As SqlParameter = Cmd.CreateParameter()
      Parm.ParameterName = name
      Parm.SqlDbType = type
      Parm.Size = size
      Parm.Value = value
      Cmd.Parameters.Add(Parm)
    End Sub
    Public Shared Function NewObj(this As Object, Reader As SqlDataReader) As Object
      Try
        For Each pi As System.Reflection.PropertyInfo In this.GetType.GetProperties
          If pi.MemberType = Reflection.MemberTypes.Property Then
            Try
              Dim Found As Boolean = False
              For I As Integer = 0 To Reader.FieldCount - 1
                If Reader.GetName(I).ToLower = pi.Name.ToLower Then
                  Found = True
                  Exit For
                End If
              Next
              If Found Then
                If Convert.IsDBNull(Reader(pi.Name)) Then
                  Select Case Reader.GetDataTypeName(Reader.GetOrdinal(pi.Name))
                    Case "decimal"
                      CallByName(this, pi.Name, CallType.Let, "0.00")
                    Case "bit"
                      CallByName(this, pi.Name, CallType.Let, Boolean.FalseString)
                    Case Else
                      CallByName(this, pi.Name, CallType.Let, String.Empty)
                  End Select
                Else
                  CallByName(this, pi.Name, CallType.Let, Reader(pi.Name))
                End If
              End If
            Catch ex As Exception
            End Try
          End If
        Next
      Catch ex As Exception
        Return Nothing
      End Try
      Return this
    End Function

#Region " IDisposable Support "
    Private disposedValue As Boolean = False    ' To detect redundant calls
    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
      If Not Me.disposedValue Then
        If disposing Then
          ' TODO: free unmanaged resources when explicitly called
        End If

        ' TODO: free shared unmanaged resources
      End If
      Me.disposedValue = True
    End Sub
    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
      ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
      Dispose(True)
      GC.SuppressFinalize(Me)
    End Sub
#End Region
  End Class

End Namespace

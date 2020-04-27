''' <summary>
''' Several functions to connect to a network (with username and password) an copy a file on there.
''' Note that you have to choose a network-drive-letter (e.g.: X:)
''' </summary>
''' <remarks>It's just a little bit of code, there are no error handlers. A CMD is needed. Maybe later I'll build up this dll.</remarks>
Public Class ConnectToNetworkFunctions

  ''' <summary>
  ''' Connects the system to a network drive and waits until the system is connected
  ''' </summary>
  ''' <param name="serverString">UNC Path to servershare (e.g.: "\\mySystem\c$\temp\")</param>
  ''' <param name="driveLetter">Local driveletter for temporary connection (e.g.: "X:")</param>
  ''' <returns>True if connection establisht, otherwise false</returns>
  ''' <remarks></remarks>
  Public Overloads Shared Function connectToNetwork(ByVal serverString As String, ByVal driveLetter As String) As Boolean
    Try
      Dim connectionProcessStartInfo As New ProcessStartInfo("cmd.exe", String.Format("/C net use {0} {1} /PERSISTENT:NO", driveLetter, serverString))
      connectionProcessStartInfo.WindowStyle = ProcessWindowStyle.Hidden
      Process.Start(connectionProcessStartInfo).WaitForExit()
      Return True
    Catch ex As Exception
    End Try
    Return False
  End Function

  ''' <summary>
  ''' Connects the system to a network drive and waits until the system is connected
  ''' </summary>
  ''' <param name="serverString">UNC Path to servershare (e.g.: "\\mySystem\c$\temp\")</param>
  ''' <param name="driveLetter">Local driveletter for temporary connection (e.g.: "X:")</param>
  ''' <param name="username">Username on the remote share (e.g.: "Adminstrator")</param>
  ''' <param name="password">Password for the user (e.g.: "123456")</param>
  ''' <returns>True if connection establisht, otherwise false</returns>
  ''' <remarks></remarks>
  Public Overloads Shared Function connectToNetwork(ByVal serverString As String, ByVal driveLetter As String, ByVal username As String, ByVal password As String) As Boolean
    Try
      Dim connectionProcessStartInfo As New ProcessStartInfo("cmd.exe", String.Format("/C net use {0} {1} {2} /USER:{3} /PERSISTENT:NO", driveLetter, serverString, password, username))
      connectionProcessStartInfo.WindowStyle = ProcessWindowStyle.Hidden
      Process.Start(connectionProcessStartInfo).WaitForExit()
      Return True
    Catch ex As Exception
    End Try
    Return False
  End Function

  ''' <summary>
  ''' Disconnects the system from a network drive, that was initialited before by connectToNetwork
  ''' </summary>
  ''' <param name="driveLetter">Driveletter to disconnect (e.g.: "X:")</param>
  ''' <returns>True if connection closed, otherwise false</returns>
  ''' <remarks></remarks>
  Public Shared Function disconnectFromNetwork(ByVal driveLetter As String) As Boolean
    Try
      Dim connectionProcessStartInfo As New ProcessStartInfo("cmd.exe", String.Format("/C net use {0} /DELETE", driveLetter))
      connectionProcessStartInfo.WindowStyle = ProcessWindowStyle.Hidden
      Process.Start(connectionProcessStartInfo).WaitForExit()
      Return True
    Catch ex As Exception
    End Try
    Return False
  End Function

  ''' <summary>
  ''' Copy a file onto a remote share with password security. All parameters are needed! After the copy is done, the connection will be disconnected
  ''' </summary>
  ''' <param name="sourceFilePath">Path to the source file you want to copy (e.g.: "C:\Temp\Testfile.doc")</param>
  ''' <param name="remoteFileName">Filename on the remote computer (e.g.: "TestfileCopy.doc")</param>
  ''' <param name="serverString">UNC Path to servershare (e.g.: "\\mySystem\c$\temp\")</param>
  ''' <param name="driveLetter">Local driveletter for temporary connection (e.g.: "X:")</param>
  ''' <param name="username">Username on the remote share (e.g.: "Adminstrator")</param>
  ''' <param name="password">Password for the user (e.g.: "123456")</param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Shared Function copyFileToSecurityNetworkShare(ByVal sourceFilePath As String, ByVal remoteFileName As String, ByVal serverString As String, ByVal driveLetter As String, ByVal username As String, ByVal password As String) As Boolean
    If System.IO.File.Exists(sourceFilePath) Then
      If serverString.Substring(serverString.Length - 1) = "\"c Then serverString = serverString.Substring(0, serverString.Length - 1)
      connectToNetwork(serverString, driveLetter, username, password)
      System.IO.File.Copy(sourceFilePath, driveLetter & "\" & remoteFileName, True)
      disconnectFromNetwork(driveLetter)
      Return True
    Else
      Throw New Exception("Sourcefile doesn't exist! Check path in your code!")
      Return False
    End If
  End Function
End Class
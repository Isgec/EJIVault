Imports System.Xml
Imports System.Xml.Serialization
Namespace EJI
  <Serializable>
  Public Class ConfigFile
    Implements ICloneable
    Public Function Clone() As Object Implements ICloneable.Clone
      Return MyBase.MemberwiseClone()
    End Function

    Public Property BaaNLive As Boolean = False
    Public Property JoomlaLive As Boolean = False
    Public Property Testing As Boolean = False
    'Derived Property
    Public Property StartupPath As String = ""
    Public Property TempPath As String = ""
    Public Property SerializedAt As String = ""
    Public Shared Function GetFile(ByVal FilePath As String) As ConfigFile
      Dim tmp As ConfigFile = Nothing
      If IO.File.Exists(FilePath) Then
        Dim rd As XmlReader = Nothing
        Try
          tmp = New ConfigFile
          rd = XmlReader.Create(FilePath)
          While rd.Read
            If rd.NodeType = XmlNodeType.Element Then
              Select Case rd.Name
                Case "TempPath"
                  Try
                    rd.Read()
                    tmp.TempPath = rd.Value.Trim
                  Catch ex As Exception
                  End Try
                Case "BaaNLive"
                  Try
                    rd.Read()
                    tmp.BaaNLive = Convert.ToBoolean(rd.Value.Trim)
                  Catch ex As Exception
                  End Try
                Case "JoomlaLive"
                  Try
                    rd.Read()
                    tmp.JoomlaLive = Convert.ToBoolean(rd.Value.Trim)
                  Catch ex As Exception
                  End Try
                Case "Testing"
                  Try
                    rd.Read()
                    tmp.Testing = Convert.ToBoolean(rd.Value.Trim)
                  Catch ex As Exception
                  End Try
              End Select
            End If
          End While
          rd.Close()
        Catch ex As Exception
        End Try
      End If
      Return tmp
    End Function
    Public Shared Function Serialize(ByVal jpConfig As ConfigFile, ByVal SerializeAt As String) As ConfigFile
      jpConfig.SerializedAt = SerializeAt
      Dim oSrz As XmlSerializer = New XmlSerializer(jpConfig.GetType)
      Dim oSW As IO.StreamWriter = New IO.StreamWriter(SerializeAt)
      oSrz.Serialize(oSW, jpConfig)
      oSW.Close()
      Return jpConfig
    End Function
    Public Shared Function DeSerialize(Optional ByVal jpConfig As ConfigFile = Nothing, Optional ByVal SerializedAt As String = "") As ConfigFile
      Dim FileName As String = ""
      If jpConfig IsNot Nothing Then
        FileName = jpConfig.SerializedAt
      Else
        FileName = SerializedAt
      End If
      If IO.File.Exists(FileName) Then
        jpConfig = New ConfigFile
        Dim oFS As IO.FileStream = New IO.FileStream(FileName, IO.FileMode.Open)
        Dim oSrz As XmlSerializer = New XmlSerializer(jpConfig.GetType)
        jpConfig = CType(oSrz.Deserialize(oFS), ConfigFile)
        oFS.Close()
      End If
      Return jpConfig
    End Function

  End Class

End Namespace

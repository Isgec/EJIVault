Imports System.Reflection
Module modMain
  Public conf As EJI.ConfigFile = Nothing
  Sub Initialize()
    If conf Is Nothing Then
      Dim assemblyPath As String = IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly.Location)
      conf = EJI.ConfigFile.GetFile(assemblyPath & "\ejiConfig.xml")
      conf.StartupPath = assemblyPath
      EJI.DBCommon.BaaNLive = conf.BaaNLive
      EJI.DBCommon.JoomlaLive = conf.JoomlaLive
    End If
  End Sub
End Module

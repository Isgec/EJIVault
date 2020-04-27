Imports System.Reflection
Namespace EJI
  Module ejiMain
    Public conf As ConfigFile = Nothing
    Sub main()
      Dim assemblyPath As String = IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly.Location)
      conf = ConfigFile.GetFile(assemblyPath & "\Settings.xml")
      conf.StartupPath = assemblyPath
      DBCommon.BaaNLive = conf.BaaNLive
      DBCommon.JoomlaLive = conf.JoomlaLive
    End Sub
  End Module
End Namespace

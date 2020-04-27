Namespace EJI
  Public Class ediException
    Inherits Exception
    Public Property ediCode As Integer = 0
    Public Property ediMessage As String = ""
    Sub New(code As Integer, msg As String)
      MyBase.New(msg)
      ediCode = code
      ediMessage = msg
    End Sub
  End Class

End Namespace


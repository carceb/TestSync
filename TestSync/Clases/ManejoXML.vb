Imports System.Xml
Public Class ManejoXML
    Public Sub New()
    End Sub
    Public Function ObtenerValorXML(nombreNodo As String) As String
        Dim objectLibrary As New Library
        Dim resultadoValorNodo As String = ""
        Dim confFilePath As String
        Try
            confFilePath = My.Application.Info.DirectoryPath & "\Configuracion.xml"
            If (IO.File.Exists(confFilePath)) Then
                Dim document As XmlReader = New XmlTextReader(confFilePath)
                While (document.Read())
                    Dim type = document.NodeType
                    If (type = XmlNodeType.Element) Then
                        If (document.Name = nombreNodo) Then
                            resultadoValorNodo = document.ReadInnerXml.ToString()
                        End If
                    End If
                End While
                document.Close()
            End If
            ObtenerValorXML = resultadoValorNodo
        Catch ex As Exception
            objectLibrary.WriteErrorLog(ex.Message)
            ObtenerValorXML = ""
        End Try
    End Function
End Class

Imports System.Xml
Public Class ManejoXML
    Public Sub New()
    End Sub
    Public Function ObtenerValorXML(nombreNodo As String) As String
        Dim resultadoValorNodo As String = ""
        If (IO.File.Exists("Configuracion.xml")) Then
            Dim document As XmlReader = New XmlTextReader("Configuracion.xml")
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
    End Function
End Class

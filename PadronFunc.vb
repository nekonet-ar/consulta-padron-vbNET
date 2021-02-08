Imports System
Imports System.Windows.Forms

'este módulo es importante para poder leer los nodos del xml
Imports System.Xml

'acá está comentado porqué es la importación del módulo de conexión a la base de datos
'Imports FuncDBParamAfip

Public Class PadronFunc

    Public denominacion As String = ""
    Public direccion As String = ""
    Public localidad As String = ""
    Public provincia As String = ""
    Public tipoPers As String = ""
    Public EstadoFiscal As String = ""
    Public ImpIVAID As String = ""
    Public ImpIVADesc As String = ""
    Public IMPGral As String = ""
    Public Actividades As String = ""
    Public ErrorPubl As String = ""
    Public CatMonotrID As String = ""

    Public Padron As Object, ok As Object

    Dim ObjWS As New AutenticarWS.AutenticarWS

    'y lo comentado acá es la declaración de la función que consulta a la base de datos
    'Dim ObjFunDB As New FDBAfip

    Public Function ConsultarPadron(ByVal Cuit As String, ByVal Cert As String, ByVal key As String, ByVal CuitRepres As String)
        Dim m_xmld As XmlDocument
        Dim m_nodelist As XmlNodeList
        Dim m_xmld2 As XmlDocument
        m_xmld = New XmlDocument()
        m_xmld2 = New XmlDocument()
        Dim xmltext As String
        Dim ImpIVAID As String = ""
        Dim TieneNom As Boolean = False
        Dim Terrores As Boolean = False
        Dim Idimp As Integer = 0
        Dim IdAct As Integer = 0
        Dim descImp As String = ""
        Dim descAct As String = ""
        Try
            ObjWS.Autenticar(Cert, key, "ws_sr_padron_a5")
            Dim id_persona As String = Cuit
            Padron = CreateObject("WSSrPadronA5")
            Padron.SetTicketAcceso(ObjWS.ta)
            Padron.token = ObjWS.Token
            Padron.sign = ObjWS.Sign
            Padron.Cuit = CuitRepres
            Padron.Conectar()

            ok = Padron.Consultar(id_persona)
            xmltext = (Padron.xmlresponse)

            'Cargamos el archivo
            m_xmld.LoadXml(xmltext)

            m_nodelist = m_xmld.GetElementsByTagName("idImpuesto")
            For Each item As XmlNode In m_nodelist
                Try
                    ImpIVAID = item.InnerXml
                    If ImpIVAID = "20" Then
                        ImpIVADesc = "MONOTRIBUTO"
                        Exit For
                    ElseIf ImpIVAID = "30" Then
                        ImpIVADesc = "RESP. INSCRIPTO"
                        Exit For
                    ElseIf ImpIVAID = "32" Then
                        ImpIVADesc = "EXENTO"
                        Exit For
                    End If
                Catch ex As Exception

                End Try
            Next
            If ImpIVAID = "" Then
                ImpIVADesc = "CONS. FINAL"
            End If

            If ImpIVADesc <> "CONS. FINAL" Then
                denominacion = Padron.denominacion
                tipoPers = Padron.tipo_persona
                EstadoFiscal = Padron.estado
                direccion = Padron.direccion
                localidad = Padron.localidad
                provincia = Padron.provincia
            End If

            If ImpIVADesc = "MONOTRIBUTO" Or "RESP. INSCRIPTO" Or ImpIVADesc = "EXENTO" Then

                xmltext = (Padron.xmlresponse)
                m_xmld.LoadXml(xmltext)
                m_nodelist = m_xmld.GetElementsByTagName("IdImpuesto")
                For Each item2 As XmlNode In m_nodelist
                    Try
                        Idimp = item2.InnerText
                        'en esta parte obtengo el ID y consultaba a una base de datos externa mía la descripción del mismo (según tabla de afip)
                        'descImp = ObjFunDB.DescripcionImp(Idimp)
                        IMPGral = Idimp & "-" & descImp & ";"
                    Catch ex As Exception

                    End Try
                Next
                xmltext = (Padron.xmlresponse)
                m_xmld.LoadXml(xmltext)
                m_nodelist = m_xmld.GetElementsByTagName("IdActividad")
                For Each item2 As XmlNode In m_nodelist
                    Try
                        IdAct = item2.InnerText
                        'en esta parte obtengo el ID y consultaba a una base de datos externa mía la descripción del mismo (según tabla de afip)
                        'descAct = ObjFunDB.DescripcionAct(Idimp)
                        Actividades = IdAct & "-" & descAct & ";"
                    Catch ex As Exception

                    End Try
                Next

            End If
            If ImpIVADesc = "MONOTRIBUTO" Then
                xmltext = (Padron.xmlresponse)
                m_xmld.LoadXml(xmltext)
                m_nodelist = m_xmld.GetElementsByTagName("idCategoria")
                For Each item2 As XmlNode In m_nodelist
                    CatMonotrID = item2.InnerText
                    Exit For
                Next
            End If
        Catch ex As Exception
            xmltext = (Padron.xmlresponse)
            m_xmld.LoadXml(xmltext)
            m_nodelist = m_xmld.GetElementsByTagName("error")
            For Each item As XmlNode In m_nodelist
                ErrorPubl = item.InnerText
                Exit For
            Next
            If ErrorPubl = "" Then
                m_nodelist = m_xmld.GetElementsByTagName("faultstring")
                For Each item As XmlNode In m_nodelist
                    ErrorPubl = item.InnerText
                    Exit For
                Next
            End If
            ErrorPubl = ""
            Terrores = True
        End Try
        If Terrores = True Then
            Try
                denominacion = Padron.denominacion
            Catch ex As Exception

            End Try
            If denominacion <> "" Then
                ImpIVADesc = "CONS. FINAL"
            End If
        End If

    End Function
End Class

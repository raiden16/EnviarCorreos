Imports Sap.Data.Hana
Imports SAPbobsCOM
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Net.Mail
Imports System.IO
Imports System.Xml

Module Module1

    Public SBOCompany As SAPbobsCOM.Company
    Dim pdf, xml, pdfSAP, xmlSAP As String

    Sub Main()

        Conectar()
        Documentos()
        Desconectar()

    End Sub

    Public Function Conectar()

        Try

            SBOCompany = New SAPbobsCOM.Company

            SBOCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB
            SBOCompany.Server = My.Settings.Server
            SBOCompany.LicenseServer = My.Settings.LicenseServer
            SBOCompany.DbUserName = My.Settings.DbUserName
            SBOCompany.DbPassword = My.Settings.DbPassword

            SBOCompany.CompanyDB = My.Settings.CompanyDB

            SBOCompany.UserName = My.Settings.UserName
            SBOCompany.Password = My.Settings.Password

            SBOCompany.Connect()

        Catch ex As Exception

            Dim stError As String
            stError = "Error al tratar de hacer conexión con SAP B1. " & ex.Message
            Setlog(stError, " ", " ", " ", " ", " ")
            'MsgBox(stError)

        End Try

    End Function

    Public Function Documentos()

        'MsgBox("Se conecto con exito")
        Dim oRecSettxb As SAPbobsCOM.Recordset
        Dim stQuerytxb As String
        Dim DocEntry, DocNum, ReportId, CardCode, CardName, DocDate, EmailC, Tipo As String

        Try

            oRecSettxb = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            stQuerytxb = "call Consulta_EnvioCorreo()"
            oRecSettxb.DoQuery(stQuerytxb)

            If oRecSettxb.RecordCount > 0 Then

                oRecSettxb.MoveFirst()

                For cont As Integer = 0 To oRecSettxb.RecordCount - 1

                    DocEntry = oRecSettxb.Fields.Item("DocEntry").Value
                    DocNum = oRecSettxb.Fields.Item("DocNum").Value
                    ReportId = oRecSettxb.Fields.Item("ReportID").Value
                    CardCode = oRecSettxb.Fields.Item("CardCode").Value
                    CardName = oRecSettxb.Fields.Item("CardName").Value
                    DocDate = oRecSettxb.Fields.Item("CreateDate").Value
                    EmailC = oRecSettxb.Fields.Item("E_Mail").Value
                    Tipo = oRecSettxb.Fields.Item("Tipo").Value

                    ValidarDoc(DocEntry, ReportId, Tipo, DocDate, CardCode, DocNum, EmailC)

                    oRecSettxb.MoveNext()

                Next

            End If

        Catch ex As Exception

            Dim stError As String
            stError = "Error en Facturas. " & ex.Message
            Setlog(stError, DocNum, EmailC, " ", CardCode, Tipo)
            MsgBox(stError)

        End Try

    End Function


    Public Function ValidarDoc(ByVal DocEntry As String, ByVal ReportID As String, ByVal Tipo As String, ByVal DocDate As Date, ByVal CardCode As String, ByVal DocNum As String, ByVal EmailC As String)

        'MsgBox("Exportar Documento Exitoso")
        Dim Ruta, RutaSAP As String
        pdf = Nothing
        xml = Nothing
        pdfSAP = Nothing
        xmlSAP = Nothing

        Try

            If Tipo = "FC" Then

                Ruta = My.Settings.Ruta & "\" & DocDate.ToString("yyyy-MM") & "\" & CardCode & "\IN"
                RutaSAP = My.Settings.RutaSAP & "\" & DocDate.ToString("yyyy-MM") & "\" & CardCode & "\IN"

            ElseIf Tipo = "NC" Then

                Ruta = My.Settings.Ruta & "\" & DocDate.ToString("yyyy-MM") & "\" & CardCode & "\CM"
                RutaSAP = My.Settings.RutaSAP & "\" & DocDate.ToString("yyyy-MM") & "\" & CardCode & "\CM"

            ElseIf Tipo = "PR" Then

                Ruta = My.Settings.Ruta & "\" & DocDate.ToString("yyyy-MM") & "\" & CardCode & "\RC"
                RutaSAP = My.Settings.RutaSAP & "\" & DocDate.ToString("yyyy-MM") & "\" & CardCode & "\RC"

            End If

            Dim dir As New System.IO.DirectoryInfo(Ruta)

            Dim fileList = dir.GetFiles("*.pdf", System.IO.SearchOption.TopDirectoryOnly)

            Dim FileQuery = From file In fileList
                            Where file.Extension = ".pdf" And file.Name.Trim.ToString.EndsWith(ReportID & ".pdf") And file.Name.Trim.ToString.StartsWith(ReportID & ".pdf")
                            Order By file.CreationTime
                            Select file

            pdf = Ruta & "\" & ReportID & ".pdf"
            pdfSAP = RutaSAP & "\" & ReportID & ".pdf"

            Dim fileList1 = dir.GetFiles("*.xml", System.IO.SearchOption.TopDirectoryOnly)

            Dim fileQuery1 = From file In fileList1
                             Where file.Extension = ".xml" And file.Name.Trim.ToString.EndsWith(ReportID & ".xml") And file.Name.Trim.ToString.StartsWith(ReportID & ".xml")
                             Order By file.CreationTime
                             Select file

            xml = Ruta & "\" & ReportID & ".xml"
            xmlSAP = RutaSAP & "\" & ReportID & ".xml"

            If FileQuery.Count > 0 And fileQuery1.Count > 0 Then

                If EmailC <> "" Then

                    UpdatePDFXML(DocNum, pdfSAP, xmlSAP, Tipo)
                    EnviarCorreo(DocNum, EmailC, pdf, xml, Tipo, pdfSAP, xmlSAP, CardCode)

                Else

                    Dim stError As String
                    stError = "El socio de negocios no tiene asignado un correo electronico"
                    Setlog(stError, DocNum, EmailC, " ", CardCode, Tipo)

                End If

            ElseIf FileQuery.Count = 0 And fileQuery1.Count > 0 Then

                ExportarPDF(DocEntry, ReportID, Tipo, DocDate, CardCode, DocNum, EmailC)

            ElseIf fileQuery1.Count = 0 Then

                Dim stError As String
                If Tipo = "FC" Then
                    stError = "A la factura no se le ha creado un xml"
                    Setlog(stError, DocNum, EmailC, " ", CardCode, Tipo)
                ElseIf Tipo = "NC" Then
                    stError = "A la nota de credito no se le ha creado un xml"
                    Setlog(stError, DocNum, EmailC, " ", CardCode, Tipo)
                ElseIf Tipo = "PR" Then
                    stError = "Al pago recibido no se le ha creado un xml"
                    Setlog(stError, DocNum, EmailC, " ", CardCode, Tipo)
                End If

            End If

        Catch ex As Exception

            Dim stError As String
            stError = "Error en ValidarDoc. " & ex.Message
            Setlog(stError, DocNum, EmailC, " ", CardCode, Tipo)
            'MsgBox(stError)

        End Try

    End Function


    Public Function ExportarPDF(ByVal DocEntry As String, ByVal ReportId As String, ByVal Tipo As String, ByVal DocDate As Date, ByVal CardCode As String, ByVal DocNum As String, ByVal EmailC As String)

        'MsgBox("Consulta de Documentos exitosa")
        Dim reportDocument As ReportDocument
        Dim diskFileDestinationOption As DiskFileDestinationOptions


        Try

            reportDocument = New ReportDocument

            If Tipo = "FC" Then

                reportDocument.Load("C:\TareasProgramadas\EnvioCorreos\FC2.rpt")
                'MsgBox("Carga de Documento Exitosa")

            ElseIf Tipo = "NC" Then

                reportDocument.Load("C:\TareasProgramadas\EnvioCorreos\NC2.rpt")

            ElseIf Tipo = "PR" Then

                reportDocument.Load("C:\TareasProgramadas\EnvioCorreos\PR2.rpt")

            End If


            Dim count As Integer = reportDocument.DataSourceConnections.Count
            reportDocument.DataSourceConnections(0).SetLogon(My.Settings.DbUserName, My.Settings.DbPassword)

            reportDocument.SetParameterValue(0, DocEntry)

            reportDocument.ExportOptions.ExportDestinationType = ExportDestinationType.DiskFile
            reportDocument.ExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat
            diskFileDestinationOption = New DiskFileDestinationOptions

            If Tipo = "FC" Then

                diskFileDestinationOption.DiskFileName = My.Settings.Ruta & "\" & DocDate.ToString("yyyy-MM") & "\" & CardCode & "\IN\" & ReportId & ".pdf"
                'MsgBox("Asignacion de direccion Exitosa")

            ElseIf Tipo = "NC" Then

                diskFileDestinationOption.DiskFileName = My.Settings.Ruta & "\" & DocDate.ToString("yyyy-MM") & "\" & CardCode & "\CM\" & ReportId & ".pdf"

            ElseIf Tipo = "PR" Then

                diskFileDestinationOption.DiskFileName = My.Settings.Ruta & "\" & DocDate.ToString("yyyy-MM") & "\" & CardCode & "\RC\" & ReportId & ".pdf"

            End If

            reportDocument.ExportOptions.ExportDestinationOptions = diskFileDestinationOption
            reportDocument.ExportOptions.ExportFormatOptions = New PdfRtfWordFormatOptions

            reportDocument.Export()
            'MsgBox("Exportacion de Documento Exitosa")
            reportDocument.Close()
            reportDocument.Dispose()
            GC.SuppressFinalize(reportDocument)

            UpdatePDFXML(DocNum, pdfSAP, xmlSAP, Tipo)

            If EmailC <> "" Then

                EnviarCorreo(DocNum, EmailC, pdf, xml, Tipo, pdfSAP, xmlSAP, CardCode)

            Else

                Dim stError As String
                stError = "El socio de negocios no tiene asignado un correo electronico"
                Setlog(stError, DocNum, EmailC, " ", CardCode, Tipo)

            End If

        Catch ex As Exception

            Dim stError As String
            stError = "Error en ExportarPDF. " & ex.Message
            Setlog(stError, DocNum, " ", " ", CardCode, Tipo)
            'MsgBox(stError)

        End Try

    End Function


    Public Function UpdatePDFXML(ByVal DocNum As String, ByVal pdfSAP As String, ByVal xmlSAP As String, ByVal Tipo As String)

        Dim oRecSettxb1, oRecSettxb2 As SAPbobsCOM.Recordset
        Dim stQuerytxb1, stQuerytxb2 As String

        Try

            If Tipo = "FC" Then

                oRecSettxb1 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                stQuerytxb1 = "Update OINV set ""U_XML""='" & xmlSAP & "' where ""DocNum""=" & DocNum
                oRecSettxb1.DoQuery(stQuerytxb1)

                oRecSettxb2 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                stQuerytxb2 = "Update OINV set ""U_PDF""='" & pdfSAP & "' where ""DocNum""=" & DocNum
                oRecSettxb2.DoQuery(stQuerytxb2)

            ElseIf Tipo = "NC" Then

                oRecSettxb1 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                stQuerytxb1 = "Update ORIN set ""U_XML""='" & xmlSAP & "' where ""DocNum""=" & DocNum
                oRecSettxb1.DoQuery(stQuerytxb1)

                oRecSettxb2 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                stQuerytxb2 = "Update ORIN set ""U_PDF""='" & pdfSAP & "' where ""DocNum""=" & DocNum
                oRecSettxb2.DoQuery(stQuerytxb2)

            ElseIf Tipo = "PR" Then

                oRecSettxb1 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                stQuerytxb1 = "Update ORCT set ""U_XML""='" & xmlSAP & "' where ""DocNum""=" & DocNum
                oRecSettxb1.DoQuery(stQuerytxb1)

                oRecSettxb2 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                stQuerytxb2 = "Update ORCT set ""U_PDF""='" & pdfSAP & "' where ""DocNum""=" & DocNum
                oRecSettxb2.DoQuery(stQuerytxb2)

            End If

        Catch ex As Exception

            Dim stError As String
            stError = "Error en UpdatePDFXML. " & ex.Message
            Setlog(stError, DocNum, " ", " ", "", Tipo)
            'MsgBox(stError)

        End Try

    End Function


    Public Function EnviarCorreo(ByVal DocNum As String, ByVal EmailC As String, ByVal pdf As String, ByVal xml As String, ByVal Tipo As String, ByVal pdfSAP As String, ByVal xmlSAP As String, ByVal CardCode As String)

        'MsgBox("Validacion de Documentos exitosa")
        Dim message As New MailMessage
        Dim smtp As New SmtpClient
        Dim oRecSettxb As SAPbobsCOM.Recordset
        Dim stQuerytxb As String
        Dim EmailU, Pass, EmailCC, Subject, Body, smtpService, Puerto, SegSSL As String

        Try

            oRecSettxb = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            stQuerytxb = "Select ""U_Email"",""U_Password"",""U_EmailCC"",""U_Subject"",""U_Body"",""U_SMTP"",""U_Puerto"",""U_SeguridadSSL"" from ""@CORREOTEKNO"" where ""Name""='Automático'"
            oRecSettxb.DoQuery(stQuerytxb)

            If oRecSettxb.RecordCount > 0 Then

                oRecSettxb.MoveFirst()

                EmailU = oRecSettxb.Fields.Item("U_Email").Value
                Pass = oRecSettxb.Fields.Item("U_Password").Value
                EmailCC = oRecSettxb.Fields.Item("U_EmailCC").Value

                Subject = oRecSettxb.Fields.Item("U_Subject").Value
                Body = oRecSettxb.Fields.Item("U_Body").Value
                smtpService = oRecSettxb.Fields.Item("U_SMTP").Value
                Puerto = oRecSettxb.Fields.Item("U_Puerto").Value
                SegSSL = oRecSettxb.Fields.Item("U_SeguridadSSL").Value

                'Limpiamos correo destinatario, correo copia y archivos adjuntos
                message.To.Clear()
                message.CC.Clear()
                message.Attachments.Clear()

                'Llenamos encabezado de correo
                message.From = New MailAddress(EmailU)
                EmailC = ArreglarTexto(EmailC, ";", ",")
                message.To.Add(EmailC)
                If EmailCC.Count > 0 Then
                    message.CC.Add(EmailCC)
                End If
                If Tipo = "FC" Then
                    message.Subject = Subject & " Factura " & DocNum
                ElseIf Tipo = "NC" Then
                    message.Subject = Subject & " Nota de Crédito " & DocNum
                ElseIf Tipo = "PR" Then
                    message.Subject = Subject & " Pago Recibido " & DocNum
                End If

                'Llenamos el cuerpo del correo y prioridad
                message.Body = Body
                message.Priority = MailPriority.Normal

                'Adjuntamos archivos xml y pdf
                Dim attxml As New Net.Mail.Attachment(xml)
                message.Attachments.Add(attxml)

                Dim attpdf As New Net.Mail.Attachment(pdf)
                message.Attachments.Add(attpdf)

                'Llenamos datos de smtp
                smtp.Host = smtpService
                smtp.Credentials = New Net.NetworkCredential(EmailU, Pass)
                smtp.Port = Puerto
                smtp.EnableSsl = SegSSL

                'Enviamos Correo
                smtp.Send(message)

                UpdateCorreoEnviado(DocNum, Tipo)

            End If

        Catch ex As Exception

            Dim stError As String
            stError = "Error en EnviarCorreo. " & ex.Message
            Setlog(stError, DocNum, EmailC, EmailU, CardCode, Tipo)
            'MsgBox(stError)

        End Try

    End Function


    Public Function UpdateCorreoEnviado(ByVal DocNum As String, ByVal Tipo As String)

        Dim oRecSettxb1 As SAPbobsCOM.Recordset
        Dim stQuerytxb1 As String

        Try

            oRecSettxb1 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            stQuerytxb1 = "Update OINV set ""U_TekEnviado""='Y' where ""DocNum""=" & DocNum
            oRecSettxb1.DoQuery(stQuerytxb1)

            If Tipo = "FC" Then

                oRecSettxb1 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                stQuerytxb1 = "Update OINV set ""U_TekEnviado""='Y' where ""DocNum""=" & DocNum
                oRecSettxb1.DoQuery(stQuerytxb1)

            ElseIf Tipo = "NC" Then

                oRecSettxb1 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                stQuerytxb1 = "Update ORIN set ""U_TekEnviado""='Y' where ""DocNum""=" & DocNum
                oRecSettxb1.DoQuery(stQuerytxb1)

            End If

        Catch ex As Exception

            Dim stError As String
            stError = "Error en UpdatePDFXML. " & ex.Message
            Setlog(stError, DocNum, " ", " ", "", Tipo)
            'MsgBox(stError)

        End Try

    End Function


    Public Function Setlog(ByVal stError As String, ByVal DocNum As String, ByVal EmailC As String, ByVal EmailU As String, ByVal CardCode As String, ByVal Tipo As String)

        Dim oRecSettxb As SAPbobsCOM.Recordset
        Dim stQuerytxb As String

        Try

            stError = ArreglarTexto(stError, "'", " ")
            oRecSettxb = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            stQuerytxb = "insert into ""LOG_ENVIAREMAIL"" values ('" & Tipo & "','" & DocNum & "','" & EmailC & "','" & EmailU & "','" & CardCode & "','" & stError & "',current_date)"
            oRecSettxb.DoQuery(stQuerytxb)

        Catch ex As Exception

            'MsgBox(stError)

        End Try

    End Function

    Public Function ArreglarTexto(ByVal TextoOriginal As String, ByVal QuitarCaracter As String, ByVal PonerCaracter As String)

        TextoOriginal = TextoOriginal.Replace(QuitarCaracter, PonerCaracter)
        Return TextoOriginal

    End Function

    Public Function Desconectar()

        Try

            SBOCompany.Disconnect()

        Catch ex As Exception

            Dim stError As String
            stError = "Error al tratar de hacer conexión con SAP B1. " & ex.Message
            Setlog(stError, " ", " ", " ", " ", " ")
            'MsgBox(stError)

        End Try

    End Function

End Module

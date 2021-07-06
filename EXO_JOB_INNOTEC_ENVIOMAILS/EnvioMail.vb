Imports System.Data.SqlClient
Imports System.IO
Imports CrystalDecisions.CrystalReports.Engine
Imports EXO_Log
Imports SAPbobsCOM

Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports

Public Class EnvioMail

    Public Shared Sub Envio()

        Dim oCompany As SAPbobsCOM.Company = Nothing
        Dim oDBSAP As SqlConnection = Nothing
        Dim dtDocumentos As DataTable = New System.Data.DataTable()
        Dim olog As EXO_Log.EXO_Log = Nothing

        Try

            Dim fichlog As String = Conexiones.GetXMLValue("Ficheros", "Log") & "Log_" & Format(Now.Year, "0000") & Format(Now.Month, "00") & Format(Now.Day, "00") & ".txt"
            '_" & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00") & ".txt"
            olog = New EXO_Log.EXO_Log(fichlog, 50, EXO_Log.EXO_Log.GestionFichero.continuo)


            Conexiones.Connect_SQLServer(oDBSAP, "SQLSAP")

            'lanzo la query para ver si hay algo pendiente. De ser así conecto ocompany y paso dt a la funcion
            Dim sql As String = "select * from " + Conexiones._sSchema + ".dbo.[@EXO_SATIMP] WHERE isnull(U_EXO_ENV,'N')='N'"

            Conexiones.FillDtDB(oDBSAP, dtDocumentos, sql)
            olog.escribeMensaje("SQL:" + sql)
            If dtDocumentos.Rows.Count > 0 Then

                Conexiones.Connect_Company(oCompany, Conexiones._sSchema, olog)

                For Each row As DataRow In dtDocumentos.Rows

                    GestionarFila(oCompany, row, oDBSAP, olog)
                Next
            End If

        Catch ex As Exception
        Finally
            Conexiones.Disconnect_SQLServer(oDBSAP)
            Conexiones.Disconnect_Company(oCompany)
        End Try


    End Sub

    Private Shared Sub GestionarFila(oCompany As Company, row As DataRow, oDBSAP As SqlConnection, olog As EXO_Log.EXO_Log)

        Dim pdfParte As String = ""
        Dim pdfChekList As String = ""
        Dim pdfAlbaran As String = ""
        Dim pdfFormato As String = ""
        Dim bContinuar As Boolean = True
        Try

            'comprobamos parte trabajo, hacemos pdf y anexamos
            If row.Item("U_EXO_PART").ToString() = "Y" Then
                If GenerarPDF(pdfParte, Conexiones.GetXMLValue("Ficheros", "ParteTrabajo"), Conexiones.GetXMLValue("Ficheros", "Pdfs"), row.Item("U_EXO_AVISO").ToString, row.Item("U_EXO_BD").ToString, "ParteTrabajo", oDBSAP, olog) Then
                Else
                    bContinuar = False
                End If
            End If

            'comprobamos checklist, hacemos pdf y anexamos
            If row.Item("U_EXO_CL").ToString() = "Y" Then
                olog.escribeMensaje("entramos a checklist")
                If GenerarPDF(pdfChekList, Conexiones.GetXMLValue("Ficheros", "CheckList"), Conexiones.GetXMLValue("Ficheros", "Pdfs"), row.Item("U_EXO_AVISO").ToString, row.Item("U_EXO_BD").ToString, "CheckList", oDBSAP, olog) Then
                Else
                    bContinuar = False
                End If

            End If

            'comprobamos formato, hacemos pdf y anexamos
            If row.Item("U_EXO_FORM").ToString() = "Y" Then
                If GenerarPDF(pdfFormato, Conexiones.GetXMLValue("Ficheros", "FormatoC0030"), Conexiones.GetXMLValue("Ficheros", "Pdfs"), row.Item("U_EXO_AVISO").ToString, row.Item("U_EXO_BD").ToString, "FormatoC0030", oDBSAP, olog) Then
                Else
                    bContinuar = False
                End If
            End If

            'comprobamos albaran, hacemos pdf y anexamos
            If row.Item("U_EXO_ALB").ToString() = "Y" Then
                'buscamos el albaran asociado a la llamada
                If GenerarPDF(pdfAlbaran, Conexiones.GetXMLValue("Ficheros", "Albaran"), Conexiones.GetXMLValue("Ficheros", "Pdfs"), row.Item("U_EXO_AVISO").ToString, row.Item("U_EXO_BD").ToString, "Albaran", oDBSAP, olog) Then
                Else
                    bContinuar = False
                End If
            End If

            If bContinuar Then
                If EnviarEmail(row.Item("U_EXO_MAIL").ToString(), row.Item("U_EXO_AVISO").ToString, pdfParte, pdfChekList, pdfFormato, pdfAlbaran, oDBSAP, olog) Then
                    ActualizarRegistroEnvioMails(oCompany, row.Item("Code").ToString)
                End If
            End If


        Catch ex As Exception
            olog.escribeMensaje("error generando documentos. Registro " + row.Item("Code").ToString + " " + ex.Message)
        Finally

        End Try

    End Sub

    Private Shared Function GenerarPDF(ByRef DocumentoPdf As String, ByVal strRutaInforme As String, ByVal sRutaFicheros As String, docentry As String, empresa As String, sTextoTipoDoc As String, oDBSAP As SqlConnection, olog As EXO_Log.EXO_Log) As Boolean

        Dim oCRReport As ReportDocument = Nothing
        'Dim Idx As Integer = 0
        'Dim Idx2 As Integer = 0
        'Dim strCadena As String = ""
        Dim sFilePDF As String = sTextoTipoDoc & "_" & docentry
        Dim strNombrePDF As String = sTextoTipoDoc & "_" & docentry
        Dim Sql As String = ""

        GenerarPDF = False
        ' Dim oCRReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument()

        Try

            'generar el rpt
            'ver donde esta las rutas del addon

            oCRReport = New ReportDocument


            oCRReport.Load(strRutaInforme)

            If sTextoTipoDoc = "FormatoC0030" Then
                'necesito la oins
                Dim ssql As String = " select t2.internalSN " +
                    " from " + empresa + ".dbo.OCLG t0 " +
                    " INNER JOIN " + empresa + ".dbo.OSCL T2 on T0.parentId=T2.callID And T0.parentType=191" +
                    " where  t0.ClgCode='" + docentry + "'"
                Dim InternalSN As String = Conexiones.ExecuteSqlString(oDBSAP, ssql)

                oCRReport.SetParameterValue("nserie", InternalSN)

            Else
                oCRReport.SetParameterValue("DocKey@", docentry)
            End If

            'ALGO

            'PONER USUARIO Y CONTRASEÑA

            Dim conrepor As CrystalDecisions.Shared.DataSourceConnections = oCRReport.DataSourceConnections
            conrepor(0).SetConnection(Conexiones._sServer, empresa, Conexiones._sUserBD, Conexiones._sPassBD)

            For Each subReport As ReportDocument In oCRReport.Subreports
                ' refUI.SBOApp.StatusBar.SetText("preparando subreport logon ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                subReport.SetDatabaseLogon(Conexiones._sUserBD, Conexiones._sPassBD, Conexiones._sServer, empresa)
                ' refUI.SBOApp.StatusBar.SetText("subreport logon ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Next

            sFilePDF = sRutaFicheros & strNombrePDF & ".pdf"
            DocumentoPdf = sFilePDF

            oCRReport.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, sFilePDF)

            olog.escribeMensaje("Pdf creado : " & strNombrePDF, EXO_Log.EXO_Log.Tipo.informacion)

            GenerarPDF = True

        Catch ex As Exception

            olog.escribeMensaje("Crear PDF exception: " + ex.Message)

        Finally
            oCRReport.Close()
            oCRReport.Dispose()
            GC.Collect()
        End Try
    End Function

    Private Shared Function EnviarEmail(dirmail As String, Actividad As String, Parte As String, CheckList As String, FormatoC0030 As String, Albaran As String, oDBSAP As SqlConnection, olog As EXO_Log.EXO_Log) As Boolean

        Dim correo As New System.Net.Mail.MailMessage()
        Dim adjunto As System.Net.Mail.Attachment

        Dim StrFirma As String = ""
        Dim htmbody As New System.Text.StringBuilder()
        Dim cuerpo As String = ""

        Try

            correo.From = New System.Net.Mail.MailAddress(Conexiones.GetXMLValue("Mail", "USUARIOMAIL"), Conexiones.GetXMLValue("Mail", "NOMBRE"))
            correo.To.Add(dirmail)

            If Parte <> "" Then
                adjunto = New System.Net.Mail.Attachment(Parte)
                correo.Attachments.Add(adjunto)
            End If

            If CheckList <> "" Then
                adjunto = New System.Net.Mail.Attachment(CheckList)
                correo.Attachments.Add(adjunto)
            End If

            If FormatoC0030 <> "" Then
                adjunto = New System.Net.Mail.Attachment(FormatoC0030)
                correo.Attachments.Add(adjunto)
            End If

            If Albaran <> "" Then
                adjunto = New System.Net.Mail.Attachment(Albaran)
                correo.Attachments.Add(adjunto)
            End If

            If Conexiones.GetXMLValue("Mail", "HTML") = "Y" Then
                Dim FicheroCab As String = Conexiones.GetXMLValue("Mail", "RutaBody")

                Dim srCAB As StreamReader = New StreamReader(FicheroCab)

                cuerpo = srCAB.ReadToEnd()
            Else
                cuerpo = "Estimado cliente, " + Chr(13) + Chr(13)

                cuerpo = cuerpo + "Le adjuntamos documentación referente al aviso " + Actividad + "." + Chr(13)
                cuerpo = cuerpo + "Para más información contacte con su gestora o en el email XXXXXX.com. " + Chr(13) + Chr(13)
                cuerpo = cuerpo + "Nuestro horario de atención al cliente es de lunes a viernes XXXX0h a XXXX0h." + Chr(13)
                cuerpo = cuerpo + "Este email se ha generado automáticamente, no responda al mismo." + Chr(13) + Chr(13)
                'cuerpo = cuerpo + "Quicesa S.A."

            End If

            correo.Subject = "Asistencia " & Actividad.ToString

            correo.Body = cuerpo
            correo.IsBodyHtml = True
            correo.Priority = System.Net.Mail.MailPriority.Normal

            Dim smtp As New System.Net.Mail.SmtpClient

            smtp.Host = Conexiones.GetXMLValue("Mail", "SMTP")
            smtp.Port = Conexiones.GetXMLValue("Mail", "PUERTO")
            smtp.UseDefaultCredentials = False
            smtp.Credentials = New System.Net.NetworkCredential(Conexiones.GetXMLValue("Mail", "USUARIOMAIL"), Conexiones.GetXMLValue("Mail", "PWD"))
            smtp.EnableSsl = True

            'smtp.Host = "smtp.office365.com"
            'smtp.Port = 587
            'smtp.UseDefaultCredentials = False
            'smtp.Credentials = New System.Net.NetworkCredential("administracion@landesa.com", "KXhF3cPe")
            'smtp.EnableSsl = True


            'smtp.Host = "exch.quicesa.com"
            'smtp.Port = 587
            'Dim credentials As System.Net.NetworkCredential = New System.Net.NetworkCredential("exch.quicesa.com\facturas@quicesa.com", "12345678", "exch.quicesa.com")

            'smtp.Credentials = credentials
            'smtp.UseDefaultCredentials = False
            'smtp.EnableSsl = True

            'smtp.DeliveryMethod = Net.Mail.SmtpDeliveryMethod.Network



            smtp.Send(correo)
            correo.Dispose()

            olog.escribeMensaje("Correo enviado: " & dirmail & " " & Actividad, EXO_Log.EXO_Log.Tipo.informacion)

            Return True

        Catch ex As Exception

            EnviarEmail = False

            olog.escribeMensaje("Error enviando correo: " & dirmail & " " & Actividad & " " & ex.Message, EXO_Log.EXO_Log.Tipo.error)

        End Try

        Return False

    End Function

    Private Shared Sub ActualizarRegistroEnvioMails(oCompany As Company, Code As String)
        Try
            Dim oUserTable As SAPbobsCOM.UserTable

            oUserTable = oCompany.UserTables.Item("EXO_SATIMP")
            oUserTable.GetByKey(Code)
            oUserTable.UserFields.Fields.Item("U_EXO_ENV").Value = "Y"

            If oUserTable.Update() = 0 Then

            Else

            End If
        Catch ex As Exception

        End Try
    End Sub

End Class

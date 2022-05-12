Imports System.IO
Imports System.Xml
Imports System.Data.SqlClient

Public Class Conexiones

#Region "Variables globales"

    Public Shared _sSchema As String = ""
    Public Shared _sUserBD As String = ""
    Public Shared _sPassBD As String = ""
    Public Shared _sServer As String = ""
    Public Shared _sRutaFicheros As String = ""
#End Region

    'prueba
#Region "Connect to Company"

    Public Shared Sub Connect_Company(ByRef oCompany As SAPbobsCOM.Company, ByVal _sEmpresa As String, ByRef olog As EXO_Log.EXO_Log)
        Dim myStream As Stream = Nothing
        Dim Reader As XmlTextReader = Nothing

        Try
            'Conectar DI SAP
            myStream = File.OpenRead(My.Application.Info.DirectoryPath.ToString & "\Connections.xml")
            Reader = New XmlTextReader(myStream)
            myStream = Nothing
            While Reader.Read
                Select Case Reader.NodeType
                    Case XmlNodeType.Element
                        Select Case Reader.Name.ToString.Trim
                            Case "DI"
                                oCompany = New SAPbobsCOM.Company

                                oCompany.language = SAPbobsCOM.BoSuppLangs.ln_Spanish
                                oCompany.Server = Reader.GetAttribute("Server").ToString.Trim
                                oCompany.LicenseServer = Reader.GetAttribute("LicenseServer").ToString.Trim
                                oCompany.UserName = Reader.GetAttribute("UserName").ToString.Trim
                                oCompany.Password = Reader.GetAttribute("Password").ToString.Trim
                                oCompany.UseTrusted = False
                                oCompany.DbPassword = Reader.GetAttribute("DbPassword").ToString.Trim
                                oCompany.DbUserName = Reader.GetAttribute("DbUserName").ToString.Trim
                                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
                                oCompany.CompanyDB = _sEmpresa

                                If oCompany.Connect <> 0 Then
                                    olog.escribeMensaje("Error en la conexión a la compañia:" & oCompany.GetLastErrorDescription.Trim)
                                End If
                        End Select
                End Select
            End While


        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            myStream = Nothing
            Reader.Close()
            Reader = Nothing
        End Try
    End Sub

    Public Shared Sub Disconnect_Company(ByRef oCompany As SAPbobsCOM.Company)
        Try
            If Not oCompany Is Nothing Then
                If oCompany.Connected = True Then
                    oCompany.Disconnect()
                End If
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            If oCompany IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCompany)
            oCompany = Nothing
        End Try
    End Sub

#End Region

#Region "Connect to SQL Server"

    Public Shared Function GetXMLValue(Tipo As String, ByVal Valor As String) As String
        Dim myStream As Stream = Nothing
        Dim Reader As XmlTextReader = Nothing
        Dim res As String = ""
        Try
            'Conectar SQL
            myStream = File.OpenRead(My.Application.Info.DirectoryPath.ToString & "\Connections.xml")
            Reader = New XmlTextReader(myStream)
            myStream = Nothing
            While Reader.Read
                Select Case Reader.NodeType
                    Case XmlNodeType.Element
                        Select Case Reader.Name.ToString.Trim
                            Case Tipo


                                res = Reader.GetAttribute(Valor).ToString.Trim.Trim
                                Return res
                        End Select
                End Select
            End While

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            myStream = Nothing
            Reader.Close()
            Reader = Nothing
        End Try
        Return res
    End Function


    Public Shared Sub Connect_SQLServer(ByRef db As SqlConnection, ByVal sTipoSQL As String)
        Dim myStream As Stream = Nothing
        Dim Reader As XmlTextReader = Nothing

        Try
            'Conectar SQL
            myStream = File.OpenRead(My.Application.Info.DirectoryPath.ToString & "\Connections.xml")
            Reader = New XmlTextReader(myStream)
            myStream = Nothing
            While Reader.Read
                Select Case Reader.NodeType
                    Case XmlNodeType.Element
                        Select Case Reader.Name.ToString.Trim
                            Case sTipoSQL
                                If db Is Nothing OrElse db.State = ConnectionState.Closed Then
                                    db = New SqlConnection
                                    db.ConnectionString = "Database=" & Reader.GetAttribute("Db").ToString.Trim.Trim & ";Data Source=" & Reader.GetAttribute("Server").ToString.Trim & ";User Id=" & Reader.GetAttribute("DbUser").ToString & ";Password=" & Reader.GetAttribute("DbPwd").ToString

                                    db.Open()
                                    'variable a añadir en las sql hana
                                    _sSchema = Reader.GetAttribute("Db").ToString.Trim.Trim
                                    _sPassBD = Reader.GetAttribute("DbPwd").ToString
                                    _sUserBD = Reader.GetAttribute("DbUser").ToString
                                    _sServer = Reader.GetAttribute("Server").ToString.Trim
                                    _sRutaFicheros = Reader.GetAttribute("Ruta").ToString.Trim
                                End If

                        End Select
                End Select
            End While

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            myStream = Nothing
            Reader.Close()
            Reader = Nothing
        End Try
    End Sub

    Public Shared Sub Disconnect_SQLServer(ByRef db As SqlConnection)
        Try
            If Not db Is Nothing AndAlso db.State = ConnectionState.Open Then
                db.Close()
                db.Dispose()
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            db = Nothing
        End Try
    End Sub

    Public Shared Sub FillDtDB(ByRef db As SqlConnection, ByRef dt As System.Data.DataTable, ByVal strConsulta As String)
        Dim cmd As SqlCommand = Nothing
        Dim da As SqlDataAdapter = Nothing

        Try
            cmd = New SqlCommand(strConsulta, db)

            cmd.CommandTimeout = 0
            da = New SqlDataAdapter
            da.SelectCommand = cmd
            da.Fill(dt)

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            If Not cmd Is Nothing Then
                cmd.Dispose()
            End If
            If Not da Is Nothing Then
                da.Dispose()
            End If
        End Try
    End Sub

    Public Shared Sub ExecuteSQLDB(ByRef db As SqlConnection, ByVal sSQL As String)
        Dim cmd As SqlCommand = Nothing

        Try
            cmd = New SqlCommand(sSQL, db)
            cmd.ExecuteNonQuery()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            If Not cmd Is Nothing Then
                cmd.Dispose()
            End If
        End Try
    End Sub

    Public Shared Function ExecuteSqlString(ByRef db As SqlConnection, ByVal sSQL As String) As String
        Dim cmd As SqlCommand = Nothing
        Dim res As Object
        Try
            cmd = New SqlCommand(sSQL, db)
            res = cmd.ExecuteScalar()
            If IsDBNull(res) Then
                res = ""
            End If
        Catch ex As Exception
            Throw ex
        Finally
            If Not cmd Is Nothing Then
                cmd.Dispose()
            End If
        End Try
        Return res
    End Function

    Public Shared Sub ExecuteSQLDB(ByRef db As SqlConnection, ByRef oTransaction As SqlTransaction, ByVal sSQL As String)
        Dim cmd As SqlCommand = Nothing

        Try
            cmd = New SqlCommand(sSQL, db)
            cmd.Transaction = oTransaction
            cmd.ExecuteNonQuery()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            If Not cmd Is Nothing Then
                cmd.Dispose()
            End If
        End Try
    End Sub

    Public Shared Function GetValueDB(ByRef db As SqlConnection, ByRef sTabla As String, ByRef sCampo As String, ByRef sCondicion As String) As String
        Dim dt As System.Data.DataTable = Nothing
        Dim sSQL As String = ""
        Dim cmd As SqlCommand = Nothing
        Dim da As SqlDataAdapter = Nothing

        Try
            If sCondicion = "" Then
                sSQL = "SELECT " & sCampo & " FROM " & sTabla
            Else
                sSQL = "SELECT " & sCampo & " FROM " & sTabla & " WHERE " & sCondicion
            End If

            dt = New System.Data.DataTable("Tabla")

            cmd = New SqlCommand(sSQL, db)
            cmd.CommandTimeout = 0

            da = New SqlDataAdapter

            da.SelectCommand = cmd
            da.Fill(dt)

            If dt.Rows.Count <= 0 Then
                Return ""
            Else
                If Not IsDBNull(dt.Rows.Item(0).Item(0).ToString) Then
                    Return dt.Rows.Item(0).Item(0).ToString
                Else
                    Return ""
                End If
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            If Not dt Is Nothing Then
                dt.Dispose()
            End If
            If Not cmd Is Nothing Then
                cmd.Dispose()
            End If
            If Not da Is Nothing Then
                da.Dispose()
            End If
        End Try
    End Function

    Public Shared Function GetValueDB(ByRef db As SqlConnection, ByRef oTransaction As SqlTransaction, ByRef sTabla As String, ByRef sCampo As String, ByRef sCondicion As String) As String
        Dim dt As System.Data.DataTable = Nothing
        Dim sSQL As String = ""
        Dim cmd As SqlCommand = Nothing
        Dim da As SqlDataAdapter = Nothing

        Try
            If sCondicion = "" Then
                sSQL = "SELECT " & sCampo & " FROM " & sTabla
            Else
                sSQL = "SELECT " & sCampo & " FROM " & sTabla & " WHERE " & sCondicion
            End If

            dt = New System.Data.DataTable("Tabla")

            cmd = New SqlCommand(sSQL, db)
            cmd.Transaction = oTransaction
            cmd.CommandTimeout = 0

            da = New SqlDataAdapter

            da.SelectCommand = cmd
            da.Fill(dt)

            If dt.Rows.Count <= 0 Then
                Return ""
            Else
                If Not IsDBNull(dt.Rows.Item(0).Item(0).ToString) Then
                    Return dt.Rows.Item(0).Item(0).ToString
                Else
                    Return ""
                End If
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            If Not dt Is Nothing Then
                dt.Dispose()
            End If
            If Not cmd Is Nothing Then
                cmd.Dispose()
            End If
            If Not da Is Nothing Then
                da.Dispose()
            End If
        End Try
    End Function

    Public Shared Function Datos_FTP(ByVal sTipo As String, ByVal sDato As String) As String
        Dim myStream As Stream = Nothing
        Dim Reader As XmlTextReader = Nothing

        Try
            myStream = File.OpenRead(My.Application.Info.DirectoryPath.ToString & "\Connections.xml")
            Reader = New XmlTextReader(myStream)
            myStream = Nothing
            While Reader.Read
                Select Case Reader.NodeType
                    Case XmlNodeType.Element
                        Select Case Reader.Name.ToString.Trim
                            Case sTipo
                                Datos_FTP = Reader.GetAttribute(sDato).ToString.Trim
                        End Select
                End Select
            End While

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            myStream = Nothing
            Reader.Close()
            Reader = Nothing
        End Try
    End Function

#End Region

End Class
Imports ExportDStoExcel

Namespace DataExportWin
    Public Class clnDataExport
        'Public Shared Sub Exporta(ByVal ds As DataSet, ByVal strNomeArquivo As String, ByVal intColunaInicial As Integer)
        '    ExportDS.ExportToExcel(ds, strNomeArquivo, intColunaInicial)
        'End Sub
        'Public Shared Sub Exporta(ByVal dt As DataTable, ByVal strNomeArquivo As String, ByVal intColunaInicial As Integer)
        '    ExportDS.ExportToExcel(dt, strNomeArquivo, intColunaInicial)
        'End Sub
        Public Shared Sub Exporta(ByVal source As System.Data.DataSet, ByVal fileName As String, ByVal columnInitial As Integer)

            Try
                System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

                Dim excelDoc As System.IO.StreamWriter
                excelDoc = New System.IO.StreamWriter(fileName)

                Dim startExcelXML As String = "<xml version>" & System.Environment.NewLine & "<Workbook xmlns=""urn:schemas-microsoft-com:office:spreadsheet""" & System.Environment.NewLine & " xmlns:o=""urn:schemas-microsoft-com:office:office""" & System.Environment.NewLine & " xmlns:x=""urn:schemas-    microsoft-com:office:excel""" & System.Environment.NewLine & " xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"">" & System.Environment.NewLine & " <Styles>" & System.Environment.NewLine & " <Style ss:ID=""Default"" ss:Name=""Normal"">" & System.Environment.NewLine & " <Alignment ss:Vertical=""Bottom""/>" & System.Environment.NewLine & " <Borders/>" & System.Environment.NewLine & " <Font/>" & System.Environment.NewLine & " <Interior/>" & System.Environment.NewLine & " <NumberFormat/>" & System.Environment.NewLine & " <Protection/>" & System.Environment.NewLine & " </Style>" & System.Environment.NewLine & " <Style ss:ID=""BoldColumn"">" & System.Environment.NewLine & " <Font x:Family=""Swiss"" ss:Bold=""1""/>" & System.Environment.NewLine & " </Style>" & System.Environment.NewLine & " <Style     ss:ID=""StringLiteral"">" & System.Environment.NewLine & " <NumberFormat ss:Format=""@""/>" & System.Environment.NewLine & " </Style>" & System.Environment.NewLine & " <Style ss:ID=""Decimal"">" & System.Environment.NewLine & " <NumberFormat ss:Format=""0.0000""/>" & System.Environment.NewLine & " </Style>" & System.Environment.NewLine & " <Style ss:ID=""Integer"">" & System.Environment.NewLine & " <NumberFormat ss:Format=""0""/>" & System.Environment.NewLine & " </Style>" & System.Environment.NewLine & " <Style ss:ID=""DateLiteral"">" & System.Environment.NewLine & " <NumberFormat ss:Format=""mm/dd/yyyy;@""/>" & System.Environment.NewLine & " </Style>" & System.Environment.NewLine & " </Styles>" & System.Environment.NewLine
                Const endExcelXML As String = "</Workbook>"

                Dim rowCount As Integer = 0
                Dim sheetCount As Integer = 1

                excelDoc.Write(startExcelXML)
                excelDoc.Write("<Worksheet ss:Name=""Sheet" + sheetCount.ToString() + """>")
                excelDoc.Write("<Table>")
                excelDoc.Write("<Row>")

                For x As Integer = columnInitial To source.Tables(0).Columns.Count - 1
                    excelDoc.Write("<Cell ss:StyleID=""BoldColumn""><Data ss:Type=""String"">")
                    excelDoc.Write(source.Tables(0).Columns(x).ColumnName)
                    excelDoc.Write("</Data></Cell>")
                Next
                excelDoc.Write("</Row>")

                For Each x As System.Data.DataRow In source.Tables(0).Rows
                    rowCount += 1
                    If rowCount = 64000 Then 'if the number of rows is > 64000 create a new page to continue output
                        rowCount = 0
                        sheetCount += 1
                        excelDoc.Write("</Table>")
                        excelDoc.Write(" </Worksheet>")
                        excelDoc.Write("<Worksheet ss:Name=""Sheet" & sheetCount & """>")
                        excelDoc.Write("<Table>")
                    End If

                    excelDoc.Write("<Row>") 'ID=" + rowCount + "

                    For y As Integer = columnInitial To source.Tables(0).Columns.Count - 1
                        Dim rowType As System.Type
                        rowType = x(y).GetType()
                        Select Case rowType.ToString()
                            Case "System.String"
                                Dim XMLstring As String = x(y).ToString()
                                XMLstring = XMLstring.Trim()
                                XMLstring = XMLstring.Replace("&", "&amp;")
                                XMLstring = XMLstring.Replace(">", "&gt;")
                                XMLstring = XMLstring.Replace("<", "&lt;")
                                excelDoc.Write("<Cell ss:StyleID=""StringLiteral""><Data ss:Type=""String"">")
                                excelDoc.Write(XMLstring)
                                excelDoc.Write("</Data></Cell>")
                            Case "System.DateTime"
                                'Excel has a specific Date Format of YYYY-MM-DD followed by the letter 'T' then hh:mm:sss.lll Example 2005-01-31T24:01:21.000
                                'The Following Code puts the date stored in XMLDate to the format above
                                Dim XMLDate As DateTime = Convert.ToDateTime(x(y))
                                Dim XMLDatetoString As System.Text.StringBuilder = New System.Text.StringBuilder("")   'Excel Converted Date

                                XMLDatetoString.Append(XMLDate.Year.ToString())
                                XMLDatetoString.Append("-")
                                If XMLDate.Month < 10 Then
                                    XMLDatetoString.Append("0")
                                    XMLDatetoString.Append(XMLDate.Month.ToString())
                                Else
                                    XMLDatetoString.Append(XMLDate.Month.ToString())
                                End If
                                XMLDatetoString.Append("-")
                                If XMLDate.Day < 10 Then
                                    XMLDatetoString.Append("0")
                                    XMLDatetoString.Append(XMLDate.Day.ToString())
                                Else
                                    XMLDatetoString.Append(XMLDate.Day.ToString())
                                End If
                                XMLDatetoString.Append("T")
                                If XMLDate.Hour < 10 Then
                                    XMLDatetoString.Append("0")
                                    XMLDatetoString.Append(XMLDate.Hour.ToString())
                                Else
                                    XMLDatetoString.Append(XMLDate.Hour.ToString())
                                End If
                                XMLDatetoString.Append(":")
                                If XMLDate.Minute < 10 Then
                                    XMLDatetoString.Append("0")
                                    XMLDatetoString.Append(XMLDate.Minute.ToString())
                                Else
                                    XMLDatetoString.Append(XMLDate.Minute.ToString())
                                End If
                                XMLDatetoString.Append(":")
                                If XMLDate.Second < 10 Then
                                    XMLDatetoString.Append("0")
                                    XMLDatetoString.Append(XMLDate.Second.ToString())
                                Else
                                    XMLDatetoString.Append(XMLDate.Second.ToString())
                                End If
                                XMLDatetoString.Append(".000")
                                excelDoc.Write("<Cell ss:StyleID=""DateLiteral""><Data ss:Type=""DateTime"">")
                                excelDoc.Write(XMLDatetoString)
                                excelDoc.Write("</Data></Cell>")
                            Case "System.Boolean"
                                excelDoc.Write("<Cell ss:StyleID=""StringLiteral""><Data ss:Type=""String"">")
                                excelDoc.Write(x(y).ToString())
                                excelDoc.Write("</Data></Cell>")
                            Case "System.Int16"
                                excelDoc.Write("<Cell ss:StyleID=""Integer""><Data ss:Type=""Number"">")
                                excelDoc.Write(x(y).ToString())
                                excelDoc.Write("</Data></Cell>")
                            Case "System.Int32"
                                excelDoc.Write("<Cell ss:StyleID=""Integer""><Data ss:Type=""Number"">")
                                excelDoc.Write(x(y).ToString())
                                excelDoc.Write("</Data></Cell>")
                            Case "System.Int64"
                                excelDoc.Write("<Cell ss:StyleID=""Integer""><Data ss:Type=""Number"">")
                                excelDoc.Write(x(y).ToString())
                                excelDoc.Write("</Data></Cell>")
                            Case "System.Byte"
                                excelDoc.Write("<Cell ss:StyleID=""Integer""><Data ss:Type=""Number"">")
                                excelDoc.Write(x(y).ToString())
                                excelDoc.Write("</Data></Cell>")
                            Case "System.Decimal"
                                excelDoc.Write("<Cell ss:StyleID=""Decimal""><Data ss:Type=""Number"">")
                                excelDoc.Write(x(y).ToString())
                                excelDoc.Write("</Data></Cell>")
                            Case "System.Double"
                                excelDoc.Write("<Cell ss:StyleID=""Double""><Data ss:Type=""Number"">")
                                excelDoc.Write(x(y).ToString())
                                excelDoc.Write("</Data></Cell>")
                            Case "System.DBNull"
                                excelDoc.Write("<Cell ss:StyleID=""StringLiteral""><Data ss:Type=""String"">")
                                excelDoc.Write("")
                                excelDoc.Write("</Data></Cell>")
                            Case Else
                                Throw New Exception(rowType.ToString() + " not handled.")
                        End Select
                    Next
                    excelDoc.Write("</Row>")
                Next

                excelDoc.Write("</Table>")
                excelDoc.Write(" </Worksheet>")
                excelDoc.Write(endExcelXML)
                excelDoc.Close()
            Catch Ex As Exception
                Throw Ex
            Finally
                System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("pt-BR")
            End Try

        End Sub

        Public Shared Sub Exporta(ByVal dt As System.Data.DataTable, ByVal CaminhoArquivo As String, ByVal columnInitial As Integer)

            Try
                System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")

                Dim excelDoc As System.IO.StreamWriter
                excelDoc = New System.IO.StreamWriter(CaminhoArquivo)

                Dim rowCount As Integer = 0
                Dim sheetCount As Integer = 1

                Const endExcelXML As String = "</Workbook>"
                Dim startExcelXML As New System.Text.StringBuilder

                startExcelXML.Append("<xml version>" & System.Environment.NewLine)
                startExcelXML.Append("  <Workbook xmlns=""urn:schemas-microsoft-com:office:spreadsheet"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:x=""urn:schemas-microsoft-com:office:excel"" xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"">" & System.Environment.NewLine)
                startExcelXML.Append("      <Styles>" & System.Environment.NewLine)
                startExcelXML.Append("          <Style ss:ID=""Default"" ss:Name=""Normal"">" & System.Environment.NewLine)
                startExcelXML.Append("              <Alignment ss:Vertical=""Bottom""/>" & System.Environment.NewLine)
                startExcelXML.Append("              <Borders/>" & System.Environment.NewLine)
                startExcelXML.Append("              <Font/>" & System.Environment.NewLine)
                startExcelXML.Append("              <Interior/>" & System.Environment.NewLine)
                startExcelXML.Append("              <NumberFormat/>" & System.Environment.NewLine)
                startExcelXML.Append("              <Protection/>" & System.Environment.NewLine)
                startExcelXML.Append("          </Style>" & System.Environment.NewLine)
                startExcelXML.Append("          <Style ss:ID=""BoldColumn"">" & System.Environment.NewLine)
                startExcelXML.Append("              <Font x:Family=""Swiss"" ss:Bold=""1""/>" & System.Environment.NewLine)
                startExcelXML.Append("          </Style>" & System.Environment.NewLine)
                startExcelXML.Append("          <Style ss:ID=""StringLiteral"">" & System.Environment.NewLine)
                startExcelXML.Append("              <NumberFormat ss:Format=""@""/>" & System.Environment.NewLine)
                startExcelXML.Append("          </Style>" & System.Environment.NewLine)
                startExcelXML.Append("          <Style ss:ID=""Decimal"">" & System.Environment.NewLine)
                startExcelXML.Append("              <NumberFormat ss:Format=""0.0000""/>" & System.Environment.NewLine)
                startExcelXML.Append("          </Style>" & System.Environment.NewLine)
                startExcelXML.Append("          <Style ss:ID=""Double"">" & System.Environment.NewLine)
                startExcelXML.Append("              <NumberFormat ss:Format=""0.0000""/>" & System.Environment.NewLine)
                startExcelXML.Append("          </Style>" & System.Environment.NewLine)
                startExcelXML.Append("          <Style ss:ID=""Integer"">" & System.Environment.NewLine)
                startExcelXML.Append("              <NumberFormat ss:Format=""0""/>" & System.Environment.NewLine)
                startExcelXML.Append("          </Style>" & System.Environment.NewLine)
                startExcelXML.Append("          <Style ss:ID=""DateLiteral"">" & System.Environment.NewLine)
                startExcelXML.Append("              <NumberFormat ss:Format=""mm/dd/yyyy;@""/>" & System.Environment.NewLine)
                startExcelXML.Append("          </Style>" & System.Environment.NewLine)
                startExcelXML.Append("      </Styles>" & System.Environment.NewLine)


                excelDoc.Write(startExcelXML.ToString())
                excelDoc.Write("<Worksheet ss:Name=""Planilha_" + sheetCount.ToString() + """>")
                excelDoc.Write("<Table>")
                excelDoc.Write("<Row>")

                For x As Integer = columnInitial To dt.Columns.Count - 1
                    excelDoc.Write("<Cell ss:StyleID=""BoldColumn""><Data ss:Type=""String"">")
                    excelDoc.Write(dt.Columns(x).ColumnName)
                    excelDoc.Write("</Data></Cell>")
                Next
                excelDoc.Write("</Row>")

                For Each x As System.Data.DataRow In dt.Rows
                    rowCount += 1
                    If rowCount = 64000 Then 'if the number of rows is > 64000 create a new page to continue output
                        rowCount = 0
                        sheetCount += 1
                        excelDoc.Write("</Table>")
                        excelDoc.Write(" </Worksheet>")
                        excelDoc.Write("<Worksheet ss:Name=""Sheet" & sheetCount & """>")
                        excelDoc.Write("<Table>")
                    End If

                    excelDoc.Write("<Row>") 'ID=" + rowCount + "

                    For y As Integer = columnInitial To dt.Columns.Count - 1
                        Dim rowType As System.Type
                        rowType = x(y).GetType()
                        Select Case rowType.ToString()
                            Case "System.String"
                                Dim XMLstring As String = x(y).ToString()
                                XMLstring = XMLstring.Trim()
                                XMLstring = XMLstring.Replace("&", "&amp;")
                                XMLstring = XMLstring.Replace(">", "&gt;")
                                XMLstring = XMLstring.Replace("<", "&lt;")
                                excelDoc.Write("<Cell ss:StyleID=""StringLiteral""><Data ss:Type=""String"">")
                                excelDoc.Write(XMLstring)
                                excelDoc.Write("</Data></Cell>")
                            Case "System.DateTime"
                                'Excel has a specific Date Format of YYYY-MM-DD followed by the letter 'T' then hh:mm:sss.lll Example 2005-01-31T24:01:21.000
                                'The Following Code puts the date stored in XMLDate to the format above
                                Dim XMLDate As DateTime = Convert.ToDateTime(x(y))
                                Dim XMLDatetoString As System.Text.StringBuilder = New System.Text.StringBuilder("")   'Excel Converted Date

                                XMLDatetoString.Append(XMLDate.Year.ToString())
                                XMLDatetoString.Append("-")
                                If XMLDate.Month < 10 Then
                                    XMLDatetoString.Append("0")
                                    XMLDatetoString.Append(XMLDate.Month.ToString())
                                Else
                                    XMLDatetoString.Append(XMLDate.Month.ToString())
                                End If
                                XMLDatetoString.Append("-")
                                If XMLDate.Day < 10 Then
                                    XMLDatetoString.Append("0")
                                    XMLDatetoString.Append(XMLDate.Day.ToString())
                                Else
                                    XMLDatetoString.Append(XMLDate.Day.ToString())
                                End If
                                XMLDatetoString.Append("T")
                                If XMLDate.Hour < 10 Then
                                    XMLDatetoString.Append("0")
                                    XMLDatetoString.Append(XMLDate.Hour.ToString())
                                Else
                                    XMLDatetoString.Append(XMLDate.Hour.ToString())
                                End If
                                XMLDatetoString.Append(":")
                                If XMLDate.Minute < 10 Then
                                    XMLDatetoString.Append("0")
                                    XMLDatetoString.Append(XMLDate.Minute.ToString())
                                Else
                                    XMLDatetoString.Append(XMLDate.Minute.ToString())
                                End If
                                XMLDatetoString.Append(":")
                                If XMLDate.Second < 10 Then
                                    XMLDatetoString.Append("0")
                                    XMLDatetoString.Append(XMLDate.Second.ToString())
                                Else
                                    XMLDatetoString.Append(XMLDate.Second.ToString())
                                End If
                                XMLDatetoString.Append(".000")
                                excelDoc.Write("<Cell ss:StyleID=""DateLiteral""><Data ss:Type=""DateTime"">")
                                excelDoc.Write(XMLDatetoString)
                                excelDoc.Write("</Data></Cell>")
                            Case "System.Boolean"
                                excelDoc.Write("<Cell ss:StyleID=""StringLiteral""><Data ss:Type=""String"">")
                                excelDoc.Write(x(y).ToString())
                                excelDoc.Write("</Data></Cell>")
                            Case "System.Int16"
                                excelDoc.Write("<Cell ss:StyleID=""Integer""><Data ss:Type=""Number"">")
                                excelDoc.Write(x(y).ToString())
                                excelDoc.Write("</Data></Cell>")
                            Case "System.Int32"
                                excelDoc.Write("<Cell ss:StyleID=""Integer""><Data ss:Type=""Number"">")
                                excelDoc.Write(x(y).ToString())
                                excelDoc.Write("</Data></Cell>")
                            Case "System.Int64"
                                excelDoc.Write("<Cell ss:StyleID=""Integer""><Data ss:Type=""Number"">")
                                excelDoc.Write(x(y).ToString())
                                excelDoc.Write("</Data></Cell>")
                            Case "System.Byte"
                                excelDoc.Write("<Cell ss:StyleID=""Integer""><Data ss:Type=""Number"">")
                                excelDoc.Write(x(y).ToString())
                                excelDoc.Write("</Data></Cell>")
                            Case "System.Decimal"
                                excelDoc.Write("<Cell ss:StyleID=""Decimal""><Data ss:Type=""Number"">")
                                excelDoc.Write(x(y).ToString())
                                excelDoc.Write("</Data></Cell>")
                            Case "System.Double"
                                excelDoc.Write("<Cell ss:StyleID=""Double""><Data ss:Type=""Number"">")
                                excelDoc.Write(x(y).ToString())
                                excelDoc.Write("</Data></Cell>")
                            Case "System.DBNull"
                                excelDoc.Write("<Cell ss:StyleID=""StringLiteral""><Data ss:Type=""String"">")
                                excelDoc.Write("")
                                excelDoc.Write("</Data></Cell>")
                            Case Else
                                Throw New Exception(rowType.ToString() + " not handled.")
                        End Select
                    Next
                    excelDoc.Write("</Row>")
                Next

                excelDoc.Write("        </Table>")
                excelDoc.Write("    </Worksheet>")
                excelDoc.Write(endExcelXML)
                excelDoc.Close()

            Catch Ex As Exception
                Throw Ex
            Finally
                System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("pt-BR")
            End Try

        End Sub
    End Class
End Namespace

Namespace DataExportWeb
    Public Class clnDataExport
        Public Overloads Shared Sub Exporta(ByVal Grid As System.Web.UI.WebControls.DataGrid, ByVal strNomeArquivo As String, ByVal rpsHttpResponse As System.web.HttpResponse)
            ExportDS.ExportToExcelWeb(Grid, strNomeArquivo, rpsHttpResponse)
        End Sub
        Public Overloads Shared Sub Exporta(ByVal Grid As System.Web.UI.WebControls.GridView, ByVal strNomeArquivo As String, ByVal rpsHttpResponse As System.Web.HttpResponse)
            Dim oResponse As System.Web.HttpResponse = rpsHttpResponse

            oResponse.Clear()

            oResponse.ContentEncoding = System.Text.Encoding.Default
            oResponse.AddHeader("Content-Disposition", "attachment; filename=" & strNomeArquivo & ".xls")

            oResponse.ContentType = "application/vnd.ms-excel"

            Dim stringWrite As New System.IO.StringWriter

            Dim htmlWrite As New System.Web.UI.HtmlTextWriter(stringWrite)

            Dim dg As New System.Web.UI.WebControls.GridView

            dg.DataSource = Grid.DataSource
            dg.DataBind()

            dg.RenderControl(htmlWrite)

            oResponse.Write(stringWrite.ToString)

            oResponse.End()
        End Sub
        Public Overloads Shared Sub Exporta(ByVal dt As System.Data.DataTable, ByVal strNome As String, ByVal rpsHttpResponse As System.Web.HttpResponse)

            'Dim g As New Guid

            Dim oResponse As System.Web.HttpResponse = rpsHttpResponse

            oResponse.Clear()

            oResponse.ContentEncoding = System.Text.Encoding.Default
            oResponse.AddHeader("Content-Disposition", "attachment; filename=" & strNome & ".xls")

            oResponse.ContentType = "application/vnd.ms-excel"

            Dim stringWrite As New System.IO.StringWriter

            Dim htmlWrite As New System.Web.UI.HtmlTextWriter(stringWrite)

            Dim dg As New System.Web.UI.WebControls.DataGrid

            dg.DataSource = dt

            dg.DataBind()

            dg.RenderControl(htmlWrite)

            oResponse.Write(stringWrite.ToString)

            oResponse.End()

        End Sub
    End Class
End Namespace

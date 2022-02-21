Imports System.Management
Imports System.Text
Imports System.Timers
Imports iTextSharp.text
Imports iTextSharp.text.pdf

Public Class Form1

    Dim _aplicaciones As String = ""
    Dim _version As String = ""
    Dim _creador As String = ""
    Dim _date As String = ""

    Dim _curp As String = "__________________________________"
    Dim _nombre As String = "__________________________________"
    Dim _fecha As String = "_______/_____________/__________"
    Dim _hora As String = "_____:_____"



    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Dim _nameSpace$ = "root\CIMV2"

        Dim wql = "SELECT * FROM WIN32_PROCESSOR"

        Dim _strbuilder As New StringBuilder

        Using _moSearcher As New ManagementObjectSearcher(_nameSpace, wql)

            For Each _mobject As ManagementObject In _moSearcher.Get
                Label7.Text = $"{_mobject("Name")}"
            Next

        End Using


        Dim wql2 = "SELECT * FROM WIN32_NetworkAdapter Where AdapterType='Ethernet 802.3'"

        Dim _strbuilder2 As New StringBuilder

        Using _moSearcher2 As New ManagementObjectSearcher(_nameSpace, wql2)

            For Each _mobject2 As ManagementObject In _moSearcher2.Get
                Label11.Text = $"{_mobject2("MACAddress")}"
            Next

        End Using


        Dim wql4 = "SELECT * FROM WIN32_BIOS"

        Dim _strbuilder4 As New StringBuilder

        Using _moSearcher4 As New ManagementObjectSearcher(_nameSpace, wql4)

            For Each _mobject4 As ManagementObject In _moSearcher4.Get
                Label18.Text = $"{_mobject4("SerialNumber")}"
            Next

        End Using



        Dim memo = My.Computer.Info.TotalPhysicalMemory
        Dim memo2 = memo * (9.31 * (10 ^ -10))

        Label4.Text = My.Computer.Name
        Label5.Text = My.User.Name
        Label6.Text = My.Computer.Info.OSFullName + " Version: " + My.Computer.Info.OSVersion
        Label9.Text = memo2

        _hora = DateTime.Now.ToString("HH:mm:ss")
        _fecha = DateTime.Now.ToString("dddd,dd,MMMM,yyy")

        Label25.Text = _hora
        Label26.Text = _fecha




    End Sub

    Private Sub Label8_Click(sender As Object, e As EventArgs) Handles Label8.Click

    End Sub

    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click

    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs) Handles Label10.Click

    End Sub

    Private Sub Label9_Click(sender As Object, e As EventArgs) Handles Label9.Click

    End Sub

    Private Sub Label12_Click(sender As Object, e As EventArgs) Handles Label12.Click

    End Sub

    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click

    End Sub

    Private Sub Label14_Click(sender As Object, e As EventArgs) Handles Label14.Click

    End Sub

    Private Sub Label13_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Dim _nameSpace$ = "root\CIMV2"

        Dim wql3 = "SELECT * FROM WIN32_Product"

        Dim _strbuilder3 As New StringBuilder


        Using _moSearcher3 As New ManagementObjectSearcher(_nameSpace, wql3)


            For Each _mobject3 As ManagementObject In _moSearcher3.Get
                _aplicaciones = _aplicaciones + $"{_mobject3("Name")}" & vbCrLf
                _version = _version + $"{_mobject3("version")}" & vbCrLf
                _creador = _creador + $"{_mobject3("vendor")}" & vbCrLf
                _date = _date + $"{_mobject3("installdate")}" & vbCrLf



            Next

        End Using


        Label13.Text = _aplicaciones
        Label15.Text = _version
        Label16.Text = _creador
        Label17.Text = _date

        Label13.Visible = True
        Label15.Visible = True
        Label16.Visible = True
        Label17.Visible = True


    End Sub

    Private Sub TableLayoutPanel1_Paint(sender As Object, e As PaintEventArgs)

    End Sub

    Private Sub Label19_Click(sender As Object, e As EventArgs) Handles Label19.Click

    End Sub

    Private Sub Label20_Click(sender As Object, e As EventArgs) Handles Label20.Click

    End Sub

    Private Sub Label21_Click(sender As Object, e As EventArgs) Handles Label21.Click

    End Sub

    Private Sub Label22_Click(sender As Object, e As EventArgs) Handles Label22.Click

    End Sub


    Private Sub Label23_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim SaveFileDialog As New SaveFileDialog
        Dim ruta As String
        With SaveFileDialog
            .Title = "Guardar"
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            .Filter = "Archivo pdf (*.pdf)|*.pdf"
            .FileName = "Constancia de Software de Uso por Equipo del Equipo " + My.Computer.Name
            .OverwritePrompt = True
            .CheckPathExists = True
        End With

        If SaveFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            ruta = SaveFileDialog.FileName
        Else
            ruta = String.Empty
            Exit Sub
        End If

        Try
            Dim document As New iTextSharp.text.Document(PageSize.LETTER)
            document.PageSize.Rotate()

            document.AddAuthor(Label1.ToString)
            document.AddTitle("Crear pdf")


            Dim writer As PdfWriter = PdfWriter.GetInstance(document, New System.IO.FileStream _
            (ruta, System.IO.FileMode.Create))
            writer.ViewerPreferences = PdfWriter.PageLayoutSinglePage

            document.Open()
            Dim cb As PdfContentByte = writer.DirectContent
            Dim bf As BaseFont = BaseFont.CreateFont(BaseFont.COURIER_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED)

            cb.SetFontAndSize(bf, 10)

            Dim pdfTable As New PdfPTable(2)

            Dim PdfTitulo As New PdfPCell(New Phrase("Constancia de Software de Uso por Equipo", New Font(Font.Name = "Tahoma", 20, Font.Bold)))
            PdfTitulo.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER
            PdfTitulo.Colspan = 4
            PdfTitulo.Border = 0
            PdfTitulo.FixedHeight = 60
            pdfTable.AddCell(PdfTitulo)



            Dim Table2 As PdfPTable = New PdfPTable(4)
            Dim Table1 As PdfPTable = New PdfPTable(4)
            Dim Table3 As PdfPTable = New PdfPTable(4)
            Dim Table4 As PdfPTable = New PdfPTable(4)
            Dim Table5 As PdfPTable = New PdfPTable(4)
            Dim Table6 As PdfPTable = New PdfPTable(4)

            Dim width1 As Single() = New Single() {1.0F, 1.0F, 1.0F, 1.0F}
            Dim width2 As Single() = New Single() {2.5F, 0.5F, 1.0F, 0.5F}
            Dim width3 As Single() = New Single() {2.5F, 0.5F, 1.0F, 0.5F}
            Dim width4 As Single() = New Single() {1.8F, 0.2F, 1.8F, 0.2F}
            Dim width5 As Single() = New Single() {1.0F, 1.0F, 1.0F, 1.0F}
            Dim width6 As Single() = New Single() {1.0F, 1.0F, 1.0F, 1.0F}

            Table1.WidthPercentage = 95
            Table2.WidthPercentage = 95
            Table3.WidthPercentage = 95
            Table4.WidthPercentage = 95
            Table5.WidthPercentage = 95
            Table6.WidthPercentage = 95

            Dim CVacio As PdfPCell = New PdfPCell(New Phrase(""))

            Dim Col11 As PdfPCell
            Dim Col12 As PdfPCell
            Dim Col13 As PdfPCell
            Dim Col14 As PdfPCell

            Dim Col1 As PdfPCell
            Dim Col2 As PdfPCell
            Dim Col3 As PdfPCell
            Dim Col4 As PdfPCell

            Dim Col31 As PdfPCell
            Dim Col32 As PdfPCell
            Dim Col33 As PdfPCell
            Dim Col34 As PdfPCell

            Dim Col41 As PdfPCell
            Dim Col42 As PdfPCell
            Dim Col43 As PdfPCell
            Dim Col44 As PdfPCell

            Dim Col51 As PdfPCell
            Dim Col52 As PdfPCell
            Dim Col53 As PdfPCell
            Dim Col54 As PdfPCell

            Dim Col61 As PdfPCell
            Dim Col62 As PdfPCell
            Dim Col63 As PdfPCell
            Dim Col64 As PdfPCell

            CVacio.Border = 0
            Table1.SetWidths(width1)
            Table2.SetWidths(width2)
            Table3.SetWidths(width3)

            Table1.AddCell(CVacio)
            Table1.AddCell(CVacio)
            Table1.AddCell(CVacio)
            Table1.AddCell(CVacio)

            Table2.AddCell(CVacio)
            Table2.AddCell(CVacio)
            Table2.AddCell(CVacio)
            Table2.AddCell(CVacio)

            Table3.AddCell(CVacio)
            Table3.AddCell(CVacio)
            Table3.AddCell(CVacio)
            Table3.AddCell(CVacio)

            Table4.AddCell(CVacio)
            Table4.AddCell(CVacio)
            Table4.AddCell(CVacio)
            Table4.AddCell(CVacio)

            Table5.AddCell(CVacio)
            Table5.AddCell(CVacio)
            Table5.AddCell(CVacio)
            Table5.AddCell(CVacio)

            Table6.AddCell(CVacio)
            Table6.AddCell(CVacio)
            Table6.AddCell(CVacio)
            Table6.AddCell(CVacio)

            Col11 = New PdfPCell(New Phrase("Equipo ID: " & Me.Label18.Text, New Font(bf, 6)))
            Col11.Border = 0
            Table1.AddCell(Col11)

            Col12 = New PdfPCell(New Phrase("", New Font(bf, 6)))
            Col12.Border = 0
            Table1.AddCell(Col12)

            Col13 = New PdfPCell(New Phrase("Direccion MAC: " & Me.Label11.Text, New Font(bf, 6)))
            Col13.Border = 0
            Table1.AddCell(Col13)

            Col14 = New PdfPCell(New Phrase("", New Font(bf, 6)))
            Col14.Border = 0
            Table1.AddCell(Col14)



            Col11 = New PdfPCell(New Phrase("Nombre del Equipo: " & Me.Label4.Text, New Font(bf, 6)))
            Col11.Border = 0
            Table1.AddCell(Col11)

            Col12 = New PdfPCell(New Phrase("", New Font(bf, 6)))
            Col12.Border = 0
            Table1.AddCell(Col12)

            Col13 = New PdfPCell(New Phrase("Nombre del Usuario: " & Me.Label5.Text, New Font(bf, 6)))
            Col13.Border = 0
            Table1.AddCell(Col13)

            Col14 = New PdfPCell(New Phrase("", New Font(bf, 6)))
            Col14.Border = 0
            Table1.AddCell(Col14)



            Col11 = New PdfPCell(New Phrase("Procesador: " & Me.Label7.Text, New Font(bf, 6)))
            Col11.Border = 0
            Table1.AddCell(Col11)

            Col12 = New PdfPCell(New Phrase("", New Font(bf, 6)))
            Col12.Border = 0
            Table1.AddCell(Col12)

            Col13 = New PdfPCell(New Phrase("Sistema Operativo: " & Me.Label6.Text, New Font(bf, 6)))
            Col13.Border = 0
            Table1.AddCell(Col13)

            Col14 = New PdfPCell(New Phrase("", New Font(bf, 6)))
            Col14.Border = 0
            Table1.AddCell(Col14)



            Col11 = New PdfPCell(New Phrase("Memoria Instalada: " & Me.Label9.Text + "GB", New Font(bf, 6)))
            Col11.Border = 0
            Table1.AddCell(Col11)

            Col12 = New PdfPCell(New Phrase("", New Font(bf, 6)))
            Col12.Border = 0
            Table1.AddCell(Col12)

            Col13 = New PdfPCell(New Phrase("", New Font(bf, 6)))
            Col13.Border = 0
            Table1.AddCell(Col13)

            Col14 = New PdfPCell(New Phrase("", New Font(bf, 6)))
            Col14.Border = 0
            Table1.AddCell(Col14)

            Col51 = New PdfPCell(New Phrase("____________________________________________________________________", New Font(bf, 6)))
            Col51.Border = 0
            Table5.AddCell(Col51)

            Col52 = New PdfPCell(New Phrase("____________________________________________________________________", New Font(bf, 6)))
            Col52.Border = 0
            Table5.AddCell(Col52)

            Col53 = New PdfPCell(New Phrase("____________________________________________________________________", New Font(bf, 6)))
            Col53.Border = 0
            Table5.AddCell(Col53)

            Col54 = New PdfPCell(New Phrase("____________________________________________________________________", New Font(bf, 6)))
            Col54.Border = 0
            Table5.AddCell(Col54)

            Col31 = New PdfPCell(New Phrase("Programa", New Font(bf, 5)))
            Col31.Border = 0
            Table3.AddCell(Col31)

            Col32 = New PdfPCell(New Phrase("Version", New Font(bf, 5)))
            Col32.Border = 0
            Table3.AddCell(Col32)

            Col33 = New PdfPCell(New Phrase("Desarrolladora", New Font(bf, 5)))
            Col33.Border = 0
            Table3.AddCell(Col33)

            Col34 = New PdfPCell(New Phrase("Se instalo el", New Font(bf, 5)))
            Col34.Border = 0
            Table3.AddCell(Col34)



            Dim _nameSpace$ = "root\CIMV2"

            Dim wql3 = "SELECT * FROM WIN32_Product"

            Dim _strbuilder3 As New StringBuilder

            Using _moSearcher3 As New ManagementObjectSearcher(_nameSpace, wql3)

                For Each _mobject3 As ManagementObject In _moSearcher3.Get


                    Col1 = New PdfPCell(New Phrase($"{_mobject3("Name")}", New Font(bf, 5)))
                    Col1.Border = 0
                    Table2.AddCell(Col1)

                    Col2 = New PdfPCell(New Phrase($"{_mobject3("version")}", New Font(bf, 5)))
                    Col2.Border = 0
                    Table2.AddCell(Col2)

                    Col3 = New PdfPCell(New Phrase($"{_mobject3("vendor")}", New Font(bf, 5)))
                    Col3.Border = 0
                    Table2.AddCell(Col3)

                    Col4 = New PdfPCell(New Phrase($"{_mobject3("installdate")}", New Font(bf, 5)))
                    Col4.Border = 0
                    Table2.AddCell(Col4)



                Next

            End Using

            Col61 = New PdfPCell(New Phrase("____________________________________________________________________", New Font(bf, 6)))
            Col61.Border = 0
            Table6.AddCell(Col61)

            Col62 = New PdfPCell(New Phrase("____________________________________________________________________", New Font(bf, 6)))
            Col62.Border = 0
            Table6.AddCell(Col62)

            Col63 = New PdfPCell(New Phrase("____________________________________________________________________", New Font(bf, 6)))
            Col63.Border = 0
            Table6.AddCell(Col63)

            Col64 = New PdfPCell(New Phrase("____________________________________________________________________", New Font(bf, 6)))
            Col64.Border = 0
            Table6.AddCell(Col64)



            Col41 = New PdfPCell(New Phrase("CURP: " + _curp & Me.TextBox1.Text, New Font(bf, 6)))
            Col41.Border = 0
            Table4.AddCell(Col41)

            Col42 = New PdfPCell(New Phrase("", New Font(bf, 6)))
            Col42.Border = 0
            Table4.AddCell(Col42)

            Col43 = New PdfPCell(New Phrase("Fecha: " + _fecha, New Font(bf, 6)))
            Col43.Border = 0
            Table4.AddCell(Col43)

            Col44 = New PdfPCell(New Phrase("", New Font(bf, 6)))
            Col44.Border = 0
            Table4.AddCell(Col44)




            Col41 = New PdfPCell(New Phrase("Nombre: " + _nombre & Me.TextBox2.Text, New Font(bf, 6)))
            Col41.Border = 0
            Table4.AddCell(Col41)

            Col42 = New PdfPCell(New Phrase(" ", New Font(bf, 6)))
            Col42.Border = 0
            Table4.AddCell(Col42)

            Col43 = New PdfPCell(New Phrase("Hora: " + _hora, New Font(bf, 6)))
            Col43.Border = 0
            Table4.AddCell(Col43)

            Col44 = New PdfPCell(New Phrase("                                                                                                                                                                                                                                                                                                                                                    __________________________________             Firma", New Font(bf, 6)))
            Col44.Border = 0
            Table4.AddCell(Col44)



            document.Add(pdfTable)
            document.Add(Table1)
            document.Add(Table5)
            document.Add(Table3)
            document.Add(Table2)
            document.Add(Table6)
            document.Add(Table4)





            document.Close()


            MessageBox.Show("Constancia de Software de Uso por Equipo del Equipo " + My.Computer.Name, "Creacion Finalizada", MessageBoxButtons.OK)


        Catch ex As Exception
            MessageBox.Show("Error de la generacion", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub Label26_Click(sender As Object, e As EventArgs) Handles Label26.Click

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

        If TextBox1.Enabled = False Then

            TextBox1.Enabled = True
            TextBox2.Enabled = True
            Label24.Visible = True
            Label25.Visible = True
            Label26.Visible = True
            Label31.Visible = True

            _nombre = ""
            _curp = ""
            _hora = Label25.Text
            _fecha = Label26.Text

        Else
            TextBox1.Enabled = False
            TextBox1.Text = ""
            TextBox2.Enabled = False
            TextBox2.Text = ""
            Label24.Visible = False
            Label25.Visible = False
            Label26.Visible = False
            Label31.Visible = False

            _nombre = "__________________________________"
            _curp = "__________________________________"
            _fecha = "_______/_____________/__________"
            _hora = "_____:_____"

        End If
    End Sub
End Class

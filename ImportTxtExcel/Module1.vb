Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel

Module Module1

    Sub Main()

        Console.WriteLine("=> BOSS....!")
        Console.WriteLine("=> SALCE <=")
        Console.WriteLine("")

        Dim sourcePath = Environment.CurrentDirectory & "\Source"

        If Directory.Exists(sourcePath) Then

            Dim files = Directory.GetFiles(sourcePath, "*TM_GL_SALES_XPT*", SearchOption.TopDirectoryOnly)

            Console.WriteLine("=> AGUARDE...")

            Console.WriteLine("")

            For Each filePath In files

                Dim fileName As String = Path.GetFileNameWithoutExtension(filePath)
                Dim lines As List(Of String) = File.ReadLines(filePath).ToList()

                Console.WriteLine("=> Localizando File: " & fileName)

                Dim xlApp As Excel.Application
                Dim excelBook As Excel.Workbook
                Dim excelWorksheet As Excel.Worksheet

                xlApp = CreateObject("Excel.Application")
                xlApp.Visible = True
                excelBook = xlApp.Workbooks.Add
                excelWorksheet = excelBook.ActiveSheet
                excelWorksheet.Name = fileName.Substring(0, 31)

                Dim contagemCelulas As Integer = 1

                For Each line In lines
                    excelWorksheet.Cells(contagemCelulas, 1).Value = line
                    contagemCelulas += 1
                Next

                xlApp.ActiveWindow.DisplayGridlines = False
                xlApp.ActiveWindow.DisplayFormulas = False
                xlApp.ActiveWindow.DisplayHeadings = True

                xlApp.ScreenUpdating = True
                'xlApp.Visible = False

                Dim NovoCaminho = Environment.CurrentDirectory & "\" & Today.Day & "-" & Today.Month & "-" & Today.Year
                Directory.CreateDirectory(NovoCaminho)

                Dim location As String = NovoCaminho & "\" & fileName & ".xlsx"
                excelWorksheet.SaveAs(location)

                Console.WriteLine("=> Finalizada Converção...")

                '~~> Fecha o arquivo
                excelBook.Close()

                '~~> Fecha Excel Application
                xlApp.Quit()

                '~~> Realiza a limpeza
                releaseObject(xlApp)
                releaseObject(excelBook)
                releaseObject(excelWorksheet)

                FinalizaExcel()

                File.Delete(filePath)

                Console.WriteLine("")

            Next

            If files.Count = 0 Then
                Console.WriteLine("=> Não possivel localizar Ffcheiros!")
                Console.WriteLine("=> Verifique se a pasta Source contem ficheiros")
            End If

            Console.WriteLine("....")
            Console.WriteLine("=> PROCESSO FINALIZADO.")

        Else

            Directory.CreateDirectory(sourcePath)
            Console.WriteLine("=> Criado o Directório 'Source'")

        End If

        Console.ReadLine()

    End Sub

    '~~> Release the objects
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub FinalizaExcel()
        ' The excel is created and opened for insert value. We most close this excel using this system
        Dim pro() As Process = System.Diagnostics.Process.GetProcessesByName("EXCEL")
        Dim ii As Process
        For Each ii In pro
            ii.Kill()
        Next
    End Sub

    Public Enum MSApplications
        WORD
        ACCESS
        EXCEL
    End Enum


End Module

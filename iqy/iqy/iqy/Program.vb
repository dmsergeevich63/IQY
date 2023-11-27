Imports System.IO
Imports Microsoft.Office.Interop.Excel

Module Program
    Sub ConvertFiles()
        Dim Filename, Pathname As String
        Dim ExcelApp As New Application()
        Dim wb As Microsoft.Office.Interop.Excel.Workbook
        Pathname = "C:\iqy\"
        Filename = Dir(Pathname & "*.iqy")
        Do While Filename <> ""
            wb = ExcelApp.Workbooks.Open(Pathname & Filename)
            wb.SaveAs(Pathname & Filename & ".xlsx", FileFormat:=51)
            wb.Close()
            Filename = Dir()
        Loop
        ExcelApp.Quit()
    End Sub

    Sub Main()
        ConvertFiles()
    End Sub
End Module
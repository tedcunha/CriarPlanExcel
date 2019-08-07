Imports NPOI.SS.UserModel

Public Class SheetHelper
    Public Function GetCell(sheet As ISheet, linha As Integer, coluna As Integer) As ICell

        '' Caminho do Exemplo
        ''https://www.devmedia.com.br/excel-x-net-framework-gerando-planilhas-xls-sem-o-uso-de-interop-com/27784

        Dim row As IRow
        Dim indiceLinha As Integer = (linha - 1)

        row = sheet.GetRow(indiceLinha)
        If row Is Nothing Then row = sheet.CreateRow(indiceLinha)

        Dim cell As ICell
        Dim indiceColuna As Integer = (coluna - 1)

        cell = row.GetCell(indiceColuna)
        If cell Is Nothing Then cell = row.CreateCell(indiceColuna)

        Return cell
    End Function
End Class

Imports System.IO
Imports NPOI.HSSF.UserModel
Imports NPOI.SS.UserModel

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim _Produtos As ArrayList = New ArrayList
        Dim _nomeArquivo As String
        Dim file As FileStream = New FileStream("C:\Ricardo\Estudo\vbnet\CriarExcel\CriarPlanExcel\Template_CatalogoProdutos.xls", FileMode.Open, FileAccess.Read)
        Dim _workbookCatalogo As HSSFWorkbook

        _Produtos = ObterCatalogo()
        _workbookCatalogo = New HSSFWorkbook(file)

        Dim sheetCatalogo As ISheet
        sheetCatalogo = _workbookCatalogo.GetSheet("Catalogo")

        Dim numeroProximaLinha As Integer = 3
        Dim I As Integer

        Dim teste As SheetHelper = New SheetHelper

        For I = 0 To (_Produtos.Count() - 1)

            teste.GetCell(sheetCatalogo, numeroProximaLinha, 1).SetCellValue(_Produtos(I).CodigoBarras.ToString())

            numeroProximaLinha += 1
        Next

        file = New FileStream("C:\Ricardo\Estudo\vbnet\CriarExcel\CriarPlanExcel\teste.xls", FileMode.Create)
        _workbookCatalogo.Write(file)
        file.Close()

    End Sub

    Public Function ObterCatalogo() As ArrayList

        ''Iniciando a Coleção dos Objetos do Produto
        Dim _Produtos As clsProduto
        Dim _arrayProduto As ArrayList = New ArrayList

        _Produtos = New clsProduto
        With _Produtos
            .CodigoBarras = "7890000000111"
            .NomeProduto = "Iron Maiden - Powerslave"
            .Categoria = "CDs"
            .DataCadastro = New DateTime(2012, 9, 29)
            .QtdEstoque = 37
            .PrecoVenda = 44.9
        End With
        _arrayProduto.Add(_Produtos)

        _Produtos = New clsProduto
        With _Produtos
            .CodigoBarras = "7890000000222"
            .NomeProduto = "Metallica - Black Album"
            .Categoria = "CDs"
            .DataCadastro = New DateTime(2012, 9, 10)
            .QtdEstoque = 45
            .PrecoVenda = 39.95
        End With
        _arrayProduto.Add(_Produtos)

        _Produtos = New clsProduto
        With _Produtos
            .CodigoBarras = "7890000000777"
            .NomeProduto = "A Arte da Guerra"
            .Categoria = "Livros"
            .DataCadastro = New DateTime(2012, 7, 27)
            .QtdEstoque = 10
            .PrecoVenda = 10.0
        End With
        _arrayProduto.Add(_Produtos)

        Return _arrayProduto
    End Function


End Class

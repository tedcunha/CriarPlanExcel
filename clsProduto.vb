Public Class clsProduto
    Private _CodigoBarras As String
    Private _NomeProduto As String
    Private _Categoria As String
    Private _DataCadastro As DateTime
    Private _QtdEstoque As Integer
    Private _PrecoVenda As Double

    Public Property CodigoBarras() As String
        Get
            Return _CodigoBarras
        End Get
        Set(ByVal value As String)
            _CodigoBarras = value
        End Set
    End Property

    Public Property NomeProduto() As String
        Get
            Return _NomeProduto
        End Get
        Set(ByVal value As String)
            _NomeProduto = value
        End Set
    End Property

    Public Property Categoria() As String
        Get
            Return _Categoria
        End Get
        Set(ByVal value As String)
            _Categoria = value
        End Set
    End Property

    Public Property DataCadastro() As DateTime
        Get
            Return _DataCadastro
        End Get
        Set(ByVal value As DateTime)
            _DataCadastro = value
        End Set
    End Property

    Public Property QtdEstoque() As Integer
        Get
            Return _QtdEstoque
        End Get
        Set(ByVal value As Integer)
            _QtdEstoque = value
        End Set
    End Property

    Public Property PrecoVenda() As Double
        Get
            Return _PrecoVenda
        End Get
        Set(ByVal value As Double)
            _PrecoVenda = value
        End Set
    End Property
End Class

section Section1;

shared Cliente = let
    Fonte = Excel.Workbook(File.Contents("C:\Users\reyna\OneDrive\Treinamento\Treinamento.xlsx"), null, true),
    Cliente_Sheet = Fonte{[Item="Cliente",Kind="Sheet"]}[Data],
    #"Cabeçalhos Promovidos" = Table.PromoteHeaders(Cliente_Sheet, [PromoteAllScalars=true]),
    #"Tipo Alterado" = Table.TransformColumnTypes(#"Cabeçalhos Promovidos",{{"Cliente_ID", Int64.Type}, {"Cliente", type text}, {"CPF", Int64.Type}, {"CEP", type text}, {"Estado", type text}, {"Cidade", type text}, {"Bairro", type text}, {"Rua", type text}, {"Número", Int64.Type}, {"Latitude", type number}, {"Longitude", type number}})
in
    #"Tipo Alterado";

shared Produtos = let
    Fonte = Excel.Workbook(File.Contents("C:\Users\reyna\OneDrive\Treinamento\Treinamento.xlsx"), null, true),
    Produtos_Sheet = Fonte{[Item="Produtos",Kind="Sheet"]}[Data],
    #"Cabeçalhos Promovidos" = Table.PromoteHeaders(Produtos_Sheet, [PromoteAllScalars=true]),
    #"Tipo Alterado" = Table.TransformColumnTypes(#"Cabeçalhos Promovidos",{{"Produto_ID", Int64.Type}, {"Produto", type text}, {"Preço", Int64.Type}, {"Categoria", type text}})
in
    #"Tipo Alterado";

shared Vendas = let
    Fonte = Excel.Workbook(File.Contents("C:\Users\reyna\OneDrive\Treinamento\Treinamento.xlsx"), null, true),
    Vendas_Sheet = Fonte{[Item="Vendas",Kind="Sheet"]}[Data],
    #"Cabeçalhos Promovidos" = Table.PromoteHeaders(Vendas_Sheet, [PromoteAllScalars=true]),
    #"Tipo Alterado" = Table.TransformColumnTypes(#"Cabeçalhos Promovidos",{{"Venda_ID", Int64.Type}, {"Vendedor_ID", Int64.Type}, {"Cliente_ID", Int64.Type}, {"Produto_ID", Int64.Type}, {"DataVenda", type date}, {"Quantidade", Int64.Type}, {"Valor Total", Int64.Type}, {"Valor Por Unidade", Int64.Type}})
in
    #"Tipo Alterado";

shared Vendedor = let
    Fonte = Excel.Workbook(File.Contents("C:\Users\reyna\OneDrive\Treinamento\Treinamento.xlsx"), null, true),
    Vendedor_Sheet = Fonte{[Item="Vendedor",Kind="Sheet"]}[Data],
    #"Cabeçalhos Promovidos" = Table.PromoteHeaders(Vendedor_Sheet, [PromoteAllScalars=true]),
    #"Tipo Alterado" = Table.TransformColumnTypes(#"Cabeçalhos Promovidos",{{"Vendedor_ID", Int64.Type}, {"Vendedor", type text}})
in
    #"Tipo Alterado";
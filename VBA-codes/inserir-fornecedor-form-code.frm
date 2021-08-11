Private Sub InserirFornecedores_Click()
    ' Cadastra os materiais nas células corretas
    Call Cadastrar
    
    ' Formata as células
    Call Formatar
    
    ' Esconde o formulário após preenchimento
    InserirFornecedor.Hide
    
    ' Limpa o formulário após preenchimento
    Call Limpar
End Sub
Sub Cadastrar()
    Dim RefCadastro As Range
    
    Sheets("Informações fornecedores").Select
    If Range("B13").Value = "" Then
            Set RefCadastro = Range("B13")
        Else
            Set RefCadastro = Range("B13").End(xlDown).Offset(1, 0)
        End If
    
    RefCadastro.Value = InserirFornecedor.InserirEmpresa.Value
    RefCadastro.Offset(0, 1).Value = InserirFornecedor.InserirContato.Value
    RefCadastro.Offset(0, 2).Value = InserirFornecedor.InserirEndereço.Value
    RefCadastro.Offset(0, 3).Value = InserirFornecedor.InserirCidade.Value
    RefCadastro.Offset(0, 4).Value = InserirFornecedor.InserirEstado.Value
    RefCadastro.Offset(0, 5).Value = InserirFornecedor.InserirCEP.Value
    RefCadastro.Offset(0, 6).Value = InserirFornecedor.InserirEmail.Value
    
End Sub

Sub Limpar()
    ' Define objeto
    Dim Objeto As Control
    
    ' Faz um loop para inserir valor vazio em todos objetos
    For Each Objeto In InserirFornecedor.Controls
        ' Pula a cada erro que acontecer nas labels/buttons
        On Error Resume Next
        Objeto.Value = ""
    Next

End Sub

Sub Formatar()

    Sheets("Informações fornecedores").Select
    Columns("B:H").Select
    With Selection.Font
        .Name = "DINPro-Light"
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Selection.Font.Size = 12

End Sub

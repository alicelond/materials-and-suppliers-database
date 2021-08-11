Private Sub InserirProduto_Click()
    ' Cadastra os materiais nas células corretas
    Call Cadastrar
    
    ' Formata as células
    Call Formatar
    
    ' Esconde o formulário após preenchimento
    InserirMaterial.Hide
    
    ' Limpa o formulário após preenchimento
    Call Limpar
End Sub

Sub Cadastrar()
    Dim RefCadastro As Range
    
    Sheets("Informações Produtos").Select
    If Range("B9").Value = "" Then
            Set RefCadastro = Range("B9")
        Else
            Set RefCadastro = Range("B9").End(xlDown).Offset(1, 0)
        End If
    
    RefCadastro.Value = InserirMaterial.InserirTipoElemento.Value
    RefCadastro.Offset(0, 1).Value = InserirMaterial.TextBox1.Value
    RefCadastro.Offset(0, 2).Value = InserirMaterial.InserirEmpresa.Value
    RefCadastro.Offset(0, 3).Value = InserirMaterial.InserirNomeMaterial.Value
    RefCadastro.Offset(0, 4).Value = InserirMaterial.InserirEspecificações.Value
    RefCadastro.Offset(0, 5).Value = InserirMaterial.InserirLinhaProduto.Value
    RefCadastro.Offset(0, 6).Value = InserirMaterial.InserirDimensões.Value
    RefCadastro.Offset(0, 7).Value = InserirMaterial.InserirEspessura.Value
    RefCadastro.Offset(0, 8).Value = InserirMaterial.InserirPlacasFolhas.Value
    RefCadastro.Offset(0, 9).Value = InserirMaterial.InserirDensidadeSuperficial.Value
    RefCadastro.Offset(0, 10).Value = InserirMaterial.InserirPreço.Value
    RefCadastro.Offset(0, 11).Value = InserirMaterial.InserirRw.Value
    RefCadastro.Offset(0, 12).Value = InserirMaterial.InserirC.Value
    RefCadastro.Offset(0, 13).Value = InserirMaterial.InserirCtr.Value
    RefCadastro.Offset(0, 14).Value = InserirMaterial.InserirDntw.Value
    RefCadastro.Offset(0, 15).Value = InserirMaterial.InserirD2mntw.Value
    RefCadastro.Offset(0, 16).Value = InserirMaterial.InserirLntw.Value
    RefCadastro.Offset(0, 17).Value = InserirMaterial.InserirLnw.Value
    RefCadastro.Offset(0, 18).Value = InserirMaterial.InserirDeltaLw.Value
    RefCadastro.Offset(0, 19).Value = InserirMaterial.InserirNRC.Value
    RefCadastro.Offset(0, 20).Value = InserirMaterial.InserirAlfaW.Value
    RefCadastro.Offset(0, 21).Value = InserirMaterial.InserirSRA.Value
    RefCadastro.Offset(0, 22).Value = InserirMaterial.InserirCAC.Value
    RefCadastro.Offset(0, 23).Value = InserirMaterial.InserirExpedidorLaudo.Value
    RefCadastro.Offset(0, 24).Value = InserirMaterial.InserirIdentificaçãoLaudo.Value
    RefCadastro.Offset(0, 25).Value = InserirMaterial.Inserir50Hz.Value
    RefCadastro.Offset(0, 26).Value = InserirMaterial.Inserir63Hz.Value
    RefCadastro.Offset(0, 27).Value = InserirMaterial.Inserir80Hz.Value
    RefCadastro.Offset(0, 28).Value = InserirMaterial.Inserir100Hz.Value
    RefCadastro.Offset(0, 29).Value = InserirMaterial.Inserir125Hz.Value
    RefCadastro.Offset(0, 30).Value = InserirMaterial.Inserir160Hz.Value
    RefCadastro.Offset(0, 31).Value = InserirMaterial.Inserir200Hz.Value
    RefCadastro.Offset(0, 32).Value = InserirMaterial.Inserir250Hz.Value
    RefCadastro.Offset(0, 33).Value = InserirMaterial.Inserir315Hz.Value
    RefCadastro.Offset(0, 34).Value = InserirMaterial.Inserir400Hz.Value
    RefCadastro.Offset(0, 35).Value = InserirMaterial.Inserir500Hz.Value
    RefCadastro.Offset(0, 36).Value = InserirMaterial.Inserir630Hz.Value
    RefCadastro.Offset(0, 37).Value = InserirMaterial.Inserir800Hz.Value
    RefCadastro.Offset(0, 38).Value = InserirMaterial.Inserir1kHz.Value
    RefCadastro.Offset(0, 39).Value = InserirMaterial.Inserir1k25Hz.Value
    RefCadastro.Offset(0, 40).Value = InserirMaterial.Inserir1k6Hz.Value
    RefCadastro.Offset(0, 41).Value = InserirMaterial.Inserir2kHz.Value
    RefCadastro.Offset(0, 42).Value = InserirMaterial.Inserir2k5Hz.Value
    RefCadastro.Offset(0, 43).Value = InserirMaterial.Inserir3k15Hz.Value
    RefCadastro.Offset(0, 44).Value = InserirMaterial.Inserir4kHz.Value
    RefCadastro.Offset(0, 45).Value = InserirMaterial.Inserir5kHz.Value
    
End Sub

Sub Limpar()
    ' Define objeto
    Dim Objeto As Control
    
    ' Faz um loop para inserir valor vazio em todos objetos
    For Each Objeto In InserirMaterial.Controls
        ' Pula a cada erro que acontecer nas labels/buttons
        On Error Resume Next
        Objeto.Value = ""
    Next

End Sub

Sub Formatar()

    Sheets("Informações Produtos").Select
    Columns("B:AU").Select
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



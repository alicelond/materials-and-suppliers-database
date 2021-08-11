Private Sub InserirSistVertical_Click()
    ' Cadastra os materiais nas células corretas
    Call Cadastrar
    
    ' Formata as células
    Call Formatar
    
    ' Esconde o formulário após preenchimento
    SistVertical.Hide
    
    ' Limpa o formulário após preenchimento
    Call Limpar
End Sub
Sub Cadastrar()
    Dim RefCadastro As Range
    
    Sheets("Sistemas de vedação vertical").Select
    If Range("B9").Value = "" Then
            Set RefCadastro = Range("B9")
        Else
            Set RefCadastro = Range("B9").End(xlDown).Offset(1, 0)
        End If
    
    RefCadastro.Value = SistVertical.InserirBloco.Value
    RefCadastro.Offset(0, 1).Value = SistVertical.InserirAcabamentoExterno.Value
    RefCadastro.Offset(0, 2).Value = SistVertical.InserirAcabamentoInterno.Value
    RefCadastro.Offset(0, 3).Value = SistVertical.InserirCaixilho.Value
    RefCadastro.Offset(0, 4).Value = SistVertical.InserirInfoAdicional.Value
    RefCadastro.Offset(0, 5).Value = SistVertical.InserirSoftware.Value
    RefCadastro.Offset(0, 6).Value = SistVertical.InserirRw.Value
    RefCadastro.Offset(0, 7).Value = SistVertical.InserirC.Value
    RefCadastro.Offset(0, 8).Value = SistVertical.InserirCtr.Value
    RefCadastro.Offset(0, 9).Value = SistVertical.InserirDntw.Value
    RefCadastro.Offset(0, 10).Value = SistVertical.InserirD2mntw.Value
    
End Sub

Sub Limpar()
    ' Define objeto
    Dim Objeto As Control
    
    ' Faz um loop para inserir valor vazio em todos objetos
    For Each Objeto In SistVertical.Controls
        ' Pula a cada erro que acontecer nas labels/buttons
        On Error Resume Next
        Objeto.Value = ""
    Next

End Sub

Sub Formatar()

    Sheets("Sistemas de vedação vertical").Select
    Columns("B:M").Select
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


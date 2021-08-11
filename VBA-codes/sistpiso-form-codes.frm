Private Sub InserirSistPisos_Click()
    ' Cadastra os materiais nas células corretas
    Call Cadastrar
    
    ' Formata as células
    Call Formatar
    
    ' Esconde o formulário após preenchimento
    SistPiso.Hide
    
    ' Limpa o formulário após preenchimento
    Call Limpar
End Sub
Sub Cadastrar()
    Dim RefCadastro As Range
    
    Sheets("Sistemas de pisos").Select
    If Range("B9").Value = "" Then
            Set RefCadastro = Range("B9")
        Else
            Set RefCadastro = Range("B9").End(xlDown).Offset(1, 0)
        End If
    
    RefCadastro.Value = SistPiso.InserirLaje.Value
    RefCadastro.Offset(0, 1).Value = SistPiso.InserirRegularização.Value
    RefCadastro.Offset(0, 2).Value = SistPiso.InserirImpermeabilização.Value
    RefCadastro.Offset(0, 3).Value = SistPiso.InserirManta.Value
    RefCadastro.Offset(0, 4).Value = SistPiso.InserirRevestimento.Value
    RefCadastro.Offset(0, 5).Value = SistPiso.InserirInformaçõesAdicionais.Value
    RefCadastro.Offset(0, 6).Value = SistPiso.InserirSoftware.Value
    RefCadastro.Offset(0, 7).Value = SistPiso.InserirRw.Value
    RefCadastro.Offset(0, 8).Value = SistPiso.InserirC.Value
    RefCadastro.Offset(0, 9).Value = SistPiso.InserirCtr.Value
    RefCadastro.Offset(0, 10).Value = SistPiso.InserirLnw.Value
    RefCadastro.Offset(0, 11).Value = SistPiso.InserirLntw.Value
    
End Sub

Sub Limpar()
    ' Define objeto
    Dim Objeto As Control
    
    ' Faz um loop para inserir valor vazio em todos objetos
    For Each Objeto In SistPiso.Controls
        ' Pula a cada erro que acontecer nas labels/buttons
        On Error Resume Next
        Objeto.Value = ""
    Next

End Sub

Sub Formatar()

    Sheets("Sistemas de pisos").Select
    Columns("B:N").Select
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


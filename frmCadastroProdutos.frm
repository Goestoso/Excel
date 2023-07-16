VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCadastroProdutos 
   Caption         =   "Cadastro de Produtos"
   ClientHeight    =   4275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8190
   OleObjectBlob   =   "frmCadastroProdutos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCadastroProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbDescricao_Change()
    sbAparenciaAlteracao
End Sub






Private Sub cmdCancelar_Click()
    sbAparenciaNormal
End Sub







Private Sub cmdIncluir_Click()
    cmbDescricao = vbNullString
    txtTamanho = vbNullString
    txtPreco = vbNullString
    txtQuantidades = vbNullString
    
    cmdSalvar.Caption = "Incluir"
End Sub








Private Sub cmdSalvar_Click()
    
    'Verificação se a decrição do produto está preenchida
    If cmbDescricao = vbNullString Then
        MsgBox "A descrição do produto deve ser selecionada!"
        Exit Sub
    End If
    
    'Verificação do tamanho do tênis
    If Not IsNumeric(txtTamanho) Then
        MsgBox "O tamanho do tênis tem que ser numérico!"
        Exit Sub
    Else
        If CLng(txtTamanho) < 16 Or CLng(txtTamanho) > 36 Then
            MsgBox "O tamanho do tênis tem que estar entre 16 e 36!"
            Exit Sub
        End If
    End If
    
    
    'Verificação do preço unitário do tênis
    If Not IsNumeric(txtPreco) Then
        MsgBox "O Preço do tênis tem que ser numérico!"
        Exit Sub
    Else
        If CLng(txtPreco) < 1 Or CLng(txtPreco) > 200 Then
            MsgBox "O preço do tênis tem que estar entre 1 e 200!"
            Exit Sub
        End If
    End If
    
    
    'Verificação da quantidade em estoque
    If Not IsNumeric(txtQuantidades) Then
        MsgBox "A quantidade em estoque do tênis tem que ser numérico!"
        Exit Sub
    Else
        If CLng(txtQuantidades) < 0 Or CLng(txtQuantidades) > 999 Then
            MsgBox "A quantidade em estoque do tênis tem que estar entre 0 e 999!"
            Exit Sub
        End If
    End If
    
    
    'Se for um item novo
    If cmdSalvar.Caption = "Incluir" Then
        Planilha1.Cells(ActiveCell.Row, 1).Interior.Color = vbWhite
        Planilha1.Cells(ActiveCell.Row, 2).Interior.Color = vbWhite
        Planilha1.Cells(ActiveCell.Row, 3).Interior.Color = vbWhite
        Planilha1.Cells(ActiveCell.Row, 4).Interior.Color = vbWhite
        Planilha1.Cells(ActiveCell.Row, 5).Interior.Color = vbWhite
        Planilha1.Cells(ActiveCell.Row, 6).Interior.Color = vbWhite
        
        Range("A2").Select
        Selection.End(xlDown).Select
        ActiveCell.Offset(1, 0).Select
    End If
    
    Planilha1.Cells(ActiveCell.Row, 1) = cmbDescricao
    Planilha1.Cells(ActiveCell.Row, 2) = CLng(txtTamanho)
    Planilha1.Cells(ActiveCell.Row, 3) = CCur(txtPreco)
    Planilha1.Cells(ActiveCell.Row, 4) = CLng(txtQuantidades)
    
    sbAparenciaNormal
End Sub







Private Sub txtPreco_Change()
    sbAparenciaAlteracao
End Sub





Private Sub txtQuantidades_Change()
    sbAparenciaAlteracao
End Sub





Private Sub txtTamanho_Change()
    sbAparenciaAlteracao
End Sub





Private Sub cmdSair_Click()
    Unload Me
End Sub






Private Sub UserForm_Activate()
    cmbDescricao.AddItem "Tênis Infantil Nika Vermelho"
    cmbDescricao.AddItem "Tênis Infantil Nika Rosa"
    cmbDescricao.AddItem "Tênis Infantil Nika Azul"
    cmbDescricao.AddItem "Tênis Infantil Atitas Vermelho"
    cmbDescricao.AddItem "Tênis Infantil Atitas Rosa"
    cmbDescricao.AddItem "Tênis Infantil Atitas Azul"
    
    sbAparenciaNormal
End Sub







Private Sub sbAparenciaAlteracao()

    cmdSair.Visible = False
    cmdIncluir.Visible = False
    cmdSalvar.Visible = True
    cmdCancelar.Visible = True

End Sub









Private Sub sbAparenciaNormal()
    
    If Trim$(Planilha1.Cells(ActiveCell.Row, 1)) = vbNullString Then
        MsgBox "Não é possível carregar dados com a descrição do produto em branco!"
        Unload Me
        Exit Sub
    End If
    
    cmbDescricao = Planilha1.Cells(ActiveCell.Row, 1)
    txtTamanho = Planilha1.Cells(ActiveCell.Row, 2)
    txtPreco = FormatCurrency(Planilha1.Cells(ActiveCell.Row, 3), 2)
    txtQuantidades = Planilha1.Cells(ActiveCell.Row, 4)
    
    cmdIncluir.Visible = True
    cmdSair.Visible = True
    
    cmdCancelar.Visible = False
    cmdSalvar.Visible = False
    
    cmdSalvar.Caption = "Salvar"
End Sub



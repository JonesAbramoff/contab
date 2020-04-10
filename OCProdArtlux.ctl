VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.UserControl OCProdArtlux 
   ClientHeight    =   5040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   ScaleHeight     =   5040
   ScaleWidth      =   6000
   Begin VB.Frame Frame2 
      Caption         =   "Identificação"
      Height          =   915
      Left            =   60
      TabIndex        =   9
      Top             =   15
      Width           =   5850
      Begin VB.Label Quantidade 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   4425
         TabIndex        =   15
         Top             =   195
         Width           =   1290
      End
      Begin VB.Label Label1 
         Caption         =   "Quantidade:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   3315
         TabIndex        =   14
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label ProdutoDesc 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1035
         TabIndex        =   13
         Top             =   540
         Width           =   4680
      End
      Begin VB.Label Label1 
         Caption         =   "Descrição:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   60
         TabIndex        =   12
         Top             =   585
         Width           =   900
      End
      Begin VB.Label Produto 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1035
         TabIndex        =   11
         Top             =   195
         Width           =   2085
      End
      Begin VB.Label Label1 
         Caption         =   "Produto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   255
         TabIndex        =   10
         Top             =   240
         Width           =   825
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalhamento"
      Height          =   4020
      Left            =   60
      TabIndex        =   3
      Top             =   915
      Width           =   5850
      Begin VB.TextBox DataFim 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   245
         Left            =   3945
         TabIndex        =   8
         Top             =   1605
         Width           =   945
      End
      Begin VB.TextBox DataIni 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   245
         Left            =   3855
         TabIndex        =   7
         Top             =   990
         Width           =   945
      End
      Begin VB.ComboBox Usuario 
         Height          =   315
         ItemData        =   "OCProdArtlux.ctx":0000
         Left            =   420
         List            =   "OCProdArtlux.ctx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   945
         Width           =   1515
      End
      Begin MSMask.MaskEdBox QuantProd 
         Height          =   225
         Left            =   1935
         TabIndex        =   4
         Top             =   1005
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox QuantPreProd 
         Height          =   225
         Left            =   2880
         TabIndex        =   5
         Top             =   1005
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin VB.CommandButton BotaoOK 
         Caption         =   "OK"
         Height          =   540
         Left            =   1950
         Picture         =   "OCProdArtlux.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3405
         Width           =   855
      End
      Begin VB.CommandButton BotaoCancela 
         Caption         =   "Cancelar"
         Height          =   540
         Left            =   2880
         Picture         =   "OCProdArtlux.ctx":015E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3405
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid GridD 
         Height          =   435
         Left            =   30
         TabIndex        =   0
         Top             =   210
         Width           =   5790
         _ExtentX        =   10213
         _ExtentY        =   767
         _Version        =   393216
         Rows            =   8
         Cols            =   6
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   2
      End
   End
End
Attribute VB_Name = "OCProdArtlux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()
 
Dim iAlterado As Integer
Dim gobjOC As ClassOCArtlux
Dim gobjOCAux As ClassOCArtlux

Dim objGridD As AdmGrid
Dim iGrid_Usuario_Col As Integer
Dim iGrid_QuantPreProd_Col As Integer
Dim iGrid_QuantProd_Col As Integer
Dim iGrid_DataIni_Col As Integer
Dim iGrid_DataFim_Col As Integer

Private Sub BotaoCancela_Click()
    giRetornoTela = vbCancel
    Unload Me
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
   
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr
        
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206710)
            
    End Select
    
    iAlterado = 0
    
    Exit Sub
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, UnloadMode, Cancel, iTelaCorrenteAtiva)
End Sub

Public Sub Form_Unload(Cancel As Integer)

    'Libera as variaveis globais
    Set gobjOC = Nothing
    Set gobjOCAux = Nothing
    Set objGridD = Nothing
    
End Sub

Private Sub GridD_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridD, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridD, iAlterado)
    End If

End Sub

Private Sub GridD_EnterCell()

    Call Grid_Entrada_Celula(objGridD, iAlterado)

End Sub

Private Sub GridD_GotFocus()

    Call Grid_Recebe_Foco(objGridD)

End Sub

Private Sub GridD_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long
Dim iItemAtual As Integer
Dim iLinhasExistentesAnt As Integer
    
On Error GoTo Erro_GridD_KeyDown

    'Guarda o número de linhas existentes e a linha atual
    iLinhasExistentesAnt = objGridD.iLinhasExistentes
    iItemAtual = GridD.Row
    
    lErro = Remove_Linha(objGridD, iItemAtual, KeyCode)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call Grid_Trata_Tecla1(KeyCode, objGridD)

    Exit Sub

Erro_GridD_KeyDown:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206724)

    End Select

    Exit Sub

End Sub

Private Sub GridD_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridD, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridD, iAlterado)
    End If

End Sub

Private Sub GridD_LeaveCell()

    Call Saida_Celula(objGridD)

End Sub

Private Sub GridD_Validate(Cancel As Boolean)
        
    Call Grid_Libera_Foco(objGridD)
        
End Sub

Private Sub GridD_RowColChange()

    Call Grid_RowColChange(objGridD)

End Sub

Private Sub GridD_Scroll()

    Call Grid_Scroll(objGridD)

End Sub

Private Sub QuantPreProd_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub QuantPreProd_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridD)
End Sub

Private Sub QuantPreProd_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridD)
End Sub

Private Sub QuantPreProd_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridD.objControle = QuantPreProd
    lErro = Grid_Campo_Libera_Foco(objGridD)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantProd_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub QuantProd_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridD)
End Sub

Private Sub QuantProd_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridD)
End Sub

Private Sub QuantProd_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridD.objControle = QuantProd
    lErro = Grid_Campo_Libera_Foco(objGridD)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Usuario_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Usuario_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridD)
End Sub

Private Sub Usuario_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridD)
End Sub

Private Sub Usuario_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridD.objControle = Usuario
    lErro = Grid_Campo_Libera_Foco(objGridD)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Inicializa_Grid_D(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Alocação

Dim iIndice As Integer

    Set objGridD.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Usuário")
    objGridInt.colColuna.Add ("Saída")
    objGridInt.colColuna.Add ("Entrada")
    objGridInt.colColuna.Add ("Início")
    objGridInt.colColuna.Add ("Fim")
    
    'Controles que participam do Grid
    objGridInt.colCampo.Add (Usuario.Name)
    objGridInt.colCampo.Add (QuantPreProd.Name)
    objGridInt.colCampo.Add (QuantProd.Name)
    objGridInt.colCampo.Add (DataIni.Name)
    objGridInt.colCampo.Add (DataFim.Name)

    'Colunas da Grid
    iGrid_Usuario_Col = 1
    iGrid_QuantPreProd_Col = 2
    iGrid_QuantProd_Col = 3
    iGrid_DataIni_Col = 4
    iGrid_DataFim_Col = 5

    'Grid do GridInterno
    objGridInt.objGrid = GridD

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 8
    
    objGridInt.objGrid.Rows = 101

    'Largura da primeira coluna
    GridD.ColWidth(0) = 300
    
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridD)

    Inicializa_Grid_D = SUCESSO

    Exit Function

End Function

Public Function Trata_Parametros(ByVal objOC As ClassOCArtlux) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objOCProd As ClassOCProdArtlux

On Error GoTo Erro_Trata_Parametros

    Set gobjOC = objOC
    
    Set objGridD = New AdmGrid
    
    lErro = Usuarios_Carrega()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Inicializa o grid de alocacoes
    lErro = Inicializa_Grid_D(objGridD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Traz_Tela(objOC)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Set gobjOCAux = New ClassOCArtlux
    Call gobjOCAux.Copiar(gobjOC)

    iAlterado = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
               
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206711)
            
    End Select
    
    iAlterado = 0
    
    Exit Function
    
End Function

Public Function Traz_Tela(objOC As ClassOCArtlux) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objOCProd As ClassOCProdArtlux
Dim sProdMask As String

On Error GoTo Erro_Traz_Tela

    For Each objOCProd In objOC.colItens
    
        iLinha = iLinha + 1

        GridD.TextMatrix(iLinha, iGrid_Usuario_Col) = objOCProd.sUsuMontagem
        
        If objOCProd.dQuantidadePreProd > 0 Then
            GridD.TextMatrix(iLinha, iGrid_QuantPreProd_Col) = Formata_Estoque(objOCProd.dQuantidadePreProd)
        Else
            GridD.TextMatrix(iLinha, iGrid_QuantPreProd_Col) = ""
        End If
        
        If objOCProd.dQuantidadeProd > 0 Then
            GridD.TextMatrix(iLinha, iGrid_QuantProd_Col) = Formata_Estoque(objOCProd.dQuantidadeProd)
        Else
            GridD.TextMatrix(iLinha, iGrid_QuantProd_Col) = ""
        End If
        
        If objOCProd.dtDataIniMontagem <> DATA_NULA Then
            GridD.TextMatrix(iLinha, iGrid_DataIni_Col) = Format(objOCProd.dtDataIniMontagem, "dd/mm/yyyy")
        Else
            GridD.TextMatrix(iLinha, iGrid_DataIni_Col) = ""
        End If
        
        If objOCProd.dtDataFimMontagem <> DATA_NULA Then
            GridD.TextMatrix(iLinha, iGrid_DataFim_Col) = Format(objOCProd.dtDataFimMontagem, "dd/mm/yyyy")
        Else
            GridD.TextMatrix(iLinha, iGrid_DataFim_Col) = ""
        End If
    Next
    
    Call Mascara_RetornaProdutoTela(objOC.sProduto, sProdMask)
    
    Produto.Caption = sProdMask
    ProdutoDesc.Caption = objOC.sProdutoDesc
    Quantidade.Caption = Formata_Estoque(objOC.dQuantidade)
    
    objGridD.iLinhasExistentes = objOC.colItens.Count
    
    Traz_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Tela:

    Traz_Tela = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
               
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206712)
            
    End Select
    
    Exit Function
    
End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        'Verifica qual a coluna do Grid em questão
        Select Case objGridInt.objGrid.Col
        
            Case iGrid_QuantProd_Col
                lErro = Saida_Celula_QuantProd(objGridInt)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            Case iGrid_QuantPreProd_Col
                lErro = Saida_Celula_QuantPreProd(objGridInt)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            Case iGrid_Usuario_Col
                lErro = Saida_Celula_Usuario(objGridInt)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
         
        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 206713

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 206713
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206714)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_QuantPreProd(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_QuantPreProd

    Set objGridInt.objControle = QuantPreProd

    If Len(Trim(QuantPreProd.ClipText)) > 0 Then
    
        lErro = Valor_NaoNegativo_Critica(QuantPreProd.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        QuantPreProd.Text = Formata_Estoque(StrParaDbl(QuantPreProd.Text))
        
        GridD.TextMatrix(GridD.Row, iGrid_DataIni_Col) = Format(Date, "dd/mm/yyyy")
        
    Else
                       
        GridD.TextMatrix(GridD.Row, iGrid_DataIni_Col) = ""
                       
    End If
    
    gobjOCAux.colItens(GridD.Row).dQuantidadePreProd = StrParaDbl(QuantPreProd.Text)
    gobjOCAux.colItens(GridD.Row).dtDataIniMontagem = StrParaDate(GridD.TextMatrix(GridD.Row, iGrid_DataIni_Col))

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 39381
    
    Saida_Celula_QuantPreProd = SUCESSO

    Exit Function
    
Erro_Saida_Celula_QuantPreProd:

    Saida_Celula_QuantPreProd = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206715)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
    End Select
 
    Exit Function
 
End Function

Private Function Saida_Celula_QuantProd(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_QuantProd

    Set objGridInt.objControle = QuantProd

    If Len(Trim(QuantProd.ClipText)) > 0 Then
    
        lErro = Valor_NaoNegativo_Critica(QuantProd.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                       
        QuantProd.Text = Formata_Estoque(StrParaDbl(QuantProd.Text))
                       
        GridD.TextMatrix(GridD.Row, iGrid_DataFim_Col) = Format(Date, "dd/mm/yyyy")
        
    Else
                       
        GridD.TextMatrix(GridD.Row, iGrid_DataFim_Col) = ""
                       
    End If
    
    If gobjOCAux.colItens(GridD.Row).dQuantidadeProd - gobjOCAux.colItens(GridD.Row).dQuantidadePreProd > QTDE_ESTOQUE_DELTA Then gError 206719
    
    gobjOCAux.colItens(GridD.Row).dQuantidadeProd = StrParaDbl(QuantProd.Text)
    gobjOCAux.colItens(GridD.Row).dtDataFimMontagem = StrParaDate(GridD.TextMatrix(GridD.Row, iGrid_DataFim_Col))

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 39381
    
    Call Trata_Alteracao_Qtd_Prod(gobjOCAux.colItens(GridD.Row))
        
    Saida_Celula_QuantProd = SUCESSO

    Exit Function
    
Erro_Saida_Celula_QuantProd:

    Saida_Celula_QuantProd = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_QTDENT_MAIOR_QTDSAIDA", gErr)
            
        Case 206719
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206716)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
    End Select
 
    Exit Function
 
End Function

Private Function Saida_Celula_Usuario(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objOCProd As ClassOCProdArtlux
Dim dQtdePreProd As Double

On Error GoTo Erro_Saida_Celula_Usuario

    Set objGridInt.objControle = Usuario

    If Len(Trim(Usuario.Text)) > 0 Then
        
        dQtdePreProd = 0
        For Each objOCProd In gobjOCAux.colItens
            dQtdePreProd = dQtdePreProd + objOCProd.dQuantidadePreProd
        Next
  
        GridD.TextMatrix(GridD.Row, iGrid_DataIni_Col) = Format(Date, "dd/mm/yyyy")
        GridD.TextMatrix(GridD.Row, iGrid_QuantPreProd_Col) = Formata_Estoque(gobjOCAux.dQuantidade - dQtdePreProd)
        
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
           objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
           Set objOCProd = New ClassOCProdArtlux
           gobjOCAux.colItens.Add objOCProd
        End If
        gobjOCAux.colItens(GridD.Row).sUsuMontagem = Usuario.Text
        gobjOCAux.colItens(GridD.Row).dQuantidadePreProd = StrParaDbl(GridD.TextMatrix(GridD.Row, iGrid_QuantPreProd_Col))
        gobjOCAux.colItens(GridD.Row).dtDataIniMontagem = StrParaDate(GridD.TextMatrix(GridD.Row, iGrid_DataIni_Col))
        gobjOCAux.colItens(GridD.Row).dQuantidadeProd = StrParaDbl(GridD.TextMatrix(GridD.Row, iGrid_QuantProd_Col))
        gobjOCAux.colItens(GridD.Row).dtDataFimMontagem = StrParaDate(GridD.TextMatrix(GridD.Row, iGrid_DataFim_Col))
                       
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Saida_Celula_Usuario = SUCESSO

    Exit Function
    
Erro_Saida_Celula_Usuario:

    Saida_Celula_Usuario = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 206717)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
    End Select
 
    Exit Function
 
End Function

Private Sub BotaoOK_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoOK_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Unload Me

    iAlterado = 0

    Exit Sub

Erro_BotaoOK_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 206718)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objOCProd As ClassOCProdArtlux
Dim iSeq As Integer
Dim dQtdPreProd As Double
Dim dQtdProd As Double
Dim sUsuMontagem As String

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    iSeq = 0
    For Each objOCProd In gobjOCAux.colItens
        iSeq = iSeq + 1
        objOCProd.iSeq = iSeq
        If objOCProd.dQuantidadePreProd = 0 Then gError 206720
        dQtdPreProd = dQtdPreProd + objOCProd.dQuantidadePreProd
        dQtdProd = dQtdProd + objOCProd.dQuantidadeProd
        
        If Len(Trim(sUsuMontagem)) = 0 Then
            sUsuMontagem = objOCProd.sUsuMontagem
        Else
            If sUsuMontagem <> objOCProd.sUsuMontagem Then sUsuMontagem = "Vários"
        End If
    Next
    
    gobjOCAux.dQuantidadeProd = dQtdProd
    gobjOCAux.dQuantidadePreProd = dQtdPreProd
    gobjOCAux.sUsuMontagem = sUsuMontagem
    
    If dQtdPreProd - gobjOCAux.dQuantidade > QTDE_ESTOQUE_DELTA Then gError 206727
    
    lErro = CF("OrdensDeCorteArtlux_Grava", gobjOCAux)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    giRetornoTela = vbOK
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
    
        Case 206720
            Call Rotina_Erro(vbOKOnly, "ERRO_QTDENT_NAO_INFORMADA", gErr)
            
        Case 206727
            Call Rotina_Erro(vbOKOnly, "ERRO_QTDENT_MAIOR_QTDCORTE", gErr)
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206721)
            
    End Select

    Exit Function

End Function

Private Function Trata_Alteracao_Qtd_Prod(ByVal objOCProd As ClassOCProdArtlux) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objOCProdAux As New ClassOCProdArtlux

On Error GoTo Erro_Trata_Alteracao_Qtd_Prod

    'Se produziu algo e não foi tudo acerta o grid para lançar saldo restante em nova linha
    If objOCProd.dQuantidadeProd > QTDE_ESTOQUE_DELTA And Abs(objOCProd.dQuantidadePreProd - objOCProd.dQuantidadeProd) > QTDE_ESTOQUE_DELTA Then
    
        objOCProdAux.dQuantidadePreProd = objOCProd.dQuantidadePreProd - objOCProd.dQuantidadeProd
        objOCProd.dQuantidadePreProd = objOCProd.dQuantidadeProd
        objOCProdAux.dtDataIniMontagem = objOCProd.dtDataIniMontagem
        objOCProdAux.sUsuMontagem = objOCProd.sUsuMontagem
        objOCProdAux.dtDataFimMontagem = DATA_NULA
        
        gobjOCAux.colItens.Add objOCProdAux
        
        lErro = Traz_Tela(gobjOCAux)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
            
    Trata_Alteracao_Qtd_Prod = SUCESSO
    
    Exit Function
    
Erro_Trata_Alteracao_Qtd_Prod:

    Trata_Alteracao_Qtd_Prod = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206722)
            
    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_LOCALIZACAO_PRODUTO
    Set Form_Load_Ocx = Me
    Caption = "Produção"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "OCProdArtlux"
    
End Function

Public Sub Show()
'    Parent.Show
'    Parent.SetFocus
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Controls
Public Property Get Controls() As Object
    Set Controls = UserControl.Controls
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Height() As Long
    Height = UserControl.Height
End Property

Public Property Get Width() As Long
    Width = UserControl.Width
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ActiveControl
Public Property Get ActiveControl() As Object
    Set ActiveControl = UserControl.ActiveControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Private Sub Unload(objme As Object)
   ' Parent.UnloadDoFilho
    
   RaiseEvent Unload
    
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

'***** fim do trecho a ser copiado ******

Private Function Usuarios_Carrega() As Long
'Carrega a ListBox

Dim lErro As Long
Dim objUsu As New ClassUsuProdArtlux
Dim colUsuarios As New Collection
Dim objUsuarios As New ClassUsuarios
Dim colUsu As New Collection

On Error GoTo Erro_Usuarios_Carrega

    Usuario.Clear

    'Le todos os Usuarios da Colecao
    lErro = CF("Usuarios_Le_Todos", colUsuarios)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Le todos os Compradores da Filial Empresa
    lErro = CF("UsuProdArtlux_Le_Todos", colUsu)
    If lErro <> SUCESSO And lErro <> 50126 Then gError ERRO_SEM_MENSAGEM

    For Each objUsu In colUsu
        For Each objUsuarios In colUsuarios
            If objUsu.sCodUsuario = objUsuarios.sCodUsuario Then
                If objUsu.iAcessoMontagem = MARCADO Then
                    Usuario.AddItem objUsu.sCodUsuario
                End If
            End If
        Next
    Next

    Usuarios_Carrega = SUCESSO

    Exit Function

Erro_Usuarios_Carrega:

    Usuarios_Carrega = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206723)

    End Select

    Exit Function

End Function

Public Function Remove_Linha(ByVal objGridInt As AdmGrid, ByVal iLinha As Integer, ByVal iKeyCode As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_Remove_Linha

    If iKeyCode = vbKeyDelete Then
        gobjOCAux.colItens.Remove (iLinha)
    End If
    
    Remove_Linha = SUCESSO
        
    Exit Function

Erro_Remove_Linha:

    Remove_Linha = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 206725)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable
              
    Select Case objControl.Name

        Case Usuario.Name
            If Len(Trim(GridD.TextMatrix(iLinha, iGrid_Usuario_Col))) = 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
            
        Case QuantPreProd.Name
            If Len(Trim(GridD.TextMatrix(iLinha, iGrid_Usuario_Col))) <> 0 And StrParaDbl(GridD.TextMatrix(iLinha, iGrid_QuantProd_Col)) = 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
            
        Case QuantProd.Name
            If StrParaDbl(GridD.TextMatrix(iLinha, iGrid_QuantPreProd_Col)) <> 0 And StrParaDbl(GridD.TextMatrix(iLinha, iGrid_QuantProd_Col)) = 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
     
        Case Else
            objControl.Enabled = False
            
    End Select
        
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 206726)

    End Select

    Exit Sub

End Sub

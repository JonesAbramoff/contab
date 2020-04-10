VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl BorderoChequesPre1Ocx 
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8445
   ScaleHeight     =   5550
   ScaleWidth      =   8445
   Begin VB.CheckBox CheckPago 
      Height          =   240
      Left            =   3960
      TabIndex        =   26
      Top             =   1335
      Width           =   1320
   End
   Begin VB.TextBox DataDeposito 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   540
      TabIndex        =   25
      Top             =   1020
      Width           =   1470
   End
   Begin VB.TextBox Numero 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   2490
      TabIndex        =   24
      Top             =   1020
      Width           =   1155
   End
   Begin VB.TextBox ContaCorrente 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   3720
      TabIndex        =   23
      Top             =   990
      Width           =   1575
   End
   Begin VB.TextBox Agencia 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   510
      TabIndex        =   22
      Top             =   750
      Width           =   885
   End
   Begin VB.TextBox Banco 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   1485
      TabIndex        =   21
      Top             =   750
      Width           =   1065
   End
   Begin VB.TextBox Filial 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   2640
      TabIndex        =   20
      Top             =   750
      Width           =   945
   End
   Begin VB.TextBox Cliente 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   3690
      TabIndex        =   19
      Top             =   750
      Width           =   1530
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   225
      Left            =   4800
      TabIndex        =   17
      Top             =   750
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      Enabled         =   0   'False
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
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox FilialEmpresa 
      Height          =   225
      Left            =   6765
      TabIndex        =   18
      Top             =   750
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      Enabled         =   0   'False
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
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin VB.CommandButton BotaoDesmarcar 
      Caption         =   "Desmarcar Todas"
      Height          =   555
      Left            =   6630
      Picture         =   "BorderoChequesPre1Ocx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4590
      Width           =   1530
   End
   Begin VB.CommandButton BotaoMarcar 
      Caption         =   "Marcar Todas"
      Height          =   555
      Left            =   4950
      Picture         =   "BorderoChequesPre1Ocx.ctx":11E2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4590
      Width           =   1530
   End
   Begin VB.Frame Frame1 
      Caption         =   "Total"
      Height          =   1005
      Left            =   210
      TabIndex        =   13
      Top             =   4290
      Width           =   2220
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   15
         Top             =   615
         Width           =   510
      End
      Begin VB.Label TotalCheques 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   735
         TabIndex        =   1
         Top             =   585
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Qtde.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   105
         TabIndex        =   14
         Top             =   270
         Width           =   540
      End
      Begin VB.Label QtdCheques 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   735
         TabIndex        =   0
         Top             =   255
         Width           =   780
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selecionado"
      Height          =   1005
      Left            =   2580
      TabIndex        =   10
      Top             =   4290
      Width           =   2205
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Qtde.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   255
         Width           =   540
      End
      Begin VB.Label QtdChequesSelecionados 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   765
         TabIndex        =   2
         Top             =   225
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   585
         Width           =   510
      End
      Begin VB.Label TotalChequesSelecionados 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   765
         TabIndex        =   3
         Top             =   570
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture7 
      Height          =   555
      Left            =   5490
      ScaleHeight     =   495
      ScaleWidth      =   2685
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   90
      Width           =   2745
      Begin VB.CommandButton BotaoSeguir 
         Height          =   345
         Left            =   1117
         Picture         =   "BorderoChequesPre1Ocx.ctx":21FC
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   90
         Width           =   930
      End
      Begin VB.CommandButton BotaoVoltar 
         Height          =   345
         Left            =   150
         Picture         =   "BorderoChequesPre1Ocx.ctx":298E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   90
         Width           =   885
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   2160
         Picture         =   "BorderoChequesPre1Ocx.ctx":30EC
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridBorderoPag2 
      Height          =   3330
      Left            =   210
      TabIndex        =   16
      Top             =   810
      Width           =   7950
      _ExtentX        =   14023
      _ExtentY        =   5874
      _Version        =   393216
      Rows            =   7
      Cols            =   4
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
   End
End
Attribute VB_Name = "BorderoChequesPre1Ocx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim gobjBorderoChequePre As ClassBorderoChequePre

Dim objGrid As AdmGrid
Dim iGrid_FilialEmpresa_Col As Integer
Dim iGrid_Cliente_Col As Integer
'Dim iGrid_NumTitulo_Col As Integer
'Dim iGrid_Parcela_Col As Integer
Dim iGrid_Filial_Col As Integer
Dim iGrid_Banco_Col As Integer
Dim iGrid_Agencia_Col As Integer
Dim iGrid_ContaCorrente_Col As Integer
Dim iGrid_Numero_Col As Integer
Dim iGrid_DataDeposito_Col As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_CheckPago_Col As Integer


Private Sub BotaoFechar_Click()

    'Fecha a tela
    Unload Me

End Sub

Private Sub BotaoDesmarcar_Click()
'Desmarca todos os cheques marcados no Grid

Dim iLinha As Integer
Dim objChequePre As New ClassChequePre


    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGrid.iLinhasExistentes

        'Desmarca na tela o cheque em questão
        GridBorderoPag2.TextMatrix(iLinha, iGrid_CheckPago_Col) = PAGO_NAO_CHECADO

        'Passa a linha do Grid para o Obj
        Set objChequePre = gobjBorderoChequePre.colchequepre.Item(iLinha)

        'Desmarca no Obj o cheque em questão
        objChequePre.iChequeSel = 0

    Next

    'Atualiza na tela os checkbox desmarcados
    Call Grid_Refresh_Checkbox(objGrid)

    'Limpa na tela os campos Qtd de cheques selecionados e Valor total dos cheques selecionados
    QtdChequesSelecionados.Caption = CStr(0)
    TotalChequesSelecionados.Caption = CStr(Format(0, "Standard"))
    gobjBorderoChequePre.iQuantChequesSel = 0

End Sub


Private Sub BotaoMarcar_Click()
'Marca todos os cheques no Grid

Dim iLinha As Integer
Dim dTotalChequesSelecionados As Double
Dim iNumChequesSelecionados As Integer
Dim objChequePre As ClassChequePre

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGrid.iLinhasExistentes

        'Marca na tela o cheque em questão
        GridBorderoPag2.TextMatrix(iLinha, iGrid_CheckPago_Col) = MARCADO

        'Passa a linha do Grid para o Obj
        Set objChequePre = gobjBorderoChequePre.colchequepre.Item(iLinha)

        'Marca no Obj o cheque em questão
        objChequePre.iChequeSel = 1

        dTotalChequesSelecionados = dTotalChequesSelecionados + CDbl(GridBorderoPag2.TextMatrix(iLinha, iGrid_Valor_Col))
        iNumChequesSelecionados = iNumChequesSelecionados + 1

    Next

    'Atualiza na tela os checkbox marcados
    Call Grid_Refresh_Checkbox(objGrid)

    'Atualiza na tela os campos Qtd de cheques selecionados e Valor total dos cheques selecionados
    QtdChequesSelecionados.Caption = CStr(iNumChequesSelecionados)
    TotalChequesSelecionados.Caption = CStr(Format(dTotalChequesSelecionados, "Standard"))
    gobjBorderoChequePre.iQuantChequesSel = iNumChequesSelecionados

End Sub

Private Sub BotaoSeguir_Click()

Dim lErro As Long
Dim dValor As Double
Dim iLinha As Integer

On Error GoTo Erro_BotaoSeguir_Click

    iLinha = 0

    'Ao menos um cheque tem que estar marcado p/pagto
    If CInt(QtdChequesSelecionados.Caption) = 0 Then gError 80331

    If Len(Trim(TotalChequesSelecionados.Caption)) <> 0 Then
        
        dValor = CDbl(TotalChequesSelecionados.Caption)
        
        'Início acrescentado por rafael menezes em 16/10/2002
        'guarda o total de cheques selecionados no grid
        gobjBorderoChequePre.dValorChequesSelecionados = dValor
        'Fim acrescentado por rafael menezes em 16/10/2002
        
    End If
          
    'Chama a tela do passo seguinte
    Call Chama_Tela("BorderoChequesPre2", gobjBorderoChequePre)

    'Fecha a tela
    Unload Me

    Exit Sub

Erro_BotaoSeguir_Click:

    Select Case gErr


        Case 80331
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SEM_CHEQUE_SELECIONADO", gErr, Error)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143633)

    End Select

    Exit Sub

End Sub

Private Sub BotaoVoltar_Click()

Dim lErro As Long
Dim dValor As Double

On Error GoTo Erro_BotaoVoltar_Click

    If Len(Trim(TotalChequesSelecionados.Caption)) <> 0 Then
        dValor = CDbl(TotalChequesSelecionados.Caption)
    End If
    
    If gobjBorderoChequePre.dValorChequesSelecionados <> dValor Then
        gobjBorderoChequePre.dValorChequesSelecionados = dValor
    End If
    
    'Chama a tela do passo anterior
    Call Chama_Tela("BorderoChequesPre", gobjBorderoChequePre)

    'Limpa o grid da tela em questão
    Call Grid_Limpa(objGrid)

    'Fecha a tela
    Unload Me

    Exit Sub

Erro_BotaoVoltar_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143634)

    End Select

    Exit Sub

End Sub

Private Sub CheckPago_Click()

Dim iLinha As Integer
Dim dTotalChequesSelecionados As Double
Dim iNumChequesSelecionados As Integer
Dim objChequePre As ClassChequePre

    iLinha = GridBorderoPag2.Row

    'Passa a linha do Grid para o Obj
    Set objChequePre = gobjBorderoChequePre.colchequepre.Item(iLinha)

    'Passa para o Obj se o cheque em questão foi marcado ou desmarcado
    objChequePre.iChequeSel = CInt(GridBorderoPag2.TextMatrix(iLinha, iGrid_CheckPago_Col))

    'Se o cheque foi marcado
    If GridBorderoPag2.TextMatrix(iLinha, iGrid_CheckPago_Col) = MARCADO Then

        'Acrescenta o novo cheque no somatório de Qtd de Títulos selecionados e Valor total de Cheques selecionados
        gobjBorderoChequePre.iQuantChequesSel = gobjBorderoChequePre.iQuantChequesSel + 1
        QtdChequesSelecionados.Caption = CStr(CInt(QtdChequesSelecionados.Caption) + 1)
        TotalChequesSelecionados.Caption = CStr(Format(CDbl(TotalChequesSelecionados.Caption) + CDbl(GridBorderoPag2.TextMatrix(iLinha, iGrid_Valor_Col)), "Standard"))

    Else

        'Subtrai o cheque do somatório de Qtd de Títulos selecionados e Valor total de Títulos selecionados
        gobjBorderoChequePre.iQuantChequesSel = gobjBorderoChequePre.iQuantChequesSel - 1
        QtdChequesSelecionados.Caption = CStr(CInt(QtdChequesSelecionados.Caption) - 1)
        TotalChequesSelecionados.Caption = CStr(Format(CDbl(TotalChequesSelecionados.Caption) - CDbl(GridBorderoPag2.TextMatrix(iLinha, iGrid_Valor_Col)), "Standard"))

    End If

End Sub

Private Sub CheckPago_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub CheckPago_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub CheckPago_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = CheckPago
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Form_Load()

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Form_Load

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143635)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente.

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    'Inicializa variáveis para saída de célula
    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then gError 80332

    'Finaliza variáveis para saída de célula
    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 80333

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 80332
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 80333
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143636)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set objGrid = Nothing

    Set gobjBorderoChequePre = Nothing

End Sub

Private Sub GridBorderoPag2_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Private Sub GridBorderoPag2_GotFocus()

    Call Grid_Recebe_Foco(objGrid)

End Sub

Private Sub GridBorderoPag2_EnterCell()

    Call Grid_Entrada_Celula(objGrid, iAlterado)

End Sub

Private Sub GridBorderoPag2_LeaveCell()

    Call Saida_Celula(objGrid)

End Sub

Private Sub GridBorderoPag2_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGrid)

End Sub

Private Sub GridBorderoPag2_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Private Sub GridBorderoPag2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Faz com que apareca um PopupMenu o botao direito do mouse acionado sobre o grid

'Comentado por Luiz Nogueira em 22/04/04. Não existe código pra abrir a tela original
'    'Verifica se foi o botao direito do mouse que foi pressionado
'    If Button = vbRightButton Then
'
'        'Seta objTela como a Tela de Baixas a Receber
'        Set PopUpMenuGrid.objTela = Me
'
'        'Chama o Menu PopUp
'        PopupMenu PopUpMenuGrid.mnuGrid, vbPopupMenuRightButton
'
'        'Limpa o objTela
'        Set PopUpMenuGrid.objTela = Nothing
'
'    End If

End Sub

Private Sub GridBorderoPag2_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGrid)

End Sub

Private Sub GridBorderoPag2_RowColChange()

    Call Grid_RowColChange(objGrid)

End Sub

Private Sub GridBorderoPag2_Scroll()

    Call Grid_Scroll(objGrid)

End Sub

Private Function Inicializa_Grid_BorderoPag2(objGridInt As AdmGrid, iRegistros As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Inicializa_Grid_BorderoPag2

    'Tela em questão
    Set objGrid.objForm = Me

    'Títulos do grid
    objGridInt.colColuna.Add (" ")

    If giFilialEmpresa = EMPRESA_TODA Then
        objGridInt.colColuna.Add ("Filial Empresa")
    End If
    
    FilialEmpresa.Visible = False
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Filial")
'    objGridInt.colColuna.Add ("No.do Título")
'    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("Banco")
    objGridInt.colColuna.Add ("Agência")
    objGridInt.colColuna.Add ("Conta Corrente")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Data de Depósito")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Depositar")
    

    If giFilialEmpresa = EMPRESA_TODA Then
        'campos de edição do grid
        objGridInt.colCampo.Add (FilialEmpresa.Name)
    End If
    objGridInt.colCampo.Add (Cliente.Name)
    objGridInt.colCampo.Add (Filial.Name)
'    objGridInt.colCampo.Add (NumTitulo.Name)
'    objGridInt.colCampo.Add (Parcela.Name)
    objGridInt.colCampo.Add (Banco.Name)
    objGridInt.colCampo.Add (Agencia.Name)
    objGridInt.colCampo.Add (ContaCorrente.Name)
    objGridInt.colCampo.Add (Numero.Name)
    objGridInt.colCampo.Add (DataDeposito.Name)
    objGridInt.colCampo.Add (Valor.Name)
    objGridInt.colCampo.Add (CheckPago.Name)

    If giFilialEmpresa = EMPRESA_TODA Then
        iGrid_FilialEmpresa_Col = 1
        iGrid_Cliente_Col = 2
        iGrid_Filial_Col = 3
'        iGrid_NumTitulo_Col = 4
'        iGrid_Parcela_Col = 5
        iGrid_Banco_Col = 4
        iGrid_Agencia_Col = 5
        iGrid_ContaCorrente_Col = 6
        iGrid_Numero_Col = 7
        iGrid_DataDeposito_Col = 8
        iGrid_Valor_Col = 9
        iGrid_CheckPago_Col = 10

    Else
        iGrid_Cliente_Col = 1
        iGrid_Filial_Col = 2
'        iGrid_NumTitulo_Col = 3
'        iGrid_Parcela_Col = 4
        iGrid_Banco_Col = 3
        iGrid_Agencia_Col = 4
        iGrid_ContaCorrente_Col = 5
        iGrid_Numero_Col = 6
        iGrid_DataDeposito_Col = 7
        iGrid_Valor_Col = 8
        iGrid_CheckPago_Col = 9

    End If

    objGridInt.objGrid = GridBorderoPag2

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 9

    'Todas as linhas do grid
    If objGridInt.iLinhasVisiveis >= iRegistros + 1 Then
        objGridInt.objGrid.Rows = objGridInt.iLinhasVisiveis + 1
    Else
        objGridInt.objGrid.Rows = iRegistros + 1
    End If

    GridBorderoPag2.ColWidth(0) = 400

    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iProibidoIncluir = 1
    objGridInt.iProibidoExcluir = 1

    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_BorderoPag2 = SUCESSO

    Exit Function

Erro_Inicializa_Grid_BorderoPag2:

    Inicializa_Grid_BorderoPag2 = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143637)

    End Select

    Exit Function

End Function

Function Trata_Parametros(Optional objBorderoChequePre As ClassBorderoChequePre) As Long
'Traz os dados dos Cheques a pagar para a Tela

Dim objChequePre As ClassChequePre
Dim iLinha As Integer, lErro As Long
Dim dTotal As Double
Dim dValorSelecionado As Double
Dim lQuantSelecionado As Long
Dim objFilialEmpresa As New AdmFiliais
Dim objCliente As New ClassCliente
Dim sNomeFilial As String
Dim objFilialCliente As New ClassFilialCliente

On Error GoTo Erro_Trata_Parametros

    Set gobjBorderoChequePre = objBorderoChequePre

    'Passa para a tela os dados dos cheques selecionados
    QtdChequesSelecionados.Caption = CStr(gobjBorderoChequePre.iQuantChequesSel)
    TotalChequesSelecionados.Caption = CStr(Format(gobjBorderoChequePre.dValorChequesSelecionados, "Standard"))

    Set objGrid = New AdmGrid

    lErro = Inicializa_Grid_BorderoPag2(objGrid, gobjBorderoChequePre.colchequepre.Count)
    If lErro <> SUCESSO Then gError 80334

    iLinha = 0

    'Percorre todos os cheques da Coleção passada por parâmetro
    For Each objChequePre In gobjBorderoChequePre.colchequepre
        
        iLinha = iLinha + 1

        Set objCliente = New ClassCliente
        
        objFilialEmpresa.iCodFilial = giFilialEmpresa
        objCliente.lCodigo = objChequePre.lCliente

        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO Then gError 80335

        If objCliente.lCodigo <> 0 Then
            lErro = CF("Cliente_Le", objCliente)
            If lErro <> SUCESSO Then gError 80339
                        
            objFilialCliente.iCodFilial = objChequePre.iFilial
            objFilialCliente.lCodCliente = objCliente.lCodigo
            
            lErro = CF("FilialCliente_Le", objFilialCliente)
            If lErro <> SUCESSO Then gError 118005
            
            sNomeFilial = CStr(objChequePre.iFilial) & SEPARADOR & CStr(objFilialCliente.sNome)
        Else
        
            objCliente.sNomeReduzido = "Não Especificado"
            sNomeFilial = "Não Especificado"
        End If
        
        'Passa para a tela os dados do cheque em questão
        
        If giFilialEmpresa = EMPRESA_TODA Then
        
            GridBorderoPag2.TextMatrix(iLinha, iGrid_FilialEmpresa_Col) = objChequePre.iFilialEmpresa & SEPARADOR & objFilialEmpresa.sNome
            
        End If
        
            GridBorderoPag2.TextMatrix(iLinha, iGrid_Cliente_Col) = objCliente.sNomeReduzido
            GridBorderoPag2.TextMatrix(iLinha, iGrid_Filial_Col) = sNomeFilial
'            GridBorderoPag2.TextMatrix(iLinha, iGrid_NumTitulo_Col) = objChequePre.lNumTitulo
'            GridBorderoPag2.TextMatrix(iLinha, iGrid_Parcela_Col) = objChequePre.iNumParcela
            GridBorderoPag2.TextMatrix(iLinha, iGrid_Banco_Col) = objChequePre.iBanco
            GridBorderoPag2.TextMatrix(iLinha, iGrid_Agencia_Col) = objChequePre.sAgencia
            GridBorderoPag2.TextMatrix(iLinha, iGrid_ContaCorrente_Col) = objChequePre.sContaCorrente
            GridBorderoPag2.TextMatrix(iLinha, iGrid_Numero_Col) = objChequePre.lNumero
            GridBorderoPag2.TextMatrix(iLinha, iGrid_DataDeposito_Col) = Format(objChequePre.dtDataDeposito, "dd/mm/yyyy")
            GridBorderoPag2.TextMatrix(iLinha, iGrid_Valor_Col) = Format(objChequePre.dValor, "Standard")
            GridBorderoPag2.TextMatrix(iLinha, iGrid_CheckPago_Col) = objChequePre.iChequeSel

        If objChequePre.iChequeSel = 1 Then
            lQuantSelecionado = lQuantSelecionado + 1
            dValorSelecionado = dValorSelecionado + objChequePre.dValor
        End If

        'Soma ao total o valor do cheque em questão
        dTotal = dTotal + objChequePre.dValor

    Next

    'Passa para o Obj o número de cheques passados pela Coleção
    objGrid.iLinhasExistentes = iLinha

    'Passa para a tela o somatório da Qtd de cheques e do Número total de cheques
    QtdCheques.Caption = CStr(objGrid.iLinhasExistentes)
    TotalCheques.Caption = CStr(Format(dTotal, "Standard"))
    TotalChequesSelecionados = Format(dValorSelecionado, "Standard")
    QtdChequesSelecionados = CStr(lQuantSelecionado)

    'Atualiza na tela os checkbox marcados
    Call Grid_Refresh_Checkbox(objGrid)

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 80334, 80335, 80339, 118005

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143638)

    End Select

    Exit Function

End Function

'Public Sub mnuGridConsultaDocOriginal_Click()
''Chama a tela de consulta de Títulos a Receber quando essa opção for selecionada no grid
'
'Dim lErro As Long
'Dim objTituloReceber As New ClassTituloReceber
'Dim objParcRec As New ClassParcelaReceber
'Dim objChequePre As New ClassChequePre
'
'On Error GoTo Erro_mnuGridConsultaDocOriginal_Click
'
'    'Se não foi selecionada uma linha do grid => erro
'    If GridBorderoPag2.Row <= 0 Then gError 79935
'
'    'Se a linha selecionada não contém dados => erro
'    If Len(Trim(GridBorderoPag2.TextMatrix(GridBorderoPag2.Row, iGrid_Numero_Col))) <= 0 Then gError 79936
'
'    'Seta o objChequePre com os dados do cheque da liha selecionada
'    Set objChequePre = gobjBorderoChequePre.colchequepre.Item(GridBorderoPag2.Row)
'
'    'Transfere para objTitulo receber os dados necessários para a leitura do título
'    objParcRec.lNumIntCheque = objChequePre.lNumIntCheque
'
'    'Lê no BD o NumIntDoc do título que contém a parcela que está vinculada ao cheque pré
'    lErro = CF("ChequePre_Obtem_TituloReceber",objParcRec, objTituloReceber)
'    If lErro <> SUCESSO And lErro <> 79939 Then gError 79942
'
'    'Se não encontrou o título => erro
'    If lErro = 79939 Then gError 79941
'
'    Call Chama_Tela("TituloReceber_Consulta", objTituloReceber)
'
'    Exit Sub
'
'Erro_mnuGridConsultaDocOriginal_Click:
'
'    Select Case gErr
'
'        Case 79935
'            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
'
'        Case 79936
'            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_PREENCHIDA", gErr)
'
'        Case 79942
'
'        Case 79941
'            Call Rotina_Erro(vbOKOnly, "ERRO_TITULOREC_CHEQUEPRE_NAO_ENCONTRADO", gErr, objChequePre.lCliente, objChequePre.lNumero, objChequePre.dValor)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143639)
'
'    End Select
'
'    Exit Sub
'
'End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BORDERO_PAGT_P2
    Set Form_Load_Ocx = Me
    Caption = "Bordero de Cheques Pré - Passo 2"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "BorderoChequesPre1Ocx"

End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
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

Private Sub TotalChequesSelecionados_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalChequesSelecionados, Source, X, Y)
End Sub

Private Sub TotalChequesSelecionados_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalChequesSelecionados, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub QtdChequesSelecionados_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QtdChequesSelecionados, Source, X, Y)
End Sub

Private Sub QtdChequesSelecionados_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QtdChequesSelecionados, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub QtdCheques_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QtdCheques, Source, X, Y)
End Sub

Private Sub QtdCheques_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QtdCheques, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub TotalCheques_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalCheques, Source, X, Y)
End Sub

Private Sub TotalCheques_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalCheques, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub


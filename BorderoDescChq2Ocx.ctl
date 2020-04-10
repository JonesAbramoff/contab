VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl BorderoDescChq2Ocx 
   ClientHeight    =   5370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8145
   LockControls    =   -1  'True
   ScaleHeight     =   5370
   ScaleWidth      =   8145
   Begin MSMask.MaskEdBox FilialEmpresa 
      Height          =   225
      Left            =   4155
      TabIndex        =   25
      Top             =   1305
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
   Begin MSMask.MaskEdBox Valor 
      Height          =   225
      Left            =   5460
      TabIndex        =   24
      Top             =   870
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
   Begin VB.PictureBox Picture7 
      Height          =   555
      Left            =   5460
      ScaleHeight     =   495
      ScaleWidth      =   2490
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   75
      Width           =   2550
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   2040
         Picture         =   "BorderoDescChq2Ocx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoVoltar 
         Height          =   345
         Left            =   60
         Picture         =   "BorderoDescChq2Ocx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   90
         Width           =   885
      End
      Begin VB.CommandButton BotaoSeguir 
         Height          =   345
         Left            =   1020
         Picture         =   "BorderoDescChq2Ocx.ctx":08DC
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   90
         Width           =   930
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selecionado"
      Height          =   1005
      Left            =   2445
      TabIndex        =   20
      Top             =   4275
      Width           =   2205
      Begin VB.Label TotalChequesSelecionados 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   765
         TabIndex        =   8
         Top             =   570
         Width           =   1335
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
         TabIndex        =   22
         Top             =   585
         Width           =   510
      End
      Begin VB.Label QtdChequesSelecionados 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   765
         TabIndex        =   7
         Top             =   225
         Width           =   780
      End
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
         TabIndex        =   21
         Top             =   255
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Total"
      Height          =   1005
      Left            =   75
      TabIndex        =   17
      Top             =   4275
      Width           =   2220
      Begin VB.Label QtdCheques 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   735
         TabIndex        =   5
         Top             =   255
         Width           =   780
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
         TabIndex        =   19
         Top             =   270
         Width           =   540
      End
      Begin VB.Label TotalCheques 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   735
         TabIndex        =   6
         Top             =   585
         Width           =   1335
      End
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
         TabIndex        =   18
         Top             =   615
         Width           =   510
      End
   End
   Begin VB.CommandButton BotaoMarcar 
      Caption         =   "Marcar Todas"
      Height          =   555
      Left            =   495
      Picture         =   "BorderoDescChq2Ocx.ctx":106E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   105
      Width           =   1530
   End
   Begin VB.CommandButton BotaoDesmarcar 
      Caption         =   "Desmarcar Todas"
      Height          =   555
      Left            =   2205
      Picture         =   "BorderoDescChq2Ocx.ctx":2088
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   105
      Width           =   1530
   End
   Begin VB.TextBox Banco 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   1425
      TabIndex        =   16
      Top             =   870
      Width           =   1065
   End
   Begin VB.TextBox Agencia 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   2520
      TabIndex        =   15
      Top             =   870
      Width           =   885
   End
   Begin VB.TextBox ContaCorrente 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   3420
      TabIndex        =   14
      Top             =   885
      Width           =   1020
   End
   Begin VB.TextBox Numero 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   4485
      TabIndex        =   13
      Top             =   900
      Width           =   945
   End
   Begin VB.TextBox DataDeposito 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   6585
      TabIndex        =   2
      Top             =   870
      Width           =   1035
   End
   Begin VB.CheckBox CheckIncluir 
      Height          =   240
      Left            =   420
      TabIndex        =   4
      Top             =   825
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid GridCheque 
      Height          =   3330
      Left            =   90
      TabIndex        =   3
      Top             =   795
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
   Begin MSMask.MaskEdBox ValorCredito 
      Height          =   300
      Left            =   6195
      TabIndex        =   9
      Top             =   4635
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   15
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Valor a Creditar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   4740
      TabIndex        =   26
      Top             =   4665
      Width           =   1395
   End
End
Attribute VB_Name = "BorderoDescChq2Ocx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_Caption As String
Event Unload()

Dim objGrid As AdmGrid
Dim gobjBorderoDescChq As ClassBorderoDescChq
Public iAlterado As Integer

'colunas do grid
Dim iGrid_FilialEmpresa_Col As Integer
Dim iGrid_CheckIncluir_Col As Integer
Dim iGrid_Banco_Col As Integer
Dim iGrid_Agencia_Col As Integer
Dim iGrid_ContaCorrente_Col As Integer
Dim iGrid_Numero_Col As Integer
Dim iGrid_DataDeposito_Col As Integer
Dim iGrid_Valor_Col As Integer

Public Sub Form_Load()

On Error GoTo Erro_Form_Load

    'instancia os obj globais
    Set objGrid = New AdmGrid
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub
    
Erro_Form_Load:
    
    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143720)
    
    End Select
    
    Exit Sub

End Sub

Public Function Trata_Parametros(Optional objBorderoDescChq As ClassBorderoDescChq) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If objBorderoDescChq Is Nothing Then gError 109236
    
    'inicializa o grid
    lErro = Inicializa_GridCheque(objGrid, objBorderoDescChq.colchequepre)
    If lErro <> SUCESSO Then gError 109189
    
    'traz os dados do bordero para a tela
    Call Traz_BorderoDescChq_Tela(objBorderoDescChq)
    
    'faz o objeto global dessa tela apontar para o obj recebido como parâmetro
    Set gobjBorderoDescChq = objBorderoDescChq

    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:
    
    Trata_Parametros = gErr
    
    Select Case gErr
        
        Case 109189
        
        Case 109236
            Call Rotina_Erro(vbOKOnly, "ERRO_OBJBORDERODESCCHQ_NAO_CRIADO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143721)
    
    End Select
    
    Exit Function

End Function

Private Sub CheckIncluir_Click()

Dim objChequePre As ClassChequePre
        
    'faz o objCheque Pre apontar para o elemento da coleção referente à linha do grid
    Set objChequePre = gobjBorderoDescChq.colchequepre.Item(GridCheque.Row)
    
    'marca ou desmarca o cheque na coleção
    objChequePre.iChequeSel = CInt(GridCheque.TextMatrix(GridCheque.Row, iGrid_CheckIncluir_Col))
    
    'se foi marcado
    If objChequePre.iChequeSel = MARCADO Then
    
        'incrementa a quantidade e adiciona o cheque ao total
        QtdChequesSelecionados.Caption = CStr(StrParaInt(QtdChequesSelecionados.Caption) + 1)
        TotalChequesSelecionados.Caption = Format(StrParaDbl(TotalChequesSelecionados.Caption) + objChequePre.dValor, "STANDARD")
        
    'se foi desmarcado
    Else
    
        'decrementa a quantidade e subtrai o cheque do total
        QtdChequesSelecionados.Caption = CStr(StrParaInt(QtdChequesSelecionados.Caption) - 1)
        TotalChequesSelecionados.Caption = Format(StrParaDbl(TotalChequesSelecionados.Caption) - objChequePre.dValor, "STANDARD")
        
    End If

End Sub

Private Sub CheckIncluir_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGrid)
End Sub

Private Sub CheckIncluir_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
End Sub

Private Sub CheckIncluir_Validate(Cancel As Boolean)
    
    Set objGrid.objControle = CheckIncluir
    Call Grid_Libera_Foco(objGrid)

End Sub

Private Sub BotaoDesmarcar_Click()
'Desmarca todos os cheques marcados no Grid

Dim iLinha As Integer
Dim objChequePre As New ClassChequePre

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGrid.iLinhasExistentes

        'Desmarca na tela o cheque em questão
        GridCheque.TextMatrix(iLinha, iGrid_CheckIncluir_Col) = DESMARCADO

        'Passa a linha do Grid para o Obj
        Set objChequePre = gobjBorderoDescChq.colchequepre.Item(iLinha)

        'Desmarca no Obj o cheque em questão
        objChequePre.iChequeSel = 0

    Next

    'Atualiza na tela os checkbox desmarcados
    Call Grid_Refresh_Checkbox(objGrid)

    'Limpa na tela os campos Qtd de cheques selecionados e Valor total dos cheques selecionados
    QtdChequesSelecionados.Caption = CStr(0)
    TotalChequesSelecionados.Caption = CStr(Format(0, "Standard"))

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
        GridCheque.TextMatrix(iLinha, iGrid_CheckIncluir_Col) = MARCADO

        'Passa a linha do Grid para o Obj
        Set objChequePre = gobjBorderoDescChq.colchequepre.Item(iLinha)

        'Marca no Obj o cheque em questão
        objChequePre.iChequeSel = 1

        dTotalChequesSelecionados = dTotalChequesSelecionados + CDbl(GridCheque.TextMatrix(iLinha, iGrid_Valor_Col))
        iNumChequesSelecionados = iNumChequesSelecionados + 1

    Next

    'Atualiza na tela os checkbox marcados
    Call Grid_Refresh_Checkbox(objGrid)

    'Atualiza na tela os campos Qtd de cheques selecionados e Valor total dos cheques selecionados
    QtdChequesSelecionados.Caption = CStr(iNumChequesSelecionados)
    TotalChequesSelecionados.Caption = CStr(Format(dTotalChequesSelecionados, "Standard"))

End Sub

Private Sub BotaoVoltar_Click()

On Error GoTo Erro_BotaoVoltar_Click

    'Chama a tela do passo anterior
    Call Chama_Tela("BorderoDescChq1", gobjBorderoDescChq)

    'Fecha a tela
    Unload Me

    Exit Sub

Erro_BotaoVoltar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143722)

    End Select

    Exit Sub

End Sub

Private Sub BotaoSeguir_Click()

On Error GoTo Erro_BotaoSeguir_Click

    'se não houver nenhum cheque selecionado-> erro
    If StrParaInt(QtdChequesSelecionados.Caption) = 0 Then gError 109231
    
    'guarda a quantidade e o total de cheques selecionados
    gobjBorderoDescChq.iQuantChequesSel = StrParaInt(QtdChequesSelecionados.Caption)
    gobjBorderoDescChq.dValorChequesSel = StrParaDbl(TotalChequesSelecionados.Caption)
    gobjBorderoDescChq.iFilialEmpresa = giFilialEmpresa
    
    'se valor de crédito = 0-> erro
    If Len(Trim(ValorCredito.Text)) = 0 Then gError 109232
    
    'se o total selecionado for menor que o total de crédito-> erro
    If StrParaDbl(TotalChequesSelecionados.Caption) < StrParaDbl(ValorCredito.Text) Then gError 109233
    
    'guarda o valor de crédito
    gobjBorderoDescChq.dValorCredito = StrParaDbl(ValorCredito.Text)
    
    If gobjBorderoDescChq.dValorCredito = 0 Then gError 126240
    
    'chama a próxima tela
    Call Chama_Tela("BorderoDescChq3", gobjBorderoDescChq)
    
    Unload Me

    Exit Sub
    
Erro_BotaoSeguir_Click:
    
    Select Case gErr
    
        Case 109231
            Call Rotina_Erro(vbOKOnly, "ERRO_SEM_CHEQUE_SELECIONADO", gErr)
        
        Case 109232, 126240
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_CREDITADO_NAO_INFORMADO", gErr)
        
        Case 109233
            Call Rotina_Erro(vbOKOnly, "ERRO_TOTAL_CHEQUESPRE_MENOR_TOTAL_CREDITAR", gErr, TotalChequesSelecionados.Caption, ValorCredito.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143723)

    End Select

    Exit Sub

End Sub

Private Sub ValorCredito_GotFocus()
    Call MaskEdBox_TrataGotFocus(ValorCredito, iAlterado)
End Sub

Private Sub ValorCredito_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorCredito_Validate

    If Len(Trim(ValorCredito.Text)) <> 0 Then
    
        lErro = Valor_Positivo_Critica(ValorCredito.Text)
        If lErro <> SUCESSO Then gError 109230
        
    End If

    Exit Sub
    
Erro_ValorCredito_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 109230

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143724)

    End Select
    
    Exit Sub
    
End Sub

Private Sub ValorCredito_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub GridCheque_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Private Sub GridCheque_GotFocus()
    Call Grid_Recebe_Foco(objGrid)
End Sub

Private Sub GridCheque_EnterCell()
    Call Grid_Entrada_Celula(objGrid, iAlterado)
End Sub

Private Sub GridCheque_LeaveCell()
    Call Saida_Celula(objGrid)
End Sub

Private Sub GridCheque_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGrid)
End Sub

Private Sub GridCheque_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Private Sub GridCheque_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Faz com que apareca um PopupMenu o botao direito do mouse acionado sobre o grid

    'Verifica se foi o botao direito do mouse que foi pressionado
    If Button = vbRightButton Then
        
        'Seta objTela como a Tela de Baixas a Receber
        Set PopUpMenuGrid.objTela = Me
        
        'Chama o Menu PopUp
        PopupMenu PopUpMenuGrid.mnuGrid, vbPopupMenuRightButton
        
        'Limpa o objTela
        Set PopUpMenuGrid.objTela = Nothing
        
    End If

End Sub

Public Sub mnuGridConsultaDocOriginal_Click()

Dim objCheque As New ClassChequePre

On Error GoTo Erro_mnuGridConsultaDocOriginal_Click

    If GridCheque.Row Then
        objCheque.lNumIntCheque = gobjBorderoDescChq.colchequepre.Item(GridCheque.Row).lNumIntCheque
        objCheque.iFilialEmpresa = gobjBorderoDescChq.colchequepre.Item(GridCheque.Row).iFilialEmpresa
        
        Call Chama_Tela("ChequePre", objCheque)
        
    End If

    Exit Sub

Erro_mnuGridConsultaDocOriginal_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143725)

    End Select

    Exit Sub

End Sub

Private Sub GridCheque_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGrid)

End Sub

Private Sub GridCheque_RowColChange()

    Call Grid_RowColChange(objGrid)

End Sub

Private Sub GridCheque_Scroll()

    Call Grid_Scroll(objGrid)

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objGrid = Nothing
    Set gobjBorderoDescChq = Nothing

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente.

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    'Inicializa variáveis para saída de célula
    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then gError 109191

    'Finaliza variáveis para saída de célula
    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 109192

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 109191
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 109192
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143725)

    End Select

    Exit Function

End Function

Private Sub Limpa_Tela_BorderoDescChq()

    QtdCheques.Caption = ""
    TotalCheques.Caption = ""
    QtdChequesSelecionados.Caption = ""
    TotalChequesSelecionados.Caption = ""

    Call Limpa_Tela(Me)

End Sub

Private Function Traz_BorderoDescChq_Tela(ByVal objBorderoDescChq As ClassBorderoDescChq) As Long

Dim objChequePre As ClassChequePre
Dim iIndice As Integer
Dim dTotalCheques As Double
Dim dTotalChequesSelecionados As Double
Dim iQtdChequesSelecionados As Integer
Dim colFiliaisEmpresa As New Collection
Dim objFilialEmpresa As AdmFiliais
Dim lErro As Long

On Error GoTo Erro_Traz_BorderoDescChq_Tela

    Call Limpa_Tela_BorderoDescChq
    
    'se for empresa toda, carrega a coleção de filiais dessa empresa
    If giFilialEmpresa = EMPRESA_TODA Then
        
        lErro = CF("FiliaisEmpresas_Le_Empresa", glEmpresa, colFiliaisEmpresa)
        If lErro <> SUCESSO Then gError 109190
    
    End If

    'preenche o grid com os dados de cada cheque pre da coleção
    For Each objChequePre In objBorderoDescChq.colchequepre
    
        objGrid.iLinhasExistentes = objGrid.iLinhasExistentes + 1
        
        'se for empresa toda, preencher a filial também
        If giFilialEmpresa = EMPRESA_TODA Then
        
            'busca o nome da filial na coleção de filiais através do código
            For Each objFilialEmpresa In colFiliaisEmpresa
                
                'se achou
                If objFilialEmpresa.iCodFilial = objChequePre.iFilialEmpresa Then
                    
                    'preenche
                    GridCheque.TextMatrix(objGrid.iLinhasExistentes, iGrid_FilialEmpresa_Col) = objChequePre.iFilialEmpresa & SEPARADOR & objFilialEmpresa.sNome
                    Exit For
                
                End If
            
            Next
        
        End If
        
        'resto das colunas a serem preenchidas
        GridCheque.TextMatrix(objGrid.iLinhasExistentes, iGrid_Agencia_Col) = objChequePre.sAgencia
        GridCheque.TextMatrix(objGrid.iLinhasExistentes, iGrid_Banco_Col) = objChequePre.iBanco
        GridCheque.TextMatrix(objGrid.iLinhasExistentes, iGrid_ContaCorrente_Col) = objChequePre.sContaCorrente
        GridCheque.TextMatrix(objGrid.iLinhasExistentes, iGrid_DataDeposito_Col) = Format(objChequePre.dtDataDeposito, "dd/mm/yyyy")
        GridCheque.TextMatrix(objGrid.iLinhasExistentes, iGrid_Numero_Col) = objChequePre.lNumero
        GridCheque.TextMatrix(objGrid.iLinhasExistentes, iGrid_Valor_Col) = Format(objChequePre.dValor, "STANDARD")
        GridCheque.TextMatrix(objGrid.iLinhasExistentes, iGrid_CheckIncluir_Col) = objChequePre.iChequeSel
        
        'se o cheque estiver selecionado,
        If objChequePre.iChequeSel = MARCADO Then
            
            'acumula o seu total e quantidade
            dTotalChequesSelecionados = dTotalChequesSelecionados + objChequePre.dValor
            iQtdChequesSelecionados = iQtdChequesSelecionados + 1
        
        End If
        
        'acumula o somatório dos cheques
        dTotalCheques = dTotalCheques + objChequePre.dValor
    
    Next
    
    'preenche os totalizadores
    QtdCheques.Caption = CStr(objBorderoDescChq.colchequepre.Count)
    QtdChequesSelecionados.Caption = CStr(iQtdChequesSelecionados)
    TotalCheques.Caption = Format(dTotalCheques, "STANDARD")
    TotalChequesSelecionados.Caption = Format(dTotalChequesSelecionados, "STANDARD")
    ValorCredito.Text = Format(objBorderoDescChq.dValorCredito, "STANDARD")
    
    Call Grid_Refresh_Checkbox(objGrid)
    
    Traz_BorderoDescChq_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_BorderoDescChq_Tela:
    
    Traz_BorderoDescChq_Tela = gErr
    
    Select Case gErr
    
        Case 109190
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143726)
            
    End Select
    
    Exit Function

End Function

Private Function Inicializa_GridCheque(ByVal objGrid As AdmGrid, ByVal colchequepre As Collection) As Long

On Error GoTo Erro_Inicializa_GridCheque

    'associa o grid à tela
    Set objGrid.objForm = Me
    
    'associa o gridint ao grid
    objGrid.objGrid = GridCheque
    
    'título da 1ª coluna
    objGrid.colColuna.Add (" ")
    
    'se estiver acessando a empresa_toda
    If giFilialEmpresa = EMPRESA_TODA Then
        
        'colocar largura manual
        objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
        
        'títulos extra
        objGrid.colColuna.Add ("Filial Empresa")
        objGrid.colCampo.Add (FilialEmpresa)
        
        'colunas numeradas para o caso da empresa_toda
        iGrid_FilialEmpresa_Col = 1
        iGrid_CheckIncluir_Col = 2
        iGrid_Banco_Col = 3
        iGrid_Agencia_Col = 4
        iGrid_ContaCorrente_Col = 5
        iGrid_Numero_Col = 6
        iGrid_DataDeposito_Col = 7
        iGrid_Valor_Col = 8
    
    Else
        
        'colocar largura automática
        objGrid.iGridLargAuto = GRID_LARGURA_AUTOMATICA
        
        'colunas numeradas para o caso de uma filial específica
        iGrid_CheckIncluir_Col = 1
        iGrid_Banco_Col = 2
        iGrid_Agencia_Col = 3
        iGrid_ContaCorrente_Col = 4
        iGrid_Numero_Col = 5
        iGrid_DataDeposito_Col = 6
        iGrid_Valor_Col = 7
        FilialEmpresa.Visible = False

    End If
    
    'títulos das colunas
    objGrid.colColuna.Add ("Selecionado")
    objGrid.colColuna.Add ("Banco")
    objGrid.colColuna.Add ("Agência")
    objGrid.colColuna.Add ("Conta")
    objGrid.colColuna.Add ("Número")
    objGrid.colColuna.Add ("Bom Para")
    objGrid.colColuna.Add ("Valor")
    
    'campos associados às colunas
    objGrid.colCampo.Add (CheckIncluir.Name)
    objGrid.colCampo.Add (Banco.Name)
    objGrid.colCampo.Add (Agencia.Name)
    objGrid.colCampo.Add (ContaCorrente.Name)
    objGrid.colCampo.Add (Numero.Name)
    objGrid.colCampo.Add (DataDeposito.Name)
    objGrid.colCampo.Add (Valor.Name)
    
    'Linhas visíveis
    objGrid.iLinhasVisiveis = 9
    
    'se tiver menos cheques na coleção do que linhas no grid, setar a quantidade de linhas do grid conforme a quantidade de linhas visíveis
    If objGrid.iLinhasVisiveis >= colchequepre.Count Then
        GridCheque.Rows = objGrid.iLinhasVisiveis + 1
    
    'caso contrário, setar a quantidade de linhas do grid em função da quantidade de cheques na coleção
    Else
        GridCheque.Rows = colchequepre.Count + 1
    
    End If
    
    'acerta a largura da primeira coluna
    GridCheque.ColWidth(0) = 400
    
    'não deixa incluir linhas no grid
    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    'não deixa excluir linhas no grid
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    
    'chama a inicialização do grid
    Call Grid_Inicializa(objGrid)
    
    Inicializa_GridCheque = SUCESSO
    
    Exit Function
    
Erro_Inicializa_GridCheque:

    Inicializa_GridCheque = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143727)
    
    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    ' ???Parent.HelpContextID = IDH_BORDERO_DESCCHQ1
    Set Form_Load_Ocx = Me
    Caption = "Borderô de Desconto de Cheques - Passo 2"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "BorderoDescChq2"
    
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

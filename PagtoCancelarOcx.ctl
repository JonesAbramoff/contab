VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl PagtoCancelarOcx 
   ClientHeight    =   4530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8775
   KeyPreview      =   -1  'True
   ScaleHeight     =   4530
   ScaleWidth      =   8775
   Begin VB.CommandButton BotaoConsultaTitPag 
      Height          =   585
      Left            =   6660
      Picture         =   "PagtoCancelarOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   1530
      Width           =   1965
   End
   Begin VB.PictureBox Picture4 
      Height          =   555
      Left            =   6900
      ScaleHeight     =   495
      ScaleWidth      =   1635
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   180
      Width           =   1695
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1125
         Picture         =   "PagtoCancelarOcx.ctx":2F16
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "PagtoCancelarOcx.ctx":3094
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "PagtoCancelarOcx.ctx":35C6
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parcelas pagas"
      Height          =   2220
      Left            =   150
      TabIndex        =   20
      Top             =   2220
      Width           =   8550
      Begin MSFlexGridLib.MSFlexGrid GridParcelas 
         Height          =   1605
         Left            =   90
         TabIndex        =   14
         Top             =   270
         Width           =   8340
         _ExtentX        =   14711
         _ExtentY        =   2831
         _Version        =   393216
         Rows            =   5
         Cols            =   9
      End
      Begin MSMask.MaskEdBox Fornecedor 
         Height          =   225
         Left            =   210
         TabIndex        =   6
         Top             =   225
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         AllowPrompt     =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   20
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
      Begin MSMask.MaskEdBox DataVencimento 
         Height          =   225
         Left            =   4950
         TabIndex        =   11
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox ValorParcela 
         Height          =   225
         Left            =   6075
         TabIndex        =   12
         Top             =   225
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Parcela 
         Height          =   225
         Left            =   3690
         TabIndex        =   10
         Top             =   210
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         AllowPrompt     =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "99"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NumeroDoc 
         Height          =   225
         Left            =   2850
         TabIndex        =   9
         Top             =   210
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         AllowPrompt     =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "999999"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox SiglaDocumento 
         Height          =   225
         Left            =   2220
         TabIndex        =   8
         Top             =   210
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         AllowPrompt     =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   4
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
      Begin MSMask.MaskEdBox ValorPago 
         Height          =   225
         Left            =   7065
         TabIndex        =   13
         Top             =   240
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Filial 
         Height          =   225
         Left            =   1470
         TabIndex        =   7
         Top             =   210
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         AllowPrompt     =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "99"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataEmissao 
         Height          =   225
         Left            =   4410
         TabIndex        =   31
         Top             =   210
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dados do Pagamento"
      Height          =   1365
      Left            =   105
      TabIndex        =   19
      Top             =   75
      Width           =   6390
      Begin VB.ComboBox ContaCorrente 
         Height          =   315
         Left            =   915
         TabIndex        =   0
         Top             =   323
         Width           =   1695
      End
      Begin VB.ComboBox TipoMeioPagto 
         Height          =   315
         Left            =   915
         TabIndex        =   2
         Top             =   855
         Width           =   1710
      End
      Begin MSComCtl2.UpDown UpDownDataMovimento 
         Height          =   300
         Left            =   4920
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   330
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataMovimento 
         Height          =   300
         Left            =   3825
         TabIndex        =   1
         Top             =   330
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.CommandButton BotaoTrazer 
         Caption         =   "Trazer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5145
         TabIndex        =   5
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton BotaoSelecionar 
         Caption         =   "Selecionar ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3030
         TabIndex        =   4
         Top             =   840
         Width           =   1860
      End
      Begin MSMask.MaskEdBox Numero 
         Height          =   300
         Left            =   3840
         TabIndex        =   3
         Top             =   855
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
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
         Left            =   3285
         TabIndex        =   23
         Top             =   390
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
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
         Left            =   390
         TabIndex        =   24
         Top             =   915
         Width           =   450
      End
      Begin VB.Label NumeroLabel 
         Caption         =   "Número:"
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
         Height          =   255
         Left            =   3060
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   25
         Top             =   885
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Conta:"
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
         Left            =   255
         TabIndex        =   26
         Top             =   383
         Width           =   570
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Atributos"
      Height          =   570
      Left            =   105
      TabIndex        =   21
      Top             =   1530
      Width           =   6390
      Begin VB.Label Valor 
         Height          =   225
         Left            =   975
         TabIndex        =   27
         Top             =   225
         Width           =   1275
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   420
         TabIndex        =   28
         Top             =   210
         Width           =   495
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Beneficiario"
      Height          =   15
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "PagtoCancelarOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Dim objGridParcelas As AdmGrid
Dim iGrid_Fornecedor_Col As Integer
Dim iGrid_Filial_Col As Integer
Dim iGrid_SiglaDocumento_Col As Integer
Dim iGrid_NumeroDoc_Col As Integer
Dim iGrid_Parcela_Col As Integer
Dim iGrid_DataVencimento_Col As Integer
Dim iGrid_DataEmissao_Col As Integer
Dim iGrid_ValorParcela_Col As Integer
Dim iGrid_ValorPago_Col As Integer
Dim glSequencial As Long

Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1

Private Sub BotaoConsultaTitPag_Click()

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objTituloPagar As New ClassTituloPagar

On Error GoTo Erro_BotaoConsultaTitPag_Click

    'Verifica se uma linha do Grid foi selecionada
    If GridParcelas.Row <= 0 Then gError 79865
    
    'Critica se a linha selecionada está preenchida
    If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Fornecedor_Col))) = 0 Then gError 79866
    
    'Guarda no obj o nome reduzido do fornecedor
    objFornecedor.sNomeReduzido = GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Fornecedor_Col)
    
    'Lê os dados do fornecedor e obtém o código do mesmo
    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then gError 79867
    
    'Se não encontrou => erro
    If lErro = 6681 Then gError 79868
    
    'Guarda no obj os dados que serão usados para obter os dados do título que será consultado
    With objTituloPagar
    
        .lFornecedor = objFornecedor.lCodigo
        .iFilial = GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Filial_Col)
        .dtDataEmissao = StrParaDate(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_DataEmissao_Col))
        .sSiglaDocumento = GridParcelas.TextMatrix(GridParcelas.Row, iGrid_SiglaDocumento_Col)
        .lNumTitulo = GridParcelas.TextMatrix(GridParcelas.Row, iGrid_NumeroDoc_Col)
    
    End With
    
    'Chama a tela de consulta de títulos a pagar
    Call Chama_Tela("TituloPagar_Consulta", objTituloPagar)
    
    Exit Sub
    
Erro_BotaoConsultaTitPag_Click:

    Select Case gErr
    
        Case 79867
                
        Case 79865
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 79866
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_PREENCHIDA", gErr)

        Case 79868
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INEXISTENTE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164217)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoSelecionar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoSelecionar_Click
    
    lErro = Pagtos_Lista()
    If lErro <> SUCESSO Then Error 49536
    
    Exit Sub

Erro_BotaoSelecionar_Click:

    Select Case Err

        Case 49536
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164218)

    End Select

    Exit Sub

End Sub

Private Sub ContaCorrente_Change()

    iAlterado = REGISTRO_ALTERADO
    glSequencial = 0
    
End Sub

Private Sub ContaCorrente_Click()

    iAlterado = REGISTRO_ALTERADO
    glSequencial = 0

End Sub

Private Function Traz_MovContaCorrente_Tela()
'Lê o movimento da conta corrente a partir de dados da tela e preenche os outros campos com os atributos deste movimento

Dim lErro As Long
Dim objMovContaCorrente As New ClassMovContaCorrente

On Error GoTo Erro_Traz_MovContaCorrente_Tela

    'Preenche o objeto com os principais campos da tela
    objMovContaCorrente.iCodConta = Codigo_Extrai(ContaCorrente.Text)
    objMovContaCorrente.dtDataMovimento = CDate(DataMovimento.Text)
    objMovContaCorrente.iTipoMeioPagto = Codigo_Extrai(TipoMeioPagto.Text)
    If objMovContaCorrente.iTipoMeioPagto <> DINHEIRO Then
        objMovContaCorrente.lNumero = CLng(Numero.Text)
    Else
        If glSequencial = 0 Then Error 15994
        objMovContaCorrente.lNumero = glSequencial
        objMovContaCorrente.lSequencial = glSequencial
    
    End If
    
    'Chama MovContaCorrente_Le_MeioPagto
    lErro = CF("MovContaCorrente_Le_MeioPagto_PagTit", objMovContaCorrente)
    If lErro <> SUCESSO Then Error 22157

    glSequencial = objMovContaCorrente.lSequencial
    
    Valor.Caption = Format(objMovContaCorrente.dValor, "Standard")

    'Chama GridParcelas_Preenche(objMovContaCorrente)
    lErro = GridParcelas_Preenche(objMovContaCorrente)
    If lErro <> SUCESSO Then Error 22158

    Traz_MovContaCorrente_Tela = SUCESSO

    Exit Function

Erro_Traz_MovContaCorrente_Tela:

    Traz_MovContaCorrente_Tela = Err
    
    Select Case Err

        'Erros já tratados
        Case 22157, 22158

        Case 15994
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTO_DINHEIRO_SEM_SEQUENCIAL", Err)
            Call Grid_Limpa(objGridParcelas)
            Valor.Caption = " "
        
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164219)

    End Select

    Exit Function

End Function

Private Sub ContaCorrente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim iCodigo As Integer

On Error GoTo Erro_ContaCorrente_Validate

    'Verifica se foi preenchida a ComboBox ContaCorrente
    If Len(Trim(ContaCorrente.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox ContaCorrente
    If ContaCorrente.Text = ContaCorrente.List(ContaCorrente.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(ContaCorrente, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 22149

    'Não existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objContaCorrenteInt.iCodigo = iCodigo

        'Tenta ler ContaCorrente com esse código no BD
        lErro = CF("ContaCorrenteInt_Le", iCodigo, objContaCorrenteInt)
        If lErro <> SUCESSO And lErro <> 11807 Then Error 22153
        
        'Não encontrou Conta Corrente no BD
        If lErro <> SUCESSO Then Error 22154

        'Se alguma Filial tiver sido selecionada
        If giFilialEmpresa <> EMPRESA_TODA Then

            If objContaCorrenteInt.iFilialEmpresa <> giFilialEmpresa Then Error 43534
        
        End If
        
        'Encontrou ContaCorrente no BD, coloca no Text da Combo
        ContaCorrente.Text = CStr(objContaCorrenteInt.iCodigo) & SEPARADOR & objContaCorrenteInt.sDescricao

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 43115

    Exit Sub

Erro_ContaCorrente_Validate:

    Cancel = True


    Select Case Err

        Case 22149, 22153
    
        Case 22154  'Não encontrou Conta Corrente no BD
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, objContaCorrenteInt.iCodigo)
            
        Case 43115
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE1", Err, ContaCorrente.Text)
            
        Case 43534
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_PERTENCE_FILIAL", Err, ContaCorrente.Text, giFilialEmpresa)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164220)
    
        End Select
    
    Exit Sub

End Sub

Private Sub DataMovimento_Change()

    iAlterado = REGISTRO_ALTERADO
    glSequencial = 0

End Sub

Private Sub DataMovimento_GotFocus()
Dim lSeqAux As Long
    
    lSeqAux = glSequencial
    Call MaskEdBox_TrataGotFocus(DataMovimento, iAlterado)
    glSequencial = lSeqAux
    
End Sub

Private Sub DataVencimento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoNumero = New AdmEvento
    
    'Carrega a ComboBox ContaCorrente com Código-NomeReduzido de Contas Correntes Internas
    lErro = Carrega_ContaCorrente()
    If lErro <> SUCESSO Then Error 22123

    'Carrega a combo dos Tipos de Meios de Pagamentos
    lErro = Carrega_TipoMeioPagto()
    If lErro <> SUCESSO Then Error 22124
    
    Set objGridParcelas = New AdmGrid
    
    'Inicializa o Grid
    lErro = Inicializa_Grid(objGridParcelas)
    If lErro <> SUCESSO Then Error 22125

    glSequencial = 0

    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 22123, 22124, 22125

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164221)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Carrega_ContaCorrente() As Long
'Carrega as Contas Correntes na Combo de ContasCorrentes

Dim lErro As Long
Dim colCodigoNomeRed As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome

On Error GoTo Erro_Carrega_ContaCorrente

    'Lê o nome e o código de todas a contas correntes
    lErro = CF("ContasCorrentesInternas_Le_CodigosNomesRed", colCodigoNomeRed)
    If lErro <> SUCESSO Then Error 22126

    For Each objCodigoNome In colCodigoNomeRed

        'Insere na combo de contas correntes
        ContaCorrente.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        ContaCorrente.ItemData(ContaCorrente.NewIndex) = objCodigoNome.iCodigo

    Next

    Carrega_ContaCorrente = SUCESSO

    Exit Function

Erro_Carrega_ContaCorrente:

    Carrega_ContaCorrente = Err

    Select Case Err

        Case 22126

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164222)

    End Select

    Exit Function

End Function

Private Function Carrega_TipoMeioPagto() As Long
'Carrega na Combo TipoMeioPagto os Tipos de Meios de Pagamentos

Dim lErro As Long
Dim colTipoMeioPagto As New Collection
Dim objTipoMeioPagto As ClassTipoMeioPagto
Dim iIndice As Integer

On Error GoTo Erro_Carrega_TipoMeioPagto

    'Lê todos os Tipos de Pagamento
    lErro = CF("TipoMeioPagto_Le_Todos", colTipoMeioPagto)
    If lErro <> SUCESSO Then Error 22169

    For Each objTipoMeioPagto In colTipoMeioPagto

        'Verifica se está ativo
        If objTipoMeioPagto.iInativo = TIPOMEIOPAGTO_ATIVO Then

            If giTipoVersao <> VERSAO_LIGHT Or objTipoMeioPagto.iTipo <> BORDERO Then
            
                'Coloca na combo
                TipoMeioPagto.AddItem CStr(objTipoMeioPagto.iTipo) & SEPARADOR & objTipoMeioPagto.sDescricao
                TipoMeioPagto.ItemData(TipoMeioPagto.NewIndex) = objTipoMeioPagto.iTipo

            End If
            
        End If

    Next
    
    For iIndice = 1 To TipoMeioPagto.ListCount - 1
    
        If TipoMeioPagto.ItemData(iIndice) = Cheque Then
        
            TipoMeioPagto.ListIndex = iIndice
            Exit For
        End If
    Next
    
    Carrega_TipoMeioPagto = SUCESSO

    Exit Function

Erro_Carrega_TipoMeioPagto:

    Carrega_TipoMeioPagto = Err

    Select Case Err

        'Erro já tratado
        Case 22169

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164223)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid(objGridInt As AdmGrid) As Long

    Set objGridParcelas = New AdmGrid

    'Tela em questão
    Set objGridParcelas.objForm = Me

    'Títulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("Emissão")
    objGridInt.colColuna.Add ("Vencimento")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Valor Pago")

   'Campos de edição do grid
    objGridInt.colCampo.Add (Fornecedor.Name)
    objGridInt.colCampo.Add (Filial.Name)
    objGridInt.colCampo.Add (SiglaDocumento.Name)
    objGridInt.colCampo.Add (NumeroDoc.Name)
    objGridInt.colCampo.Add (Parcela.Name)
    objGridInt.colCampo.Add (DataEmissao.Name)
    objGridInt.colCampo.Add (DataVencimento.Name)
    objGridInt.colCampo.Add (ValorParcela.Name)
    objGridInt.colCampo.Add (ValorPago.Name)

    'Indica onde estão situadas as colunas do grid
    iGrid_Fornecedor_Col = 1
    iGrid_Filial_Col = 2
    iGrid_SiglaDocumento_Col = 3
    iGrid_NumeroDoc_Col = 4
    iGrid_Parcela_Col = 5
    iGrid_DataEmissao_Col = 6
    iGrid_DataVencimento_Col = 7
    iGrid_ValorParcela_Col = 8
    iGrid_ValorPago_Col = 9

    objGridInt.objGrid = GridParcelas

    'todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_PARCELAS_BORDERO

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 5

    GridParcelas.ColWidth(0) = 400

    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    Call Grid_Inicializa(objGridInt)

    Exit Function

End Function

Function Trata_Parametros(Optional objMovContaCorrente As ClassMovContaCorrente) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se algum movimento foi passado por parametro
    If Not (objMovContaCorrente Is Nothing) Then

        'Lê MovimentoContaCorrente no BD
        lErro = CF("MovContaCorrente_Le", objMovContaCorrente)
        If lErro <> SUCESSO And lErro <> 11893 Then Error 22135

        'Se não existe
        If lErro <> AD_SQL_SUCESSO Then Error 22136

        'Verifica se é um movto de um tipo válido
        If objMovContaCorrente.iTipo <> MOVCCI_PAGTO_TITULO_POR_BORDERO And objMovContaCorrente.iTipo <> MOVCCI_PAGTO_TITULO_POR_CHEQUE And objMovContaCorrente.iTipo <> MOVCCI_PAGTO_TITULO_POR_DINHEIRO Then Error 22139

        'Chama Preenche_Tela
        lErro = Preenche_Tela(objMovContaCorrente)
        If lErro <> SUCESSO Then Error 22162

    End If

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        'Erro  já tratado
        Case 22162

        Case 22135

        Case 22136
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMMOVTO_INEXISTENTE", Err, objMovContaCorrente.lNumMovto)

        Case 22139
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CANCELAMENTO_PAG_NAO_SE_APLICA_AO_MOV", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164224)

    End Select

    Exit Function

End Function

Function Preenche_Tela(objMovContaCorrente As ClassMovContaCorrente) As Long
'Traz os dados do movimento para tela

Dim lErro As Long
Dim objContaCorrenteInt As New ClassContasCorrentesInternas

On Error GoTo Erro_Preenche_Tela

    'Lê a  Conta Corrente a partir do Código
    lErro = CF("ContaCorrenteInt_Le", objMovContaCorrente.iCodConta, objContaCorrenteInt)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 22137

    'Não encontrou a Conta Corrente --> erro
    If lErro = 11807 Then Error 22138

    'Preenche o Grid
    Call GridParcelas_Preenche(objMovContaCorrente)

    'Exibe os dados na tela que não fazem parte do Grid
    If objMovContaCorrente.iCodConta = 0 Then
        ContaCorrente.Text = ""
    Else
        ContaCorrente.Text = CStr(objMovContaCorrente.iCodConta)
        Call ContaCorrente_Validate(bSGECancelDummy)
    End If

    If objMovContaCorrente.iTipoMeioPagto = 0 Then
        TipoMeioPagto.Text = ""
    Else
        TipoMeioPagto.Text = CStr(objMovContaCorrente.iTipoMeioPagto)
        Call TipoMeioPagto_Validate(bSGECancelDummy)
    End If

    DataMovimento.PromptInclude = False
    DataMovimento.Text = Format(objMovContaCorrente.dtDataMovimento, "dd/MM/yy")
    DataMovimento.PromptInclude = True

    If objMovContaCorrente.lNumero = 0 Then
        Numero.Text = ""
    Else
        Numero.Text = CStr(objMovContaCorrente.lNumero)
        Call Numero_Validate(bSGECancelDummy)
    End If

    Valor.Caption = Format(objMovContaCorrente.dValor, "Standard")

    'Zerar iAlterado
    iAlterado = 0

    Preenche_Tela = SUCESSO

    Exit Function

Erro_Preenche_Tela:

    Preenche_Tela = Err

    Select Case Err

        'Erro  já  tratado
        Case 22137

        Case 22138
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, objMovContaCorrente.iCodConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164225)

    End Select

    Exit Function

End Function

Private Function GridParcelas_Preenche(objMovContaCorrente As ClassMovContaCorrente) As Long
'Preenche o Grid de Parcelas

Dim lErro As Long
Dim iLinha As Integer
Dim objInfoParcPag As ClassInfoParcPag
Dim colInfoParcPag As New Collection

On Error GoTo Erro_GridParcelas_Preenche

    'Limpa o Grid
    Call Grid_Limpa(objGridParcelas)

    iLinha = 0

    'Lê no BD as parcelas
    lErro = CF("ParcelasPag_Le_MovContaCorrente", objMovContaCorrente, colInfoParcPag)
    If lErro <> SUCESSO Then Error 22148
        
    If colInfoParcPag.Count >= objGridParcelas.objGrid.Rows Then
        Call Refaz_Grid(objGridParcelas, colInfoParcPag.Count)
    End If

    'Percorre todas as parcelas da Coleção
    For Each objInfoParcPag In colInfoParcPag

        iLinha = iLinha + 1

        'Passa para a tela os dados da parcela em questão
        GridParcelas.TextMatrix(iLinha, iGrid_Fornecedor_Col) = objInfoParcPag.sNomeRedForn
        GridParcelas.TextMatrix(iLinha, iGrid_Filial_Col) = objInfoParcPag.iFilialForn
        GridParcelas.TextMatrix(iLinha, iGrid_SiglaDocumento_Col) = objInfoParcPag.sSiglaDocumento
        GridParcelas.TextMatrix(iLinha, iGrid_NumeroDoc_Col) = objInfoParcPag.lNumTitulo
        GridParcelas.TextMatrix(iLinha, iGrid_Parcela_Col) = objInfoParcPag.iNumParcela
        If objInfoParcPag.dtDataEmissao <> DATA_NULA Then GridParcelas.TextMatrix(iLinha, iGrid_DataEmissao_Col) = Format(objInfoParcPag.dtDataEmissao, "dd/MM/yyyy")
        GridParcelas.TextMatrix(iLinha, iGrid_DataVencimento_Col) = Format(objInfoParcPag.dtDataVencimento, "dd/MM/yyyy")
        GridParcelas.TextMatrix(iLinha, iGrid_ValorParcela_Col) = Format(objInfoParcPag.dValorOriginal, "Standard")
        GridParcelas.TextMatrix(iLinha, iGrid_ValorPago_Col) = Format(objInfoParcPag.dValor - objInfoParcPag.dValorDesconto + (objInfoParcPag.dValorJuros + objInfoParcPag.dValorMulta), "Standard")

    Next

    'Passa para o Obj o número de parcelas passadas pela Coleção
    objGridParcelas.iLinhasExistentes = iLinha

    Exit Function

Erro_GridParcelas_Preenche:

    GridParcelas_Preenche = Err

    Select Case Err

        'Erro  já  tratado
        Case 22148

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164226)

    End Select

    Exit Function

End Function

Private Sub Fornecedor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Filial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Numero_Change()

    iAlterado = REGISTRO_ALTERADO
    glSequencial = 0
    
End Sub

Private Sub Numero_GotFocus()
Dim lSeqAux As Long
    
    lSeqAux = glSequencial
    Call MaskEdBox_TrataGotFocus(Numero, iAlterado)
    glSequencial = lSeqAux

End Sub

Private Sub NumeroDoc_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NumeroLabel_Click()

Dim lErro As Long

On Error GoTo Erro_NumeroLabel_Click
    
    lErro = Pagtos_Lista()
    If lErro <> SUCESSO Then Error 49535
    
    Exit Sub

Erro_NumeroLabel_Click:

    Select Case Err

        Case 49535
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164227)

    End Select

    Exit Sub

End Sub

Function Pagtos_Lista() As Long

Dim lErro As Long
Dim objMovContaCorrente As New ClassMovContaCorrente
Dim colSelecao As New Collection

On Error GoTo Erro_Pagtos_Lista

    'Verifica se a ContaCorrente está preenchida
    If Len(Trim(ContaCorrente.Text)) = 0 Then Error 49533

    'Verifica se o meio de pagto está preenchido
    If Len(Trim(TipoMeioPagto.Text)) = 0 Then Error 49534

    'Chama Move_Tela_Memoria
    lErro = Move_Tela_Memoria(objMovContaCorrente)
    If lErro <> SUCESSO Then Error 49532

    'Adiciona em ColSelecao
    colSelecao.Add objMovContaCorrente.iCodConta
    colSelecao.Add objMovContaCorrente.iTipoMeioPagto

    'Chama tela PagtoLista
    Call Chama_Tela("PagtoLista", colSelecao, objMovContaCorrente, objEventoNumero)
    
    Pagtos_Lista = SUCESSO
    
    Exit Function

Erro_Pagtos_Lista:
    
    Pagtos_Lista = Err
    
    Select Case Err

        'Erro  já  tratado
        Case 49532

        Case 49533
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_INFORMADA", Err)

        Case 49534
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MEIOPAGTO_NAO_INFORMADO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164228)

    End Select

    Exit Function

End Function

Function Move_Tela_Memoria(objMovContaCorrente As ClassMovContaCorrente) As Long

Dim lErro As Long

    'Move os dados da tela para objMovcontacorrente
    If Len(Trim(ContaCorrente.Text)) > 0 Then objMovContaCorrente.iCodConta = Codigo_Extrai(ContaCorrente.Text)
    If Len(Trim(DataMovimento.ClipText)) = 0 Then
        objMovContaCorrente.dtDataMovimento = DATA_NULA
    Else
        objMovContaCorrente.dtDataMovimento = CDate(DataMovimento.Text)
    End If
    If Len(Trim(TipoMeioPagto.Text)) > 0 Then objMovContaCorrente.iTipoMeioPagto = Codigo_Extrai(TipoMeioPagto.Text)
    If Len(Trim(Numero.Text)) > 0 Then objMovContaCorrente.lNumero = CLng(Numero.Text)
    If Len(Trim(Valor.Caption)) > 0 Then objMovContaCorrente.dValor = CDbl(Valor)

    Move_Tela_Memoria = SUCESSO

    Exit Function

End Function

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objMovContaCorrente As ClassMovContaCorrente

    Set objMovContaCorrente = obj1

    'Chama  Preenche_Tela
    Call Preenche_Tela(objMovContaCorrente)

    glSequencial = objMovContaCorrente.lSequencial
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Me.Show

End Sub

Private Sub Parcela_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub SiglaDocumento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoMeioPagto_Change()

    iAlterado = REGISTRO_ALTERADO
    glSequencial = 0

End Sub

Private Sub TipoMeioPagto_Click()

Dim lErro  As Integer

    iAlterado = REGISTRO_ALTERADO
    glSequencial = 0

    If TipoMeioPagto.ItemData(TipoMeioPagto.ListIndex) = DINHEIRO Then
    
        BotaoSelecionar.Visible = True
        BotaoTrazer.Visible = False
        
    Else
    
        BotaoSelecionar.Visible = False
        BotaoTrazer.Visible = True
    
    End If
    
End Sub

Private Sub TipoMeioPagto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objTipoMeioPagto As New ClassTipoMeioPagto
Dim iCodigo As Integer

On Error GoTo Erro_TipoMeioPagto_Validate

    'Verifica se foi preenchida a ComboBox TipoMeioPagto
    If Len(Trim(TipoMeioPagto.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox TipoMeioPagto
    If TipoMeioPagto.Text = TipoMeioPagto.List(TipoMeioPagto.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(TipoMeioPagto, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 22128

    'Não existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Or lErro = 6731 Then Error 22127
    
    Exit Sub

Erro_TipoMeioPagto_Validate:

    Cancel = True


    Select Case Err

        Case 22127
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_MEIO_PAGAMENTO_NAO_CADASTRADO", Err)
    
        Case 22128
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164229)

    End Select

    Exit Sub

End Sub

Private Sub DataMovimento_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataMovimento_Validate

    'Verifica se a data está preenchida
    If Len(Trim(DataMovimento.ClipText)) > 0 Then

        'Verifica se a data final é válida
        lErro = Data_Critica(DataMovimento.Text)
        If lErro <> SUCESSO Then Error 22131

    End If

    Exit Sub

Erro_DataMovimento_Validate:

    Cancel = True


    Select Case Err

        'Erro  já  tratado
        Case 22151

        Case 22131

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164230)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objMovContaCorrente As New ClassMovContaCorrente
Dim colTipos As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "MovimentosContaCorrente"

    'Lê os dados da Tela
    lErro = Move_Tela_Memoria(objMovContaCorrente)
    If lErro <> SUCESSO Then Error 22168

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "CodConta", objMovContaCorrente.iCodConta, 0, "CodConta"
    colCampoValor.Add "TipoMeioPagto", objMovContaCorrente.iTipoMeioPagto, 0, "TipoMeioPagto"
    colCampoValor.Add "Numero", objMovContaCorrente.lNumero, 0, "Numero"
    colCampoValor.Add "DataMovimento", objMovContaCorrente.dtDataMovimento, 0, "DataMovimento"
    colCampoValor.Add "Valor", objMovContaCorrente.dValor, 0, "Valor"

   'Filtros para o Sistema de Setas
    colSelecao.Add "Excluido", OP_DIFERENTE, MOVCONTACORRENTE_EXCLUIDO
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    lErro = CF("TiposMovimento_Le_NaoPagto", colTipos)
    If lErro <> SUCESSO Then Error 22147

    For iIndice = 1 To colTipos.Count
    
        'Para cada tipo selecionado, adiciona filtro Tipo<>Tipo
        colSelecao.Add "Tipo", OP_DIFERENTE, colTipos(iIndice)
        
    Next

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        'Erros já tratados
        Case 22147, 22168

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164231)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objMovContaCorrente As New ClassMovContaCorrente

On Error GoTo Erro_Tela_Preenche

    objMovContaCorrente.iCodConta = colCampoValor.Item("CodConta").vValor

    If objMovContaCorrente.iCodConta > 0 Then

        'Carrega objMovContaCorrente com os dados passados em colCampoValor
        objMovContaCorrente.iTipoMeioPagto = colCampoValor.Item("TipoMeioPagto").vValor
        objMovContaCorrente.lNumero = colCampoValor.Item("Numero").vValor
        objMovContaCorrente.dtDataMovimento = colCampoValor.Item("DataMovimento").vValor
        objMovContaCorrente.dValor = colCampoValor.Item("Valor").vValor

        'Traz dados do  Pagamento para a Tela
        lErro = Preenche_Tela(objMovContaCorrente)
        If lErro <> SUCESSO Then Error 22146

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        'Erro já tratado
        Case 22146

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164232)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objGridParcelas = Nothing
    Set objEventoNumero = Nothing
    
    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
    
    glSequencial = 0

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    If glSequencial <> 0 Then
    
        Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
        
    End If
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    If glSequencial <> 0 Then
    
        'Testa se deseja salvar mudanças
        lErro = Teste_Salva(Me, iAlterado)
        If lErro <> SUCESSO Then Error 22130
        
    End If
    
    Call Limpa_Tela_PagtoCancelar

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        'Erros já tratados
        Case 22130, 22055

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164233)

    End Select

End Sub

Sub Limpa_Tela_PagtoCancelar()

Dim lErro As Long

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Limpa TextBox e MaskedEditBox
    Call Limpa_Tela(Me)

    'Limpa os textos das Combos
    ContaCorrente.Text = ""
    TipoMeioPagto.Text = ""

    'Atualiza data
    DataMovimento.PromptInclude = False
    DataMovimento.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataMovimento.PromptInclude = True
    
    'Limpa o Label Valor
    Valor.Caption = ""

    'Limpa GridParcelas
    Call Grid_Limpa(objGridParcelas)

    'linhas visiveis do grid
    objGridParcelas.iLinhasExistentes = 0

    iAlterado = 0
    glSequencial = 0

End Sub

Private Sub UpDownDataMovimento_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataMovimento_DownClick

    'Verifica se a data foi preenchida
    If Len(Trim(DataMovimento.ClipText)) > 0 Then

        'Diminui a data
        lErro = Data_Up_Down_Click(DataMovimento, DIMINUI_DATA)
        If lErro <> SUCESSO Then Error 22132

    End If

    Exit Sub

Erro_UpDownDataMovimento_DownClick:

    Select Case Err

        'Erro já tratado
        Case 22132

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164234)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataMovimento_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataMovimento_UpClick

    'verifica se a data foi preenchida
    If Len(Trim(DataMovimento.ClipText)) > 0 Then

        'Aumenta a data
        lErro = Data_Up_Down_Click(DataMovimento, AUMENTA_DATA)
        If lErro <> SUCESSO Then Error 22133

    End If

    Exit Sub

Erro_UpDownDataMovimento_UpClick:

    Select Case Err

        'Erro já tratado
        Case 22133

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164235)

    End Select

    Exit Sub

End Sub

Private Sub ValorPago_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorParcela_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Numero_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objTipoMeioPagto As New ClassTipoMeioPagto
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Numero_Validate

    'Verifica se o Número foi preenchido
    If Len(Trim(Numero.ClipText)) = 0 Then Exit Sub

    'Critica se é Long Positivo
    lErro = Valor_NaoNegativo_Critica(Numero.Text)
    If lErro <> SUCESSO Then Error 22134

    Exit Sub

Erro_Numero_Validate:

    Cancel = True


    Select Case Err

    Case 22134

    Case 22150

    Case Else

        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164236)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama a função de gravacao
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 22140

    Call Limpa_Tela_PagtoCancelar
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        'Erro já tratado
        Case 22140

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164237)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objMovContaCorrente As New ClassMovContaCorrente

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se a ContaCorrente está preenchida
    If Len(Trim(ContaCorrente.Text)) = 0 Then Error 22141

    'Verifica se o meio de pagto está preenchido
    If Len(Trim(TipoMeioPagto.Text)) = 0 Then Error 22142

    'Verifica se o Número está preenchido
    If Codigo_Extrai(TipoMeioPagto.Text) <> DINHEIRO Then
        If Len(Trim(Numero.Text)) = 0 Then Error 22143
    Else
        If glSequencial = 0 Then Error 15996
    End If

    'Verifica se a DataMovimento está preenchida
    If Len(Trim(DataMovimento.ClipText)) = 0 Then Error 22144

    'Preenche o objeto com os principais campos da tela
    'Lê os dados da Tela PagtoLista
    lErro = Move_Tela_Memoria(objMovContaCorrente)
    If lErro <> SUCESSO Then Error 22170

    If objMovContaCorrente.iTipoMeioPagto = DINHEIRO Then
        objMovContaCorrente.lNumero = glSequencial
        objMovContaCorrente.lSequencial = glSequencial
    End If

    'Lê no BD
    lErro = CF("MovContaCorrente_Le_MeioPagto_PagTit", objMovContaCorrente)
    If lErro <> SUCESSO Then Error 22145

    'Pede confirmação para cancelamento de Pagamento
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_PAGAMENTO")

    If vbMsgRes = vbYes Then

        'Cancela o Pagamento
        lErro = CF("MovContaCorrente_Pagto_Cancelar", objMovContaCorrente)
        If lErro <> SUCESSO Then Error 22155

        'Limpa a Tela
        Call Limpa_Tela_PagtoCancelar
        
    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 15996
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTO_DINHEIRO_SEM_SEQUENCIAL", Err)

        'Erro já tratado
        Case 22155, 22170

        Case 22141
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_INFORMADA", Err)

        Case 22142
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MEIOPAGTO_NAO_INFORMADO", Err)

        Case 22143
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DOCUMENTO_NAO_PREENCHIDO", Err)

        Case 22144
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", Err)

        Case 22145
            lErro = Rotina_Erro(vbOKOnly, "NAO_EXISTE_PAG_PARA_SER_CANCELADO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164238)

    End Select

    Exit Function

End Function
    
Private Sub BotaoTrazer_Click()

Dim lErro As Long
Dim iTipo As Integer

On Error GoTo Erro_BotaoTrazer_Click

    'Verifica se Numero, TipoMeioPagto, DataMovimento e ContaCorrente estão preenchidos
    If Len(Trim(TipoMeioPagto.Text)) = 0 Then Error 22151
    If Len(Trim(DataMovimento.ClipText)) = 0 Then Error 22152
    If Len(Trim(ContaCorrente.Text)) = 0 Then Error 22156
        
    iTipo = Codigo_Extrai(TipoMeioPagto.Text)
        
    If iTipo <> DINHEIRO Then If Len(Trim(Numero.Text)) = 0 Then Error 22129
        
    'Chama Traz_MovContaCorrente_Tela
    lErro = Traz_MovContaCorrente_Tela
    If lErro <> SUCESSO Then Error 22150

    Exit Sub
    
Erro_BotaoTrazer_Click:

    Select Case Err
    
        Case 22129
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INFORMADO", Err, iTipo)
        
        Case 22150
            Call Grid_Limpa(objGridParcelas)
            Valor.Caption = ""
        
        Case 22151
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_NAO_INFORMADO", Err)
        
        Case 22152
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", Err)
            
        Case 22156
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_INFORMADA", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164239)
            
    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CANCELAR_PAGAMENTOS
    Set Form_Load_Ocx = Me
    Caption = "Cancelar Pagamento"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "PagtoCancelar"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Numero Then
            Call NumeroLabel_Click
        End If
    
    End If
    
End Sub


Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub NumeroLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumeroLabel, Source, X, Y)
End Sub

Private Sub NumeroLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumeroLabel, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Valor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Valor, Source, X, Y)
End Sub

Private Sub Valor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Valor, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Sub Refaz_Grid(ByVal objGridInt As AdmGrid, ByVal iNumLinhas As Integer)
    objGridInt.objGrid.Rows = iNumLinhas + 1

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)
End Sub

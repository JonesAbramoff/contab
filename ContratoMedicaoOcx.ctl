VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ContratoMedicaoOcx 
   ClientHeight    =   6090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   KeyPreview      =   -1  'True
   ScaleHeight     =   6090
   ScaleWidth      =   9600
   Begin VB.Frame Frame3 
      Caption         =   "Identificação"
      Height          =   705
      Left            =   150
      TabIndex        =   32
      Top             =   735
      Width           =   9315
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   2445
         Picture         =   "ContratoMedicaoOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Numeração Automática"
         Top             =   270
         Width           =   300
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   1620
         TabIndex        =   34
         Top             =   270
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown 
         Height          =   300
         Left            =   7665
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   300
         Left            =   6570
         TabIndex        =   36
         Top             =   270
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label CodigoLabel 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
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
         Left            =   780
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   38
         Top             =   315
         Width           =   660
      End
      Begin VB.Label DataLabel 
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
         Left            =   5925
         TabIndex        =   37
         Top             =   330
         Width           =   480
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Contrato"
      Height          =   1125
      Left            =   165
      TabIndex        =   23
      Top             =   1515
      Width           =   9300
      Begin VB.ComboBox FilCliContrato 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   6540
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   660
         Width           =   1815
      End
      Begin MSMask.MaskEdBox CliContrato 
         Height          =   315
         Left            =   1605
         TabIndex        =   25
         Top             =   690
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483644
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Contrato 
         Height          =   315
         Left            =   1590
         TabIndex        =   26
         Top             =   195
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         PromptChar      =   " "
      End
      Begin VB.Label FilCliContratoLabel 
         AutoSize        =   -1  'True
         Caption         =   "Filial:"
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
         Left            =   5955
         TabIndex        =   30
         Top             =   720
         Width           =   465
      End
      Begin VB.Label ContratoLabel 
         Caption         =   "Contrato:"
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
         Height          =   210
         Left            =   630
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   29
         Top             =   270
         Width           =   795
      End
      Begin VB.Label CliContratoLabel 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
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
         Left            =   780
         TabIndex        =   28
         Top             =   720
         Width           =   660
      End
      Begin VB.Label DescContrato 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2925
         TabIndex        =   27
         Top             =   210
         Width           =   5415
      End
   End
   Begin VB.PictureBox Picture 
      Height          =   555
      Left            =   7365
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   75
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ContratoMedicaoOcx.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "ContratoMedicaoOcx.ctx":0244
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "ContratoMedicaoOcx.ctx":03CE
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "ContratoMedicaoOcx.ctx":0900
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Itens"
      Height          =   3315
      Left            =   165
      TabIndex        =   4
      Top             =   2730
      Width           =   9315
      Begin MSMask.MaskEdBox DataCobranca 
         Height          =   300
         Left            =   6780
         TabIndex        =   31
         Top             =   540
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         MaxLength       =   8
         Format          =   "dd/mm/yy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataRefFim 
         Height          =   300
         Left            =   6675
         TabIndex        =   21
         Top             =   1005
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         MaxLength       =   8
         Format          =   "dd/mm/yy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataRefIni 
         Height          =   300
         Left            =   5670
         TabIndex        =   22
         Top             =   585
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         MaxLength       =   8
         Format          =   "dd/mm/yy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Item 
         Height          =   225
         Left            =   2580
         TabIndex        =   15
         Top             =   975
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.CommandButton BotaoFat 
         Caption         =   "Faturamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   7755
         TabIndex        =   3
         ToolTipText     =   "Abre a Nota Fiscal gerada se o item tiver sido faturado."
         Top             =   2820
         Width           =   1425
      End
      Begin VB.ComboBox UM 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2670
         TabIndex        =   14
         Top             =   1260
         Width           =   630
      End
      Begin VB.ComboBox Status 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   420
         TabIndex        =   13
         Top             =   930
         Width           =   1455
      End
      Begin MSMask.MaskEdBox ValorCobrar 
         Height          =   225
         Left            =   5100
         TabIndex        =   12
         Top             =   675
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.CheckBox Cobrar 
         Caption         =   "Cobrar"
         Height          =   210
         Left            =   5460
         TabIndex        =   10
         Tag             =   "1"
         Top             =   1305
         Width           =   705
      End
      Begin MSMask.MaskEdBox Custo 
         Height          =   225
         Left            =   4110
         TabIndex        =   11
         Top             =   585
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox QtdContratada 
         Height          =   225
         Left            =   3795
         TabIndex        =   8
         Top             =   1350
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Quantidade 
         Height          =   225
         Left            =   975
         TabIndex        =   9
         Top             =   1380
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   225
         Left            =   3120
         TabIndex        =   5
         Top             =   555
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ItemDescricao 
         Height          =   225
         Left            =   3600
         TabIndex        =   6
         Top             =   945
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Produto 
         Height          =   225
         Left            =   2355
         TabIndex        =   7
         Top             =   750
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.CommandButton BotaoLimparGrid 
         Caption         =   "Limpar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2835
         TabIndex        =   2
         Top             =   2820
         Width           =   1350
      End
      Begin VB.CommandButton BotaoTrazer 
         Caption         =   "Trazer Itens Contratados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   180
         TabIndex        =   1
         Top             =   2820
         Width           =   2565
      End
      Begin MSFlexGridLib.MSFlexGrid GridItens 
         Height          =   1500
         Left            =   195
         TabIndex        =   0
         Top             =   255
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   2646
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "ContratoMedicaoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim iAlterado As Integer
Dim iCodigoAlterado As Integer
Dim iContratoAlterado As Integer

Private iFrameAtual As Integer

'HElp
Const IDH_RASTROPRODNFFAT = 0

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoContrato As AdmEvento
Attribute objEventoContrato.VB_VarHelpID = -1
Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1

Dim objGridItens As AdmGrid
Dim iGrid_Produto_Col As Integer
Dim iGrid_Descricao_Col As Integer
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_QtdContratada_Col As Integer
Dim iGrid_Custo_Col As Integer
Dim iGrid_Cobrar_Col As Integer
Dim iGrid_ValorCobrar_Col As Integer
Dim iGrid_Status_Col As Integer
Dim iGrid_UM_Col As Integer
Dim iGrid_Item_Col As Integer
Dim iGrid_DataCobranca_Col As Integer
Dim iGrid_DataRefIni_Col As Integer
Dim iGrid_DataRefFim_Col As Integer

Private Sub Limpa_Tela_Medicao()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Medicao

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Call Limpa_Tela(Me)
       
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
        
    FilCliContrato.Clear
    DescContrato.Caption = ""
       
    Call Grid_Limpa(objGridItens)
    
    iAlterado = 0
    iCodigoAlterado = 0
    iContratoAlterado = 0
    
    Exit Sub

Erro_Limpa_Tela_Medicao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155072)

    End Select

    Exit Sub

End Sub

Private Function Move_GridItens_Memoria(objMedicaoContrato As ClassMedicaoContrato, objContrato As ClassContrato) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objItensDeMedicao As ClassItensMedCtr
Dim objItensDeContrato As ClassItensDeContrato

On Error GoTo Erro_Move_GridItens_Memoria

    For iIndice = 1 To objGridItens.iLinhasExistentes
        
        Set objItensDeMedicao = New ClassItensMedCtr

        With objItensDeMedicao
        
            .dCusto = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Custo_Col))
            .dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
            .dVlrCobrar = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_ValorCobrar_Col))
            
            If Codigo_Extrai(GridItens.TextMatrix(iIndice, iGrid_Status_Col)) = COBRADO Then
                .iStatus = COBRADO
            Else
                .iStatus = NAO_COBRADO
            End If
            
            .lMedicao = objMedicaoContrato.lCodigo
                                              
            For Each objItensDeContrato In objContrato.colItens
                If objItensDeContrato.iSeq = StrParaInt(GridItens.TextMatrix(iIndice, iGrid_Item_Col)) Then
                    .lNumIntItensContrato = objItensDeContrato.lNumIntDoc
                End If
            Next
            
            .dtDataRefIni = StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataRefIni_Col))
            .dtDataRefFim = StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataRefFim_Col))
            .dtDataCobranca = StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataCobranca_Col))
                                              
        End With
            
        'Se houve alteração no item de contrato = > Grava o Item de Medição
        If objItensDeMedicao.dCusto <> 0 Or objItensDeMedicao.dQuantidade <> 0 Or objItensDeMedicao.dVlrCobrar <> 0 Then
            objMedicaoContrato.colItens.Add objItensDeMedicao
        End If
    
    Next

    Move_GridItens_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItens_Memoria:

    Move_GridItens_Memoria = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155073)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(objMedicaoContrato As ClassMedicaoContrato) As Long

Dim lErro As Long
Dim objContrato As New ClassContrato

On Error GoTo Erro_Move_Tela_Memoria

    With objMedicaoContrato

        .dtData = StrParaDate(Data.Text)
        .lCodigo = StrParaLong(Codigo.Text)
        
        objContrato.sCodigo = Contrato.Text
        objContrato.iFilialEmpresa = giFilialEmpresa
        
        lErro = CF("Contrato_Le", objContrato)
        If lErro <> SUCESSO And lErro <> 129332 Then gError 129837
        
        'Contrato Não Cadastrado
        If lErro = 129332 Then gError 132911
        
        If objContrato.iTipo <> CONTRATOS_RECEBER Then gError 132904
        
        .lNumIntContrato = objContrato.lNumIntDoc
    
    End With
    
    lErro = Move_GridItens_Memoria(objMedicaoContrato, objContrato)
    If lErro <> SUCESSO Then gError 129630

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 129630

        Case 132904
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_NAO_RECEBER", gErr, objContrato.sCodigo, objContrato.iFilialEmpresa)

        Case 132911
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_NAO_CADASTRADO", gErr, objContrato.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155074)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long
Dim lNumAlmoxarifados As Long

On Error GoTo Erro_Form_Load
    
    iFrameAtual = 1

    Set objGridItens = New AdmGrid
   
    Set objEventoContrato = New AdmEvento
    Set objEventoCodigo = New AdmEvento
    
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
   
    'Inicializa o grid de Itens de Contrato
    lErro = Inicializa_GridItens(objGridItens)
    If lErro <> SUCESSO Then gError 129631

    lErro = Carrega_ComboStatus(Status)
    If lErro <> SUCESSO Then gError 129632
    
    If Len(Trim(gobjFAT.sFormatoPrecoUnitario)) <> 0 Then
        Custo.Format = gobjFAT.sFormatoPrecoUnitario
        Valor.Format = gobjFAT.sFormatoPrecoUnitario
        ValorCobrar.Format = gobjFAT.sFormatoPrecoUnitario
    End If

    iAlterado = 0
    iCodigoAlterado = 0
    iContratoAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 129631, 129632

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155075)

    End Select

    iAlterado = 0
    iCodigoAlterado = 0
    iContratoAlterado = 0
    
    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, UnloadMode, Cancel, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objGridItens = Nothing
    Set objEventoContrato = Nothing
    Set objEventoCodigo = Nothing
    
    'Fecha o Comando de Setas
    Call ComandoSeta_Liberar(Me.Name)


End Sub

Private Sub BotaoFat_Click()

Dim objNF As New ClassNFiscal
Dim lErro As Long
Dim objItemNF As New ClassItemNF
Dim objContrato As New ClassContrato
Dim objItensDeContrato As New ClassItensDeContrato
Dim iLinha As Integer
Dim bAchou As Boolean
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objItemMedicao As New ClassItensMedCtr

On Error GoTo Erro_BotaoFat_Click

    If GridItens.Row = 0 Then gError 129947
    
    iLinha = GridItens.Row

    objContrato.sCodigo = Contrato.Text
    objContrato.iFilialEmpresa = giFilialEmpresa
    
    bAchou = True

    'Le o contrato
    lErro = CF("Contrato_Le", objContrato)
    If lErro <> SUCESSO And lErro <> 129332 Then gError 129948
    If lErro = 129332 Then bAchou = False
       
    'Se o contrato está cadastrado => Le o item de contrato
    If bAchou Then
        
        If objContrato.iTipo <> CONTRATOS_RECEBER Then gError 132905
    
        objItensDeContrato.iSeq = StrParaInt(GridItens.TextMatrix(iLinha, iGrid_Item_Col))
        objItensDeContrato.lNumIntContrato = objContrato.lNumIntDoc
    
        'Le o item de contrato
        lErro = CF("ItensDeContrato_Le2", objItensDeContrato)
        If lErro <> SUCESSO And lErro <> 129266 Then gError 129949
        If lErro = 129266 Then bAchou = False
    
    End If

    'Se o Item de contrato está cadastrado => Le os dados sobre a fatura
    If bAchou Then
    
        objItemNF.objCobrItensContrato.lMedicao = StrParaLong(Codigo.Text)
        objItemNF.objCobrItensContrato.lNumIntItensContrato = objItensDeContrato.lNumIntDoc
    
        objItemMedicao.lMedicao = objItemNF.objCobrItensContrato.lMedicao
        objItemNF.objCobrItensContrato.colMedicoes.Add objItemMedicao
    
        'Le os dados sobre a fatura
        lErro = CF("ItensDeContrato_Le_DadosFatura", objNF, objItemNF)
        If lErro <> SUCESSO And lErro <> 129904 And lErro <> 129907 And lErro <> 129908 Then gError 129950
        If lErro <> SUCESSO Then bAchou = False
        
    End If
    
    'Se encontrou os dados da fatura => Abre a tela de NFiscal
    If bAchou Then
    
        objTipoDocInfo.iCodigo = objNF.iTipoNFiscal

        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO Then gError 129879
    
        lErro = Chama_Tela(objTipoDocInfo.sNomeTelaNFiscal, objNF)
        If lErro <> SUCESSO Then gError 129951
        
    End If
    
    If Not bAchou Then gError 136036
    
    Exit Sub
    
Erro_BotaoFat_Click:

    Select Case gErr
    
        Case 129879
    
        Case 129948 To 129951
        
        Case 129947
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 132905
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_NAO_RECEBER", gErr, objContrato.sCodigo, objContrato.iFilialEmpresa)
        
        Case 136036
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_NAO_FATURADO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155076)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimparGrid_Click()

    Call Grid_Limpa(objGridItens)

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lNumProx As Long
Dim lTransacao As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Abre transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 129633

    'Obtem o identificador da Medição
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_PROX_ITENSDEMEDICAO", lNumProx)
    If lErro <> SUCESSO Then gError 129634
    
     'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 129635
   
    'Traz o identificador para tela
    Codigo.PromptInclude = False
    Codigo.Text = CStr(lNumProx)
    Codigo.PromptInclude = True
    
    Exit Sub
    
Erro_BotaoProxNum_Click:

    Select Case gErr
    
        Case 129633
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 129634
        
        Case 129635
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
   
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155077)
        
    End Select
    
    'Desfaz alterações
    Call Transacao_Rollback
    
    Exit Sub
    
End Sub

Private Sub BotaoTrazer_Click()

Dim objContrato As New ClassContrato
Dim lErro As Long

On Error GoTo Erro_BotaoTrazer_Click

    If Len(Trim(Contrato.Text)) = 0 Then Exit Sub
    
    objContrato.sCodigo = Contrato.Text
    objContrato.iFilialEmpresa = giFilialEmpresa

    'Traz todos os itens do contrato para tela
    lErro = Traz_GridItens_Tela(objContrato)
    If lErro <> SUCESSO Then gError 131056
    
    Exit Sub
    
Erro_BotaoTrazer_Click:

    Select Case gErr
    
        Case 131056

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155078)

    End Select
    
    Exit Sub
    
End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO
    iCodigoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objMedicaoContrato As New ClassMedicaoContrato

On Error GoTo Erro_Codigo_Validate

    objMedicaoContrato.lCodigo = StrParaLong(Codigo.ClipText)

    If objMedicaoContrato.lCodigo <> 0 And iCodigoAlterado <> 0 Then
    
        lErro = Traz_Medicao_Tela(objMedicaoContrato)
        If lErro <> SUCESSO Then gError 131057

        Call ComandoSeta_Fechar(Me.Name)
        
    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 131057

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155079)

    End Select
    
    Exit Sub

End Sub

Private Sub CodigoLabel_Click()

Dim objMedicaoContrato As New ClassMedicaoContrato
Dim colSelecao As New Collection

    If Len(Trim(Codigo.Text)) > 0 Then
        objMedicaoContrato.lCodigo = StrParaLong(Codigo.Text)
    End If
    
    Call Chama_Tela("MedicaoContratosLista", colSelecao, objMedicaoContrato, objEventoCodigo)

End Sub

Private Sub Contrato_Change()

    iAlterado = REGISTRO_ALTERADO
    iContratoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Data_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim objMedicaoContrato As ClassMedicaoContrato
Dim bCancel As Boolean

    Set objMedicaoContrato = obj1

    Codigo.PromptInclude = False
    Codigo.Text = objMedicaoContrato.lCodigo
    Codigo.PromptInclude = True

    Call Codigo_Validate(bCancel)

    Exit Sub

End Sub

Private Sub objEventoContrato_evSelecao(obj1 As Object)

Dim objContrato As ClassContrato
Dim bCancel As Boolean

    Set objContrato = obj1

    Contrato.Text = objContrato.sCodigo
    Call Contrato_Validate(bCancel)

    Exit Sub
    
End Sub

Private Sub Contrato_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objContrato As New ClassContrato

On Error GoTo Erro_Contrato_Validate
    
    If Len(Trim(Contrato.Text)) <> 0 And iContratoAlterado = REGISTRO_ALTERADO Then

        objContrato.sCodigo = Contrato.Text
        objContrato.iFilialEmpresa = giFilialEmpresa
        
        lErro = Traz_Contrato_Tela(objContrato)
        If lErro <> SUCESSO Then gError 129636
         
        Call BotaoLimparGrid_Click
            
        Call ComandoSeta_Fechar(Me.Name)
        
    End If

    Exit Sub

Erro_Contrato_Validate:

    Cancel = True

    Select Case gErr

        Case 129636

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155080)

    End Select

End Sub

Private Sub ContratoLabel_Click()

Dim objContrato As New ClassContrato
Dim colSelecao As New Collection

    If Len(Trim(Contrato.Text)) > 0 Then
        objContrato.sCodigo = Contrato.Text
    End If

    Call Chama_Tela("ContratosLista", colSelecao, objContrato, objEventoContrato)

End Sub

Private Function Traz_Contrato_Tela(objContrato As ClassContrato) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_Contrato_Tela

    lErro = CF("Contrato_Le", objContrato)
    If lErro <> SUCESSO And lErro <> 129332 Then gError 129637
    
    'Contrato Não Cadastrado
    If lErro = 129332 Then gError 131060
    
    If objContrato.iTipo <> CONTRATOS_RECEBER Then gError 132906
    
    With objContrato
            
        If .lCliente <> 0 Then
            Call Cliente_Formata(.lCliente)
            Call Filial_Formata(FilCliContrato, .iFilCli)
        Else
            CliContrato.Text = ""
            FilCliContrato.Text = ""
        End If
       
        Contrato.Text = .sCodigo
        DescContrato.Caption = .sDescricao
   
    End With
    
    iContratoAlterado = 0
         
    Traz_Contrato_Tela = SUCESSO

    Exit Function

Erro_Traz_Contrato_Tela:
     
    Traz_Contrato_Tela = gErr

    Select Case gErr
    
        Case 129637
        
        Case 131060
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_NAO_CADASTRADO", gErr, Contrato.Text)

        Case 132906
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_NAO_RECEBER", gErr, objContrato.sCodigo, objContrato.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155081)

    End Select

    Exit Function

End Function

Private Function Traz_Contrato_Tela2(objContrato As ClassContrato) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_Contrato_Tela2

    lErro = CF("Contrato_Le2", objContrato)
    If lErro <> SUCESSO And lErro <> 129332 Then gError 131058
        
    'Contrato Não Cadastrado
    If lErro = 129332 Then gError 131059

    If objContrato.iTipo <> CONTRATOS_RECEBER Then gError 132907

    With objContrato
            
        If .lCliente <> 0 Then
            Call Cliente_Formata(.lCliente)
            Call Filial_Formata(FilCliContrato, .iFilCli)
        Else
            CliContrato.Text = ""
            FilCliContrato.Text = ""
        End If
       
        Contrato.Text = .sCodigo
        DescContrato.Caption = .sDescricao
   
    End With
    
    iContratoAlterado = 0
         
    Traz_Contrato_Tela2 = SUCESSO

    Exit Function

Erro_Traz_Contrato_Tela2:
     
    Traz_Contrato_Tela2 = gErr

    Select Case gErr
    
        Case 131058
        
        Case 131059
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_NAO_CADASTRADO", gErr, Contrato.Text)

        Case 132907
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_NAO_RECEBER", gErr, objContrato.sCodigo, objContrato.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155082)

    End Select

    Exit Function

End Function

Private Function Traz_GridItens_Tela(objContrato As ClassContrato) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_GridItens_Tela

    lErro = CF("Contrato_Le", objContrato)
    If lErro <> SUCESSO And lErro <> 129332 Then gError 129638
    
    If lErro = SUCESSO And objContrato.iTipo <> CONTRATOS_RECEBER Then gError 132908
    
    lErro = Carrega_GridItens_Contratos(objContrato)
    If lErro <> SUCESSO Then gError 129639
         
    Traz_GridItens_Tela = SUCESSO

    Exit Function

Erro_Traz_GridItens_Tela:
     
    Traz_GridItens_Tela = gErr

    Select Case gErr
    
        Case 129638, 129639

        Case 132908
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_NAO_RECEBER", gErr, objContrato.sCodigo, objContrato.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155083)

    End Select

    Exit Function

End Function

Private Function Traz_Medicao_Tela(objMedicaoContrato As ClassMedicaoContrato) As Long

Dim lErro As Long
Dim objContrato As New ClassContrato

On Error GoTo Erro_Traz_Medicao_Tela
    
    lErro = CF("MedicaoContrato_Le", objMedicaoContrato)
    If lErro <> SUCESSO And lErro <> 129622 Then gError 129640
    
    'Não existe regitros de medição gravados com esse código
    If lErro = 129622 Then Exit Function
    
    Call Limpa_Tela_Medicao
    
    With objMedicaoContrato
    
        If .dtData <> DATA_NULA Then Data.Text = Format(.dtData, "dd/mm/yy")
            
        Codigo.PromptInclude = False
        Codigo.Text = .lCodigo
        Codigo.PromptInclude = True
    
        objContrato.lNumIntDoc = .lNumIntContrato

    End With
    
    lErro = Traz_Contrato_Tela2(objContrato)
    If lErro <> SUCESSO Then gError 129641
    
    lErro = Carrega_GridItens_Contratos(objContrato)
    If lErro <> SUCESSO Then gError 129642
    
    lErro = Carrega_GridItens(objMedicaoContrato)
    If lErro <> SUCESSO Then gError 129643
         
    iAlterado = 0
    iCodigoAlterado = 0
    iContratoAlterado = 0
         
    Traz_Medicao_Tela = SUCESSO

    Exit Function

Erro_Traz_Medicao_Tela:

    Traz_Medicao_Tela = gErr

    Select Case gErr
    
        Case 129640 To 129643
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155084)

    End Select

    Exit Function

End Function

Public Sub Cliente_Formata(lCliente As Long)

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Formata

    CliContrato.Text = lCliente

    'Busca o Cliente no BD
    lErro = TP_Cliente_Le(CliContrato, objcliente, iCodFilial)
    If lErro <> SUCESSO Then gError 129644

    lErro = CF("FiliaisClientes_Le_Cliente", objcliente, colCodigoNome)
    If lErro <> SUCESSO Then gError 129645

    'Preenche ComboBox de Filiais
    Call CF("Filial_Preenche", FilCliContrato, colCodigoNome)

    Exit Sub

Erro_Cliente_Formata:

    Select Case gErr

        Case 129644, 129645

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155085)

    End Select

    Exit Sub

End Sub

Public Sub Filial_Formata(objFilial As Object, iFilial As Integer)

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Formata

    objFilial.Text = CStr(iFilial)
    sCliente = CliContrato.Text
    objFilialCliente.iCodFilial = iFilial

    'Pesquisa se existe Filial com o código extraído
    lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
    If lErro <> SUCESSO And lErro <> 17660 Then gError 129646

    If lErro = 17660 Then gError 129647

    'Coloca na tela a Filial lida
    objFilial.Text = iFilial & SEPARADOR & objFilialCliente.sNome

    Exit Sub

Erro_Filial_Formata:

    Select Case gErr

        Case 129646

        Case 129647
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, objFilial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155086)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Private Function Inicializa_GridItens(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Alocação

Dim iIndice As Integer

    Set objGridItens.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Cobrar")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Qtde Contrato")
    objGridInt.colColuna.Add ("UM")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Custo")
    objGridInt.colColuna.Add ("Valor Cobrar")
    objGridInt.colColuna.Add ("Cobrança")
    objGridInt.colColuna.Add ("Dt Ref Ini")
    objGridInt.colColuna.Add ("Dt Ref Fim")
    objGridInt.colColuna.Add ("Status")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Item.Name)
    objGridInt.colCampo.Add (Cobrar.Name)
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (ItemDescricao.Name)
    objGridInt.colCampo.Add (QtdContratada.Name)
    objGridInt.colCampo.Add (UM.Name)
    objGridInt.colCampo.Add (Valor.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (Custo.Name)
    objGridInt.colCampo.Add (ValorCobrar.Name)
    objGridInt.colCampo.Add (DataCobranca.Name)
    objGridInt.colCampo.Add (DataRefIni.Name)
    objGridInt.colCampo.Add (DataRefFim.Name)
    objGridInt.colCampo.Add (Status.Name)

    'Colunas da Grid
    iGrid_Item_Col = 1
    iGrid_Cobrar_Col = 2
    iGrid_Produto_Col = 3
    iGrid_Descricao_Col = 4
    iGrid_QtdContratada_Col = 5
    iGrid_UM_Col = 6
    iGrid_Valor_Col = 7
    iGrid_Quantidade_Col = 8
    iGrid_Custo_Col = 9
    iGrid_ValorCobrar_Col = 10
    iGrid_DataCobranca_Col = 11
    iGrid_DataRefIni_Col = 12
    iGrid_DataRefFim_Col = 13
    iGrid_Status_Col = 14
 
    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = 6000

    objGridInt.objGrid.Rows = 500

    objGridInt.iLinhasVisiveis = 5
    
    'Habilita a execução da Rotina_Grid_Enable
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
       
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridItens)

    Inicializa_GridItens = SUCESSO

    Exit Function

End Function

Public Function Trata_Parametros(Optional objMedicaoContrato As ClassMedicaoContrato) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objMedicaoContrato Is Nothing) Then

        lErro = Traz_Medicao_Tela(objMedicaoContrato)
        If lErro <> SUCESSO Then gError 129648
        
    End If

    iAlterado = 0
    iCodigoAlterado = 0
    iContratoAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
    
        Case 129648
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155087)

    End Select

    Exit Function

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objMedicaoContrato As New ClassMedicaoContrato

On Error GoTo Erro_Gravar_Registro

    If Len(Trim(Codigo.Text)) = 0 Then gError 129649
    If Len(Trim(Contrato.Text)) = 0 Then gError 129650
    If Len(Trim(Data.ClipText)) = 0 Then gError 129651
    
    lErro = Valida_Grid_Itens()
    If lErro <> SUCESSO Then gError 129653
       
    lErro = Move_Tela_Memoria(objMedicaoContrato)
    If lErro <> SUCESSO Then gError 129654
    
    GL_objMDIForm.MousePointer = vbHourglass
       
    lErro = CF("MedicaoContrato_Grava", objMedicaoContrato)
    If lErro <> SUCESSO Then gError 129655
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
        
        Case 129649
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_MEDICAOCONTRATO_PREENCHIDO", gErr)

        Case 129650
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CONTRATO_PREENCHIDO", gErr)
        
        Case 129651
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_MEDICAOCONTRATO_PREENCHIDA", gErr)

        Case 129653 To 129655
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155088)

    End Select

    Exit Function

End Function

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then
            Call CodigoLabel_Click
        ElseIf Me.ActiveControl Is Contrato Then
            Call ContratoLabel_Click
        End If
          
    End If

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RASTROPRODNFFAT
    Set Form_Load_Ocx = Me
    Caption = "Medição de Contrato a Receber"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ContratoMedicao"

End Function

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

Private Sub Data_GotFocus()

     Call MaskEdBox_TrataGotFocus(Data, iAlterado)

End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    'Verifica se a Data de Emissao foi digitada
    If Len(Trim(Data.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(Data.Text)
    If lErro <> SUCESSO Then gError 129657

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case gErr

        'se houve erro de crítica, segura o foco
        Case 129657

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155089)

    End Select

    Exit Sub

End Sub

'####################################################
'INÍCIO DOS BOTÕES UPDOWN
'####################################################
Private Sub UpDown_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 129661

    Exit Sub

Erro_UpDown_DownClick:

    Select Case gErr

        Case 129661

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155090)

    End Select

    Exit Sub

End Sub

Private Sub UpDown_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 129662

    Exit Sub

Erro_UpDown_UpClick:

    Select Case gErr

        Case 129662

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155091)

    End Select

    Exit Sub

End Sub
'####################################################
'FIM DOS BOTÕES UPDOWN
'####################################################



'#######################################################################
'INÍCIO DAS ROTINAS RELACIONADAS AS SETAS DO SISTEMA
'#######################################################################
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objMedicaoContrato As New ClassMedicaoContrato

On Error GoTo Erro_Tela_Preenche

    objMedicaoContrato.lCodigo = colCampoValor.Item("Codigo").vValor

    If objMedicaoContrato.lCodigo > 0 Then

        lErro = Traz_Medicao_Tela(objMedicaoContrato)
        If lErro <> SUCESSO Then gError 129663

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 129663

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155092)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objMedicaoContrato As New ClassMedicaoContrato

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "MedicaoContrato"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objMedicaoContrato)
    If lErro <> SUCESSO Then gError 129664

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objMedicaoContrato.lCodigo, 0, "Codigo"
      
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 129664

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155093)

    End Select

    Exit Sub

End Sub
'#######################################################################
'FIM DAS ROTINAS RELACIONADAS AS SETAS DO SISTEMA
'#######################################################################

'#######################################################################
'INÍCIO DAS FUNÇÕES DE SAÍDA DE CÉLULA
'#######################################################################
Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
                
        If objGridInt.objGrid Is GridItens Then
        
            Select Case GridItens.Col
    
                Case iGrid_Cobrar_Col
    
                    lErro = Saida_Celula_Cobrar(objGridInt)
                    If lErro <> SUCESSO Then gError 129665
                        
                Case iGrid_Descricao_Col
    
                    lErro = Saida_Celula_Descricao(objGridInt)
                    If lErro <> SUCESSO Then gError 129666
                        
                Case iGrid_Status_Col
    
                    lErro = Saida_Celula_Status(objGridInt)
                    If lErro <> SUCESSO Then gError 129667
                          
                Case iGrid_Produto_Col
    
                    lErro = Saida_Celula_Produto(objGridInt)
                    If lErro <> SUCESSO Then gError 129668
             
                Case iGrid_QtdContratada_Col
    
                    lErro = Saida_Celula_QtdContratada(objGridInt)
                    If lErro <> SUCESSO Then gError 129669
             
                Case iGrid_Quantidade_Col
    
                    lErro = Saida_Celula_Quantidade(objGridInt)
                    If lErro <> SUCESSO Then gError 129670
             
                Case iGrid_UM_Col
    
                    lErro = Saida_Celula_UM(objGridInt)
                    If lErro <> SUCESSO Then gError 129671
                
                Case iGrid_Valor_Col
    
                    lErro = Saida_Celula_Valor(objGridInt)
                    If lErro <> SUCESSO Then gError 129672
                
                Case iGrid_Custo_Col
    
                    lErro = Saida_Celula_Custo(objGridInt)
                    If lErro <> SUCESSO Then gError 129673
                
                Case iGrid_ValorCobrar_Col
    
                    lErro = Saida_Celula_ValorCobrar(objGridInt)
                    If lErro <> SUCESSO Then gError 129674
                 
                Case iGrid_Item_Col
    
                    lErro = Saida_Celula_Item(objGridInt)
                    If lErro <> SUCESSO Then gError 129675
            
                Case iGrid_DataRefIni_Col
    
                    lErro = Saida_Celula_DataRefIni(objGridInt)
                    If lErro <> SUCESSO Then gError 136067
                
                Case iGrid_DataRefFim_Col
    
                    lErro = Saida_Celula_DataRefFim(objGridInt)
                    If lErro <> SUCESSO Then gError 136068
                    
                Case iGrid_DataCobranca_Col
    
                    lErro = Saida_Celula_DataCobranca(objGridInt)
                    If lErro <> SUCESSO Then gError 136069
                    
             End Select
                
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 129676

        iAlterado = REGISTRO_ALTERADO

    End If
    
    Saida_Celula = SUCESSO
    
    Exit Function

Erro_Saida_Celula:
    
    Saida_Celula = gErr
    
    Select Case gErr

        Case 129665 To 129675, 136067 To 136069

        Case 129676
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155094)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    If Len(Trim(Quantidade.ClipText)) <> 0 Then

        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 131061

        Quantidade.Text = Formata_Estoque(Quantidade.Text)

        GridItens.TextMatrix(GridItens.Row, iGrid_ValorCobrar_Col) = Format(StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Valor_Col)) * StrParaDbl(Quantidade.Text), ValorCobrar.Format)
   
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129677

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 129677, 131061
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 155095)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_QtdContratada(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_QtdContratada

    Set objGridInt.objControle = QtdContratada

    If Len(Trim(QtdContratada.ClipText)) <> 0 Then

        lErro = Valor_Positivo_Critica(QtdContratada.Text)
        If lErro <> SUCESSO Then gError 129678

        QtdContratada.Text = Formata_Estoque(QtdContratada.Text)

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129679

    Saida_Celula_QtdContratada = SUCESSO

    Exit Function

Erro_Saida_Celula_QtdContratada:

    Saida_Celula_QtdContratada = gErr

    Select Case gErr

        Case 129678 To 129679
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 155096)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Valor(objGridInt As AdmGrid) As Long
'Faz a crítica da celula Valor do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridInt.objControle = Valor

    'Verifica se valor está preenchido
    If Len(Trim(Valor.Text)) > 0 Then
    
        'Critica se valor é positivo
        lErro = Valor_Positivo_Critica(Valor.Text)
        If lErro <> SUCESSO Then gError 129680

        Valor.Text = Format(Valor.Text, Valor.Format)
        
        'Acrescenta uma linha no Grid se for o caso
        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
              
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129681
        
    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = gErr

    Select Case gErr

        Case 129680, 129681
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155097)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Custo(objGridInt As AdmGrid) As Long
'Faz a crítica da celula Custo do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Custo

    Set objGridInt.objControle = Custo

    'Verifica se Custo está preenchido
    If Len(Trim(Custo.Text)) > 0 Then
    
        'Critica se Custo é positivo
        lErro = Valor_Positivo_Critica(Custo.Text)
        If lErro <> SUCESSO Then gError 129682

        Custo.Text = Format(Custo.Text, Custo.Format)
        
        'Acrescenta uma linha no Grid se for o caso
        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
              
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129683

    Saida_Celula_Custo = SUCESSO

    Exit Function

Erro_Saida_Celula_Custo:

    Saida_Celula_Custo = gErr

    Select Case gErr

        Case 129682 To 129683
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155098)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ValorCobrar(objGridInt As AdmGrid) As Long
'Faz a crítica da celula ValorCobrar do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_ValorCobrar

    Set objGridInt.objControle = ValorCobrar

    'Verifica se ValorCobrar está preenchido
    If Len(Trim(ValorCobrar.Text)) > 0 Then
    
        'Critica se ValorCobrar é positivo
        lErro = Valor_Positivo_Critica(ValorCobrar.Text)
        If lErro <> SUCESSO Then gError 129684

        ValorCobrar.Text = Format(ValorCobrar.Text, ValorCobrar.Format)
        
        'Acrescenta uma linha no Grid se for o caso
        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
              
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129685

    Saida_Celula_ValorCobrar = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorCobrar:

    Saida_Celula_ValorCobrar = gErr

    Select Case gErr

        Case 129684 To 129685
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155099)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Cobrar(objGridInt As AdmGrid) As Long
'Faz a crítica da celula Cobrar do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Cobrar

    Set objGridInt.objControle = Cobrar

    'Verifica se valor está preenchido
    If Cobrar.Value <> 0 Then
           
        'Acrescenta uma linha no Grid se for o caso
        If GridItens.Row - GridItens.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129686

    Saida_Celula_Cobrar = SUCESSO

    Exit Function

Erro_Saida_Celula_Cobrar:

    Saida_Celula_Cobrar = gErr

    Select Case gErr

        Case 129686
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155100)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Descricao(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Descricao

    Set objGridInt.objControle = ItemDescricao

    GridItens.TextMatrix(GridItens.Row, iGrid_Descricao_Col) = ItemDescricao.Text

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129687

    Saida_Celula_Descricao = SUCESSO

    Exit Function

Erro_Saida_Celula_Descricao:

    Saida_Celula_Descricao = gErr

    Select Case gErr

        Case 129687
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 155101)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_UM(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_UM

    Set objGridInt.objControle = UM

    GridItens.TextMatrix(GridItens.Row, iGrid_UM_Col) = UM.Text

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129688

    Saida_Celula_UM = SUCESSO

    Exit Function

Erro_Saida_Celula_UM:

    Saida_Celula_UM = gErr

    Select Case gErr

        Case 129688
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 155102)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Item(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Item

    Set objGridInt.objControle = Item

    For iIndice = 1 To objGridItens.iLinhasExistentes
        If iIndice <> GridItens.Row Then
            If GridItens.TextMatrix(iIndice, iGrid_Item_Col) = Item.Text Then gError 131064
        End If
    Next

    If Len(Trim(Item.Text)) <> 0 Then
        GridItens.TextMatrix(GridItens.Row, iGrid_Item_Col) = Item.Text
    
        lErro = Traz_ItemDeContrato(Contrato.Text, Item.Text, GridItens.Row)
        If lErro <> SUCESSO Then gError 129833
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129689

    Saida_Celula_Item = SUCESSO

    Exit Function

Erro_Saida_Celula_Item:

    Saida_Celula_Item = gErr

    Select Case gErr

        Case 129689, 129833
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 131064
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_JA_EXISTENTE", gErr, Item.Text, iIndice)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 155103)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Status(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Saida_Celula_Status

    Set objGridInt.objControle = Status

        'Verifica se a Condicaopagamento foi preenchida
        If Len(Trim(Status.Text)) = 0 Then
    
        'Verifica se é uma Condicaopagamento selecionada
        If Status.Text <> Status.List(Status.ListIndex) Then
    
            'Tenta selecionar na combo
            lErro = Combo_Seleciona(Status, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 129690
            
            GridItens.TextMatrix(GridItens.Row, iGrid_Status_Col) = Status.Text

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129691

    Saida_Celula_Status = SUCESSO

    Exit Function

Erro_Saida_Celula_Status:

    Saida_Celula_Status = gErr

    Select Case gErr

        Case 129690 To 129691
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 155104)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult
Dim iIndice As Integer
Dim sProduto As String

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto

    'Verifica se o produto existe e foi preenchido
    lErro = CF("Produto_Critica", Produto.Text, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 25041 Then gError 129692

    'se o produto não estiver cadastrado
    If lErro = 25041 Then gError 129693
            
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
        lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProduto)
        If lErro <> SUCESSO Then gError 129694
    
        Produto.PromptInclude = False
        Produto.Text = sProduto
        Produto.PromptInclude = True
   
        GridItens.TextMatrix(GridItens.Row, iGrid_Descricao_Col) = objProduto.sDescricao
        GridItens.TextMatrix(GridItens.Row, iGrid_UM_Col) = objProduto.sSiglaUMEstoque
                                                              
        'ALTERAÇÃO DE LINHAS EXISTENTES
        If (GridItens.Row - GridItens.FixedRows) = objGridItens.iLinhasExistentes Then
            objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
        End If
            
    Else
        
        GridItens.TextMatrix(GridItens.Row, iGrid_UM_Col) = ""
        GridItens.TextMatrix(GridItens.Row, iGrid_Descricao_Col) = ""
        GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = ""
        GridItens.TextMatrix(GridItens.Row, iGrid_Status_Col) = ""
    
    End If
    
    For iIndice = 1 To objGridItens.iLinhasExistentes
        If iIndice <> GridItens.Row Then
            If GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text Then gError 129695
        End If
    Next
        
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129696
    
    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 129692, 129696
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 129693
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)

            If vbMsg = vbYes Then
                objProduto.sCodigo = Produto.Text
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
                 
        Case 129694
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
         
        Case 129695
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_EXISTENTE", gErr, Produto.Text, Produto.Text, iIndice)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 155105)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DataRefIni(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DataRefIni

    Set objGridInt.objControle = DataRefIni

    'verifica se a data está preenchida
    If Len(Trim(DataRefIni.ClipText)) > 0 Then

        'verifica se a data é válida
        lErro = Data_Critica(DataRefIni.Text)
        If lErro <> SUCESSO Then gError 136060
                
        'ALTERAÇÃO DE LINHAS EXISTENTES
        If (GridItens.Row - GridItens.FixedRows) = objGridItens.iLinhasExistentes Then
            objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 136061

    Saida_Celula_DataRefIni = SUCESSO

    Exit Function

Erro_Saida_Celula_DataRefIni:

    Saida_Celula_DataRefIni = gErr

    Select Case gErr

        Case 136060, 136061
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155106)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DataRefFim(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DataRefFim

    Set objGridInt.objControle = DataRefFim

    'verifica se a data está preenchida
    If Len(Trim(DataRefFim.ClipText)) > 0 Then

        'verifica se a data é válida
        lErro = Data_Critica(DataRefFim.Text)
        If lErro <> SUCESSO Then gError 136062
                
        'ALTERAÇÃO DE LINHAS EXISTENTES
        If (GridItens.Row - GridItens.FixedRows) = objGridItens.iLinhasExistentes Then
            objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 136063

    Saida_Celula_DataRefFim = SUCESSO

    Exit Function

Erro_Saida_Celula_DataRefFim:

    Saida_Celula_DataRefFim = gErr

    Select Case gErr

        Case 136062, 136063
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155107)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DataCobranca(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DataCobranca

    Set objGridInt.objControle = DataCobranca

    'verifica se a data está preenchida
    If Len(Trim(DataCobranca.ClipText)) > 0 Then

        'verifica se a data é válida
        lErro = Data_Critica(DataCobranca.Text)
        If lErro <> SUCESSO Then gError 136064
                
        'ALTERAÇÃO DE LINHAS EXISTENTES
        If (GridItens.Row - GridItens.FixedRows) = objGridItens.iLinhasExistentes Then
            objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 136065

    Saida_Celula_DataCobranca = SUCESSO

    Exit Function

Erro_Saida_Celula_DataCobranca:

    Saida_Celula_DataCobranca = gErr

    Select Case gErr

        Case 136064, 136065
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155108)

    End Select

    Exit Function

End Function
'#######################################################################
'FIM DAS FUNÇÕES DE SAÍDA DE CÉLULA
'#######################################################################

'#######################################################################
'INÍCIO DO SCRIPT DO GRID
'#######################################################################
Public Sub GridItens_GotFocus()

    Call Grid_Recebe_Foco(objGridItens)

End Sub

Public Sub GridItens_EnterCell()

    Call Grid_Entrada_Celula(objGridItens, iAlterado)

End Sub

Public Sub GridItens_LeaveCell()

    Call Saida_Celula(objGridItens)

End Sub

Private Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Public Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinha As Integer
Dim iLinhaAtual As Integer
Dim iLinhasExistentesAnterior As Integer

    Call Grid_Trata_Tecla1(KeyCode, objGridItens)

    iLinhaAtual = GridItens.Row
    
    iLinhasExistentesAnterior = objGridItens.iLinhasExistentes

End Sub

Public Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Public Sub GridItens_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridItens)

End Sub

Public Sub GridItens_RowColChange()

    Call Grid_RowColChange(objGridItens)

End Sub

Public Sub GridItens_Scroll()

    Call Grid_Scroll(objGridItens)

End Sub
'#######################################################################
'FIM DO SCRIPT DO GRID
'#######################################################################

'#######################################################################
'INÍCIO DO SCRIPT PARA CAMPOS DO GRID
'#######################################################################
Public Sub Status_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Status_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Status_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Status_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Status
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Quantidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Quantidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub UM_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UM_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub UM_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub UM_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = UM
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Custo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Custo_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Custo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Custo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Custo
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ValorCobrar_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorCobrar_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub ValorCobrar_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub ValorCobrar_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = ValorCobrar
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Valor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Valor_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Valor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Valor
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Cobrar_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Cobrar_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Cobrar_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Cobrar_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Cobrar
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub QtdContratada_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QtdContratada_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub QtdContratada_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub QtdContratada_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = QtdContratada
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ItemDescricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ItemDescricao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub ItemDescricao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub ItemDescricao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = ItemDescricao
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Produto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Produto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Item_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Item_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Item_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Item_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Item
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub DataRefIni_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataRefIni_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub DataRefIni_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub DataRefIni_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DataRefIni
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub DataRefFim_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataRefFim_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub DataRefFim_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub DataRefFim_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DataRefFim
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub DataCobranca_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataCobranca_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub DataCobranca_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub DataCobranca_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DataCobranca
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub
'#######################################################################
'FIM DO SCRIPT PARA CAMPOS DO GRID
'#######################################################################

'#######################################################################
'INÍCIO SCRIPT DE BOTÕES SUPERIORES
'#######################################################################
Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 129697

    Call Limpa_Tela_Medicao

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 129697

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 155109)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 129698

    Call Limpa_Tela_Medicao
    
    iAlterado = 0
    iCodigoAlterado = 0
    iContratoAlterado = 0
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 129698

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155110)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim objMedicaoContrato As New ClassMedicaoContrato
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'Verifica se o código foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 129699

    objMedicaoContrato.lCodigo = Codigo.Text
    
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_MedicaoContrato", objMedicaoContrato.lCodigo)

    If vbMsgRes = vbYes Then

        'Exclui a producao
        lErro = CF("MedicaoContrato_Exclui", objMedicaoContrato)
        If lErro <> SUCESSO Then gError 129700

        Call Limpa_Tela_Medicao
    End If
    
    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 129699
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 129700
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155111)
    
    End Select
    
    Exit Sub
End Sub
'#######################################################################
'FIM SCRIPT DE BOTÕES SUPERIORES
'#######################################################################


Private Function Valida_Grid_Itens() As Long

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Valida_Grid_Itens

    'Verifica se há itens no grid
    If objGridItens.iLinhasExistentes = 0 Then gError 129701

    'para cada item do grid
    For iIndice = 1 To objGridItens.iLinhasExistentes

        If StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col)) = 0 Then gError 136050
        If StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataCobranca_Col)) = DATA_NULA Then gError 136051
        If StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataRefIni_Col)) = DATA_NULA Then gError 136052
        If StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataRefFim_Col)) = DATA_NULA Then gError 136053
        If StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataRefFim_Col)) < StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataRefIni_Col)) Then gError 136054
  
    Next

    Valida_Grid_Itens = SUCESSO

    Exit Function

Erro_Valida_Grid_Itens:

    Valida_Grid_Itens = gErr

    Select Case gErr
    
        Case 129701
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_ITENS_MEDICAOCONTRATOS", gErr)

        Case 136050
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_ITEM_NAO_PREENCHIDA", gErr, iIndice)
        
        Case 136051
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAPROXCOBRANCA_NAO_PREENCHIDO", gErr, iIndice)
        
        Case 136052
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAREFINI_NAO_PREENCHIDO", gErr, iIndice)
        
        Case 136053
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAREFFIM_NAO_PREENCHIDO", gErr, iIndice)

        Case 136054
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAREFINI_MAIOR_DATAREFFIM", gErr, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155112)

    End Select

    Exit Function

End Function


Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)

Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable

    'Formata o produto do grid de itens
    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 129702

    Select Case objControl.Name

        Case Produto.Name
                
            Produto.Enabled = False

        Case UM.Name
                    
            UM.Enabled = False

        Case Quantidade.Name
                
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Or left(GridItens.TextMatrix(iLinha, 0), 1) = "#" Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
                     
        Case ItemDescricao.Name
                
            objControl.Enabled = False
            
        Case Cobrar.Name
                
            objControl.Enabled = False

        Case Valor.Name
                
            objControl.Enabled = False

        Case ValorCobrar.Name
                
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case QtdContratada.Name
                
            objControl.Enabled = False
            
        Case Item.Name
                
            'Se o produto estiver não preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
            
        Case Status.Name
                
            objControl.Enabled = False
   
        Case Custo.Name
                
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case Else
            
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
           
    End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 129702

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155113)

    End Select

    Exit Sub

End Sub

Private Function Carrega_GridItens_Contratos(objContrato As ClassContrato) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objItensDeContrato As ClassItensDeContrato
Dim sProduto As String
Dim sProdutoAux As String

On Error GoTo Erro_Carrega_GridItens_Contratos

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridItens)

    For Each objItensDeContrato In objContrato.colItens
           
        With objItensDeContrato
        
            If .iMedicao = 1 Then
            
                iIndice = iIndice + 1
               
                sProdutoAux = objItensDeContrato.sProduto
               
                lErro = Mascara_RetornaProdutoTela(sProdutoAux, sProduto)
                If lErro <> SUCESSO Then gError 129703
        
                'Mascara o produto enxuto
                Produto.PromptInclude = False
                Produto.Text = sProduto
                Produto.PromptInclude = True
    
                GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
                GridItens.TextMatrix(iIndice, iGrid_Descricao_Col) = .sDescProd
                       
                GridItens.TextMatrix(iIndice, iGrid_Cobrar_Col) = .iCobrar
                                
                GridItens.TextMatrix(iIndice, iGrid_QtdContratada_Col) = Formata_Estoque(.dQuantidade)
                GridItens.TextMatrix(iIndice, iGrid_UM_Col) = .sUM
                GridItens.TextMatrix(iIndice, iGrid_Valor_Col) = Format(.dValor, Valor.Format)
                GridItens.TextMatrix(iIndice, iGrid_Item_Col) = .iSeq
            
                Status.Text = NAO_COBRADO
                lErro = Combo_Seleciona_Grid(Status, NAO_COBRADO)
                If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 129877
                GridItens.TextMatrix(iIndice, iGrid_Status_Col) = Status.Text
                
                GridItens.TextMatrix(iIndice, iGrid_DataCobranca_Col) = Format(.dtDataProxCobranca, "dd/mm/yyyy")
                GridItens.TextMatrix(iIndice, iGrid_DataRefIni_Col) = Format(.dtDataRefIni, "dd/mm/yyyy")
                GridItens.TextMatrix(iIndice, iGrid_DataRefFim_Col) = Format(.dtDataRefFim, "dd/mm/yyyy")
           
            End If
            
        End With
            
    Next
       
    Call Grid_Refresh_Checkbox(objGridItens)

    'Atualiza o número de linhas existentes
    objGridItens.iLinhasExistentes = iIndice
    
    Carrega_GridItens_Contratos = SUCESSO
        
    Exit Function

Erro_Carrega_GridItens_Contratos:

    Call Grid_Limpa(objGridItens)
    
    Carrega_GridItens_Contratos = gErr

    Select Case gErr
    
        Case 129703, 129877
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155114)

    End Select

    Exit Function

End Function

Private Function Traz_ItemDeContrato(sContrato As String, iSeq As Integer, iLinha As Integer) As Long

Dim lErro As Long
Dim objContrato As New ClassContrato
Dim objItensDeContrato As ClassItensDeContrato
Dim sProduto As String
Dim sProdutoAux As String
Dim bAchou As Boolean

On Error GoTo Erro_Traz_ItemDeContrato

    objContrato.sCodigo = sContrato
    objContrato.iFilialEmpresa = giFilialEmpresa

    lErro = CF("Contrato_Le", objContrato)
    If lErro <> SUCESSO And lErro <> 129332 Then gError 129834

    'Contrato Não Cadastrado
    If lErro = 129332 Then gError 132910
    
    If objContrato.iTipo <> CONTRATOS_RECEBER Then gError 132909

    bAchou = False

    For Each objItensDeContrato In objContrato.colItens
           
        With objItensDeContrato
        
            If .iSeq = iSeq Then
            
                If .iMedicao <> 1 Then gError 129835
                
                bAchou = True
            
                sProdutoAux = objItensDeContrato.sProduto
               
                lErro = Mascara_RetornaProdutoTela(sProdutoAux, sProduto)
                If lErro <> SUCESSO Then gError 129836
        
                'Mascara o produto enxuto
                Produto.PromptInclude = False
                Produto.Text = sProduto
                Produto.PromptInclude = True
    
                GridItens.TextMatrix(iLinha, iGrid_Produto_Col) = Produto.Text
                GridItens.TextMatrix(iLinha, iGrid_Descricao_Col) = .sDescProd
                       
                GridItens.TextMatrix(iLinha, iGrid_Cobrar_Col) = .iCobrar
                                
                GridItens.TextMatrix(iLinha, iGrid_QtdContratada_Col) = Formata_Estoque(.dQuantidade)
                GridItens.TextMatrix(iLinha, iGrid_UM_Col) = .sUM
                GridItens.TextMatrix(iLinha, iGrid_Valor_Col) = Format(.dValor, Valor.Format)
            
                Status.Text = NAO_COBRADO
                lErro = Combo_Seleciona_Grid(Status, NAO_COBRADO)
                If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 129876
                GridItens.TextMatrix(iLinha, iGrid_Status_Col) = Status.Text
            
                GridItens.TextMatrix(iLinha, iGrid_DataCobranca_Col) = Format(.dtDataProxCobranca, "dd/mm/yyyy")
                GridItens.TextMatrix(iLinha, iGrid_DataRefIni_Col) = Format(.dtDataRefIni, "dd/mm/yyyy")
                GridItens.TextMatrix(iLinha, iGrid_DataRefFim_Col) = Format(.dtDataRefFim, "dd/mm/yyyy")
            
            End If
            
        End With
            
    Next
    
    If Not bAchou Then gError 129841
       
    Call Grid_Refresh_Checkbox(objGridItens)

    'Atualiza o número de linhas existentes
    objGridItens.iLinhasExistentes = iLinha
    
    Traz_ItemDeContrato = SUCESSO
        
    Exit Function

Erro_Traz_ItemDeContrato:
    
    Traz_ItemDeContrato = gErr

    Select Case gErr
    
        Case 129834, 129836, 129876
        
        Case 129835
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_SEM_MEDICAO", gErr, iSeq)
            
        Case 129841
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMDECONTRATO_NAO_CADASTRADO", gErr, iSeq)
        
        Case 132910
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_NAO_CADASTRADO", gErr, objContrato.sCodigo)
        
        Case 132909
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_NAO_RECEBER", gErr, objContrato.sCodigo, objContrato.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155115)

    End Select

    Exit Function

End Function

Private Function Carrega_GridItens(objMedicaoContrato As ClassMedicaoContrato) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objItensDeMedicao As ClassItensMedCtr
Dim objNF As ClassNFiscal
Dim objItemNF As ClassItemNF

On Error GoTo Erro_Carrega_GridItens

    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        For Each objItensDeMedicao In objMedicaoContrato.colItens
               
            If GridItens.TextMatrix(iIndice, iGrid_Item_Col) = objItensDeMedicao.objItensDeContrato.iSeq Then
                                
                GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItensDeMedicao.dQuantidade)
                GridItens.TextMatrix(iIndice, iGrid_ValorCobrar_Col) = Format(objItensDeMedicao.dVlrCobrar, ValorCobrar.Format)
                GridItens.TextMatrix(iIndice, iGrid_Custo_Col) = Format(objItensDeMedicao.dCusto, Custo.Format)
                            
                Set objItemNF = New ClassItemNF
                Set objNF = New ClassNFiscal
                            
                objItemNF.objCobrItensContrato.colMedicoes.Add objItensDeMedicao
                            
                objItemNF.objCobrItensContrato.lNumIntItensContrato = objItensDeMedicao.lNumIntItensContrato
                objItemNF.objCobrItensContrato.lMedicao = objItensDeMedicao.lMedicao
            
                'Verifica se o item já foi faturado
                lErro = CF("ItensDeContrato_Le_DadosFatura", objNF, objItemNF)
                If lErro <> SUCESSO And lErro <> 129904 And lErro <> 129907 And lErro <> 129908 Then gError 131065
                If lErro = SUCESSO Then
                    objItensDeMedicao.iStatus = COBRADO
                Else
                    objItensDeMedicao.iStatus = NAO_COBRADO
                End If
                Status.Text = objItensDeMedicao.iStatus
                lErro = Combo_Seleciona_Grid(Status, objItensDeMedicao.iStatus)
                If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 129704
                GridItens.TextMatrix(iIndice, iGrid_Status_Col) = Status.Text

                GridItens.TextMatrix(iIndice, iGrid_DataCobranca_Col) = Format(objItensDeMedicao.dtDataCobranca, "dd/mm/yyyy")
                GridItens.TextMatrix(iIndice, iGrid_DataRefIni_Col) = Format(objItensDeMedicao.dtDataRefIni, "dd/mm/yyyy")
                GridItens.TextMatrix(iIndice, iGrid_DataRefFim_Col) = Format(objItensDeMedicao.dtDataRefFim, "dd/mm/yyyy")

            End If
            
        Next
            
    Next
       
    Call Grid_Refresh_Checkbox(objGridItens)
    
    Carrega_GridItens = SUCESSO
        
    Exit Function

Erro_Carrega_GridItens:

    Carrega_GridItens = gErr

    Select Case gErr
    
        Case 129704, 131065
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155116)

    End Select

    Exit Function

End Function

Private Function Carrega_ComboStatus(objCombo As ComboBox) As Long

Dim lErro As Long
   
On Error GoTo Erro_Carrega_ComboStatus
    
    objCombo.AddItem COBRADO & SEPARADOR & STRING_COBRADO
    objCombo.ItemData(objCombo.NewIndex) = COBRADO
    
    objCombo.AddItem NAO_COBRADO & SEPARADOR & STRING_NAO_COBRADO
    objCombo.ItemData(objCombo.NewIndex) = NAO_COBRADO
    
    Carrega_ComboStatus = SUCESSO
    
    Exit Function
    
Erro_Carrega_ComboStatus:

    Carrega_ComboStatus = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155117)

    End Select

    Exit Function
    
End Function


'#######################################################################
'INÍCIO DO SCRIPT PARA MODO DE EDICAO
'#######################################################################
Private Sub DescContrato_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescContrato, Source, X, Y)
End Sub

Private Sub DescContrato_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescContrato, Button, Shift, X, Y)
End Sub

Private Sub CodigoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CodigoLabel, Source, X, Y)
End Sub

Private Sub CodigoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CodigoLabel, Button, Shift, X, Y)
End Sub

Private Sub Codigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Codigo, Source, X, Y)
End Sub

Private Sub ContratoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContratoLabel, Source, X, Y)
End Sub

Private Sub ContratoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContratoLabel, Button, Shift, X, Y)
End Sub

Private Sub Contrato_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Contrato, Source, X, Y)
End Sub


Private Sub CliContratoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CliContratoLabel, Source, X, Y)
End Sub

Private Sub CliContratoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CliContratoLabel, Button, Shift, X, Y)
End Sub

Private Sub CliContrato_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CliContrato, Source, X, Y)
End Sub

Private Sub FilCliContratoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilCliContratoLabel, Source, X, Y)
End Sub

Private Sub FilCliContratoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilCliContratoLabel, Button, Shift, X, Y)
End Sub

Private Sub FilCliContrato_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilCliContrato, Source, X, Y)
End Sub
'#######################################################################
'FIM DO SCRIPT PARA MODO DE EDICAO
'#######################################################################


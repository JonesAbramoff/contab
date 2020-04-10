VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ComissoesPagOcx 
   ClientHeight    =   5190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7035
   LockControls    =   -1  'True
   ScaleHeight     =   5190
   ScaleWidth      =   7035
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   555
      Left            =   2347
      Picture         =   "ComissoesPagOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4500
      Width           =   1035
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancela"
      Height          =   555
      Left            =   3652
      Picture         =   "ComissoesPagOcx.ctx":015A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4500
      Width           =   1035
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comissões geradas"
      Height          =   1455
      Left            =   135
      TabIndex        =   15
      Top             =   1320
      Width           =   4245
      Begin VB.Frame Frame1 
         Caption         =   "Pela"
         Height          =   1155
         Left            =   2175
         TabIndex        =   16
         Top             =   165
         Width           =   1815
         Begin VB.OptionButton ComissaoEmissaoOuBaixa 
            Caption         =   "Ambos"
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
            Left            =   225
            TabIndex        =   7
            Top             =   840
            Value           =   -1  'True
            Width           =   930
         End
         Begin VB.OptionButton ComissaoPelaBaixa 
            Caption         =   "Baixa (Pagto)"
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
            Left            =   225
            TabIndex        =   6
            Top             =   540
            Width           =   1470
         End
         Begin VB.OptionButton ComissaoPelaEmissao 
            Caption         =   "Emissão"
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
            Left            =   225
            TabIndex        =   5
            Top             =   255
            Width           =   1065
         End
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   300
         Left            =   1695
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   300
         Left            =   540
         TabIndex        =   3
         Top             =   360
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   300
         Left            =   1725
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   930
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   300
         Left            =   540
         TabIndex        =   4
         Top             =   930
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Até:"
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
         Left            =   120
         TabIndex        =   20
         Top             =   983
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "De:"
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
         Left            =   165
         TabIndex        =   21
         Top             =   413
         Width           =   315
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Vendedor"
      Height          =   1455
      Left            =   135
      TabIndex        =   18
      Top             =   2865
      Width           =   4245
      Begin MSMask.MaskEdBox VendedorDe 
         Height          =   300
         Left            =   480
         TabIndex        =   8
         Top             =   360
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox VendedorAte 
         Height          =   300
         Left            =   480
         TabIndex        =   9
         Top             =   885
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Até:"
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
         Left            =   75
         TabIndex        =   22
         Top             =   945
         Width           =   360
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "De:"
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
         Left            =   120
         TabIndex        =   23
         Top             =   420
         Width           =   315
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Ação desejada"
      Height          =   1095
      Left            =   135
      TabIndex        =   13
      Top             =   120
      Width           =   4245
      Begin VB.OptionButton OpcaoCancelarBaixa 
         Caption         =   "Cancelar Pagamento"
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
         Left            =   135
         TabIndex        =   2
         Top             =   720
         Width           =   2115
      End
      Begin VB.OptionButton OpcaoBaixar 
         Caption         =   "Registrar Pagamento em:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   135
         TabIndex        =   0
         Top             =   345
         Value           =   -1  'True
         Width           =   2505
      End
      Begin MSComCtl2.UpDown UpDown3 
         Height          =   300
         Left            =   3765
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataBaixa 
         Height          =   300
         Left            =   2625
         TabIndex        =   1
         Top             =   315
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
   End
   Begin VB.ListBox VendedoresList 
      Height          =   3840
      IntegralHeight  =   0   'False
      Left            =   4500
      TabIndex        =   10
      Top             =   465
      Width           =   2400
   End
   Begin VB.Label Label13 
      Caption         =   "Vendedores"
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
      Left            =   4500
      TabIndex        =   24
      Top             =   225
      Width           =   1440
   End
End
Attribute VB_Name = "ComissoesPagOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim giFocoInicial As Integer
Dim giVendDeAlt As Integer 'Verifica se houve alguma alteração no VendedorDe
Dim giVendAteAlt As Integer 'Verifica se houve alguma alteração no VendedorAte

Private Sub BotaoCancela_Click()
    
    Unload Me

End Sub

Private Sub BotaoOK_Click()
'Executa a baixa ou o cancelamento das comissões

Dim lErro As Long
Dim objComissoesPag As New ClassComissoesPag

'Dim dtDataBaixa As Date, iCodVendedorIni As String, iCodVendedorFim As Integer, dtComisGeradasDe As Date, dtComisGeradasAte As Date, iTipo As Integer

On Error GoTo Erro_BotaoOK_Click

    'Exibe a ampulheta como ponteiro do mouse
    MousePointer = vbHourglass
    
    'Critica os dados da tela para verificar se é possível
    'efetuar a baixa / cancelamento de baixa das comissões
    lErro = ComissoesPag_Critica()
    If lErro <> SUCESSO Then gError 102048
        
    'Se a opção baixar com data estiver selecionada
    If OpcaoBaixar.Value Then
        
        'Guarda a data no obj
        objComissoesPag.dtDataBaixa = CDate(DataBaixa.Text)
    
    'Senão
    Else
        
        'Guarda data nula no obj
        objComissoesPag.dtDataBaixa = DATA_NULA
        
    End If
      
    'Guarda no obj os dados restantes da tela
    objComissoesPag.iFilialEmpresa = giFilialEmpresa
    objComissoesPag.iCodVendedorIni = Codigo_Extrai(VendedorDe.Text)
    objComissoesPag.iCodVendedorFim = Codigo_Extrai(VendedorAte.Text)
    objComissoesPag.dtComisGeradasDe = CDate(DataInicial.Text)
    objComissoesPag.dtComisGeradasAte = CDate(DataFinal.Text)
    
    'Guarda no obj o tipo de comissão que será baixada ou cancelada
    If ComissaoPelaEmissao.Value Then
        objComissoesPag.iTipo = COMISSAO_EMISSAO
        
    ElseIf ComissaoPelaBaixa.Value Then
        objComissoesPag.iTipo = COMISSAO_BAIXA
        
    Else
        objComissoesPag.iTipo = COMISSAO_AMBOS
        
    End If
    
    'Efetua a gravação da baixa ou cancelamento no BD
    lErro = CF("ComissoesVendasBaixar", objComissoesPag)
    If lErro <> SUCESSO Then gError 102050
        
    'Exibe o ponteiro padrão do mouse
    MousePointer = vbDefault
    
    'Exibe um aviso para o usuário indicando que a rotina foi executada com Sucesso
    Call Rotina_Aviso(vbOKOnly, "AVISO_OPERACAO_SUCESSO")
        
    'Fecha a tela
    Unload Me
    
    Exit Sub

Erro_BotaoOK_Click:

    Select Case gErr

        Case 102048, 102050
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154414)

    End Select

    'Exibe o ponteiro padrão do mouse
    MousePointer = vbDefault
    
    Exit Sub

End Sub
    
Private Sub DataBaixa_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataBaixa)
    
End Sub

Private Sub DataBaixa_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataBaixa_Validate

    If Len(DataBaixa.ClipText) = 0 Then Exit Sub
    
    'Se a opção baixar com data estiver selecionada
    If OpcaoBaixar.Value Then
    
        'Verificar se a Data de baixa é válida
        lErro = Data_Critica(DataBaixa.Text)
        If lErro <> SUCESSO Then Error 23195
        
    End If
    
    Exit Sub
    
Erro_DataBaixa_Validate:

    Cancel = True


    Select Case Err

        Case 23195

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154415)

    End Select

    Exit Sub

End Sub

Private Sub DataFinal_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataFinal)
    
End Sub

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    If Len(DataFinal.ClipText) = 0 Then Exit Sub
    
    'Verificar se a data inicial é valida
    lErro = Data_Critica(DataFinal.Text)
    If lErro <> SUCESSO Then Error 23199
    
    Exit Sub
    
Erro_DataFinal_Validate:

    Cancel = True


    Select Case Err
    
        Case 23199
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154416)
            
    End Select
    
    Exit Sub

End Sub

Private Sub DataInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicial)

End Sub

Private Sub DataInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataInicial_Validate

    If Len(DataInicial.ClipText) = 0 Then Exit Sub
    
    'Verificar se a data inicial é valida
    lErro = Data_Critica(DataInicial.Text)
    If lErro <> SUCESSO Then Error 23198
    
    Exit Sub
    
Erro_DataInicial_Validate:

    Cancel = True


    Select Case Err
    
        Case 23198
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154417)
            
    End Select
    
    Exit Sub
    
End Sub

Public Sub Form_Load()

Dim colCodigoDescricao As New AdmColCodigoNome
Dim lErro As Long
Dim objCodigoDescricao As AdmCodigoNome

On Error GoTo Erro_Form_Load

    giFocoInicial = 1
    
    'Inicia Data de Baixa e Data Final com data corrente
    DataBaixa.Text = Format(gdtDataAtual, "dd/mm/yy") 'CStr(gdtDataAtual)
    DataFinal.Text = Format(gdtDataAtual, "dd/mm/yy") 'CStr(gdtDataAtual)
    
    'Preenche a listbox vendedores
    'Le cada codigo e Nome Reduzido da tabela Vendedores
    lErro = CF("Cod_Nomes_Le", "Vendedores", "Codigo", "NomeReduzido", STRING_VENDEDOR_NOME_REDUZIDO, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 23182

    'preenche a listbox vendedores com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao

        VendedoresList.AddItem objCodigoDescricao.iCodigo & SEPARADOR & objCodigoDescricao.sNome
        
    Next
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 23182 'Tratada na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154418)

    End Select

    Unload Me

    Exit Sub

End Sub

Private Sub OpcaoBaixar_Click()

    DataBaixa.Enabled = True
    
End Sub

Private Sub OpcaoCancelarBaixa_Click()

    DataBaixa.Enabled = False

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick
    
    'Diminui a Data
    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 23191

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 23191
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154419)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick
    
    'Aumenta a Data
    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 23192

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 23192
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154420)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick
    
    'Diminui a Data
    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 23194

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case Err

        Case 23194
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154421)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    'Aumenta a Data
    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 23193

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case Err

        Case 23193
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154422)

    End Select

    Exit Sub

End Sub

Private Sub UpDown3_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown3_DownClick

    'Se a opcão baixar com data estiver selecionada
    If DataBaixa.Enabled Then
        
        'Diminui a Data
        lErro = Data_Up_Down_Click(DataBaixa, DIMINUI_DATA)
        If lErro <> SUCESSO Then Error 23189

    End If
    
    Exit Sub

Erro_UpDown3_DownClick:

    Select Case Err

        Case 23189
            DataBaixa.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154423)

    End Select

    Exit Sub

End Sub

Private Sub UpDown3_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown3_UpClick

    'Se a opcão baixar com data estiver selecionada
    If DataBaixa.Enabled Then
        
        'Aumenta a Data
        lErro = Data_Up_Down_Click(DataBaixa, AUMENTA_DATA)
        If lErro <> SUCESSO Then Error 23190

    End If
    
    Exit Sub

Erro_UpDown3_UpClick:

    Select Case Err

        Case 23190
            DataBaixa.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154424)

    End Select

    Exit Sub

End Sub

Private Sub VendedorAte_Change()

    giVendAteAlt = 1

End Sub

Private Sub VendedorAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_VendedorAte_Validate

    giFocoInicial = 0
    
    If giVendAteAlt = 1 Then
    
        If Len(Trim(VendedorAte.Text)) = 0 Then Exit Sub
        
        'Faz a Critica para Vendedor
        lErro = TP_Vendedor_Le2(VendedorAte, objVendedor)
        If lErro <> SUCESSO Then Error 23202
        
        giVendAteAlt = 0
        
    End If
        
    Exit Sub
    
Erro_VendedorAte_Validate:

    Cancel = True


    Select Case Err
    
        Case 23202
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154425)
            
    End Select
    
    Exit Sub
        
End Sub

Private Sub VendedorDe_Change()

    giVendDeAlt = 1
    
End Sub

Private Sub VendedorDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_VendedorDe_Validate

    giFocoInicial = 1
    
    If giVendDeAlt = 1 Then
    
        If Len(Trim(VendedorDe.Text)) = 0 Then Exit Sub
        
        'Faz a Critica para Vendedor
        lErro = TP_Vendedor_Le2(VendedorDe, objVendedor)
        If lErro <> SUCESSO Then Error 23201
        
        giVendDeAlt = 0
    
    End If
        
    Exit Sub
    
Erro_VendedorDe_Validate:

    Cancel = True


    Select Case Err
    
        Case 23201
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154426)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub VendedoresList_DblClick()
Dim sListBoxItem As String
Dim lErro As Long

On Error GoTo Erro_VendedoresList_DblClick

    'Se não há Vendedor selecionado sai da rotina
    If VendedoresList.ListIndex = -1 Then Exit Sub

    'Pega o Código do Vendedor e joga no Nome do Vendedor que teve o último foco
    sListBoxItem = Trim(VendedoresList.List(VendedoresList.ListIndex))

    'Verifica se o código do Vendedor está vazio
    If Len(sListBoxItem) = 0 Then Error 23183

    If giFocoInicial = 1 Then
        VendedorDe.Text = sListBoxItem
    Else
        VendedorAte.Text = sListBoxItem
    End If

    Exit Sub

Erro_VendedoresList_DblClick:

    Select Case Err

        Case 23183
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_VAZIO", Err, Error$)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154427)

    End Select

    Exit Sub
    
End Sub

'Criada por Luiz Nogueira em 10/05/02
Private Function ComissoesPag_Critica() As Long
'Critica os parâmetros da tela, verificando se é possível efetuar
'a baixa ou o cancelamento das comissões

On Error GoTo Erro_ComissoesPag_Critica

    'Verificar se Data Inicial foi informada
    If Len(DataInicial.ClipText) = 0 Then gError 102043
    
    'Verificar se Data Final foi informada
    If Len(DataFinal.ClipText) = 0 Then gError 102044
    
    'Verificar se Data Inicial é maior que a Data Final
    If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then gError 102045
        
    'Se a opção baixar com data estiver selecionada
    If OpcaoBaixar.Value Then
        
        'Verificar se a data foi informada
        If Len(DataBaixa.ClipText) = 0 Then gError 102046
        
        'Verificar se a data de baixa é maior ou igual a data final
        If CDate(DataFinal.Text) > CDate(DataBaixa.Text) Then gError 102047
    
    End If
    
    'Se Vendedor final e inicial estiverem preenchidos
    If Len(Trim(VendedorAte.Text)) <> 0 And Len(Trim(VendedorDe.Text)) <> 0 Then

        'Verificar se o código do Vendedor Final é maior que o código do Vendedor Inicial
        If Codigo_Extrai(VendedorAte.Text) < Codigo_Extrai(VendedorDe.Text) Then gError 102049

    End If
    
    ComissoesPag_Critica = SUCESSO
    
    Exit Function
    
Erro_ComissoesPag_Critica:

    ComissoesPag_Critica = gErr
    
    Select Case gErr

        Case 102043
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr, Error$)
            DataInicial.SetFocus
        
        Case 102044
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr, Error$)
            DataFinal.SetFocus
            
        Case 102045
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr, Error$)
            DataInicial.SetFocus
        
        Case 102046
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr, Error$)
            DataBaixa.SetFocus
        
        Case 102047
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_FINAL_MAIOR", gErr, Error$)
            DataFinal.SetFocus
        
        Case 102049
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_INICIAL_MAIOR", gErr, Error$)
            VendedorDe.SetFocus
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154428)

    End Select

    Exit Function

End Function

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO
    
End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_COMISSOES_PAG
    Set Form_Load_Ocx = Me
    Caption = "Pagto de Comissões"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ComissoesPag"
    
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




Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label22_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label22, Source, X, Y)
End Sub

Private Sub Label22_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label22, Button, Shift, X, Y)
End Sub

Private Sub Label21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label21, Source, X, Y)
End Sub

Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label21, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub


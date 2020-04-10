VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.UserControl PlanoContaOcx 
   ClientHeight    =   5460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8925
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5460
   ScaleWidth      =   8925
   Begin VB.ComboBox NaturezaSped 
      Height          =   315
      ItemData        =   "PlanoContaOcx.ctx":0000
      Left            =   1050
      List            =   "PlanoContaOcx.ctx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2040
      Width           =   2985
   End
   Begin VB.CheckBox FluxoCaixa 
      Caption         =   "Check1"
      Height          =   285
      Left            =   240
      TabIndex        =   11
      Top             =   3480
      Width           =   225
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6600
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "PlanoContaOcx.ctx":0093
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "PlanoContaOcx.ctx":01ED
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1110
         Picture         =   "PlanoContaOcx.ctx":0377
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "PlanoContaOcx.ctx":08A9
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox TipoConta 
      Height          =   315
      Left            =   1050
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1530
      Width           =   1530
   End
   Begin VB.ComboBox Natureza 
      Height          =   315
      Left            =   4170
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1530
      Width           =   1665
   End
   Begin VB.ComboBox Categoria 
      Height          =   315
      Left            =   1050
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1065
      Width           =   3015
   End
   Begin VB.ListBox ListaModulo 
      Columns         =   2
      Height          =   960
      Left            =   210
      Style           =   1  'Checkbox
      TabIndex        =   12
      Top             =   4155
      Width           =   5580
   End
   Begin VB.CheckBox Ativo 
      Caption         =   "Ativa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3120
      TabIndex        =   1
      Top             =   255
      Value           =   1  'Checked
      Width           =   795
   End
   Begin MSMask.MaskEdBox Conta 
      Height          =   315
      Left            =   1050
      TabIndex        =   0
      Top             =   195
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      AllowPrompt     =   -1  'True
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox DescConta 
      Height          =   315
      Left            =   1050
      TabIndex        =   2
      Top             =   630
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   40
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox HistPadrao 
      Height          =   315
      Left            =   1860
      TabIndex        =   7
      Top             =   2550
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin VB.CheckBox DigitoVerif 
      Caption         =   "Usa Dígito Verificador"
      Enabled         =   0   'False
      Height          =   315
      Left            =   285
      TabIndex        =   19
      Top             =   4080
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame SSFrame1 
      Caption         =   "Conta Simplificada"
      Height          =   1185
      Left            =   210
      TabIndex        =   20
      Top             =   4065
      Visible         =   0   'False
      Width           =   5640
      Begin VB.CheckBox DigitoVerifSimples 
         Caption         =   "Usa Dígito Verificador"
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   21
         Top             =   735
         Visible         =   0   'False
         Width           =   2280
      End
      Begin VB.Label LabelDigitoVerificadorSimples 
         Caption         =   "Valor do Dígito Verificador:"
         Enabled         =   0   'False
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
         Left            =   2730
         TabIndex        =   22
         Top             =   840
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.Label DigitoVerificadorSimples 
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   255
         Left            =   5175
         TabIndex        =   23
         Top             =   825
         Visible         =   0   'False
         Width           =   195
      End
   End
   Begin MSMask.MaskEdBox SldIni 
      Height          =   315
      Left            =   4155
      TabIndex        =   8
      Top             =   2535
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   15
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSComctlLib.TreeView TvwContas 
      Height          =   4125
      Left            =   5910
      TabIndex        =   13
      Top             =   1260
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   7276
      _Version        =   393217
      Indentation     =   453
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox UsaContaSimples 
      Caption         =   "Usa Conta Simplificada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   255
      TabIndex        =   9
      Top             =   3030
      Width           =   2295
   End
   Begin MSMask.MaskEdBox ContaSimples 
      Height          =   315
      Left            =   4965
      TabIndex        =   10
      Top             =   3060
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   6
      Mask            =   "######"
      PromptChar      =   " "
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Nat. Sped:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   45
      TabIndex        =   37
      Top             =   2085
      Width           =   930
   End
   Begin VB.Label Label6 
      Caption         =   "Participa do Fluxo de Caixa"
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
      Left            =   540
      TabIndex        =   36
      Top             =   3495
      Width           =   2415
   End
   Begin VB.Label LabelContaSimples 
      AutoSize        =   -1  'True
      Caption         =   "Conta Simplificada:"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3255
      TabIndex        =   24
      Top             =   3120
      Width           =   1665
   End
   Begin VB.Label Label1 
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
      Left            =   390
      TabIndex        =   25
      Top             =   285
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   45
      TabIndex        =   26
      Top             =   690
      Width           =   945
   End
   Begin VB.Label Label3 
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
      Left            =   525
      TabIndex        =   27
      Top             =   1590
      Width           =   465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Natureza:"
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
      Left            =   3255
      TabIndex        =   28
      Top             =   1575
      Width           =   855
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Plano de Contas"
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
      Left            =   5910
      TabIndex        =   29
      Top             =   1020
      Width           =   1410
   End
   Begin VB.Label LabelDigitoVerificador 
      Caption         =   "Valor do Dígito Verificador:"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3015
      TabIndex        =   30
      Top             =   4080
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.Label DigitoVerificador 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   315
      Left            =   5415
      TabIndex        =   31
      Top             =   4065
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label LabelSldIni 
      Caption         =   "Saldo Inicial:"
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
      Left            =   2970
      TabIndex        =   32
      Top             =   2610
      Width           =   1170
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Histórico Padrão:"
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
      Left            =   315
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   33
      Top             =   2595
      Width           =   1485
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Categoria:"
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
      Left            =   105
      TabIndex        =   34
      Top             =   1125
      Width           =   885
   End
   Begin VB.Label Label9 
      Caption         =   "Módulos que utilizam a conta (somente para contas analíticas)"
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
      TabIndex        =   35
      Top             =   3885
      Width           =   5505
   End
End
Attribute VB_Name = "PlanoContaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_Caption As String
Event Unload()

Private WithEvents objEventoConta As AdmEvento
Attribute objEventoConta.VB_VarHelpID = -1
Private WithEvents objEventoHistPadrao As AdmEvento
Attribute objEventoHistPadrao.VB_VarHelpID = -1

Dim iAlterado As Integer

Private Sub Ativo_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub BotaoExcluir_Click()

Dim objPlanoConta As New ClassPlanoConta
Dim lErro As Long
Dim sMsg As String
Dim iTemFilho As Integer
Dim vbConfirma As VbMsgBoxResult
Dim sConta As String
'Alteracao Daniel - 25/09/2001 (iAviso por sAviso)
Dim sAviso As String
Dim iContaTemMovimento As Integer
Dim iContaPreenchida As Integer

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    sConta = String(STRING_CONTA, 0)
    
    lErro = CF("Conta_Formata", Conta.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then Error 5947
    
    If iContaPreenchida <> CONTA_PREENCHIDA Then Error 5485
    
    'le a conta na tabela PlanoConta
    lErro = CF("Conta_SelecionaUma", sConta, objPlanoConta, MODULO_CONTABILIDADE)
    If lErro <> SUCESSO And lErro <> 6030 Then Error 5589
    
    'conta não cadastrada
    If lErro = 6030 Then Error 5595
    
    lErro = CF("Conta_Critica_Possui_Movimento", sConta, iContaTemMovimento)
    If lErro <> SUCESSO Then Error 5643
    
    If iContaTemMovimento = 1 Then Error 5644
    
    'se for uma conta analitica
    If objPlanoConta.iTipoConta = CONTA_ANALITICA Then
    
        'verifica se tem documentos automaticos associados a conta
        lErro = CF("DocAuto_Le_Conta", sConta)
        If lErro <> SUCESSO And lErro <> 5578 Then Error 5580
    
        If lErro = SUCESSO Then sMsg = sMsg & " Documentos Automáticos " & Chr(10)
        
        'verifica se tem associações com centros de custo
        lErro = CF("ContaCcl_Le_Conta", sConta)
        If lErro <> SUCESSO And lErro <> 5469 Then Error 5590
    
        If lErro = SUCESSO Then sMsg = sMsg & " Associações com Centros de Custo/Lucro " & Chr(10)
        
        If Len(sMsg) > 0 Then
            sAviso = "AVISO_EXCLUSAO_CONTA_ANALITICA_COM_ASSOCIACOES"
        Else
            sAviso = "AVISO_EXCLUSAO_CONTA_ANALITICA"
        End If
        
    Else
    
        lErro = CF("PlanoConta_Tem_Filho", sConta, iTemFilho)
        If lErro <> SUCESSO Then Error 5594
        
        If iTemFilho = 1 Then
            sAviso = "AVISO_EXCLUSAO_CONTA_SINTETICA_COM_FILHOS"
        Else
            sAviso = "AVISO_EXCLUSAO_CONTA_SINTETICA"
        End If
        
    End If
    
    vbConfirma = Rotina_Aviso(vbYesNo, sAviso, sMsg)
    
    If vbConfirma = vbYes Then
    
        'exclui a conta e suas descendentes, se houverem
        lErro = CF("Conta_Exclui", sConta)
        If lErro <> SUCESSO Then Error 5528
        
        'retira a conta da árvore
        Call Exclui_Arvore_Conta(TvwContas.Nodes, objPlanoConta)
        
        Call Limpa_PlanoConta
        
        iAlterado = 0
        
    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub
    
Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 5485
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_INFORMADA", Err)
            Conta.SetFocus
            
        Case 5528, 5580, 5589, 5590, 5594, 5643, 5947
            Conta.SetFocus
            
        Case 5644
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CONTA_COM_MOVIMENTO", Err)
            Conta.SetFocus
            
        Case 5595
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", Err, sConta)
            Conta.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164982)
        
    End Select

    Exit Sub
    
End Sub

Sub Exclui_Arvore_Conta(colNodes As Nodes, objPlanoConta As ClassPlanoConta)

Dim objNode As Node
Dim sConta As String
    
     'alterado por cyntia
    sConta = objPlanoConta.sConta
    
    For Each objNode In colNodes
        If Mid(objNode.Key, 2) = sConta Then
            colNodes.Remove (objNode.Index)
            Exit For
        End If
    Next
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 14732
    
    Call Limpa_PlanoConta
    
    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case Err
    
        Case 14732
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164983)
            
    End Select
    
    Exit Sub
    
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim sConta As String
Dim objHistPadrao As New ClassHistPadrao
Dim iContaPreenchida As Integer
Dim iCategoria As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    sConta = String(STRING_CONTA, 0)

    'critica o formato da conta
    lErro = CF("Conta_Formata", Conta.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then Error 5936

    'testa se a conta está preenchida
    If iContaPreenchida <> CONTA_PREENCHIDA Then Error 5485

    'verifica se o historico padrão está cadastrado
    lErro = HistPadrao_Critica(bSGECancelDummy)
    If lErro <> SUCESSO Then Error 5486
    
    'verifica se a conta possui uma conta pai
    lErro = CF("Conta_Critica_ContaPai", sConta, MODULO_CONTABILIDADE)
    If lErro <> SUCESSO Then Error 5487
    
    If Categoria.ListIndex >= 0 Then
        iCategoria = Categoria.ItemData(Categoria.ListIndex)
    Else
        iCategoria = 0
    End If
    
    'verifica se a conta é de nivel 1 e tem a categoria preenchida.
    lErro = CF("Conta_Critica_Categoria", sConta, iCategoria)
    If lErro <> SUCESSO Then Error 9717
    
    'verifica se a conta simples não está sendo utilizada por outra conta
    lErro = ContaSimples_Critica1()
    If lErro <> SUCESSO Then Error 5488
    
    'verifica se a conta já está cadastrada
    lErro = CF("PlanoConta_Le_Conta", sConta)
    If lErro <> SUCESSO And lErro <> 10051 Then Error 5490
    
    'se estiver cadastrada ==>  alteração
    If lErro = SUCESSO Then
    
        lErro = Atualizar_Conta(sConta)
        If lErro <> SUCESSO Then Error 5491
    
    Else
        
        'se estiver cadastrando uma conta nova
        lErro = Inserir_Conta(sConta)
        If lErro <> SUCESSO Then Error 5492
    
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = Err
    
    GL_objMDIForm.MousePointer = vbDefault
       
     Select Case Err
    
        Case 32296
    
        Case 5485
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_INFORMADA", Err)
            Conta.SetFocus
            
        Case 5936
            Conta.SetFocus
            
        Case 5486, 5487, 5488, 5489, 5490, 5491, 5492, 9717
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164984)
        
    End Select

    Exit Function

End Function
        
Function Alterar_Arvore_Conta(colNodes As Nodes, objPlanoConta As ClassPlanoConta) As Long

Dim objNode As Node
Dim sConta As String
Dim sContaMascarada As String
Dim lErro As Long
Dim iAchou As Integer

On Error GoTo Erro_Alterar_Arvore_Conta
    
     'alterado por cyntia
    sConta = objPlanoConta.sConta
    
    iAchou = 0
    
    For Each objNode In colNodes
    
        If Mid(objNode.Key, 2) = sConta Then
        
        
        
            '#############################################################
            'Alterado por Wagner
            If objPlanoConta.iTipoConta = CONTA_ANALITICA Then
                sConta = "A" & objPlanoConta.sConta
            Else
                sConta = "S" & objPlanoConta.sConta
            End If
            
            If sConta <> objNode.Key Then objNode.Key = sConta
            '############################################################
        
            sContaMascarada = String(STRING_CONTA, 0)
            
            'coloca a conta no formato que é exibida na tela
            lErro = Mascara_MascararConta(objPlanoConta.sConta, sContaMascarada)
            If lErro <> SUCESSO Then Error 5949
        
            objNode.Text = sContaMascarada & SEPARADOR & objPlanoConta.sDescConta
            
            iAchou = 1
            
            Exit For
            
        End If
    Next
    
    'se não achou a conta na arvore
    If iAchou = 0 Then Error 9557
    
    Alterar_Arvore_Conta = SUCESSO

    Exit Function

Erro_Alterar_Arvore_Conta:

    Alterar_Arvore_Conta = Err

    Select Case Err

        Case 5949
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararConta", Err, objPlanoConta.sConta)
            
        Case 9557
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164985)

    End Select
    
    Exit Function

End Function
        
Private Function Atualizar_Conta(sConta As String) As Long

Dim objPlanoConta As New ClassPlanoConta
Dim lErro As Long

On Error GoTo Erro_Atualizar_Conta

    lErro = Mover_Tela_Memoria(sConta, objPlanoConta)
    If lErro <> SUCESSO Then Error 5501
    
    lErro = Trata_Alteracao(objPlanoConta, sConta)
    If lErro <> SUCESSO Then Error 32302

    lErro = CF("PlanoConta_Altera", objPlanoConta)
    If lErro <> SUCESSO Then Error 5509
    
    'alterar a definição da conta na lista de contas
    lErro = Alterar_Arvore_Conta(TvwContas.Nodes, objPlanoConta)
    If lErro <> SUCESSO And lErro <> 9557 Then Error 5950
    
    'Se a conta não estava cadastrada na arvore
    If lErro = 9557 Then
    
        'inserir a conta na lista de contas
        lErro = Inserir_Arvore_Conta(TvwContas.Nodes, objPlanoConta)
        If lErro <> SUCESSO And lErro <> 20587 And lErro <> 20588 Then Error 9558
        
    End If
    
    Atualizar_Conta = SUCESSO
    
    Exit Function
    
Erro_Atualizar_Conta:

    Atualizar_Conta = Err
    
    Select Case Err
    
        Case 5501, 5509, 5950, 9558, 32302
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164986)
        
    End Select

    Exit Function
        
End Function

Private Function Inserir_Conta(sConta As String) As Long

Dim objPlanoConta As New ClassPlanoConta
Dim lErro As Long

On Error GoTo Erro_Inserir_Conta

    'move os dados da tela para objPlanoConta
    lErro = Mover_Tela_Memoria(sConta, objPlanoConta)
    If lErro <> SUCESSO Then Error 5502
    
    lErro = Trata_Alteracao(objPlanoConta, sConta)
    If lErro <> SUCESSO Then Error 32302
    
    'insere a conta no banco de dados
    lErro = CF("PlanoConta_Insere", objPlanoConta)
    If lErro <> SUCESSO Then Error 5503
    
    'exclui da lista de contas se estiver cadastrado
    Call Exclui_Arvore_Conta(TvwContas.Nodes, objPlanoConta)
    
    'inserir a conta na lista de contas
    lErro = Inserir_Arvore_Conta(TvwContas.Nodes, objPlanoConta)
    If lErro <> SUCESSO And lErro <> 20587 And lErro <> 20588 Then Error 5951
    
    Inserir_Conta = SUCESSO
    
    Exit Function
    
Erro_Inserir_Conta:

    Inserir_Conta = Err
    
    Select Case Err
    
        Case 5502, 5503, 5951, 32302
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164987)
        
    End Select

    Exit Function
        
End Function

Private Function Inserir_Arvore_Conta(colNodes As Nodes, objPlanoConta As ClassPlanoConta) As Long
'insere a conta na lista de contas

Dim objNode As Node
Dim lErro As Long
Dim sContaMascarada As String
Dim sConta As String
Dim sContaPai As String
Dim sContaAvo As String
Dim iAchou As Integer
    
On Error GoTo Erro_Inserir_Arvore_Conta
    
    'verifica se o ramo da arvore a que se refere o nó em questão já está carregado na arvore
    'se estiver, pode inserir. Se não estiver, não pode inserir.
    'os niveis 1 e 2 estão sempre na arvore.
    'se o nivel da conta for maior do que 2 ==> testar se o avo indica a carga dos netos. Se não indicar, não inserir o no.
    If objPlanoConta.iNivelConta > 2 Then
    
        sContaAvo = String(STRING_CONTA, 0)
    
        lErro = Mascara_RetornaContaNoNivel(objPlanoConta.iNivelConta - 2, objPlanoConta.sConta, sContaAvo)
        If lErro <> SUCESSO Then Error 20586
        
        'alterado por cyntia
        sConta = sContaAvo
    
        iAchou = 0
    
        For Each objNode In colNodes
            If Mid(objNode.Key, 2) = sConta Then
                iAchou = 1
                'os netos do avo do elemento em questão não estão na arvore ==> não pode inserir o elemento na arvore
                If objNode.Tag <> NETOS_NA_ARVORE Then Error 20587
                Exit For
            End If
        Next
        
        'o avo do nó em questão ainda não está carregado ==> não pode inserir o elemento na árvore
        If iAchou = 0 Then Error 20588
        
    End If
            
        
    sContaMascarada = String(STRING_CONTA, 0)
        
    'coloca a conta no formato que é exibida na tela
    lErro = Mascara_MascararConta(objPlanoConta.sConta, sContaMascarada)
    If lErro <> SUCESSO Then Error 5952
    
    'alterado por cyntia
    'se for uma conta analitica
    If objPlanoConta.iTipoConta = CONTA_ANALITICA Then
        sConta = "A" & objPlanoConta.sConta
    Else
        sConta = "S" & objPlanoConta.sConta
    End If
    
    sContaPai = String(STRING_CONTA, 0)
        
    'retorna a conta "pai" da conta em questão, se houver
    lErro = Mascara_RetornaContaPai(objPlanoConta.sConta, sContaPai)
    If lErro <> SUCESSO Then Error 5953
        
    'se a conta possui uma conta "pai"
    If Len(Trim(sContaPai)) > 0 Then

        sContaPai = "S" & sContaPai
            
        Set objNode = colNodes.Add(colNodes.Item(sContaPai), tvwChild, sConta, sContaMascarada & SEPARADOR & objPlanoConta.sDescConta)
        colNodes.Item(sContaPai).Sorted = True

    Else
        'se a conta não possui conta "pai"
        Set objNode = colNodes.Add(, , sConta, sContaMascarada & SEPARADOR & objPlanoConta.sDescConta)
        TvwContas.Sorted = True
    End If
    
    Inserir_Arvore_Conta = SUCESSO

    Exit Function

Erro_Inserir_Arvore_Conta:

    Inserir_Arvore_Conta = Err

    Select Case Err

        Case 5952
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararConta", Err, objPlanoConta.sConta)
            
        Case 5953
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaContaPai", Err, objPlanoConta.sConta)
        
        Case 20586, 20587, 20588
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164988)

    End Select
    
    Exit Function

End Function

Private Function Mover_Tela_Memoria(sConta As String, objPlanoConta As ClassPlanoConta) As Long

Dim lErro As Long
Dim iNivel As Integer

On Error GoTo Erro_Mover_Tela_Memoria

    objPlanoConta.iFilialEmpresa = giFilialEmpresa

    objPlanoConta.sConta = sConta
    objPlanoConta.sDescConta = DescConta.Text
    
    If Len(Trim(sConta)) > 0 Then
        'retorna o nivel da conta
        lErro = Mascara_Conta_ObterNivel(sConta, iNivel)
        If lErro <> SUCESSO Then Error 5493
    
    Else
        iNivel = 0
    End If
    
    objPlanoConta.iNivelConta = iNivel
    
    If ativo.Value = vbChecked Then
        objPlanoConta.iAtivo = CONTA_ATIVA
    Else
        objPlanoConta.iAtivo = CONTA_INATIVA
    End If
    
    objPlanoConta.iTipoConta = gobjColTipoConta.TipoConta(TipoConta.Text)
    objPlanoConta.iNatureza = gobjColNaturezaConta.NaturezaConta(Natureza.Text)
    
    If Len(Trim(HistPadrao.Text)) > 0 Then
        objPlanoConta.iHistPadrao = CInt(HistPadrao.Text)
    Else
        objPlanoConta.iHistPadrao = 0
    End If
    objPlanoConta.iNaturezaSped = Codigo_Extrai(NaturezaSped.Text)
    
    If Categoria.ListIndex >= 0 Then
        objPlanoConta.iCategoria = Categoria.ItemData(Categoria.ListIndex)
    Else
        objPlanoConta.iCategoria = 0
    End If
    
    If Len(Trim(SldIni.Text)) > 0 Then
        If objPlanoConta.iNatureza = CONTA_CREDITO Then
            objPlanoConta.dSldIni = CDbl(SldIni.Text)
        Else
            objPlanoConta.dSldIni = -CDbl(SldIni.Text)
        End If
    Else
        objPlanoConta.dSldIni = 0
    End If
    
    If UsaContaSimples.Value = vbChecked Then
        objPlanoConta.iUsaContaSimples = CONTA_USA_CONTASIMPLES
        If ContaSimples.Text = "" Then
            objPlanoConta.lContaSimples = 0
        Else
            objPlanoConta.lContaSimples = CLng(ContaSimples.Text)
        End If
    Else
        objPlanoConta.iUsaContaSimples = CONTA_NAO_USA_CONTASIMPLES
        objPlanoConta.lContaSimples = 0
    End If
        
    If FluxoCaixa.Value = vbChecked Then
        objPlanoConta.iFluxoCaixa = FLUXOCAIXA_UTILIZADO
    Else
       objPlanoConta.iFluxoCaixa = FLUXOCAIXA_NAO_UTILIZADO
    End If
        
    Call Mover_Modulo_Memoria(objPlanoConta)
        
    Mover_Tela_Memoria = SUCESSO
    
    Exit Function
    
Erro_Mover_Tela_Memoria:

    Mover_Tela_Memoria = Err
    
    Select Case Err
    
        Case 5493
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_CONTA_OBTERNIVEL", Err, sConta)
            Conta.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164989)
        
    End Select

    Exit Function
        
End Function

Private Sub Mover_Modulo_Memoria(objPlanoConta As ClassPlanoConta)
'Move os modulos selecionados para dentro de objPlanoConta

Dim lErro As Long
Dim iNivel As Integer
Dim iIndice As Integer

    Set objPlanoConta.colModulo = New Collection

    For iIndice = 0 To ListaModulo.ListCount - 1
    
        If ListaModulo.Selected(iIndice) = True Then
        
            objPlanoConta.colModulo.Add gcolModulo.Sigla(ListaModulo.List(iIndice))
    
        End If
        
    Next
    
End Sub

Private Sub BotaoLimpar_Click()

Dim iLote As Integer
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 9623

    Call Limpa_PlanoConta
    
    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoLimpar_Click:

    Select Case Err
    
        Case 9623
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164990)
        
    End Select
        
    Exit Sub
    
End Sub

Private Sub Categoria_Click()
        iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Conta_Change()
        iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Conta_Validate(Cancel As Boolean)
    
Dim lErro As Long
Dim sContaFormatada As String
Dim iContaPreenchida As Integer

On Error GoTo Erro_Conta_Validate

    If Len(Conta.ClipText) > 0 Then

        'critica o formato da conta
        lErro = CF("Conta_Formata", Conta.Text, sContaFormatada, iContaPreenchida)
        If lErro <> SUCESSO Then Error 5934
    
        If iContaPreenchida = CONTA_PREENCHIDA Then
        
            lErro = CF("Conta_Critica_ContaPai", sContaFormatada, MODULO_CONTABILIDADE)
            If lErro <> SUCESSO Then Error 5935
        
        Else
            sContaFormatada = ""
        End If
        
        lErro = Habilita_Categoria(sContaFormatada)
        If lErro <> SUCESSO Then Error 9742
        
        
    End If
    
    Exit Sub
    
Erro_Conta_Validate:

    Cancel = True

    Select Case Err
    
        Case 5934, 5935, 9742
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164991)
        
    End Select

    Exit Sub
    
End Sub

Private Sub ContaSimples_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ContaSimples_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ContaSimples, iAlterado)

End Sub

Private Sub ContaSimples_Validate(Cancel As Boolean)

    Call ContaSimples_Critica1

End Sub

Function ContaSimples_Critica1() As Long

Dim objPlanoConta As New ClassPlanoConta
Dim lContaSimples As Long
Dim lErro As Long
Dim sConta As String
Dim iContaPreenchida As Integer

On Error GoTo Erro_ContaSimples_Critica1

    If Len(ContaSimples.Text) > 0 Then
    
        sConta = String(STRING_CONTA, 0)
        
        lErro = CF("Conta_Formata", Conta.Text, sConta, iContaPreenchida)
        If lErro <> SUCESSO Then Error 5990
    
        If iContaPreenchida = CONTA_PREENCHIDA Then
    
            lContaSimples = CLng(ContaSimples.Text)
    
            'le os dados da conta
            lErro = CF("PlanoConta_Le_ContaSimples", lContaSimples, objPlanoConta)
            If lErro = SUCESSO Then
        
                If objPlanoConta.sConta <> sConta Then Error 5453
                
            Else
                
                'houve erro na leitura do plano de conta.
                If lErro <> 5451 Then Error 5454
                
            End If
            
        End If
        
    End If
    
    ContaSimples_Critica1 = SUCESSO
    
    Exit Function
        
Erro_ContaSimples_Critica1:

    ContaSimples_Critica1 = Err
    
    Select Case Err
    
        Case 5453
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTASIMPLES_JA_UTILIZADA", Err, ContaSimples.Text, objPlanoConta.sConta)
            ContaSimples.SetFocus
            
        Case 5454
            ContaSimples.SetFocus
            
        Case 5990
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164992)
        
    End Select

    Exit Function

End Function

Private Sub DescConta_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_Load()

Dim objOrigem As ClassOrigemContab
Dim iIndice As Integer
Dim iLote As Integer
Dim lErro As Long
Dim iIndice1 As Integer
Dim iHabilitaSaldoInicial As Integer

On Error GoTo Erro_Form_Load

    Set objEventoConta = New AdmEvento
    Set objEventoHistPadrao = New AdmEvento
    
    'inicializa a mascara de conta
    lErro = Inicializa_Mascara_Conta()
    If lErro <> SUCESSO Then Error 5933
    
    'inicializar os tipos de conta
    For iIndice = 1 To gobjColTipoConta.Count
        TipoConta.AddItem gobjColTipoConta.Item(iIndice).sDescricao
    Next
        
    'selecionar o tipo de conta atual
    For iIndice = 0 To TipoConta.ListCount - 1
        If gobjColTipoConta.TipoConta(TipoConta.List(iIndice)) = giTipoConta Then
            TipoConta.ListIndex = iIndice
            Exit For
        End If
    Next
    
    'inicializar as naturezas de conta
    For iIndice = 1 To gobjColNaturezaConta.Count
        Natureza.AddItem gobjColNaturezaConta.Item(iIndice).sDescricao
    Next
        
    'selecionar a natureza atual
    For iIndice = 0 To Natureza.ListCount - 1
        Natureza.ListIndex = iIndice
        If gobjColNaturezaConta.NaturezaConta(Natureza.Text) = giNaturezaConta Then Exit For
    Next
    
    'Inicializa a Combobox de Categoria
    lErro = Carga_Categoria()
    If lErro <> SUCESSO Then Error 9698
    
    'Inicializa a Lista de Plano de Contas
    lErro = CF("Carga_Arvore_Conta", TvwContas.Nodes)
    If lErro <> SUCESSO Then Error 5939
    
    'Inicializa a Lista de Módulos
    Call Carga_Lista_Modulo
    
    'coloca os campos relativos a conta simplificada desabilitados
    ContaSimples.Enabled = False
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
        
Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
        
        Case 5933, 5939, 9698
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164993)
        
    End Select
    
    iAlterado = 0
    
    Exit Sub
        
End Sub

Function Trata_Parametros(Optional objPlanoConta As ClassPlanoConta) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se há uma conta selecionada, exibir seus dados
    If Not (objPlanoConta Is Nothing) Then
        
        lErro = Traz_Conta_Tela(objPlanoConta.sConta)
        If lErro <> SUCESSO Then Error 5944
        
        lErro = Habilita_Categoria(objPlanoConta.sConta)
        If lErro <> SUCESSO Then Error 55670
        
    End If

    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
    
        Case 5944, 55670
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164994)
    
    End Select
    
    iAlterado = 0
    
    Exit Function

End Function

Private Function Habilita_Saldo_Inicial() As Long
'habilita ou desabilita a alteracao do saldo inicial da conta

Dim lErro As Long
Dim iDisponivel As Integer

On Error GoTo Erro_Habilita_Saldo_Inicial

    If giFilialEmpresa = EMPRESA_TODA Then
        iDisponivel = SLDINI_NAO_DISPONIVEL
    Else
        'verifica se o saldo inicial pode ser alterado
        lErro = CF("Saldo_Inicial_Critica", iDisponivel)
        If lErro <> SUCESSO Then Error 5972
    End If
    
    'se puder habilita o campo de saldo inicial
    If iDisponivel = SLDINI_DISPONIVEL Then
        LabelSldIni.Enabled = True
        SldIni.Enabled = True
    Else
        LabelSldIni.Enabled = False
        SldIni.Enabled = False
    End If
    
    Habilita_Saldo_Inicial = SUCESSO

    Exit Function

Erro_Habilita_Saldo_Inicial:

    Habilita_Saldo_Inicial = Err

    Select Case Err

        Case 5972
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164995)

    End Select
    
    Exit Function
        
End Function

Private Function Inicializa_Mascara_Conta() As Long
'inicializa a mascara de conta

Dim sMascaraConta As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_Mascara_Conta

    'Inicializa a máscara de Conta
    sMascaraConta = String(STRING_CONTA, 0)
    
    'le a mascara das contas
    lErro = MascaraConta(sMascaraConta)
    If lErro <> SUCESSO Then Error 5932
    
    Conta.Mask = sMascaraConta
    
    Inicializa_Mascara_Conta = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Mascara_Conta:

    Inicializa_Mascara_Conta = Err
    
    Select Case Err
    
        Case 5932
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164996)
        
    End Select

    Exit Function

End Function

Private Function Traz_Conta_Tela(sConta As String) As Long
'carrega os dados da conta em questão na tela.

Dim objPlanoConta As New ClassPlanoConta
Dim sDescricao As String
Dim lErro As Long
Dim iIndice As Integer
Dim sContaEnxuta As String

On Error GoTo Erro_Traz_Conta_Tela

    Call Limpa_PlanoConta

    sContaEnxuta = String(STRING_CONTA, 0)
    
    lErro = Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
    If lErro <> SUCESSO Then Error 5945
            
    Conta.PromptInclude = False
    Conta.Text = sContaEnxuta
    Conta.PromptInclude = True

    'le a conta
    lErro = CF("PlanoConta_Le_Conta1", sConta, objPlanoConta)
    If lErro <> SUCESSO And lErro <> 6030 Then Error 5442
        
    'se encontrou a conta
    If lErro = SUCESSO Then
    
        DescConta.Text = objPlanoConta.sDescConta
        
        sDescricao = gobjColTipoConta.Descricao(objPlanoConta.iTipoConta)
        
        Categoria.ListIndex = -1
        
        'mostra a categoria da conta
        For iIndice = 0 To Categoria.ListCount - 1
            If Categoria.ItemData(iIndice) = objPlanoConta.iCategoria Then
                Categoria.ListIndex = iIndice
                Exit For
            End If
        Next
        Call Combo_Seleciona_ItemData(NaturezaSped, objPlanoConta.iNaturezaSped)
        
        'mostra o tipo da conta
        For iIndice = 0 To TipoConta.ListCount - 1
            TipoConta.ListIndex = iIndice
            If TipoConta.Text = sDescricao Then Exit For
        Next
        
        sDescricao = gobjColNaturezaConta.Descricao(objPlanoConta.iNatureza)
        
        'mostra a natureza da conta
        For iIndice = 0 To Natureza.ListCount - 1
            Natureza.ListIndex = iIndice
            If Natureza.Text = sDescricao Then Exit For
        Next
        
        lErro = Traz_Conta_Tela1(objPlanoConta)
        If lErro <> SUCESSO Then Error 10049
        
        iAlterado = 0
        
    End If
    
    Traz_Conta_Tela = SUCESSO

    Exit Function

Erro_Traz_Conta_Tela:

    Traz_Conta_Tela = Err

    Select Case Err
    
        Case 5442, 10049
        
        Case 5945
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, sConta)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164997)
        
    End Select
    
    iAlterado = 0
    
    Exit Function
        
End Function

Private Function Traz_Conta_Tela1(objPlanoConta As ClassPlanoConta) As Long
'complementação da função Traz_Conta_Tela. Traz os dados para a tela

'Alteracao Daniel - 20/09/2001
'Alterado o preenchimento do SaldoInicial para poder receber numero negativo.

Dim lErro As Long
Dim iIndice As Integer
Dim objSaldoInicialConta As New ClassSaldoInicialConta

On Error GoTo Erro_Traz_Conta_Tela1
    
    'mostra se a conta está ativa
    If objPlanoConta.iAtivo = CONTA_ATIVA Then
        ativo.Value = vbChecked
    Else
        ativo.Value = vbUnchecked
    End If
    
    If objPlanoConta.iHistPadrao > 0 Then HistPadrao.Text = CStr(objPlanoConta.iHistPadrao)
    
    If objPlanoConta.iUsaContaSimples = CONTA_USA_CONTASIMPLES Then
        UsaContaSimples.Value = vbChecked
        ContaSimples.Enabled = True
        If objPlanoConta.lContaSimples = 0 Then
            ContaSimples.Text = ""
        Else
            ContaSimples.Text = CStr(objPlanoConta.lContaSimples)
        End If
        
    Else
        UsaContaSimples.Value = vbUnchecked
        ContaSimples.Text = ""
    End If
    
    If objPlanoConta.iFluxoCaixa = FLUXOCAIXA_UTILIZADO Then
        FluxoCaixa.Value = vbChecked
    Else
        FluxoCaixa.Value = vbUnchecked
    End If
    
    If giFilialEmpresa <> EMPRESA_TODA And gcolModulo.ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then

        objSaldoInicialConta.iFilialEmpresa = giFilialEmpresa
        objSaldoInicialConta.sConta = objPlanoConta.sConta
        
        lErro = CF("SaldoInicialConta_Le", objSaldoInicialConta)
        If lErro <> SUCESSO Then Error 10342
    
        'alteracao daniel...
        If objPlanoConta.iNatureza = CONTA_DEBITO Then
            objSaldoInicialConta.dSldIni = -CDbl(objSaldoInicialConta.dSldIni)
        End If
    
        SldIni.Text = CStr(objSaldoInicialConta.dSldIni)

    End If
    
    lErro = Habilita_Categoria(objPlanoConta.sConta)
    If lErro <> SUCESSO Then Error 9743
    
    Call Traz_Modulo_Tela(objPlanoConta.colModulo)
    
    Traz_Conta_Tela1 = SUCESSO

    Exit Function

Erro_Traz_Conta_Tela1:

    Traz_Conta_Tela1 = Err

    Select Case Err
    
        Case 9743, 10342
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164998)
        
    End Select

    Exit Function
        
End Function

Private Sub Traz_Modulo_Tela(colModulo As Collection)

Dim vSigla As Variant
Dim iIndice As Integer

    Call Limpa_ListaModulo

    For Each vSigla In colModulo
        
        For iIndice = 0 To ListaModulo.ListCount - 1
    
            If ListaModulo.List(iIndice) = gcolModulo.Nome(CStr(vSigla)) Then
                ListaModulo.Selected(iIndice) = True
                Exit For
            End If
        Next
        
    Next

    ListaModulo.ListIndex = -1

End Sub

Private Sub Limpa_ListaModulo()
'limpa as seleções da lista de modulos

Dim iIndice As Integer

    For iIndice = 0 To ListaModulo.ListCount - 1
        ListaModulo.Selected(iIndice) = False
    Next
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
           
    Set objEventoConta = Nothing
    Set objEventoHistPadrao = Nothing
    
End Sub

Private Sub HistPadrao_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub HistPadrao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(HistPadrao, iAlterado)

End Sub

Private Sub HistPadrao_Validate(Cancel As Boolean)
    
    Call HistPadrao_Critica(Cancel)

End Sub

Private Function HistPadrao_Critica(Cancel As Boolean) As Long

Dim objHistPadrao As New ClassHistPadrao
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_HistPadrao_Critica

    If Len(HistPadrao.Text) > 0 Then
    
        objHistPadrao.iHistPadrao = CInt(HistPadrao.Text)
        
        If objHistPadrao.iHistPadrao > 0 Then
    
            'le os dados do historico
            lErro = CF("HistPadrao_Le", objHistPadrao)
            If lErro <> SUCESSO And lErro <> 5446 Then Error 5449
                
            'se o historico não estiver cadastrado
            If lErro = 5446 Then Error 5448
            
        End If
        
    End If
    
    HistPadrao_Critica = SUCESSO
    
    Exit Function
        
Erro_HistPadrao_Critica:

    Cancel = True
    
    HistPadrao_Critica = Err
    
    Select Case Err
    
        Case 5448
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_HISTPADRAO_INEXISTENTE", objHistPadrao.iHistPadrao)
            
            If vbMsgRes = vbYes Then
                
                'Usuário quer criar este historico
                Call Chama_Tela("HistoricoPadrao", objHistPadrao)
            Else
''                HistPadrao.SetFocus
            End If
            
            
        Case 5449
''            HistPadrao.SetFocus
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164999)
        
    End Select

    Exit Function

End Function

Private Sub ArvorePlanoConta_NodeClick(ByVal Node As MSComctlLib.Node)

        Conta.Text = right(Node.Key, Len(Node.Key) - 1)
        DescConta.Text = Node.Text

End Sub

Private Sub Natureza_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoConta_evSelecao(obj1 As Object)
    
Dim objPlanoConta As ClassPlanoConta
Dim lErro As Long

On Error GoTo Erro_objEventoConta_evSelecao

    Set objPlanoConta = obj1
    
    lErro = Traz_Conta_Tela(objPlanoConta.sConta)
    If lErro <> SUCESSO Then Error 9288
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoConta_evSelecao:

    Select Case Err
    
        Case 9288
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165000)
        
    End Select

    Exit Sub
        
End Sub

Private Sub SldIni_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub SldIni_Validate(Cancel As Boolean)

'Daniel em 20/09/2001.
'Substituida a função Valor_NaoNegativo_Critica por Valor_Double_Critica.
'Agora o Saldo Inicial pode ser negativo.

Dim lErro As Long

On Error GoTo Erro_SldIni_Validate
    
    If Len(SldIni.Text) > 0 Then
        'alteracao daniel
        lErro = Valor_Double_Critica(SldIni.Text)
        If lErro <> SUCESSO Then Error 5938
        
        SldIni.Text = Format(SldIni.Text, "Fixed")
        
    End If
    
    Exit Sub
    
Erro_SldIni_Validate:

    Cancel = True


    Select Case Err
    
        Case 5938
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165001)
        
    End Select
    
    Exit Sub
    
End Sub

Private Sub TipoConta_Click()

Dim lErro As Long

On Error GoTo Erro_TipoConta_Click

    iAlterado = REGISTRO_ALTERADO

    If gobjColTipoConta.TipoConta(TipoConta.Text) = CONTA_SINTETICA Then
        SldIni.Text = ""
        SldIni.Enabled = False
        LabelSldIni.Enabled = False
    Else
        lErro = Habilita_Saldo_Inicial()
        If lErro <> SUCESSO Then Error 5970
    End If

    Exit Sub
    
Erro_TipoConta_Click:

    Select Case Err
    
        Case 5970
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165002)
        
    End Select
    
    Exit Sub

End Sub

Private Sub TipoConta_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iNivel As Integer
Dim sContaFormatada As String
Dim iContaPreenchida As Integer

On Error GoTo Erro_TipoConta_Validate

    If Len(Conta.Text) = 0 Then Exit Sub
        
    lErro = CF("Conta_Formata", Conta.Text, sContaFormatada, iContaPreenchida)
    If lErro <> SUCESSO Then Error 10471
    
    'verifica se a conta já está cadastrada
    lErro = CF("PlanoConta_Le_Conta", sContaFormatada)
    If lErro <> SUCESSO And lErro <> 10051 Then Error 5455
        
    'nao encontrou a conta ==> não continua a crítica
    If lErro = 10051 Then Exit Sub
    
    lErro = CF("Conta_Critica_Tipo", sContaFormatada, gobjColTipoConta.TipoConta(TipoConta.Text))
    If lErro <> SUCESSO Then Error 5489
    
    Exit Sub
        
Erro_TipoConta_Validate:

    Cancel = True


    Select Case Err
    
        Case 5455, 5489
        
        Case 10471
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165003)
        
    End Select

    Exit Sub

End Sub

Private Sub TvwContas_Expand(ByVal objNode As MSComctlLib.Node)

Dim lErro As Long

On Error GoTo Erro_TvwContas_Expand

    If objNode.Tag <> NETOS_NA_ARVORE Then
    
        'move os dados do plano de contas do banco de dados para a arvore colNodes.
        lErro = CF("Carga_Arvore_Conta1", objNode, TvwContas.Nodes)
        If lErro <> SUCESSO Then Error 36771
        
    End If
    
    Exit Sub
    
Erro_TvwContas_Expand:

    Select Case Err
    
        Case 36771
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165004)
        
    End Select
        
    Exit Sub
    
End Sub

Private Sub UsaContaSimples_Click()

    iAlterado = REGISTRO_ALTERADO

    If UsaContaSimples.Value = vbChecked Then
        'habilita os campos relativos à conta simplificada
        ContaSimples.Enabled = True
        LabelContaSimples.Enabled = True
    Else
        'desabilita e limpa os campos relativos à conta simplificada
        ContaSimples.Enabled = False
        ContaSimples.Text = ""
        LabelContaSimples.Enabled = False
    End If

End Sub

Private Sub FluxoCaixa_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub





Private Sub TvwContas_NodeClick(ByVal Node As MSComctlLib.Node)

Dim sConta As String
Dim sCaracterInicial As String
Dim lPosicaoSeparador As Long
Dim lErro As Long
Dim sContaEnxuta As String

On Error GoTo Erro_TvwContas_NodeClick
    
    sConta = right(Node.Key, Len(Node.Key) - 1)
            
    lErro = Traz_Conta_Tela(sConta)
    If lErro <> SUCESSO Then Error 5946
    
    Exit Sub

Erro_TvwContas_NodeClick:

    Select Case Err
    
        Case 5946
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165005)
            
    End Select
        
    Exit Sub
    
End Sub

Private Sub objEventoHistPadrao_evSelecao(obj1 As Object)

Dim objHistPadrao As ClassHistPadrao
Dim lErro As Long
    
On Error GoTo Erro_objEventoHistPadrao_evSelecao
    
    Set objHistPadrao = obj1
    
    HistPadrao.Text = objHistPadrao.iHistPadrao
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoHistPadrao_evSelecao:

    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165006)
            
        End Select
        
    Exit Sub
        
End Sub

Private Sub Label5_Click()

Dim objHistPadrao As New ClassHistPadrao
Dim colSelecao As Collection

    If Len(HistPadrao.Text) = 0 Then
        objHistPadrao.iHistPadrao = 0
    Else
        objHistPadrao.iHistPadrao = CLng(HistPadrao.ClipText)
    End If

    Call Chama_Tela("HistPadraoLista", colSelecao, objHistPadrao, objEventoHistPadrao)

End Sub

Private Function Carga_Categoria() As Long
'inicializa a combobox de Categoria

Dim colContaCategoria As New Collection
Dim objContaCategoria As ClassContaCategoria
Dim lErro As Long

    'Le todas as categorias existentes no BD
    lErro = CF("ContaCategoria_Le_Todos", colContaCategoria)
    If lErro <> SUCESSO Then Error 9699
    
    For Each objContaCategoria In colContaCategoria
    
        Categoria.AddItem objContaCategoria.sNome
        Categoria.ItemData(Categoria.NewIndex) = objContaCategoria.iCodigo
        
    Next
    
    Carga_Categoria = SUCESSO

    Exit Function
    
Erro_Carga_Categoria:

    Carga_Categoria = Err

    Select Case Err
            
        Case 9699
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165007)

    End Select
    
    Exit Function

End Function

Function Limpa_PlanoConta() As Long

Dim lErro As Long

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Call Limpa_Tela(Me)
    
    Categoria.ListIndex = -1
    NaturezaSped.ListIndex = -1
    Categoria.Enabled = True
    UsaContaSimples.Value = 0
    FluxoCaixa.Value = 0
        
    Limpa_PlanoConta = SUCESSO

End Function

Function Habilita_Categoria(sConta As String) As Long
'se o nivel for 1 ou a conta não estiver preenchida ==> habilita a conta. Senão desabilita.

Dim lErro As Long
Dim iNivel As Integer

On Error GoTo Erro_Habilita_Categoria

    If Len(sConta) > 0 Then
        'retorna o nivel da conta
        lErro = Mascara_Conta_ObterNivel(sConta, iNivel)
        If lErro <> SUCESSO Then Error 9741
        
        If iNivel = 1 Then
            Categoria.Enabled = True
            'NaturezaSped.Enabled = True
        Else
            Categoria.Enabled = False
            Categoria.ListIndex = -1
            'NaturezaSped.Enabled = False
            'NaturezaSped.ListIndex = -1
        End If
        
    Else
        Categoria.Enabled = True
        'NaturezaSped.Enabled = True
    End If

    Habilita_Categoria = SUCESSO
    
    Exit Function
    
Erro_Habilita_Categoria:

    Habilita_Categoria = Err
    
    Select Case Err
    
        Case 9741
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165008)

    End Select
    
    Exit Function
        
End Function

Private Sub Carga_Lista_Modulo()
'carrega a lista de modulos

Dim iIndice As Integer

    For iIndice = 1 To gcolModulo.Count

        If gcolModulo.Item(iIndice).iAtivo = MODULO_ATIVO And gcolModulo.Item(iIndice).sSigla <> MODULO_CONTABILIDADE And gcolModulo.Item(iIndice).sSigla <> MODULO_ADM And gcolModulo.Item(iIndice).sSigla <> MODULO_PCP _
            And gcolModulo.Item(iIndice).sSigla <> MODULO_LOJA And gcolModulo.Item(iIndice).sSigla <> MODULO_CRM And gcolModulo.Item(iIndice).sSigla <> MODULO_PONTO_DE_VENDA And gcolModulo.Item(iIndice).sSigla <> MODULO_FOLHA And gcolModulo.Item(iIndice).sSigla <> MODULO_QUALIDADE And gcolModulo.Item(iIndice).sSigla <> MODULO_PROJETO Then

            ListaModulo.AddItem gcolModulo.Item(iIndice).sNome
            
        End If
        
    Next

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim sConta As String
Dim iContaPreenchida As Integer
Dim objPlanoConta As New ClassPlanoConta

On Error GoTo Erro_Tela_Extrai

    sConta = String(STRING_CONTA, 0)

    'Informa tabela associada à Tela
    sTabela = "PlanoConta"
    
    'critica o formato da conta
    lErro = CF("Conta_Formata", Conta.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then Error 24151

    'Le os dados da Tela de Lotes
    lErro = Mover_Tela_Memoria(sConta, objPlanoConta)
    If lErro <> SUCESSO Then Error 24149
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Conta", objPlanoConta.sConta, STRING_CONTA, "Conta"
    colCampoValor.Add "DescConta", objPlanoConta.sDescConta, STRING_CONTA_DESCRICAO, "DescConta"
    colCampoValor.Add "Ativo", objPlanoConta.iAtivo, 0, "Ativo"
    colCampoValor.Add "Categoria", objPlanoConta.iCategoria, 0, "Categoria"
    colCampoValor.Add "DigitoVerif", objPlanoConta.iDigitoVerif, 0, "DigitoVerif"
    colCampoValor.Add "DigitoVerifSimples", objPlanoConta.iDigitoVerifSimples, 0, "DigitoVerifSimples"
    colCampoValor.Add "HistPadrao", objPlanoConta.iHistPadrao, 0, "HistPadrao"
    colCampoValor.Add "Natureza", objPlanoConta.iNatureza, 0, "Natureza"
    colCampoValor.Add "NivelConta", objPlanoConta.iNivelConta, 0, "NivelConta"
    colCampoValor.Add "TipoConta", objPlanoConta.iTipoConta, 0, "TipoConta"
    colCampoValor.Add "UsaContaSimples", objPlanoConta.iUsaContaSimples, 0, "UsaContaSimples"
    colCampoValor.Add "ContaSimples", objPlanoConta.lContaSimples, 0, "ContaSimples"
    colCampoValor.Add "FluxoCaixa", objPlanoConta.iFluxoCaixa, 0, "FluxoCaixa"
    
    Exit Sub
    
Erro_Tela_Extrai:

    Select Case Err

        Case 24149, 24151
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165009)

    End Select
    
    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objPlanoConta As New ClassPlanoConta

On Error GoTo Erro_Tela_Preenche

    objPlanoConta.sConta = colCampoValor.Item("Conta").vValor

    If objPlanoConta.sConta <> "" Then
    
        lErro = Traz_Conta_Tela(objPlanoConta.sConta)
        If lErro <> SUCESSO Then Error 24150

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 24150

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165010)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

'Function Carga_Arvore_Conta(colNodes As Nodes) As Long
''copiado de PlanoConta
''move os dados do plano de contas do banco de dados para a arvore colNodes.
'
'Dim objNode As Node
'Dim colPlanoConta As New Collection
'Dim objPlanoConta As ClassPlanoConta
'Dim lErro As Long
'Dim sContaMascarada As String
'Dim sConta As String
'Dim sContaPai As String
'
'On Error GoTo Erro_Carga_Arvore_Conta
'
'    'le todas as contas de nível 0 e 1 e coloca-as em colPlanoConta
'    lErro = CF("PlanoConta_Le_Niveis0e1", colPlanoConta)
'    If lErro <> SUCESSO Then Error 13042
'
'    For Each objPlanoConta In colPlanoConta
'
'        sContaMascarada = String(STRING_CONTA, 0)
'
'        'coloca a conta no formato que é exibida na tela
'        lErro = Mascara_MascararConta(objPlanoConta.sConta, sContaMascarada)
'        If lErro <> SUCESSO Then Error 13043
'
'''        sConta = "S" & objPlanoConta.sConta
'
'        sContaPai = String(STRING_CONTA, 0)
'
'        'retorna a conta "pai" da conta em questão, se houver
'        lErro = Mascara_RetornaContaPai(objPlanoConta.sConta, sContaPai)
'        If lErro <> SUCESSO Then Error 13044
'
'        'se a conta possui uma conta "pai"
'        If Len(Trim(sContaPai)) > 0 Then
'
'            sContaPai = "S" & sContaPai
'
'            Set objNode = colNodes.Add(colNodes.Item(sContaPai), tvwChild, sConta)
'
'        Else
'            'se a conta não possui conta "pai"
'            Set objNode = colNodes.Add(, tvwLast, sConta)
'
'        End If
'
'        objNode.Text = sContaMascarada & SEPARADOR & objPlanoConta.sDescConta
'
'    Next
'
'    Carga_Arvore_Conta = SUCESSO
'
'    Exit Function
'
'Erro_Carga_Arvore_Conta:
'
'    Carga_Arvore_Conta = Err
'
'    Select Case Err
'
'        Case 13042
'
'        Case 13043
'            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararConta", Err, objPlanoConta.sConta)
'
'        Case 13044
'            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaContaPai", Err, objPlanoConta.sConta)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165011)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Carga_Arvore_Conta1(objNodeAvo As Node, colNodes As Nodes) As Long
''move os dados do plano de contas do banco de dados para a arvore colNodes.
'
'Dim objNode As Node
'Dim colPlanoConta As New Collection
'Dim objPlanoConta As ClassPlanoConta
'Dim lErro As Long
'Dim sContaMascarada As String
'Dim sConta As String
'Dim sContaAvo As String
'Dim sContaPai As String
'
'On Error GoTo Erro_Carga_Arvore_Conta1
'
'    sContaAvo = Mid(objNodeAvo.Key, 2)
'
'    'le os filhos da conta em questão e coloca-as em colPlanoConta
'    lErro = CF("PlanoConta_Le_Netos", sContaAvo, colPlanoConta)
'    If lErro <> SUCESSO Then Error 40811
'
'    For Each objPlanoConta In colPlanoConta
'
'        sContaMascarada = String(STRING_CONTA, 0)
'
'        'coloca a conta no formato que é exibida na tela
'        lErro = Mascara_MascararConta(objPlanoConta.sConta, sContaMascarada)
'        If lErro <> SUCESSO Then Error 40812
'
'        sConta = "S" & objPlanoConta.sConta
'
'        sContaPai = String(STRING_CONTA, 0)
'
'        'retorna a conta "pai" da conta em questão, se houver
'        lErro = Mascara_RetornaContaPai(objPlanoConta.sConta, sContaPai)
'        If lErro <> SUCESSO Then Error 40813
'
'        'se a conta possui uma conta "pai"
'        If Len(Trim(sContaPai)) > 0 Then
'
'            sContaPai = "S" & sContaPai
'
'            Set objNode = colNodes.Add(colNodes.Item(sContaPai), tvwChild, sConta)
'
'        Else
'            'se a conta não possui conta "pai"
'            Set objNode = colNodes.Add(, tvwLast, sConta)
'
'        End If
'
'        objNode.Text = sContaMascarada & SEPARADOR & objPlanoConta.sDescConta
'
'    Next
'
'    'coloca o tag indicando que os netos já foram carregados
'    objNodeAvo.Tag = NETOS_NA_ARVORE
'
'    Carga_Arvore_Conta1 = SUCESSO
'
'    Exit Function
'
'Erro_Carga_Arvore_Conta1:
'
'    Carga_Arvore_Conta1 = Err
'
'    Select Case Err
'
'        Case 40811
'
'        Case 40812
'            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararConta", Err, objPlanoConta.sConta)
'
'        Case 40813
'            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaContaPai", Err, objPlanoConta.sConta)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165012)
'
'    End Select
'
'    Exit Function
'
'End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_PLANO_CONTAS
    Set Form_Load_Ocx = Me
    Caption = "Plano de Contas"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "PlanoConta"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
        
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is HistPadrao Then
            Call Label5_Click
        End If
    
    End If

End Sub


Private Sub LabelDigitoVerificadorSimples_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDigitoVerificadorSimples, Source, X, Y)
End Sub

Private Sub LabelDigitoVerificadorSimples_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDigitoVerificadorSimples, Button, Shift, X, Y)
End Sub

Private Sub DigitoVerificadorSimples_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DigitoVerificadorSimples, Source, X, Y)
End Sub

Private Sub DigitoVerificadorSimples_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DigitoVerificadorSimples, Button, Shift, X, Y)
End Sub

Private Sub LabelContaSimples_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelContaSimples, Source, X, Y)
End Sub

Private Sub LabelContaSimples_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelContaSimples, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub LabelDigitoVerificador_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDigitoVerificador, Source, X, Y)
End Sub

Private Sub LabelDigitoVerificador_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDigitoVerificador, Button, Shift, X, Y)
End Sub

Private Sub DigitoVerificador_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DigitoVerificador, Source, X, Y)
End Sub

Private Sub DigitoVerificador_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DigitoVerificador, Button, Shift, X, Y)
End Sub

Private Sub LabelSldIni_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelSldIni, Source, X, Y)
End Sub

Private Sub LabelSldIni_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelSldIni, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub


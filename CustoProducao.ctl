VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl CustoProducao 
   Appearance      =   0  'Flat
   ClientHeight    =   4485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8445
   KeyPreview      =   -1  'True
   ScaleHeight     =   4485
   ScaleWidth      =   8445
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3150
      Index           =   1
      Left            =   210
      TabIndex        =   18
      Top             =   1065
      Width           =   8130
      Begin MSMask.MaskEdBox CustoFator5 
         Height          =   225
         Left            =   5220
         TabIndex        =   28
         Top             =   2100
         Width           =   1650
         _ExtentX        =   2910
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
         Format          =   "#,##0.00##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CustoFator6 
         Height          =   225
         Left            =   5775
         TabIndex        =   29
         Top             =   2415
         Width           =   1650
         _ExtentX        =   2910
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
         Format          =   "#,##0.00##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CustoFator4 
         Height          =   225
         Left            =   4740
         TabIndex        =   27
         Top             =   1950
         Width           =   1650
         _ExtentX        =   2910
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
         Format          =   "#,##0.00##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CustoFator3 
         Height          =   225
         Left            =   4785
         TabIndex        =   26
         Top             =   1620
         Width           =   1650
         _ExtentX        =   2910
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
         Format          =   "#,##0.00##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CustoFator1 
         Height          =   225
         Left            =   4770
         TabIndex        =   24
         Top             =   1035
         Width           =   1650
         _ExtentX        =   2910
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
         Format          =   "#,##0.00##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CustoFator2 
         Height          =   225
         Left            =   4905
         TabIndex        =   25
         Top             =   1305
         Width           =   1650
         _ExtentX        =   2910
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
         Format          =   "#,##0.00##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Status 
         Height          =   225
         Left            =   5235
         TabIndex        =   6
         Top             =   750
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   370
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox GastosIndiretos 
         Height          =   225
         Left            =   3585
         TabIndex        =   5
         Top             =   735
         Width           =   1650
         _ExtentX        =   2910
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
         Format          =   "#,##0.00##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox FilialEmpresa 
         Height          =   225
         Left            =   375
         TabIndex        =   3
         Top             =   780
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox GastosDiretos 
         Height          =   225
         Left            =   1890
         TabIndex        =   4
         Top             =   735
         Width           =   1650
         _ExtentX        =   2910
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
         Format          =   "#,##0.00##"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridEstoqueMes 
         Height          =   2730
         Left            =   135
         TabIndex        =   19
         Top             =   150
         Width           =   7800
         _ExtentX        =   13758
         _ExtentY        =   4815
         _Version        =   393216
         Rows            =   100
         Cols            =   5
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
         AllowUserResizing=   1
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3270
      Index           =   2
      Left            =   195
      TabIndex        =   2
      Top             =   1035
      Visible         =   0   'False
      Width           =   8130
      Begin VB.CommandButton BotaoProdutos 
         Caption         =   "Produtos"
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
         Left            =   6465
         TabIndex        =   12
         Top             =   2805
         Width           =   1560
      End
      Begin VB.ComboBox FilialEmpresaProd 
         Height          =   315
         ItemData        =   "CustoProducao.ctx":0000
         Left            =   480
         List            =   "CustoProducao.ctx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1080
         Width           =   1320
      End
      Begin VB.TextBox DescricaoItem 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   270
         Left            =   3285
         MaxLength       =   50
         TabIndex        =   9
         Top             =   1065
         Width           =   2010
      End
      Begin MSMask.MaskEdBox StatusProd 
         Height          =   225
         Left            =   6420
         TabIndex        =   11
         Top             =   1095
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox GastoProduto 
         Height          =   225
         Left            =   5355
         TabIndex        =   10
         Top             =   1095
         Width           =   990
         _ExtentX        =   1746
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
         Format          =   "#,##0.00##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Produto 
         Height          =   270
         Left            =   1860
         TabIndex        =   8
         Top             =   1050
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridGastoProduto 
         Height          =   2565
         Left            =   135
         TabIndex        =   17
         Top             =   165
         Width           =   7935
         _ExtentX        =   14023
         _ExtentY        =   4524
         _Version        =   393216
         Rows            =   100
         Cols            =   5
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
         AllowUserResizing=   1
      End
   End
   Begin VB.ComboBox Mes 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "CustoProducao.ctx":0004
      Left            =   2745
      List            =   "CustoProducao.ctx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   285
      Width           =   1545
   End
   Begin VB.ComboBox Ano 
      Height          =   315
      ItemData        =   "CustoProducao.ctx":0008
      Left            =   645
      List            =   "CustoProducao.ctx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   285
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   690
      Left            =   5295
      ScaleHeight     =   630
      ScaleWidth      =   2955
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   75
      Width           =   3015
      Begin VB.CommandButton BotaoApurar 
         Height          =   510
         Left            =   90
         Picture         =   "CustoProducao.ctx":000C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   75
         Width           =   1245
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   510
         Left            =   1440
         Picture         =   "CustoProducao.ctx":18CE
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   510
         Left            =   1950
         Picture         =   "CustoProducao.ctx":1A28
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   510
         Left            =   2460
         Picture         =   "CustoProducao.ctx":1F5A
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3645
      Left            =   135
      TabIndex        =   21
      Top             =   690
      Width           =   8235
      _ExtentX        =   14552
      _ExtentY        =   6403
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Gastos Diretos/Indiretos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Gastos de Produto"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mês:"
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
      Left            =   2175
      TabIndex        =   23
      Top             =   345
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ano:"
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
      TabIndex        =   22
      Top             =   330
      Width           =   405
   End
End
Attribute VB_Name = "CustoProducao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Problemas encontrados por Leo em 31/01/02 na tela:
'- Tela desorganizada
'- Não havia códio p/ a troca de Tabs
'- Não era possível acessar as linhas diferentes da 1ª no grid 1
'- O Grid 1 estava permitindo a adição de novas linhas
'- O Grid 2 não permite a inclusão de novas linhas a partir da 1ª
'- Os saídas de célula do grid 2 não funcionam (faltando tratamento dos controles)
'- Erro ao chamar apuração


Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1

'campos do grid de gastos diretos/indiretos
Dim iGrid_FilialEmpresa_Col As Integer
Dim iGrid_GastosDiretos_Col As Integer
Dim iGrid_GastosIndiretos_Col As Integer
Dim iGrid_Status_Col As Integer
Dim iGrid_CustoFator1_Col As Integer
Dim iGrid_CustoFator2_Col As Integer
Dim iGrid_CustoFator3_Col As Integer
Dim iGrid_CustoFator4_Col As Integer
Dim iGrid_CustoFator5_Col As Integer
Dim iGrid_CustoFator6_Col As Integer

'campos do grid de gastos de fecularia
Dim iGrid_FilialEmpresaProd_Col As Integer
Dim iGrid_Produto_Col As Integer
Dim iGrid_DescricaoItem_Col As Integer
Dim iGrid_GastoProduto_Col As Integer
Dim iGrid_StatusProd_Col As Integer


Dim iMesAtual As Integer
Dim iAnoAtual As Integer

Dim objGrid As AdmGrid
Dim objGridGastoProduto As AdmGrid

Dim iFrameAtual As Integer 'Incluido por Leo em 30/01/02

'erros por Leo:
'ERRO_FILIAL_GRIDGASTOPRODUTO_NAO_SELECIONADA : A Filial não foi selecionada.


Private Sub Ano_Click()

Dim colMeses As New Collection
Dim vMes As Variant
Dim sMes As String
Dim lErro As Long
Dim iAno As Integer
Dim iMes As Integer

On Error GoTo Erro_Ano_Click

    'se trocou o Ano e o Ano Anterior era diferente de -1 ==> tenta salvar os dados da tela
    If iAnoAtual <> StrParaInt(Ano.Text) And iAnoAtual <> -1 Then
    
        'Testa se deseja salvar mudanças
        lErro = Teste_Salva(Me, iAlterado)
        
        If lErro <> SUCESSO Then gError 92582
    

    End If

    If Ano.ListIndex = -1 Then
    
        Mes.Enabled = False
        iAnoAtual = -1
        Mes.Clear
        
    Else

        iAno = Ano.Text
        iAnoAtual = iAno
    
        Mes.Clear
    
        'Le todos os meses da tabela EstoqueMes com Ano = iAno
        lErro = CF("EstoqueMes_Le_Meses1", iAno, colMeses)
        If lErro <> SUCESSO And lErro <> 92568 Then gError 92576
        
        'se não houver nenhum mes cadastrado para o Ano em questão ==> erro
        If lErro = 92568 Then gError 92579
        
        For Each vMes In colMeses
        
            iMes = vMes
        
            'Pega o Nome do mes em questão
            lErro = MesNome(iMes, sMes)
            If lErro <> SUCESSO Then gError 92577
            
            Mes.AddItem sMes
            Mes.ItemData(Mes.NewIndex) = vMes
            
        Next
        
        Mes.Enabled = True
        
    End If
    
    iAlterado = 0
        
    Exit Sub
        
Erro_Ano_Click:

    Select Case gErr
    
        Case 92576, 92577, 92582
        
        Case 92579
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEMES_INEXISTENTE5", gErr, Ano.Text)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158660)

    End Select

    Exit Sub

End Sub

Private Sub BotaoApurar_Click()

Dim lErro As Long
Dim iAno As Integer
Dim iMes As Integer
Dim sNomeArqParam As String

On Error GoTo Erro_BotaoApurar_Click

    If iAlterado = 1 Then
    
        'Testa se deseja salvar mudanças
        lErro = Teste_Salva(Me, iAlterado)
        If lErro <> SUCESSO Then gError 92598
    
    Else
    
        lErro = CustoProducao_Valida_Gravacao()
        If lErro <> SUCESSO Then gError 94566
    
    End If
   
    'Prepara para chamar rotina batch
    lErro = Sistema_Preparar_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError 92599
    
    'Chama rotina batch que calcula custo médio de produção
    'e valoriza movimentos de materiais produzidos
    lErro = CF("Rotina_CustoMedioProducao_Calcula", sNomeArqParam, iAnoAtual, iMesAtual)
    If lErro <> SUCESSO Then gError 92600
    
    Exit Sub
        
Erro_BotaoApurar_Click:

    Select Case gErr
    
        Case 92598, 92599, 92600, 94566
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158661)

    End Select

    Exit Sub

End Sub


Private Sub Mes_Click()

Dim iMes As Integer
Dim iAno As Integer
Dim colEstoqueMes As New Collection
Dim objEstoqueMes As ClassEstoqueMes
Dim iLinha As Integer
Dim lErro As Long
Dim colEstoqueMesProduto As New Collection
Dim objEstoqueMesProduto As ClassEstoqueMesProduto
Dim objFilialEmpresa As New AdmFiliais
Dim sProdutoMascarado As String

On Error GoTo Erro_Mes_Click

    If Mes.ListIndex = -1 Then
    
        If iMesAtual <> -1 Then

            'Testa se deseja salvar mudanças
            lErro = Teste_Salva(Me, iAlterado)
            If lErro <> SUCESSO Then gError 94560

        End If

    'se trocou o Ano e o Ano Anterior era diferente de -1 ==> tenta salvar os dados da tela
    ElseIf iMesAtual <> Mes.ItemData(Mes.ListIndex) And iMesAtual <> -1 Then
    
        'Testa se deseja salvar mudanças
        lErro = Teste_Salva(Me, iAlterado)
        If lErro <> SUCESSO Then gError 92583
    

    End If

    If Mes.ListIndex = -1 Then
    
        iMesAtual = -1
        
    Else
    
        GastosDiretos.Enabled = True
        GastosIndiretos.Enabled = True
        CustoFator1.Enabled = True
        CustoFator2.Enabled = True
        CustoFator3.Enabled = True
        CustoFator4.Enabled = True
        CustoFator5.Enabled = True
        CustoFator6.Enabled = True
        
        iAno = Ano.Text

        iMes = Mes.ItemData(Mes.ListIndex)
        iMesAtual = iMes
    
        'Le todas as filiaisEmpresa da tabela EstoqueMes para o Ano/Mes em questão
        lErro = CF("EstoqueMes_Le_FiliaisEmpresa", iAno, iMes, colEstoqueMes)
        If lErro <> SUCESSO And lErro <> 92573 Then gError 92580
    
        If lErro = 92573 Then gError 92581
    
        Call Grid_Limpa(objGrid)
    
        iLinha = 0
    
        For Each objEstoqueMes In colEstoqueMes
        
            iLinha = iLinha + 1
        
            GridEstoqueMes.TextMatrix(iLinha, iGrid_FilialEmpresa_Col) = objEstoqueMes.iFilialEmpresa & SEPARADOR & objEstoqueMes.sNomeFilialEmpresa
            GridEstoqueMes.TextMatrix(iLinha, iGrid_GastosDiretos_Col) = Format(objEstoqueMes.dGastosDiretos, "Standard")
            GridEstoqueMes.TextMatrix(iLinha, iGrid_GastosIndiretos_Col) = Format(objEstoqueMes.dGastosIndiretos, "Standard")
            GridEstoqueMes.TextMatrix(iLinha, iGrid_Status_Col) = IIf(objEstoqueMes.iCustoProdApurado = CUSTO_APURADO, STRING_CUSTO_APURADO, STRING_CUSTO_NAO_APURADO)
            GridEstoqueMes.TextMatrix(iLinha, iGrid_CustoFator1_Col) = Format(objEstoqueMes.dCustoFator1, "Standard")
            GridEstoqueMes.TextMatrix(iLinha, iGrid_CustoFator2_Col) = Format(objEstoqueMes.dCustoFator2, "Standard")
            GridEstoqueMes.TextMatrix(iLinha, iGrid_CustoFator3_Col) = Format(objEstoqueMes.dCustoFator3, "Standard")
            GridEstoqueMes.TextMatrix(iLinha, iGrid_CustoFator4_Col) = Format(objEstoqueMes.dCustoFator4, "Standard")
            GridEstoqueMes.TextMatrix(iLinha, iGrid_CustoFator5_Col) = Format(objEstoqueMes.dCustoFator5, "Standard")
            GridEstoqueMes.TextMatrix(iLinha, iGrid_CustoFator6_Col) = Format(objEstoqueMes.dCustoFator6, "Standard")
            
        Next
        
        objGrid.iLinhasExistentes = iLinha
        
        Mes.Enabled = True
        
        
        'preenche uma colecao associados a EstoqueMesProduto para o Ano/Mes em questão
        lErro = CF("EstoqueMesProduto_Le", iAno, iMes, colEstoqueMesProduto)
        If lErro <> SUCESSO Then gError 92875
        
        Call Grid_Limpa(objGridGastoProduto)
    
        iLinha = 0
        
        For Each objEstoqueMesProduto In colEstoqueMesProduto
        
            iLinha = iLinha + 1
        
            objFilialEmpresa.iCodFilial = objEstoqueMesProduto.iFilialEmpresa
        
            lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
            If lErro <> SUCESSO Then gError 92889
            
            'incluido por Leo em 07/02/02
            
            lErro = Mascara_MascararProduto(objEstoqueMesProduto.sProduto, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 94322
            
            Produto.PromptInclude = False
            Produto.Text = sProdutoMascarado
            Produto.PromptInclude = True
        
            'Leo até aqui
        
            GridGastoProduto.TextMatrix(iLinha, iGrid_FilialEmpresaProd_Col) = objEstoqueMesProduto.iFilialEmpresa & SEPARADOR & objFilialEmpresa.sNome
            GridGastoProduto.TextMatrix(iLinha, iGrid_Produto_Col) = sProdutoMascarado
            GridGastoProduto.TextMatrix(iLinha, iGrid_DescricaoItem_Col) = objEstoqueMesProduto.sDescricao
            GridGastoProduto.TextMatrix(iLinha, iGrid_GastoProduto_Col) = Format(objEstoqueMesProduto.dGasto, "Standard")
            
            'Alteracao Daniel em 08/04/2002
            GridGastoProduto.TextMatrix(iLinha, iGrid_StatusProd_Col) = IIf(objEstoqueMesProduto.iCustoProdApurado = CUSTO_APURADO, STRING_CUSTO_APURADO, STRING_CUSTO_NAO_APURADO)
            'Alteracao Daniel em 08/04/2002
            
        Next
        
        objGridGastoProduto.iLinhasExistentes = iLinha 'Alterado por Leo em 30/01/02
        
    End If
        
    iAlterado = 0
    
    Exit Sub
        
Erro_Mes_Click:

    Select Case gErr
    
        Case 92580, 92583, 92889, 94560, 92875, 94322 'Alterado por Leo em 01/02/02
        
        Case 92581
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEMES_INEXISTENTE3", gErr, Ano.Text, Mes.Text)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158662)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 92584
    
    'Limpa a Tela
    lErro = CustoProducao_Limpa()
    If lErro <> SUCESSO Then gError 92615
    
    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case gErr
    
        Case 92584, 92615
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158663)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long

Dim lErro  As Long
Dim iLinha As Integer
Dim objEstoqueMes As ClassEstoqueMes
Dim objEstoqueMesProduto As ClassEstoqueMesProduto
Dim colEstoqueMes As New Collection
Dim colEstoqueMesProduto As New Collection
Dim sProduto As String
Dim iPreenchido As Integer

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    lErro = Valida_GridGastoProduto 'Incluido por leo em 31/01/02
    If lErro <> SUCESSO Then gError 94318
    
    lErro = CustoProducao_Valida_Gravacao()
    If lErro <> SUCESSO Then gError 94565
    
    'Para cada Produto atualiza o custo de Producao
    For iLinha = 1 To objGrid.iLinhasExistentes
    
        Set objEstoqueMes = New ClassEstoqueMes
        
        objEstoqueMes.iFilialEmpresa = Codigo_Extrai(GridEstoqueMes.TextMatrix(iLinha, iGrid_FilialEmpresa_Col))
        objEstoqueMes.iAno = iAnoAtual
        objEstoqueMes.iMes = iMesAtual
        objEstoqueMes.dGastosDiretos = StrParaDbl(GridEstoqueMes.TextMatrix(iLinha, iGrid_GastosDiretos_Col))
        objEstoqueMes.dGastosIndiretos = StrParaDbl(GridEstoqueMes.TextMatrix(iLinha, iGrid_GastosIndiretos_Col))
        objEstoqueMes.dCustoFator1 = StrParaDbl(GridEstoqueMes.TextMatrix(iLinha, iGrid_CustoFator1_Col))
        objEstoqueMes.dCustoFator2 = StrParaDbl(GridEstoqueMes.TextMatrix(iLinha, iGrid_CustoFator2_Col))
        objEstoqueMes.dCustoFator3 = StrParaDbl(GridEstoqueMes.TextMatrix(iLinha, iGrid_CustoFator3_Col))
        objEstoqueMes.dCustoFator4 = StrParaDbl(GridEstoqueMes.TextMatrix(iLinha, iGrid_CustoFator4_Col))
        objEstoqueMes.dCustoFator5 = StrParaDbl(GridEstoqueMes.TextMatrix(iLinha, iGrid_CustoFator5_Col))
        objEstoqueMes.dCustoFator6 = StrParaDbl(GridEstoqueMes.TextMatrix(iLinha, iGrid_CustoFator6_Col))
        
        colEstoqueMes.Add objEstoqueMes
        
    Next
    
    'Para cada Produto atualiza o custo de Producao
    For iLinha = 1 To objGridGastoProduto.iLinhasExistentes
    
        Set objEstoqueMesProduto = New ClassEstoqueMesProduto
        
        objEstoqueMesProduto.iFilialEmpresa = Codigo_Extrai(GridGastoProduto.TextMatrix(iLinha, iGrid_FilialEmpresaProd_Col))
        objEstoqueMesProduto.iAno = iAnoAtual
        objEstoqueMesProduto.iMes = iMesAtual
        lErro = CF("Produto_Formata", GridGastoProduto.TextMatrix(iLinha, iGrid_Produto_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 94319 'Alterado por Leo em 31/01/02
        objEstoqueMesProduto.sProduto = sProduto
        objEstoqueMesProduto.dGasto = StrParaDbl(GridGastoProduto.TextMatrix(iLinha, iGrid_GastoProduto_Col))
        
        colEstoqueMesProduto.Add objEstoqueMesProduto
        
    Next
    
    lErro = CF("EstoqueMes_Grava1", colEstoqueMes, colEstoqueMesProduto)
    If lErro <> SUCESSO Then gError 92590
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:
        
    Gravar_Registro = gErr
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
    
        Case 92590, 94565, 94318, 94319 'Alterado por Leo em 31/01/02
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158664)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 92591

    'Limpa a Tela
    lErro = CustoProducao_Limpa()
    If lErro <> SUCESSO Then gError 92592
    
    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 92591, 92592

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158665)

    End Select

    Exit Sub

End Sub

Function CustoProducao_Limpa() As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_CustoProducao_Limpa

    Call Grid_Limpa(objGrid)
    Call Grid_Limpa(objGridGastoProduto)
    
    iAnoAtual = -1
    Ano.ListIndex = -1
    iMesAtual = -1
    Mes.Clear
    GastosDiretos.Enabled = False
    GastosIndiretos.Enabled = False
    CustoFator1.Enabled = False
    CustoFator2.Enabled = False
    CustoFator3.Enabled = False
    CustoFator4.Enabled = False
    CustoFator5.Enabled = False
    CustoFator6.Enabled = False
    
    Exit Function
    
Erro_CustoProducao_Limpa:

    CustoProducao_Limpa = gErr
    
    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158666)

    End Select

    Exit Function
    
End Function

Public Sub Form_Load()

Dim lErro As Long
Dim objEstoqueMes As New ClassEstoqueMes
Dim objEstoqueMesAnterior As New ClassEstoqueMes
Dim objProdutosProduzidos As ClassProdutoCusto
Dim objProdutosCustoAnterior As New ClassProdutoCusto
Dim colAnos As New Collection
Dim vAno As Variant
Dim objFilial As AdmFiliais

On Error GoTo Erro_Form_Load
    
    iAnoAtual = -1
    iMesAtual = -1
    
    GastosDiretos.Enabled = False
    GastosIndiretos.Enabled = False
    CustoFator1.Enabled = False
    CustoFator2.Enabled = False
    CustoFator3.Enabled = False
    CustoFator4.Enabled = False
    CustoFator5.Enabled = False
    CustoFator6.Enabled = False
    
    'Le todos os anos da tabela EstoqueMes
    lErro = CF("EstoqueMes_Le_Anos1", colAnos)
    If lErro <> SUCESSO And lErro <> 92564 Then gError 92575
    
    'se não encontrou ------> Erro
    If lErro = 92564 Then gError 92576

    For Each vAno In colAnos
        Ano.AddItem vAno
    Next
    
    'coloca o ultimo ano encontrado como o ano default
    Ano.ListIndex = Ano.ListCount - 1

    'Inicialização do GridEstoqueMes
    Set objGrid = New AdmGrid

    Set objEventoProduto = New AdmEvento

    lErro = Inicializa_GridEstoqueMes(objGrid)
    If lErro <> SUCESSO Then gError 92609
    
    'Inicializa Máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 92885
    
    'Inicialização do GridGastoProduto
    Set objGridGastoProduto = New AdmGrid

    lErro = Inicializa_GridGastoProduto(objGridGastoProduto)
    If lErro <> SUCESSO Then gError 92860
    
    FilialEmpresaProd.Clear
    
    For Each objFilial In gcolFiliais
    
        FilialEmpresaProd.AddItem CStr(objFilial.iCodFilial) & SEPARADOR & objFilial.sNome
        FilialEmpresaProd.ItemData(FilialEmpresaProd.NewIndex) = objFilial.iCodFilial
                
    Next
    
    iFrameAtual = 1 'Incluido por Leo em 30/01/02
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
        
        Case 92575, 92609, 92860, 92885
        
        Case 92576
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEMES_INEXISTENTE6", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158667)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub FilialEmpresaProd_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub FilialEmpresaProd_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub FilialEmpresaProd_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sFornecedor As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_FilialEmpresaProd_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(FilialEmpresaProd.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If FilialEmpresaProd.Text = FilialEmpresaProd.List(FilialEmpresaProd.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialEmpresaProd, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 92876

    'Não encontrou a filial ==> erro
    If lErro <> SUCESSO Then gError 92877

    Exit Sub

Erro_FilialEmpresaProd_Validate:

    Cancel = True

    Select Case gErr

        Case 92876, 31447

        Case 92877
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, FilialEmpresaProd.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158668)

    End Select

    Exit Sub

End Sub

'Incluido por Leo em 31/01/02
Private Sub FilialEmpresaProd_KeyPress(KeyAscii As Integer)
    
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridGastoProduto)

End Sub


Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)
       
    Set objGrid = Nothing
    Set objGridGastoProduto = Nothing
    Set objEventoProduto = Nothing
    
    Unload Me
    
End Sub

Private Sub GridEstoqueMes_Click()

Dim iExecutaEntradaCelula As Integer

        Call Grid_Click(objGrid, iExecutaEntradaCelula)

        If iExecutaEntradaCelula = 1 Then
            Call Grid_Entrada_Celula(objGrid, iAlterado)
        End If
    
End Sub

Private Sub GridEstoqueMes_EnterCell()
    
    Call Grid_Entrada_Celula(objGrid, iAlterado)
    
End Sub

Private Sub GridEstoqueMes_GotFocus()
    
    Call Grid_Recebe_Foco(objGrid)

End Sub

Private Sub GridEstoqueMes_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Sub GridEstoqueMes_RowColChange()

    Call Grid_RowColChange(objGrid)

End Sub

Public Sub GridEstoqueMes_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long

On Error GoTo Erro_GridEstoqueMes_KeyDown

    Call Grid_Trata_Tecla1(KeyCode, objGrid)

    Exit Sub
    
Erro_GridEstoqueMes_KeyDown:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158669)

    End Select

    Exit Sub
    
End Sub

Private Sub GridEstoqueMes_LeaveCell()

    Call Saida_Celula(objGrid)

End Sub

Private Sub GridEstoqueMes_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGrid)
    
End Sub

Private Sub GridEstoqueMes_Scroll()

    Call Grid_Scroll(objGrid)

End Sub

Private Function Inicializa_GridEstoqueMes(objGridInt As AdmGrid) As Long
'Inicializa o Grid

Dim lErro As Long, objCamposGenericos As New ClassCamposGenericos

On Error GoTo Erro_Inicializa_GridEstoqueMes

    objCamposGenericos.lCodigo = CAMPOSGENERICOS_KIT_FATOR
    lErro = CF("CamposGenericosValores_Le_CodCampo", objCamposGenericos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("FilialEmpresa")
    objGridInt.colColuna.Add ("Gastos Diretos")
    objGridInt.colColuna.Add ("Gastos Indiretos")
    objGridInt.colColuna.Add ("Status")
    objGridInt.colColuna.Add (objCamposGenericos.colCamposGenericosValores.Item(1).sValor)
    objGridInt.colColuna.Add (objCamposGenericos.colCamposGenericosValores.Item(2).sValor)
    objGridInt.colColuna.Add (objCamposGenericos.colCamposGenericosValores.Item(3).sValor)
    objGridInt.colColuna.Add (objCamposGenericos.colCamposGenericosValores.Item(4).sValor)
    objGridInt.colColuna.Add (objCamposGenericos.colCamposGenericosValores.Item(5).sValor)
    objGridInt.colColuna.Add (objCamposGenericos.colCamposGenericosValores.Item(6).sValor)
    
    'Controles que participam do Grid
    objGridInt.colCampo.Add (FilialEmpresa.Name)
    objGridInt.colCampo.Add (GastosDiretos.Name)
    objGridInt.colCampo.Add (GastosIndiretos.Name)
    objGridInt.colCampo.Add (Status.Name)
    objGridInt.colCampo.Add (CustoFator1.Name)
    objGridInt.colCampo.Add (CustoFator2.Name)
    objGridInt.colCampo.Add (CustoFator3.Name)
    objGridInt.colCampo.Add (CustoFator4.Name)
    objGridInt.colCampo.Add (CustoFator5.Name)
    objGridInt.colCampo.Add (CustoFator6.Name)

    'Colunas do Grid
    iGrid_FilialEmpresa_Col = 1
    iGrid_GastosDiretos_Col = 2
    iGrid_GastosIndiretos_Col = 3
    iGrid_Status_Col = 4
    iGrid_CustoFator1_Col = 5
    iGrid_CustoFator2_Col = 6
    iGrid_CustoFator3_Col = 7
    iGrid_CustoFator4_Col = 8
    iGrid_CustoFator5_Col = 9
    iGrid_CustoFator6_Col = 10
    
    'Grid do GridInterno
    objGridInt.objGrid = GridEstoqueMes
    
    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 10

    'Largura da primeira coluna
    GridEstoqueMes.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Proibido Incluir Linhas
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR 'Incluido por Leo em 20/01/02

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridEstoqueMes = SUCESSO

    Exit Function
    
Erro_Inicializa_GridEstoqueMes:

    Inicializa_GridEstoqueMes = SUCESSO

    Exit Function

End Function

Private Function Inicializa_GridGastoProduto(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("FilialEmpresa")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descricao")
    objGridInt.colColuna.Add ("Gasto")
    objGridInt.colColuna.Add ("Status")
    
    'Controles que participam do Grid
    objGridInt.colCampo.Add (FilialEmpresaProd.Name)
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescricaoItem.Name)
    objGridInt.colCampo.Add (GastoProduto.Name)
    objGridInt.colCampo.Add (StatusProd.Name)

    'Colunas do Grid
    iGrid_FilialEmpresaProd_Col = 1
    iGrid_Produto_Col = 2
    iGrid_DescricaoItem_Col = 3
    iGrid_GastoProduto_Col = 4
    iGrid_StatusProd_Col = 5
    
    'Grid do GridInterno
    objGridInt.objGrid = GridGastoProduto
    
    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 6 'Alterado por Leo em 30/01/02

    'Largura da primeira coluna
    GridGastoProduto.ColWidth(0) = 400 'Alterado por Leo em 30/01/02
        
    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Irá executar a rotina grid enable
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE 'Incluido por Leo em 31/01/02
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridGastoProduto = SUCESSO

    Exit Function

End Function

Private Sub GridGastoProduto_Click()

Dim iExecutaEntradaCelula As Integer

        Call Grid_Click(objGridGastoProduto, iExecutaEntradaCelula)

        If iExecutaEntradaCelula = 1 Then
            Call Grid_Entrada_Celula(objGridGastoProduto, iAlterado)
        End If
    
End Sub

Private Sub GridGastoProduto_EnterCell()
    
    Call Grid_Entrada_Celula(objGridGastoProduto, iAlterado)
    
End Sub

Private Sub GridGastoProduto_GotFocus()
    
    Call Grid_Recebe_Foco(objGridGastoProduto)

End Sub

Private Sub GridGastoProduto_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridGastoProduto, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridGastoProduto, iAlterado)
    End If

End Sub

Sub GridGastoProduto_RowColChange()

    Call Grid_RowColChange(objGridGastoProduto)

End Sub

Public Sub GridGastoProduto_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long

On Error GoTo Erro_GridGastoProduto_KeyDown

    Call Grid_Trata_Tecla1(KeyCode, objGridGastoProduto)

    Exit Sub
    
Erro_GridGastoProduto_KeyDown:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158670)

    End Select

    Exit Sub
    
End Sub

Private Sub GridGastoProduto_LeaveCell()

    Call Saida_Celula(objGridGastoProduto)

End Sub

Private Sub GridGastoProduto_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridGastoProduto)
    
End Sub

Private Sub GridGastoProduto_Scroll()

    Call Grid_Scroll(objGridGastoProduto)

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
        
        If objGridInt Is objGrid Then
        
            Select Case GridEstoqueMes.Col
        
                Case iGrid_GastosDiretos_Col
        
                    lErro = Saida_Celula_GastosDiretos(objGridInt)
                    If lErro <> SUCESSO Then gError 92593
                        
                Case iGrid_GastosIndiretos_Col
        
                    lErro = Saida_Celula_GastosIndiretos(objGridInt)
                    If lErro <> SUCESSO Then gError 92594
                        
                Case iGrid_CustoFator1_Col
        
                    lErro = Saida_Celula_CustoFator1(objGridInt)
                    If lErro <> SUCESSO Then gError 92594
                        
                Case iGrid_CustoFator2_Col
        
                    lErro = Saida_Celula_CustoFator2(objGridInt)
                    If lErro <> SUCESSO Then gError 92594
                        
                Case iGrid_CustoFator3_Col
        
                    lErro = Saida_Celula_CustoFator3(objGridInt)
                    If lErro <> SUCESSO Then gError 92594
                        
                Case iGrid_CustoFator4_Col
        
                    lErro = Saida_Celula_CustoFator4(objGridInt)
                    If lErro <> SUCESSO Then gError 92594
                        
                Case iGrid_CustoFator5_Col
        
                    lErro = Saida_Celula_CustoFator5(objGridInt)
                    If lErro <> SUCESSO Then gError 92594
                        
                Case iGrid_CustoFator6_Col
        
                    lErro = Saida_Celula_CustoFator6(objGridInt)
                    If lErro <> SUCESSO Then gError 92594
                        
            End Select
                    
        End If
        
        If objGridInt Is objGridGastoProduto Then
                    
            Select Case GridGastoProduto.Col
                    
                Case iGrid_FilialEmpresaProd_Col
        
                    lErro = Saida_Celula_FilialEmpresaProd(objGridInt)
                    If lErro <> SUCESSO Then gError 92878
                    
                Case iGrid_Produto_Col
        
                    lErro = Saida_Celula_Produto(objGridInt)
                    If lErro <> SUCESSO Then gError 92879
                    
                Case iGrid_GastoProduto_Col
        
                    lErro = Saida_Celula_GastoProduto(objGridInt)
                    If lErro <> SUCESSO Then gError 92880
                    
            End Select
            
        End If
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 92595
       
    End If
    
    Saida_Celula = SUCESSO
    
    Exit Function

Erro_Saida_Celula:
    
    Saida_Celula = gErr
    
    Select Case gErr

        Case 92593, 92594, 92595, 92878, 92879, 92880
        
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158671)

    End Select

    Exit Function

End Function

Function Saida_Celula_GastosDiretos(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GastosDiretos

    Set objGridInt.objControle = GastosDiretos

    If Len(Trim(GastosDiretos.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(GastosDiretos)
        If lErro <> SUCESSO Then gError 92594
        
        GastosDiretos.Text = Format(GastosDiretos.Text, "Standard")

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 92595

    Saida_Celula_GastosDiretos = SUCESSO

    Exit Function

Erro_Saida_Celula_GastosDiretos:

    Saida_Celula_GastosDiretos = gErr

    Select Case gErr
        
        Case 92594, 92595
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 158672)

    End Select

    Exit Function

End Function

Function Saida_Celula_GastosIndiretos(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GastosIndiretos

    Set objGridInt.objControle = GastosIndiretos

    If Len(Trim(GastosIndiretos.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(GastosIndiretos)
        If lErro <> SUCESSO Then gError 92596
        
        GastosIndiretos.Text = Format(GastosIndiretos.Text, "Standard")

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 92597

    Saida_Celula_GastosIndiretos = SUCESSO

    Exit Function

Erro_Saida_Celula_GastosIndiretos:

    Saida_Celula_GastosIndiretos = gErr

    Select Case gErr
        
        Case 92596, 92597
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 158673)

    End Select

    Exit Function

End Function

Function Saida_Celula_FilialEmpresaProd(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_FilialEmpresaProd

    Set objGridInt.objControle = FilialEmpresaProd

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 92881

    Saida_Celula_FilialEmpresaProd = SUCESSO

    Exit Function

Erro_Saida_Celula_FilialEmpresaProd:

    Saida_Celula_FilialEmpresaProd = gErr

    Select Case gErr
        
        Case 92881
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 158674)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto

    If Len(Produto.ClipText) <> 0 Then

        lErro = CF("Produto_Critica_Estoque", Produto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25077 Then gError 92882

        If lErro = 25077 Then gError 92883

        'Descricao
        GridGastoProduto.TextMatrix(GridGastoProduto.Row, iGrid_DescricaoItem_Col) = objProduto.sDescricao
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 92884
    
    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 92882, 92884
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 92883
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, Produto.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 158675)

    End Select

    Exit Function

End Function

Function Saida_Celula_GastoProduto(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GastoProduto

    Set objGridInt.objControle = GastoProduto

    If Len(Trim(GastoProduto.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(GastoProduto)
        If lErro <> SUCESSO Then gError 92862
        
        GastoProduto.Text = Format(GastoProduto.Text, "Standard")

    End If
    
    'Trecho incluido por Leo em 31/01/02
    If Len(Trim(Produto.ClipText)) > 0 Then
    
        If (GridGastoProduto.Row - GridGastoProduto.FixedRows) = objGridGastoProduto.iLinhasExistentes Then
            objGridGastoProduto.iLinhasExistentes = objGridGastoProduto.iLinhasExistentes + 1
        End If
    
    End If
    'Leo até aqui
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 92863

    Saida_Celula_GastoProduto = SUCESSO

    Exit Function

Erro_Saida_Celula_GastoProduto:

    Saida_Celula_GastoProduto = gErr

    Select Case gErr
        
        Case 92862, 92863
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 158676)

    End Select

    Exit Function

End Function

Private Sub GastosDiretos_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub GastosDiretos_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub GastosDiretos_KeyPress(KeyAscii As Integer)
    
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub GastosDiretos_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = GastosDiretos
    
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub GastosIndiretos_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub GastosIndiretos_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub GastosIndiretos_KeyPress(KeyAscii As Integer)
    
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub GastosIndiretos_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = GastosIndiretos
    
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub GastoProduto_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub GastoProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridGastoProduto)

End Sub

Private Sub GastoProduto_KeyPress(KeyAscii As Integer)
    
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridGastoProduto)

End Sub

Private Sub GastoProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = GastoProduto
    
    lErro = Grid_Campo_Libera_Foco(objGridGastoProduto)
    If lErro <> SUCESSO Then Cancel = True

End Sub


'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CUSTO_PRODUCAO
    Set Form_Load_Ocx = Me
    Caption = "Custo de Produção"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "CustoProducao"
    
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

'Incluido por Leo em 31/01/02
Private Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

'Incluido por Leo em 31/01/02
Private Sub Produto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridGastoProduto)

End Sub

'Incluido por Leo em 31/01/02
Private Sub Produto_KeyPress(KeyAscii As Integer)
    
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridGastoProduto)

End Sub

'Incluido por Leo em 31/01/02
Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridGastoProduto.objControle = Produto
    
    lErro = Grid_Campo_Libera_Foco(objGridGastoProduto)
    If lErro <> SUCESSO Then Cancel = True

End Sub

'Função incluida por Leo em 30/01/02

Private Sub TabStrip1_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index
        
    End If


End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
    
        If Me.ActiveControl Is Produto Then
            Call BotaoProdutos_Click
        End If
    
    End If
    
End Sub


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

'**** fim do trecho a ser copiado *****

Public Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO
    
End Function

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

Private Function CustoProducao_Valida_Gravacao() As Long

Dim iLinha As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_CustoProducao_Valida_Gravacao

    If Len(Trim(Ano.Text)) = 0 Then gError 94561
    
    If Len(Trim(Mes.Text)) = 0 Then gError 94562
    
    For iLinha = 1 To objGrid.iLinhasExistentes
    
       
        If StrParaDbl(GridEstoqueMes.TextMatrix(iLinha, iGrid_GastosDiretos_Col)) = 0 Then
        
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_GASTOSDIRETOS_NAO_INFORMADO", iLinha)
            
            If vbMsgRes = vbNo Then gError 94563
            
            Exit For
        
        End If
        
        If StrParaDbl(GridEstoqueMes.TextMatrix(iLinha, iGrid_GastosIndiretos_Col)) = 0 Then
    
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_GASTOSINDIRETOS_NAO_INFORMADO", iLinha)
            
            If vbMsgRes = vbNo Then gError 94564
            
            Exit For
            
        End If
        
        If StrParaDbl(GridEstoqueMes.TextMatrix(iLinha, iGrid_GastosDiretos_Col)) - (StrParaDbl(GridEstoqueMes.TextMatrix(iLinha, iGrid_CustoFator1_Col)) + StrParaDbl(GridEstoqueMes.TextMatrix(iLinha, iGrid_CustoFator2_Col)) + StrParaDbl(GridEstoqueMes.TextMatrix(iLinha, iGrid_CustoFator3_Col)) + StrParaDbl(GridEstoqueMes.TextMatrix(iLinha, iGrid_CustoFator4_Col)) + StrParaDbl(GridEstoqueMes.TextMatrix(iLinha, iGrid_CustoFator5_Col)) + StrParaDbl(GridEstoqueMes.TextMatrix(iLinha, iGrid_CustoFator6_Col))) < -DELTA_VALORMONETARIO Then gError 201490
    
    Next
    
    For iLinha = 1 To objGridGastoProduto.iLinhasExistentes

        If StrParaDbl(GridGastoProduto.TextMatrix(iLinha, iGrid_GastoProduto_Col)) = 0 Then

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_GASTOPRODUTO_NAO_INFORMADO", iLinha)

            If vbMsgRes = vbNo Then gError 92864

            Exit For

        End If

    Next
    
    CustoProducao_Valida_Gravacao = SUCESSO
    
    Exit Function
    
Erro_CustoProducao_Valida_Gravacao:
    
    CustoProducao_Valida_Gravacao = gErr
    
    Select Case gErr
    
        Case 92864, 94563, 94564
                
        Case 94561
            Call Rotina_Erro(vbOKOnly, "ERRO_ANO_NAO_PREENCHIDO", gErr)
            
        Case 94562
            Call Rotina_Erro(vbOKOnly, "ERRO_MES_NAO_PREENCHIDO", gErr)
        
        Case 201490
            Call Rotina_Erro(vbOKOnly, "ERRO_GASTODIRETO_MENOR_PARCELAS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158677)
            
    End Select
    
    Exit Function
    
End Function

Private Sub BotaoProdutos_Click()

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim colSelecao As New Collection
Dim sSelecao As String

On Error GoTo Erro_BotaoProdutos_Click

    If GridGastoProduto.Row = 0 Then gError 92886
    
    'Verifica se existe uma filial selecionada
    If Len(Trim(GridGastoProduto.TextMatrix(GridGastoProduto.Row, iGrid_FilialEmpresaProd_Col))) = 0 Then gError 94314 'Incluido por Leo em 31/01/02
    
    If Len(Trim(GridGastoProduto.TextMatrix(GridGastoProduto.Row, iGrid_Produto_Col))) > 0 Then
        
        lErro = CF("Produto_Formata", GridGastoProduto.TextMatrix(GridGastoProduto.Row, iGrid_Produto_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 92887
        
        If iPreenchido = PRODUTO_PREENCHIDO Then objProduto.sCodigo = sProduto
    End If
    
    sSelecao = "ControleEstoque<>? AND Apropriacao=?"
    colSelecao.Add PRODUTO_CONTROLE_SEM_ESTOQUE
    colSelecao.Add PRODUTO_CUSTO_PRODUCAO
    
    Call Chama_Tela("ProdutoEstoqueLista", colSelecao, objProduto, objEventoProduto, sSelecao)

    Exit Sub

Erro_BotaoProdutos_Click:

     Select Case gErr
     
        Case 92886
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
     
        Case 92887
     
        Case 94314
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_INFORMADO", gErr, GridGastoProduto.Row)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158678)
     
     End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim sProdutoEnxuto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    If Len(Trim(GridGastoProduto.TextMatrix(GridGastoProduto.Row, iGrid_Produto_Col))) = 0 Then

        lErro = CF("Produto_Formata", GridGastoProduto.TextMatrix(GridGastoProduto.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 92888

        If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then

            sProdutoEnxuto = String(STRING_PRODUTO, 0)

            lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
            If lErro <> SUCESSO Then gError 92889

            'Lê os demais atributos do Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 92890

            If lErro = 28030 Then gError 92891
            
            Produto.PromptInclude = False
            Produto.Text = sProdutoEnxuto
            Produto.PromptInclude = True

            If Not (Me.ActiveControl Is Produto) Then
            
                GridGastoProduto.TextMatrix(GridGastoProduto.Row, iGrid_Produto_Col) = Produto.Text
                GridGastoProduto.TextMatrix(GridGastoProduto.Row, iGrid_DescricaoItem_Col) = objProduto.sDescricao
                
            End If

        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 92888, 92889, 92890

        Case 92891
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 158679)

    End Select

    Exit Sub

End Sub

'Função inserida por Leo em 31/01/02
Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iChamada As Integer)

On Error GoTo Erro_Rotina_Grid_Enable

    'Pesquisa o controle da coluna em questão
    Select Case objControl.Name

        Case FilialEmpresaProd.Name
            
            If Len(Trim(GridGastoProduto.TextMatrix(iLinha, iGrid_FilialEmpresaProd_Col))) = 0 Then
                FilialEmpresaProd.Enabled = True
            Else
                FilialEmpresaProd.Enabled = False
            End If
        
        Case Produto.Name

            If Len(Trim(GridGastoProduto.TextMatrix(iLinha, iGrid_FilialEmpresaProd_Col))) <> 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If

        
        Case GastoProduto.Name
        
            If Len(Trim(GridGastoProduto.TextMatrix(iLinha, iGrid_Produto_Col))) <> 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
        
        End Select

    
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158680)

    End Select

    Exit Sub

End Sub

'Criada por Leo em 31/01/02
Private Function Valida_GridGastoProduto() As Long
'Percorre o GridGastoProduto e em cada linha, verifica se o produto preenchido está
'relacionado com a filial selecionada na linha.

Dim lErro As Long
Dim iLinha As Integer
Dim objProdutoFilial As New ClassProdutoFilial
Dim iPreenchido As Integer
Dim sProduto As String

On Error GoTo Erro_Valida_GridGastoProduto

    If objGridGastoProduto.iLinhasExistentes = 0 Then Exit Function
    
    'Percorre as linhas do grid
    For iLinha = 1 To objGridGastoProduto.iLinhasExistentes
                
        'tenta formatar lê e formatar o produto retornando o rpoduto formatado em objProdutoFilial.sProduto
        lErro = CF("Produto_Formata", GridGastoProduto.TextMatrix(iLinha, iGrid_Produto_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then Error 64459
        
        'Caso o produto não esteja preenchido, Erro.
        If iPreenchido = PRODUTO_VAZIO Then gError 94315
        
        objProdutoFilial.sProduto = sProduto
        
        objProdutoFilial.iFilialEmpresa = Codigo_Extrai(GridGastoProduto.TextMatrix(iLinha, iGrid_FilialEmpresaProd_Col))
    
        'Verifica se o produto está relacionado com a filial preenchida
        lErro = CF("ProdutoFilial_Le", objProdutoFilial)
        If lErro <> SUCESSO And lErro <> 28261 Then gError 94316
        
        'Se não foi encontrado o produto relacionado a filial passada, Erro.
        If lErro <> SUCESSO Then gError 94317
        
    Next
    
    Exit Function
    
    Valida_GridGastoProduto = SUCESSO

Erro_Valida_GridGastoProduto:

    Valida_GridGastoProduto = gErr

    Select Case gErr

        Case 94315
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO_GRID", gErr, iLinha)
            
        Case 94316
        
        Case 94317
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTOFILIAL_INEXISTENTE_FILIALFATURAMENTO", gErr, objProdutoFilial.sProduto, objProdutoFilial.iFilialEmpresa)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158681)
    
    End Select
    
    Exit Function

End Function

Function Saida_Celula_CustoFator1(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_CustoFator1

    Set objGridInt.objControle = CustoFator1

    If Len(Trim(CustoFator1.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(CustoFator1)
        If lErro <> SUCESSO Then gError 92594
        
        CustoFator1.Text = Format(CustoFator1.Text, "Standard")

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 92595

    Saida_Celula_CustoFator1 = SUCESSO

    Exit Function

Erro_Saida_Celula_CustoFator1:

    Saida_Celula_CustoFator1 = gErr

    Select Case gErr
        
        Case 92594, 92595
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 158672)

    End Select

    Exit Function

End Function

Function Saida_Celula_CustoFator2(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_CustoFator2

    Set objGridInt.objControle = CustoFator2

    If Len(Trim(CustoFator2.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(CustoFator2)
        If lErro <> SUCESSO Then gError 92594
        
        CustoFator2.Text = Format(CustoFator2.Text, "Standard")

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 92595

    Saida_Celula_CustoFator2 = SUCESSO

    Exit Function

Erro_Saida_Celula_CustoFator2:

    Saida_Celula_CustoFator2 = gErr

    Select Case gErr
        
        Case 92594, 92595
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 158672)

    End Select

    Exit Function

End Function

Function Saida_Celula_CustoFator3(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_CustoFator3

    Set objGridInt.objControle = CustoFator3

    If Len(Trim(CustoFator3.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(CustoFator3)
        If lErro <> SUCESSO Then gError 92594
        
        CustoFator3.Text = Format(CustoFator3.Text, "Standard")

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 92595

    Saida_Celula_CustoFator3 = SUCESSO

    Exit Function

Erro_Saida_Celula_CustoFator3:

    Saida_Celula_CustoFator3 = gErr

    Select Case gErr
        
        Case 92594, 92595
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 158672)

    End Select

    Exit Function

End Function

Function Saida_Celula_CustoFator4(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_CustoFator4

    Set objGridInt.objControle = CustoFator4

    If Len(Trim(CustoFator4.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(CustoFator4)
        If lErro <> SUCESSO Then gError 92594
        
        CustoFator4.Text = Format(CustoFator4.Text, "Standard")

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 92595

    Saida_Celula_CustoFator4 = SUCESSO

    Exit Function

Erro_Saida_Celula_CustoFator4:

    Saida_Celula_CustoFator4 = gErr

    Select Case gErr
        
        Case 92594, 92595
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 158672)

    End Select

    Exit Function

End Function

Private Sub CustoFator1_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub CustoFator1_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub CustoFator1_KeyPress(KeyAscii As Integer)
    
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub CustoFator1_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = CustoFator1
    
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CustoFator2_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub CustoFator2_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub CustoFator2_KeyPress(KeyAscii As Integer)
    
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub CustoFator2_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = CustoFator2
    
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CustoFator3_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub CustoFator3_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub CustoFator3_KeyPress(KeyAscii As Integer)
    
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub CustoFator3_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = CustoFator3
    
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CustoFator4_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub CustoFator4_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub CustoFator4_KeyPress(KeyAscii As Integer)
    
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub CustoFator4_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = CustoFator4
    
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Function Saida_Celula_CustoFator6(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_CustoFator6

    Set objGridInt.objControle = CustoFator6

    If Len(Trim(CustoFator6.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(CustoFator6)
        If lErro <> SUCESSO Then gError 92594
        
        CustoFator6.Text = Format(CustoFator6.Text, "Standard")

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 92595

    Saida_Celula_CustoFator6 = SUCESSO

    Exit Function

Erro_Saida_Celula_CustoFator6:

    Saida_Celula_CustoFator6 = gErr

    Select Case gErr
        
        Case 92594, 92595
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 158672)

    End Select

    Exit Function

End Function

Function Saida_Celula_CustoFator5(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_CustoFator5

    Set objGridInt.objControle = CustoFator5

    If Len(Trim(CustoFator5.Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(CustoFator5)
        If lErro <> SUCESSO Then gError 92594
        
        CustoFator5.Text = Format(CustoFator5.Text, "Standard")

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 92595

    Saida_Celula_CustoFator5 = SUCESSO

    Exit Function

Erro_Saida_Celula_CustoFator5:

    Saida_Celula_CustoFator5 = gErr

    Select Case gErr
        
        Case 92594, 92595
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 158672)

    End Select

    Exit Function

End Function

Private Sub CustoFator5_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub CustoFator5_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub CustoFator5_KeyPress(KeyAscii As Integer)
    
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub CustoFator5_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = CustoFator5
    
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CustoFator6_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub CustoFator6_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub CustoFator6_KeyPress(KeyAscii As Integer)
    
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub CustoFator6_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = CustoFator6
    
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub


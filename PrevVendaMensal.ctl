VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl PrevVendaMensalOcx 
   ClientHeight    =   4680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7740
   ScaleHeight     =   4680
   ScaleWidth      =   7740
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5412
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   168
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "PrevVendaMensal.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1110
         Picture         =   "PrevVendaMensal.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "PrevVendaMensal.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "PrevVendaMensal.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox Filial 
      Height          =   315
      Left            =   4530
      TabIndex        =   11
      Top             =   945
      Width           =   1995
   End
   Begin VB.CheckBox CheckFixar 
      Caption         =   "Fixar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   3075
      TabIndex        =   10
      Top             =   945
      Width           =   825
   End
   Begin VB.TextBox Codigo 
      Height          =   315
      Left            =   960
      MaxLength       =   10
      TabIndex        =   9
      Top             =   300
      Width           =   1380
   End
   Begin VB.CheckBox CheckFixar 
      Caption         =   "Fixar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   -20000
      TabIndex        =   8
      Top             =   750
      Width           =   795
   End
   Begin VB.CommandButton PrevisaoVenda 
      Caption         =   "Previsões de Venda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   195
      TabIndex        =   7
      Top             =   4125
      Width           =   2010
   End
   Begin MSMask.MaskEdBox PrecoTotal 
      Height          =   225
      Left            =   3285
      TabIndex        =   5
      Top             =   2490
      Width           =   1080
      _ExtentX        =   1905
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
   Begin MSMask.MaskEdBox DataLimite 
      Height          =   225
      Left            =   4035
      TabIndex        =   6
      Top             =   2805
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox PrecoUnitario 
      Height          =   225
      Left            =   1995
      TabIndex        =   3
      Top             =   2415
      Width           =   1080
      _ExtentX        =   1905
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
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Quantidade 
      Height          =   225
      Left            =   825
      TabIndex        =   4
      Top             =   2580
      Width           =   990
      _ExtentX        =   1746
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
      Format          =   "0"
      PromptChar      =   " "
   End
   Begin VB.ComboBox Regiao 
      Height          =   315
      Left            =   -20000
      TabIndex        =   0
      Top             =   780
      Width           =   1995
   End
   Begin MSFlexGridLib.MSFlexGrid GridPrevMensal 
      Height          =   1905
      Left            =   165
      TabIndex        =   2
      Top             =   2115
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   3360
      _Version        =   393216
      Rows            =   4
      Cols            =   5
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
   End
   Begin MSMask.MaskEdBox Ano 
      Height          =   315
      Left            =   3270
      TabIndex        =   12
      Top             =   300
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Cliente 
      Height          =   315
      Left            =   960
      TabIndex        =   13
      Top             =   945
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Produto 
      Height          =   315
      Left            =   960
      TabIndex        =   14
      Top             =   1575
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin VB.Label ProdutoLabel 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   210
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   22
      Top             =   1635
      Width           =   735
   End
   Begin VB.Label Descricao 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2445
      TabIndex        =   21
      Top             =   1575
      Width           =   3510
   End
   Begin VB.Label Label5 
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   4020
      TabIndex        =   20
      Top             =   1005
      Width           =   480
   End
   Begin VB.Label LabelCliente 
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   270
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   19
      Top             =   1005
      Width           =   660
   End
   Begin VB.Label LabelCodigo 
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
      Left            =   270
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   18
      Top             =   360
      Width           =   660
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
      Left            =   2805
      TabIndex        =   17
      Top             =   360
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "U.M.:"
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
      Left            =   6045
      TabIndex        =   16
      Top             =   1635
      Width           =   480
   End
   Begin VB.Label UnidMed 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   6615
      TabIndex        =   15
      Top             =   1575
      Width           =   930
   End
   Begin VB.Label LabelRegiao 
      AutoSize        =   -1  'True
      Caption         =   "Região:"
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
      Left            =   -20000
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   1
      Top             =   840
      Width           =   675
   End
End
Attribute VB_Name = "PrevVendaMensalOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const ERRO_CODIGO_PREVVENDAMENSAL_NAO_PREENCHIDO = 0
'É obrigatório o preenchimento do codigo de Previsão de Venda Mensal.
Const ERRO_LEITURA_PREVVENDAMENSAL = 0
'Erro na leitura da tabela PrevVendaMensal.
Const ERRO_EXCLUSAO_PREVVENDAMENSAL = 0
'Erro na exclusão da tabla PrevVendaMensal.
Const ERRO_ATUALIZACAO_PREVVENDAMENSAL = 0
'Erro na atualização da tabla PrevVendaMensal.
Const ERRO_INSERCAO_PREVVENDAMENSAL = 0
'Erro na inserção da tabla PrevVendaMensal.
Const ERRO_DATAATUALIZACAO_GRID_NAO_PREENCHIDA = 0
'A data de atualização da linha %i não está preenchida
Const ERRO_REGIAO_NAO_PREENCHIDA = 0
'O preenchimento da Região Venda é obrigatório.
Const ERRO_LOCK_PREVVENDAMENSAL = 0
'Erro na tentativa de fazer lock na tabela PrevvendaMensal.
Const AVISO_EXCLUSAO_PREVVENDAMENSAL = 0

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim sProdutoAnterior As String
Dim iAlterado As Integer
Dim iClienteAlterado As Integer

Const STRING_PREVVENDAMENSAL_CODIGO = 10

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoRegiao As AdmEvento
Attribute objEventoRegiao.VB_VarHelpID = -1

'Grid
Dim objGridPrevMensal As AdmGrid
Dim iGrid_Mes_Col As Integer
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_Total_Col As Integer
Dim iGrid_Atualizacao_Col As Integer

Type typePrevVendaMensal

    dQuantidade1 As Double
    dQuantidade2 As Double
    dQuantidade3 As Double
    dQuantidade4 As Double
    dQuantidade5 As Double
    dQuantidade6 As Double
    dQuantidade7 As Double
    dQuantidade8 As Double
    dQuantidade9 As Double
    dQuantidade10 As Double
    dQuantidade11 As Double
    dQuantidade12 As Double
    dValor1 As Double
    dValor2 As Double
    dValor3 As Double
    dValor4 As Double
    dValor5 As Double
    dValor6 As Double
    dValor7 As Double
    dValor8 As Double
    dValor9 As Double
    dValor10 As Double
    dValor11 As Double
    dValor12 As Double
    dtDataAtualizacao1 As Date
    dtDataAtualizacao2 As Date
    dtDataAtualizacao3 As Date
    dtDataAtualizacao4 As Date
    dtDataAtualizacao5 As Date
    dtDataAtualizacao6 As Date
    dtDataAtualizacao7 As Date
    dtDataAtualizacao8 As Date
    dtDataAtualizacao9 As Date
    dtDataAtualizacao10 As Date
    dtDataAtualizacao11 As Date
    dtDataAtualizacao12 As Date
    
End Type

Private Sub Ano_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Ano_GotFocus()

    Call MaskEdBox_TrataGotFocus(Ano, iAlterado)

End Sub

Private Sub Ano_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Ano_Validate

    If Len(Trim(Ano.ClipText)) = 0 Then Exit Sub
    
    If StrParaInt(Ano.ClipText) < 2000 Then gError 91075

    Exit Sub

Erro_Ano_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 91075
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ANO_MENOR_2000", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165137)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objPrevVendaMensal As New ClassPrevVendaMensal
Dim vbMsgRet As VbMsgBoxResult
Dim iPreenchida As Integer
Dim sPrevVendaMensalFormatada As String

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o código foi informado
    If Len(Trim(Codigo.Text)) = 0 Then gError 91048
   
    'Verifica se o ano foi preenchido
    If Len(Trim(Ano.ClipText)) = 0 Then gError 91051
    
    'Verifica se o Regiao foi preenchido
    If Len(Trim(Regiao.Text)) = 0 Then gError 91150

    'Verifica se o Cliente foi preenchido
    If Len(Trim(Cliente.Text)) = 0 Then gError 91049
    
    'Verifica se a Filial foi preenchido
    If Len(Trim(Filial.Text)) = 0 Then gError 91151

    'Verifica se o produto foi preenchido
    If Len(Trim(Produto.ClipText)) = 0 Then gError 91050
    
    'Move dados da tela para a memória
    lErro = Move_Tela_Memoria(objPrevVendaMensal)
    If lErro <> SUCESSO Then gError 91052
    
    'Verifica se a Previsão existe
    lErro = PrevVendaMensal_Le(objPrevVendaMensal)
    If lErro <> SUCESSO And lErro <> 91152 Then gError 91053

    'Previsão mensal não está cadastrada
    If lErro = 91152 Then gError 91054

    'Pede Confirmação para exclusão ao usuário
    vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_PREVVENDAMENSAL", Codigo.Text)

    If vbMsgRet = vbYes Then

        'Exclui a previsão de venda mensal
        lErro = PrevVendaMensal_Exclui(objPrevVendaMensal)
        If lErro <> SUCESSO Then gError 91145
      
        Call Limpa_Tela_PrevVendaMensal
        

    End If
    
    lErro = ComandoSeta_Fechar(Me.Name)
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 91048
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PREVVENDAMENSAL_NAO_PREENCHIDO", gErr)
        
        Case 91049
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 91050
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case 91051
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ANO_NAO_PREENCHIDO", gErr)
        
        Case 91052, 91053, 91054, 91145
        
        Case 91150
            Call Rotina_Erro(vbOKOnly, "ERRO_REGIAO_NAO_PREENCHIDA", gErr)

        Case 91151
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165124)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()
    
    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

   'Finaliza os objEventos
    Set objEventoCodigo = Nothing
    Set objEventoProduto = Nothing
    Set objEventoCliente = Nothing
    Set objEventoRegiao = Nothing
    
    'Libera variáveis globais
    Set objGridPrevMensal = Nothing
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome
    
On Error GoTo Erro_Form_Load
    
    iAlterado = 0
    iClienteAlterado = 0
    
    'Marca Checks Para Fixar Cliente e Regiao
    CheckFixar(0).Value = vbChecked
    CheckFixar(1).Value = vbChecked
    
    'Inicializa os ObjEventos
    
    Set objEventoCodigo = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoCliente = New AdmEvento
    Set objEventoRegiao = New AdmEvento
   
    'Inicializa o GridItens
    Set objGridPrevMensal = New AdmGrid

    lErro = Inicializa_GridPrevMensal(objGridPrevMensal)
    If lErro <> SUCESSO Then Error 91148

    'Leitura dos códigos e descrições das Regiões de Venda
    lErro = CF("Cod_Nomes_Le", "RegioesVendas", "Codigo", "Descricao", STRING_REGIAO_VENDA_DESCRICAO, colCodigoNome)
    If lErro <> SUCESSO Then gError 91149

    'Preenche a ComboBox Região com código e descrição das Regiões de Venda
    For Each objCodigoNome In colCodigoNome
       Regiao.AddItem objCodigoNome.iCodigo & SEPARADOR & objCodigoNome.sNome
       Regiao.ItemData(Regiao.NewIndex) = objCodigoNome.iCodigo ' - OK ???????
    
    Next
 
   'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 91150
    
    'Formata a quantidade
    Quantidade.Format = FORMATO_ESTOQUE
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 91148, 91149, 91150

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165125)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_PREV_VENDAS
    Set Form_Load_Ocx = Me
    Caption = "Previsão de Vendas Mensal"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "PrevVendaMensal"
    
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

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim i As Integer
On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de gravação no BD
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 91055

    'Limpa a tela
    Call Limpa_Tela_PrevVendaMensal
    
    lErro = ComandoSeta_Fechar(Me.Name)
  
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 91055

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165126)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 91146

    Call Limpa_Tela_PrevVendaMensal
    
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 91146

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165127)
    End Select

    Exit Sub

End Sub

Public Function Limpa_Tela_PrevVendaMensal() As Long

Dim lErro As Long
Dim AuxCliente As String

    'Quarda o cliente
    AuxCliente = Cliente.Text
    
    'Limpa os campos da tela que não foram limpos pela função acima
    Call Limpa_Tela(Me)
    
    Cliente.Text = AuxCliente
    Call Cliente_Validate(bSGECancelDummy)
    
    Descricao.Caption = ""
    UnidMed.Caption = ""
    Codigo.Text = ""
    Ano.PromptInclude = False
    Ano.Text = ""
    Ano.PromptInclude = True
    
    If CheckFixar(0).Value = vbUnchecked Then
        Regiao.Text = ""
    End If
    
    If CheckFixar(1).Value = vbUnchecked Then
        Cliente.Text = ""
        Filial.Clear
    End If
    
    'Limpa grid
    Call Grid_Limpa(objGridPrevMensal)
    objGridPrevMensal.iLinhasExistentes = 12
     
    iAlterado = 0
    iClienteAlterado = 0
    
End Function

Private Sub Cliente_Change()

    iAlterado = REGISTRO_ALTERADO
    iClienteAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Function PrevVendaMensal_Le(objPrevVendaMensal As ClassPrevVendaMensal) As Long

Dim lErro As Long
Dim lComando As Long
Dim iFilialEmpresa As Integer

Dim sCodigo As String
Dim iCodRegiao As Integer
Dim sProduto As String
Dim tPrevVendaMensal As typePrevVendaMensal
 
On Error GoTo Erro_PrevVendaMensal_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 91056

    'Lê os dados da tabela PrevVendaMensal
    lErro = Comando_Executar(lComando, "SELECT Quantidade1, valor1, DataAtualizacao1, Quantidade2, valor2, DataAtualizacao2, Quantidade3 , valor3, DataAtualizacao3, Quantidade4, valor4, DataAtualizacao4, Quantidade5 , valor5, DataAtualizacao5, " & _
    "Quantidade6, valor6, DataAtualizacao6, Quantidade7 , valor7, DataAtualizacao7, Quantidade8, valor8, DataAtualizacao8, Quantidade9 , valor9, DataAtualizacao9, Quantidade10, valor10, DataAtualizacao10, Quantidade11, valor11, DataAtualizacao11, " & _
    "Quantidade12, valor12, DataAtualizacao12 FROM PrevVendaMensal  WHERE FilialEmpresa = ? AND Codigo = ? AND Ano = ? AND CodRegiao = ? AND Cliente = ? AND Filial = ? AND Produto = ? ", _
    tPrevVendaMensal.dQuantidade1, tPrevVendaMensal.dValor1, tPrevVendaMensal.dtDataAtualizacao1, tPrevVendaMensal.dQuantidade2, tPrevVendaMensal.dValor2, tPrevVendaMensal.dtDataAtualizacao2, tPrevVendaMensal.dQuantidade3, tPrevVendaMensal.dValor3, tPrevVendaMensal.dtDataAtualizacao3, tPrevVendaMensal.dQuantidade4, tPrevVendaMensal.dValor4, tPrevVendaMensal.dtDataAtualizacao4, tPrevVendaMensal.dQuantidade5, tPrevVendaMensal.dValor5, tPrevVendaMensal.dtDataAtualizacao5, tPrevVendaMensal.dQuantidade6, tPrevVendaMensal.dValor6, tPrevVendaMensal.dtDataAtualizacao6, tPrevVendaMensal.dQuantidade7, tPrevVendaMensal.dValor7, tPrevVendaMensal.dtDataAtualizacao7, tPrevVendaMensal.dQuantidade8, tPrevVendaMensal.dValor8, tPrevVendaMensal.dtDataAtualizacao8, _
    tPrevVendaMensal.dQuantidade9, tPrevVendaMensal.dValor9, tPrevVendaMensal.dtDataAtualizacao9, tPrevVendaMensal.dQuantidade10, tPrevVendaMensal.dValor10, tPrevVendaMensal.dtDataAtualizacao10, tPrevVendaMensal.dQuantidade11, tPrevVendaMensal.dValor11, tPrevVendaMensal.dtDataAtualizacao11, tPrevVendaMensal.dQuantidade12, tPrevVendaMensal.dValor12, tPrevVendaMensal.dtDataAtualizacao12, objPrevVendaMensal.iFilialEmpresa, objPrevVendaMensal.sCodigo, objPrevVendaMensal.iAno, objPrevVendaMensal.iCodRegiao, objPrevVendaMensal.lCliente, objPrevVendaMensal.iFilial, objPrevVendaMensal.sProduto)
    
    If lErro <> AD_SQL_SUCESSO Then gError 91057

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 91058

    'Se não encontrou --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 91152

    'carrega os dados no Grid
    objPrevVendaMensal.dQuantidade1 = tPrevVendaMensal.dQuantidade1
    objPrevVendaMensal.dValor1 = tPrevVendaMensal.dValor1
    objPrevVendaMensal.dtDataAtualizacao1 = tPrevVendaMensal.dtDataAtualizacao1
    
    objPrevVendaMensal.dQuantidade2 = tPrevVendaMensal.dQuantidade2
    objPrevVendaMensal.dValor2 = tPrevVendaMensal.dValor2
    objPrevVendaMensal.dtDataAtualizacao2 = tPrevVendaMensal.dtDataAtualizacao2
    
    objPrevVendaMensal.dQuantidade3 = tPrevVendaMensal.dQuantidade3
    objPrevVendaMensal.dValor3 = tPrevVendaMensal.dValor3
    objPrevVendaMensal.dtDataAtualizacao3 = tPrevVendaMensal.dtDataAtualizacao3
    
    objPrevVendaMensal.dQuantidade4 = tPrevVendaMensal.dQuantidade4
    objPrevVendaMensal.dValor4 = tPrevVendaMensal.dValor4
    objPrevVendaMensal.dtDataAtualizacao4 = tPrevVendaMensal.dtDataAtualizacao4
    
    objPrevVendaMensal.dQuantidade5 = tPrevVendaMensal.dQuantidade5
    objPrevVendaMensal.dValor5 = tPrevVendaMensal.dValor5
    objPrevVendaMensal.dtDataAtualizacao5 = tPrevVendaMensal.dtDataAtualizacao5
    
    objPrevVendaMensal.dQuantidade6 = tPrevVendaMensal.dQuantidade6
    objPrevVendaMensal.dValor6 = tPrevVendaMensal.dValor6
    objPrevVendaMensal.dtDataAtualizacao6 = tPrevVendaMensal.dtDataAtualizacao6
    
    objPrevVendaMensal.dQuantidade7 = tPrevVendaMensal.dQuantidade7
    objPrevVendaMensal.dValor7 = tPrevVendaMensal.dValor7
    objPrevVendaMensal.dtDataAtualizacao7 = tPrevVendaMensal.dtDataAtualizacao7
    
    objPrevVendaMensal.dQuantidade8 = tPrevVendaMensal.dQuantidade8
    objPrevVendaMensal.dValor8 = tPrevVendaMensal.dValor8
    objPrevVendaMensal.dtDataAtualizacao8 = tPrevVendaMensal.dtDataAtualizacao8
    
    objPrevVendaMensal.dQuantidade9 = tPrevVendaMensal.dQuantidade9
    objPrevVendaMensal.dValor9 = tPrevVendaMensal.dValor9
    objPrevVendaMensal.dtDataAtualizacao9 = tPrevVendaMensal.dtDataAtualizacao9
    
    objPrevVendaMensal.dQuantidade10 = tPrevVendaMensal.dQuantidade10
    objPrevVendaMensal.dValor10 = tPrevVendaMensal.dValor10
    objPrevVendaMensal.dtDataAtualizacao10 = tPrevVendaMensal.dtDataAtualizacao10
    
    objPrevVendaMensal.dQuantidade11 = tPrevVendaMensal.dQuantidade11
    objPrevVendaMensal.dValor11 = tPrevVendaMensal.dValor11
    objPrevVendaMensal.dtDataAtualizacao11 = tPrevVendaMensal.dtDataAtualizacao11
    
    objPrevVendaMensal.dQuantidade12 = tPrevVendaMensal.dQuantidade12
    objPrevVendaMensal.dValor12 = tPrevVendaMensal.dValor12
    objPrevVendaMensal.dtDataAtualizacao12 = tPrevVendaMensal.dtDataAtualizacao12
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    PrevVendaMensal_Le = SUCESSO

    Exit Function

Erro_PrevVendaMensal_Le:

    PrevVendaMensal_Le = gErr

    Select Case gErr

        
        Case 91056
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 91057, 91058
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PREVVENDAMENSAL", gErr, objPrevVendaMensal.sCodigo)

        Case 91152
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165128)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Sub Filial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LabelCliente_Click()

Dim lErro As Long
Dim objcliente As ClassCliente
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCliente_Click

    'Chama a Tela que Lista de Clientes
    Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoCliente)

    Exit Sub

Erro_LabelCliente_Click:

    Select Case gErr
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165129)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objcliente As ClassCliente

On Error GoTo Erro_objEventoCliente_evSelecao

    Set objcliente = obj1
    
    Cliente.Text = objcliente.lCodigo
    Call Cliente_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

Erro_objEventoCliente_evSelecao:

    Select Case gErr
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165130)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim lErro As Long
Dim objPrevVendaMensal As New ClassPrevVendaMensal
Dim colSelecao As Collection

On Error GoTo Erro_LabelCodigo_Click

    lErro = Move_Tela_Memoria(objPrevVendaMensal)
    If lErro <> SUCESSO Then gError 91059

    'Chama a Tela que Lista as PrevVendas
    Call Chama_Tela("PrevVendaLista", colSelecao, objPrevVendaMensal, objEventoCodigo)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case 91059

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165131)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPrevVendaMensal As ClassPrevVendaMensal

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objPrevVendaMensal = obj1
    
    '???????? Ler

    lErro = PrevVendaMensal_Le(objPrevVendaMensal)
    If lErro <> SUCESSO And lErro <> 91152 Then gError 91157

    lErro = Traz_PrevVendaMensal_Tela(objPrevVendaMensal)
    If lErro <> SUCESSO Then gError 91060

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 91060, 91157

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165132)

    End Select

    Exit Sub

End Sub

Private Sub PrevisaoVenda_Click()

Dim lErro As Long
Dim objPrevVendaMensal As New ClassPrevVendaMensal
Dim colSelecao As Collection

On Error GoTo Erro_LabelCodigo_Click

    lErro = Move_Tela_Memoria(objPrevVendaMensal)
    If lErro <> SUCESSO Then gError 91059

    'Chama a Tela que Lista as PrevVendas
    Call Chama_Tela("PrevVendaLista", colSelecao, objPrevVendaMensal, objEventoCodigo)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case 91059

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165133)

    End Select

    Exit Sub

End Sub

Private Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodFilial As Integer
Dim objcliente As New ClassCliente
Dim colFilialEmpresa As New Collection
Dim objUsuarioEmpresa As ClassUsuarioEmpresa
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Validate
   
    If iClienteAlterado <> 0 Then

        Filial.Clear
        'Se o Cliente foi preenchido
        If Len(Trim(Cliente.Text)) > 0 Then
    
            'Busca o Cliente no BD
            lErro = TP_Cliente_Le(Cliente, objcliente, iCodFilial)
            If lErro <> SUCESSO Then gError 91061
        
            lErro = CF("FiliaisClientes_Le_Cliente", objcliente, colCodigoNome)
            If lErro <> SUCESSO Then gError 91062

            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", Filial, colCodigoNome)
            Call CF("Filial_Seleciona", Filial, iCodFilial)
    
        End If
        
    End If

    iClienteAlterado = 0
    
    Exit Sub
        
Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr
    
        Case 91061, 91062
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165134)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Produto_GotFocus()

    sProdutoAnterior = Produto.Text

End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Produto_Validate

    'Se o produto foi alterado
     If sProdutoAnterior <> Produto.Text Then
        
        Descricao.Caption = ""
        UnidMed.Caption = ""

        'Caso esteja preenchido o produto estiver preenchido
        If Len(Trim(Produto.ClipText)) <> 0 Then
           
            lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 91063

            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
                objProduto.sCodigo = sProdutoFormatado
                
                lErro = Traz_Produto_Tela(objProduto)
                If lErro <> SUCESSO And lErro <> 91154 Then gError 91065
                
                'Caso não esteja Cadastrado o Prodduto no BD
                If lErro = 91154 Then gError 91067
                
            End If

        End If

        sProdutoAnterior = Produto.Text

    End If

    Exit Sub

Erro_Produto_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 91063, 91064, 91065, 91066
        
        Case 91067
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)
            If vbMsg = vbYes Then
                
                Call Chama_Tela("Produto", objProduto)

            End If
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165135)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoLabel_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_ProdutoLabel_Click

    'Verifica se o produto foi preenchido
    If Len(Produto.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 91068

        objProduto.sCodigo = sProdutoFormatado

    End If
    
    'Chama a lista de Produtos que Podem ser Vendidos
    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoProduto)
    
    Exit Sub

Erro_ProdutoLabel_Click:

    Select Case gErr

        Case 91068

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165136)

    End Select

    Exit Sub

End Sub

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objFilialCliente As New ClassFilialCliente
Dim iCodigo As Integer

On Error GoTo Erro_Filial_Validate

    'Verifica se foi preenchida a ComboBox Filial
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Filial
    If Filial.ListIndex >= 0 Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 91073

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objFilialCliente.iCodFilial = iCodigo

        'Tentativa de leitura da Filial com esse código no BD
        lErro = CF("FilialCliente_Le", objFilialCliente)
        If lErro <> SUCESSO And lErro <> 16137 Then gError 91074

        If lErro = 16137 Then gError 91075  'Não encontrou Filial no  BD

        'Encontrou Filial no BD, coloca no Text da Combo
        Filial.Text = CStr(objFilialCliente.iCodFilial) & SEPARADOR & objFilialCliente.sNome

    End If
        
    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 91076

    Exit Sub

Erro_Filial_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 91073, 91074

        Case 91076         ' - OK ????? O erro  91076 deve ter msg
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, Filial.Text)

        Case 91075
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_ENCONTRADA2", gErr, Filial.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165137)

    End Select

    Exit Sub

End Sub

Public Function Traz_Produto_Tela(objProduto As ClassProduto) As Long
'Preenche o campo do produto

Dim lErro As Long
Dim objProdutoFilial As New ClassProdutoFilial

On Error GoTo Erro_Traz_Produto_Tela
    
    'Lê o Produto que está sendo Passado
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 91077
    
    'Se ele não existir --- > ERRO
    If lErro = 28030 Then gError 91154
    
    'Se ele for Gerencial --- > ERRO
    If objProduto.iGerencial = GERENCIAL Then gError 91079
    
    'Se for Produto não vendavel ---> ERRO
    If objProduto.iFaturamento = PRODUTO_NAO_VENDAVEL Then gError 91080
    
    objProdutoFilial.sProduto = objProduto.sCodigo
    objProdutoFilial.iFilialEmpresa = giFilialEmpresa
    
    lErro = CF("ProdutoFilial_Le", objProdutoFilial)
    If lErro <> SUCESSO And lErro <> 28261 Then gError 91081
    
    'Se não encontrou
    If lErro = 28261 Then gError 91082
    
    'Preenche a Tela com o Produto e sua Descrição
    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, Produto, Descricao)
    If lErro <> SUCESSO Then gError 91083
    
    'Preenche unidade de medida
    UnidMed.Caption = objProduto.sSiglaUMVenda
        
    Traz_Produto_Tela = SUCESSO

    Exit Function

Erro_Traz_Produto_Tela:

    Traz_Produto_Tela = gErr

    Select Case gErr
        
        Case 91077, 91078, 91081, 91083, 91147
                             
        Case 91079
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)
        
        Case 91080
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PODE_SER_VENDIDO2", gErr, objProduto.sCodigo)
        
        Case 91082 'Produto não cadastrado em ProdutoFilial
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_FILIAL_INEXISTENTE", gErr, objProduto.sCodigo, giFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165138)

    End Select

    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no Banco de Dados

Dim lErro As Long
Dim objPrevVendaMensal As New ClassPrevVendaMensal

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "PrevVendaMensal"
       
    'Lê os atributos de objPrevVendaMensal que aparecem na Tela
    lErro = Move_Tela_Memoria(objPrevVendaMensal)
    If lErro <> SUCESSO Then gError 91084

    'Preenche a coleção colCampoValor, com nome do campo
    
    colCampoValor.Add "Codigo", objPrevVendaMensal.sCodigo, STRING_PREVVENDA_CODIGO, "Codigo"
    colCampoValor.Add "Ano", objPrevVendaMensal.iAno, 0, "Ano"
    colCampoValor.Add "Produto", objPrevVendaMensal.sProduto, STRING_PRODUTO, "Produto"
    colCampoValor.Add "Cliente", objPrevVendaMensal.lCliente, 0, "Cliente"
    colCampoValor.Add "CodRegiao", objPrevVendaMensal.iCodRegiao, 0, "CodRegiao"
    colCampoValor.Add "Filial", objPrevVendaMensal.iFilial, 0, "Filial"
    
    colCampoValor.Add "Quantidade1", objPrevVendaMensal.dQuantidade1, 0, "Quantidade1"
    colCampoValor.Add "Valor1", objPrevVendaMensal.dValor1, 0, "Valor1"
    colCampoValor.Add "DataAtualizacao1", objPrevVendaMensal.dtDataAtualizacao1, 0, "DataAtualizacao1"
    
    colCampoValor.Add "Quantidade2", objPrevVendaMensal.dQuantidade2, 0, "Quantidade2"
    colCampoValor.Add "Valor2", objPrevVendaMensal.dValor2, 0, "Valor2"
    colCampoValor.Add "DataAtualizacao2", objPrevVendaMensal.dtDataAtualizacao2, 0, "DataAtualizacao2"
    
    colCampoValor.Add "Quantidade3", objPrevVendaMensal.dQuantidade3, 0, "Quantidade3"
    colCampoValor.Add "Valor3", objPrevVendaMensal.dValor3, 0, "Valor3"
    colCampoValor.Add "DataAtualizacao3", objPrevVendaMensal.dtDataAtualizacao3, 0, "DataAtualizacao3"
    
    colCampoValor.Add "Quantidade4", objPrevVendaMensal.dQuantidade4, 0, "Quantidade4"
    colCampoValor.Add "Valor4", objPrevVendaMensal.dValor4, 0, "Valor4"
    colCampoValor.Add "DataAtualizacao4", objPrevVendaMensal.dtDataAtualizacao4, 0, "DataAtualizacao4"
    
    colCampoValor.Add "Quantidade5", objPrevVendaMensal.dQuantidade5, 0, "Quantidade5"
    colCampoValor.Add "Valor5", objPrevVendaMensal.dValor5, 0, "Valor5"
    colCampoValor.Add "DataAtualizacao5", objPrevVendaMensal.dtDataAtualizacao5, 0, "DataAtualizacao5"
    
    colCampoValor.Add "Quantidade6", objPrevVendaMensal.dQuantidade6, 0, "Quantidade6"
    colCampoValor.Add "Valor6", objPrevVendaMensal.dValor6, 0, "Valor6"
    colCampoValor.Add "DataAtualizacao6", objPrevVendaMensal.dtDataAtualizacao6, 0, "DataAtualizacao6"
    
    colCampoValor.Add "Quantidade7", objPrevVendaMensal.dQuantidade7, 0, "Quantidade7"
    colCampoValor.Add "Valor7", objPrevVendaMensal.dValor7, 0, "Valor7"
    colCampoValor.Add "DataAtualizacao7", objPrevVendaMensal.dtDataAtualizacao7, 0, "DataAtualizacao7"
    
    colCampoValor.Add "Quantidade8", objPrevVendaMensal.dQuantidade8, 0, "Quantidade8"
    colCampoValor.Add "Valor8", objPrevVendaMensal.dValor8, 0, "Valor8"
    colCampoValor.Add "DataAtualizacao8", objPrevVendaMensal.dtDataAtualizacao8, 0, "DataAtualizacao8"
    
    colCampoValor.Add "Quantidade9", objPrevVendaMensal.dQuantidade9, 0, "Quantidade9"
    colCampoValor.Add "Valor9", objPrevVendaMensal.dValor9, 0, "Valor9"
    colCampoValor.Add "DataAtualizacao9", objPrevVendaMensal.dtDataAtualizacao9, 0, "DataAtualizacao9"
    
    colCampoValor.Add "Quantidade10", objPrevVendaMensal.dQuantidade10, 0, "Quantidade10"
    colCampoValor.Add "Valor10", objPrevVendaMensal.dValor10, 0, "Valor10"
    colCampoValor.Add "DataAtualizacao10", objPrevVendaMensal.dtDataAtualizacao10, 0, "DataAtualizacao10"
    
    colCampoValor.Add "Quantidade11", objPrevVendaMensal.dQuantidade11, 0, "Quantidade11"
    colCampoValor.Add "Valor11", objPrevVendaMensal.dValor11, 0, "Valor11"
    colCampoValor.Add "DataAtualizacao11", objPrevVendaMensal.dtDataAtualizacao11, 0, "DataAtualizacao11"
    
    colCampoValor.Add "Quantidade12", objPrevVendaMensal.dQuantidade12, 0, "Quantidade12"
    colCampoValor.Add "Valor12", objPrevVendaMensal.dValor1, 0, "Valor12"
    colCampoValor.Add "DataAtualizacao12", objPrevVendaMensal.dtDataAtualizacao12, 0, "DataAtualizacao12"

    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr
    
        Case 91084

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165139)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do Banco de Dados

Dim lErro As Long
Dim objPrevVendaMensal As New ClassPrevVendaMensal

On Error GoTo Erro_Tela_Preenche

    'Passa os dados da coleção para objPrevVenda
    objPrevVendaMensal.sCodigo = colCampoValor.Item("Codigo").vValor
    objPrevVendaMensal.iAno = colCampoValor.Item("Ano").vValor
    objPrevVendaMensal.sProduto = colCampoValor.Item("Produto").vValor
    objPrevVendaMensal.lCliente = colCampoValor.Item("Cliente").vValor
    objPrevVendaMensal.iCodRegiao = colCampoValor.Item("CodRegiao").vValor
    objPrevVendaMensal.iFilial = colCampoValor.Item("Filial").vValor
    objPrevVendaMensal.iFilialEmpresa = giFilialEmpresa
    
    objPrevVendaMensal.dQuantidade1 = colCampoValor.Item("Quantidade1").vValor
    objPrevVendaMensal.dValor1 = colCampoValor.Item("Valor1").vValor
    objPrevVendaMensal.dtDataAtualizacao1 = colCampoValor.Item("DataAtualizacao1").vValor
    
    objPrevVendaMensal.dQuantidade2 = colCampoValor.Item("Quantidade2").vValor
    objPrevVendaMensal.dValor2 = colCampoValor.Item("Valor2").vValor
    objPrevVendaMensal.dtDataAtualizacao2 = colCampoValor.Item("DataAtualizacao2").vValor
    
    objPrevVendaMensal.dQuantidade3 = colCampoValor.Item("Quantidade3").vValor
    objPrevVendaMensal.dValor3 = colCampoValor.Item("Valor3").vValor
    objPrevVendaMensal.dtDataAtualizacao3 = colCampoValor.Item("DataAtualizacao3").vValor
    
    objPrevVendaMensal.dQuantidade4 = colCampoValor.Item("Quantidade4").vValor
    objPrevVendaMensal.dValor4 = colCampoValor.Item("Valor4").vValor
    objPrevVendaMensal.dtDataAtualizacao4 = colCampoValor.Item("DataAtualizacao4").vValor
    
    objPrevVendaMensal.dQuantidade5 = colCampoValor.Item("Quantidade5").vValor
    objPrevVendaMensal.dValor5 = colCampoValor.Item("Valor5").vValor
    objPrevVendaMensal.dtDataAtualizacao5 = colCampoValor.Item("DataAtualizacao5").vValor
    
    objPrevVendaMensal.dQuantidade6 = colCampoValor.Item("Quantidade6").vValor
    objPrevVendaMensal.dValor6 = colCampoValor.Item("Valor6").vValor
    objPrevVendaMensal.dtDataAtualizacao6 = colCampoValor.Item("DataAtualizacao6").vValor
    
    objPrevVendaMensal.dQuantidade7 = colCampoValor.Item("Quantidade7").vValor
    objPrevVendaMensal.dValor7 = colCampoValor.Item("Valor7").vValor
    objPrevVendaMensal.dtDataAtualizacao7 = colCampoValor.Item("DataAtualizacao7").vValor
    
    objPrevVendaMensal.dQuantidade8 = colCampoValor.Item("Quantidade8").vValor
    objPrevVendaMensal.dValor8 = colCampoValor.Item("Valor8").vValor
    objPrevVendaMensal.dtDataAtualizacao8 = colCampoValor.Item("DataAtualizacao8").vValor
    
    objPrevVendaMensal.dQuantidade9 = colCampoValor.Item("Quantidade9").vValor
    objPrevVendaMensal.dValor9 = colCampoValor.Item("Valor9").vValor
    objPrevVendaMensal.dtDataAtualizacao9 = colCampoValor.Item("DataAtualizacao9").vValor
    
    objPrevVendaMensal.dQuantidade10 = colCampoValor.Item("Quantidade10").vValor
    objPrevVendaMensal.dValor10 = colCampoValor.Item("Valor10").vValor
    objPrevVendaMensal.dtDataAtualizacao10 = colCampoValor.Item("DataAtualizacao10").vValor
    
    objPrevVendaMensal.dQuantidade11 = colCampoValor.Item("Quantidade11").vValor
    objPrevVendaMensal.dValor11 = colCampoValor.Item("Valor11").vValor
    objPrevVendaMensal.dtDataAtualizacao11 = colCampoValor.Item("DataAtualizacao11").vValor
    
    objPrevVendaMensal.dQuantidade12 = colCampoValor.Item("Quantidade12").vValor
    objPrevVendaMensal.dValor12 = colCampoValor.Item("Valor12").vValor
    objPrevVendaMensal.dtDataAtualizacao12 = colCampoValor.Item("DataAtualizacao12").vValor
    
    'Preenche a tela
    lErro = Traz_PrevVendaMensal_Tela(objPrevVendaMensal)
    If lErro <> SUCESSO Then gError 91085
    
    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr
        
        Case 91085
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165140)

    End Select

    Exit Sub

End Sub

Public Function Trata_Parametros(Optional objPrevVendaMensal As ClassPrevVendaMensal) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    'Se há uma previsão selecionada exibir seus dados
    If Not (objPrevVendaMensal Is Nothing) Then

        'Verifica se a previsão existe
        lErro = PrevVendaMensal_Le(objPrevVendaMensal)
        If lErro <> SUCESSO And lErro <> 91152 Then gError 91087

        If lErro = SUCESSO Then

            'A previsão está cadastrada
            lErro = Traz_PrevVendaMensal_Tela(objPrevVendaMensal)
            If lErro <> SUCESSO Then gError 91088
            
        Else

            'Previsão não está cadastrada
            Codigo.Text = objPrevVendaMensal.sCodigo

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
    
        Case 91087, 91088
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165141)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Public Function Move_Grid_Memoria(objPrevVendaMensal As ClassPrevVendaMensal) As Long
'Move as celulas do Grid para a memoria
'???? Função sem comentário

Dim lErro As Long

On Error GoTo Erro_Move_Grid_Memoria

    If Len(Trim(GridPrevMensal.TextMatrix(1, iGrid_Atualizacao_Col))) <> 0 Then
        objPrevVendaMensal.dQuantidade1 = StrParaDbl(GridPrevMensal.TextMatrix(1, iGrid_Quantidade_Col))
        objPrevVendaMensal.dValor1 = StrParaDbl(GridPrevMensal.TextMatrix(1, iGrid_Valor_Col))
        objPrevVendaMensal.dtDataAtualizacao1 = StrParaDate(GridPrevMensal.TextMatrix(1, iGrid_Atualizacao_Col))
    Else
        objPrevVendaMensal.dQuantidade1 = 0
        objPrevVendaMensal.dValor1 = 0
        objPrevVendaMensal.dtDataAtualizacao1 = DATA_NULA
    End If
    
    If Len(Trim(GridPrevMensal.TextMatrix(2, iGrid_Atualizacao_Col))) <> 0 Then
        objPrevVendaMensal.dQuantidade2 = StrParaDbl(GridPrevMensal.TextMatrix(2, iGrid_Quantidade_Col))
        objPrevVendaMensal.dValor2 = StrParaDbl(GridPrevMensal.TextMatrix(2, iGrid_Valor_Col))
        objPrevVendaMensal.dtDataAtualizacao2 = StrParaDate(GridPrevMensal.TextMatrix(2, iGrid_Atualizacao_Col))
    Else
        objPrevVendaMensal.dQuantidade2 = objPrevVendaMensal.dQuantidade1
        objPrevVendaMensal.dValor2 = objPrevVendaMensal.dValor1
        objPrevVendaMensal.dtDataAtualizacao2 = DATA_NULA
    End If
    
    If Len(Trim(GridPrevMensal.TextMatrix(3, iGrid_Atualizacao_Col))) <> 0 Then
        objPrevVendaMensal.dQuantidade3 = StrParaDbl(GridPrevMensal.TextMatrix(3, iGrid_Quantidade_Col))
        objPrevVendaMensal.dValor3 = StrParaDbl(GridPrevMensal.TextMatrix(3, iGrid_Valor_Col))
        objPrevVendaMensal.dtDataAtualizacao3 = StrParaDate(GridPrevMensal.TextMatrix(3, iGrid_Atualizacao_Col))
    Else
        objPrevVendaMensal.dQuantidade3 = objPrevVendaMensal.dQuantidade2
        objPrevVendaMensal.dValor3 = objPrevVendaMensal.dValor2
        objPrevVendaMensal.dtDataAtualizacao3 = DATA_NULA
    End If
    
    If Len(Trim(GridPrevMensal.TextMatrix(4, iGrid_Atualizacao_Col))) <> 0 Then
        objPrevVendaMensal.dQuantidade4 = StrParaDbl(GridPrevMensal.TextMatrix(4, iGrid_Quantidade_Col))
        objPrevVendaMensal.dValor4 = StrParaDbl(GridPrevMensal.TextMatrix(4, iGrid_Valor_Col))
        objPrevVendaMensal.dtDataAtualizacao4 = StrParaDate(GridPrevMensal.TextMatrix(4, iGrid_Atualizacao_Col))
    Else
        objPrevVendaMensal.dQuantidade4 = objPrevVendaMensal.dQuantidade3
        objPrevVendaMensal.dValor4 = objPrevVendaMensal.dValor3
        objPrevVendaMensal.dtDataAtualizacao4 = DATA_NULA
    End If
        
    If Len(Trim(GridPrevMensal.TextMatrix(5, iGrid_Atualizacao_Col))) <> 0 Then
        objPrevVendaMensal.dQuantidade5 = StrParaDbl(GridPrevMensal.TextMatrix(5, iGrid_Quantidade_Col))
        objPrevVendaMensal.dValor5 = StrParaDbl(GridPrevMensal.TextMatrix(5, iGrid_Valor_Col))
        objPrevVendaMensal.dtDataAtualizacao5 = StrParaDate(GridPrevMensal.TextMatrix(5, iGrid_Atualizacao_Col))
    Else
        objPrevVendaMensal.dQuantidade5 = objPrevVendaMensal.dQuantidade4
        objPrevVendaMensal.dValor5 = objPrevVendaMensal.dValor4
        objPrevVendaMensal.dtDataAtualizacao5 = DATA_NULA
    End If

    If Len(Trim(GridPrevMensal.TextMatrix(6, iGrid_Atualizacao_Col))) <> 0 Then
        objPrevVendaMensal.dQuantidade6 = StrParaDbl(GridPrevMensal.TextMatrix(6, iGrid_Quantidade_Col))
        objPrevVendaMensal.dValor6 = StrParaDbl(GridPrevMensal.TextMatrix(6, iGrid_Valor_Col))
        objPrevVendaMensal.dtDataAtualizacao6 = StrParaDate(GridPrevMensal.TextMatrix(6, iGrid_Atualizacao_Col))
    Else
       objPrevVendaMensal.dQuantidade6 = objPrevVendaMensal.dQuantidade5
       objPrevVendaMensal.dValor6 = objPrevVendaMensal.dValor5
       objPrevVendaMensal.dtDataAtualizacao6 = DATA_NULA
    End If
    
    If Len(Trim(GridPrevMensal.TextMatrix(7, iGrid_Atualizacao_Col))) <> 0 Then
        objPrevVendaMensal.dQuantidade7 = StrParaDbl(GridPrevMensal.TextMatrix(7, iGrid_Quantidade_Col))
        objPrevVendaMensal.dValor7 = StrParaDbl(GridPrevMensal.TextMatrix(7, iGrid_Valor_Col))
        objPrevVendaMensal.dtDataAtualizacao7 = StrParaDate(GridPrevMensal.TextMatrix(7, iGrid_Atualizacao_Col))
    Else
       objPrevVendaMensal.dQuantidade7 = objPrevVendaMensal.dQuantidade6
       objPrevVendaMensal.dValor7 = objPrevVendaMensal.dValor6
       objPrevVendaMensal.dtDataAtualizacao7 = DATA_NULA
    End If
    
    If Len(Trim(GridPrevMensal.TextMatrix(8, iGrid_Atualizacao_Col))) <> 0 Then
        objPrevVendaMensal.dQuantidade8 = StrParaDbl(GridPrevMensal.TextMatrix(8, iGrid_Quantidade_Col))
        objPrevVendaMensal.dValor8 = StrParaDbl(GridPrevMensal.TextMatrix(8, iGrid_Valor_Col))
        objPrevVendaMensal.dtDataAtualizacao8 = StrParaDate(GridPrevMensal.TextMatrix(8, iGrid_Atualizacao_Col))
    Else
       objPrevVendaMensal.dQuantidade8 = objPrevVendaMensal.dQuantidade7
       objPrevVendaMensal.dValor8 = objPrevVendaMensal.dValor7
       objPrevVendaMensal.dtDataAtualizacao8 = DATA_NULA
    End If
    
    If Len(Trim(GridPrevMensal.TextMatrix(9, iGrid_Atualizacao_Col))) <> 0 Then
       objPrevVendaMensal.dQuantidade9 = StrParaDbl(GridPrevMensal.TextMatrix(9, iGrid_Quantidade_Col))
       objPrevVendaMensal.dValor9 = StrParaDbl(GridPrevMensal.TextMatrix(9, iGrid_Valor_Col))
        objPrevVendaMensal.dtDataAtualizacao9 = StrParaDate(GridPrevMensal.TextMatrix(9, iGrid_Atualizacao_Col))
    Else
       objPrevVendaMensal.dQuantidade9 = objPrevVendaMensal.dQuantidade8
       objPrevVendaMensal.dValor9 = objPrevVendaMensal.dValor8
       objPrevVendaMensal.dtDataAtualizacao9 = DATA_NULA
    End If

    If Len(Trim(GridPrevMensal.TextMatrix(10, iGrid_Atualizacao_Col))) <> 0 Then
        objPrevVendaMensal.dQuantidade10 = StrParaDbl(GridPrevMensal.TextMatrix(10, iGrid_Quantidade_Col))
        objPrevVendaMensal.dValor10 = StrParaDbl(GridPrevMensal.TextMatrix(10, iGrid_Valor_Col))
        objPrevVendaMensal.dtDataAtualizacao10 = StrParaDate(GridPrevMensal.TextMatrix(10, iGrid_Atualizacao_Col))
    Else
       objPrevVendaMensal.dQuantidade10 = objPrevVendaMensal.dQuantidade9
       objPrevVendaMensal.dValor10 = objPrevVendaMensal.dValor9
       objPrevVendaMensal.dtDataAtualizacao10 = DATA_NULA
    End If
    
    If Len(Trim(GridPrevMensal.TextMatrix(11, iGrid_Atualizacao_Col))) <> 0 Then
        objPrevVendaMensal.dQuantidade11 = StrParaDbl(GridPrevMensal.TextMatrix(11, iGrid_Quantidade_Col))
        objPrevVendaMensal.dValor11 = StrParaDbl(GridPrevMensal.TextMatrix(11, iGrid_Valor_Col))
        objPrevVendaMensal.dtDataAtualizacao11 = StrParaDate(GridPrevMensal.TextMatrix(11, iGrid_Atualizacao_Col))
    Else
       objPrevVendaMensal.dQuantidade11 = objPrevVendaMensal.dQuantidade10
       objPrevVendaMensal.dValor11 = objPrevVendaMensal.dValor10
       objPrevVendaMensal.dtDataAtualizacao11 = DATA_NULA
    End If
    
   If Len(Trim(GridPrevMensal.TextMatrix(12, iGrid_Atualizacao_Col))) <> 0 Then
        objPrevVendaMensal.dQuantidade12 = StrParaDbl(GridPrevMensal.TextMatrix(12, iGrid_Quantidade_Col))
        objPrevVendaMensal.dValor12 = StrParaDbl(GridPrevMensal.TextMatrix(12, iGrid_Valor_Col))
        objPrevVendaMensal.dtDataAtualizacao12 = StrParaDate(GridPrevMensal.TextMatrix(12, iGrid_Atualizacao_Col))
    Else
       objPrevVendaMensal.dQuantidade12 = objPrevVendaMensal.dQuantidade11
       objPrevVendaMensal.dValor12 = objPrevVendaMensal.dValor11
       objPrevVendaMensal.dtDataAtualizacao12 = DATA_NULA
    End If

    Move_Grid_Memoria = SUCESSO

    Exit Function

Erro_Move_Grid_Memoria:

    Move_Grid_Memoria = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165142)

    End Select

    Exit Function

End Function

Public Function Move_Tela_Memoria(objPrevVendaMensal As ClassPrevVendaMensal) As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iPreenchido As Integer
Dim iNivel As Integer
Dim objcliente As New ClassCliente
Dim objFilialCliente As New ClassFilialCliente

On Error GoTo Erro_Move_Tela_Memoria

    objPrevVendaMensal.sCodigo = Codigo.Text
    
    If Len(Trim(Cliente.Text)) > 0 Then
    
        'Lê o Cliente a partir do Nome Reduzido
        objcliente.sNomeReduzido = Cliente.Text
        lErro = CF("Cliente_Le_NomeReduzido", objcliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 91088
        
        'Se não econtrou o Cliente, erro
        If lErro = 12348 Then gError 91089
        
        objPrevVendaMensal.lCliente = objcliente.lCodigo
        
        'Alterado por cyntia
        If objcliente.iRegiao <> 0 Then
            Regiao.Text = objcliente.iRegiao
        Else
            gError 99342
        End If
        
    End If
    
    'Verifica se o Produto foi preenchido
    If Len(Trim(Produto.ClipText)) > 0 Then
    
        'Passa para o formato do BD
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iPreenchido)
        If lErro <> SUCESSO Then gError 91090

        'testa se o codigo está preenchido
        If iPreenchido = PRODUTO_PREENCHIDO Then
            objPrevVendaMensal.sProduto = sProdutoFormatado
        End If
        
    End If

    'Se o regiao foi preenchido
     objPrevVendaMensal.iCodRegiao = Codigo_Extrai(Regiao.Text)
     objPrevVendaMensal.iFilialEmpresa = giFilialEmpresa
     objPrevVendaMensal.iFilial = Codigo_Extrai(Filial.Text)
    
    objPrevVendaMensal.iAno = StrParaInt(Ano.ClipText)
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 91088, 91090, 91091

        Case 91089
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objcliente.sNomeReduzido)
        
        Case 99342
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_RELACIONADO_VENDEDOR", gErr, objcliente.sNomeReduzido)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165143)

    End Select

    Exit Function

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objPrevVendaMensal As New ClassPrevVendaMensal
Dim sPrevVendaFormatada As String
Dim iIndice As Integer
Dim vbMsgRet As VbMsgBoxResult
 
On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
        
    'Verifica se o código foi informado
    If Len(Trim(Codigo.Text)) = 0 Then gError 91093

    'Verifica se o Ano foi informado
    If Len(Trim(Ano.ClipText)) = 0 Then gError 91094
           
    'Verifica se o Cliente foi preenchido
    If Len(Trim(Cliente.Text)) = 0 Then gError 91096
              
    'Verifica se a Filial foi preenchido
    If Len(Trim(Filial.Text)) = 0 Then gError 91098
    
    'Verifica se o produto foi informado
    If Len(Trim(Produto.ClipText)) = 0 Then gError 91095
    
    'Obriga o preenchimento da data caso o valor ou a quantidade tenham sido informados
    For iIndice = 1 To 12
        If Len(Trim(GridPrevMensal.TextMatrix(iIndice, iGrid_Quantidade_Col))) <> 0 And Len(Trim(GridPrevMensal.TextMatrix(iIndice, iGrid_Atualizacao_Col))) = 0 Or Len(Trim(GridPrevMensal.TextMatrix(iIndice, iGrid_Valor_Col))) <> 0 And Len(Trim(GridPrevMensal.TextMatrix(iIndice, iGrid_Atualizacao_Col))) = 0 Then gError 91155
    Next iIndice
    
    'Preenche objPrevVendaMensal
    lErro = Move_Tela_Memoria(objPrevVendaMensal)
    If lErro <> SUCESSO Then gError 91099
    
    lErro = Move_Grid_Memoria(objPrevVendaMensal)
    If lErro <> SUCESSO Then gError 91100
    
    'Grava Previsao de Venda no Banco de Dados
    lErro = PrevVendaMensal_Grava(objPrevVendaMensal)
    If lErro <> SUCESSO Then gError 91101

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 91093
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PREVVENDAMENSAL_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 91094
            Call Rotina_Erro(vbOKOnly, "ERRO_ANO_NAO_PREENCHIDO", gErr)
            Ano.SetFocus
        
        Case 91095
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRODUTO_NAO_PREENCHIDO", gErr)
            Produto.SetFocus
            
        Case 91096
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
            Cliente.SetFocus
            
        Case 91098
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
            Filial.SetFocus
 
        Case 91099, 91100, 91101
        
        Case 91155
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAATUALIZACAO_GRID_NAO_PREENCHIDA", gErr, iIndice)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165144)

    End Select

    Exit Function

End Function

Private Function Traz_PrevVendaMensal_Tela(objPrevVendaMensal As ClassPrevVendaMensal) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto

On Error GoTo Erro_Traz_PrevVenda_Tela

    Call Limpa_Tela_PrevVendaMensal
    
    'Preenche código de previsão
    Codigo.Text = objPrevVendaMensal.sCodigo

    'Verifica se produto retornado é válido
    objProduto.sCodigo = objPrevVendaMensal.sProduto
    
    lErro = Traz_Produto_Tela(objProduto)
    If lErro <> SUCESSO And lErro <> 91078 Then gError 91102

    'Produto não cadastrado --> Erro
    If lErro = 91078 Then gError 91103
    
    'Coloca código da Regiao em .Text e chama Validate
    Regiao.Text = CStr(objPrevVendaMensal.iCodRegiao)

    'Cliente
    Cliente.Text = objPrevVendaMensal.lCliente
    Call Cliente_Validate(bSGECancelDummy)
    
    If objPrevVendaMensal.iFilial <> 0 Then
        Filial.Text = CStr(objPrevVendaMensal.iFilial)
        Filial_Validate (bSGECancelDummy)
    End If
    
    'Ano
    Ano.Text = objPrevVendaMensal.iAno
    
    If objPrevVendaMensal.dtDataAtualizacao1 <> DATA_NULA Then
        GridPrevMensal.TextMatrix(1, iGrid_Quantidade_Col) = Format(objPrevVendaMensal.dQuantidade1, "Standard")
        GridPrevMensal.TextMatrix(1, iGrid_Valor_Col) = Format(objPrevVendaMensal.dValor1, "Standard")
        GridPrevMensal.TextMatrix(1, iGrid_Total_Col) = Format(GridPrevMensal.TextMatrix(1, iGrid_Quantidade_Col) * GridPrevMensal.TextMatrix(1, iGrid_Valor_Col), "Standard")
        GridPrevMensal.TextMatrix(1, iGrid_Atualizacao_Col) = Format(objPrevVendaMensal.dtDataAtualizacao1, "dd/mm/yyyy")
    End If
    
   If objPrevVendaMensal.dtDataAtualizacao2 <> DATA_NULA Then
        GridPrevMensal.TextMatrix(2, iGrid_Quantidade_Col) = Format(objPrevVendaMensal.dQuantidade2, "Standard")
        GridPrevMensal.TextMatrix(2, iGrid_Valor_Col) = Format(objPrevVendaMensal.dValor2, "Standard")
        GridPrevMensal.TextMatrix(2, iGrid_Total_Col) = Format(GridPrevMensal.TextMatrix(2, iGrid_Quantidade_Col) * GridPrevMensal.TextMatrix(2, iGrid_Valor_Col), "Standard")
        GridPrevMensal.TextMatrix(2, iGrid_Atualizacao_Col) = Format(objPrevVendaMensal.dtDataAtualizacao2, "dd/mm/yyyy")
    End If
    
   If objPrevVendaMensal.dtDataAtualizacao3 <> DATA_NULA Then
        GridPrevMensal.TextMatrix(3, iGrid_Quantidade_Col) = Format(objPrevVendaMensal.dQuantidade3, "Standard")
        GridPrevMensal.TextMatrix(3, iGrid_Valor_Col) = Format(objPrevVendaMensal.dValor3, "Standard")
        GridPrevMensal.TextMatrix(3, iGrid_Total_Col) = Format(GridPrevMensal.TextMatrix(3, iGrid_Quantidade_Col) * GridPrevMensal.TextMatrix(3, iGrid_Valor_Col), "Standard")
        GridPrevMensal.TextMatrix(3, iGrid_Atualizacao_Col) = Format(objPrevVendaMensal.dtDataAtualizacao3, "dd/mm/yyyy")
    End If
        
    If objPrevVendaMensal.dtDataAtualizacao4 <> DATA_NULA Then
        GridPrevMensal.TextMatrix(4, iGrid_Quantidade_Col) = Format(objPrevVendaMensal.dQuantidade4, "Standard")
        GridPrevMensal.TextMatrix(4, iGrid_Valor_Col) = Format(objPrevVendaMensal.dValor4, "Standard")
        GridPrevMensal.TextMatrix(4, iGrid_Total_Col) = Format(GridPrevMensal.TextMatrix(4, iGrid_Quantidade_Col) * GridPrevMensal.TextMatrix(4, iGrid_Valor_Col), "Standard")
        GridPrevMensal.TextMatrix(4, iGrid_Atualizacao_Col) = Format(objPrevVendaMensal.dtDataAtualizacao4, "dd/mm/yyyy")
    End If
    
    If objPrevVendaMensal.dtDataAtualizacao5 <> DATA_NULA Then
        GridPrevMensal.TextMatrix(5, iGrid_Quantidade_Col) = Format(objPrevVendaMensal.dQuantidade5, "Standard")
        GridPrevMensal.TextMatrix(5, iGrid_Valor_Col) = Format(objPrevVendaMensal.dValor5, "Standard")
        GridPrevMensal.TextMatrix(5, iGrid_Total_Col) = Format(GridPrevMensal.TextMatrix(5, iGrid_Quantidade_Col) * GridPrevMensal.TextMatrix(5, iGrid_Valor_Col), "Standard")
        GridPrevMensal.TextMatrix(5, iGrid_Atualizacao_Col) = Format(objPrevVendaMensal.dtDataAtualizacao5, "dd/mm/yyyy")
    End If
    
   If objPrevVendaMensal.dtDataAtualizacao6 <> DATA_NULA Then
        GridPrevMensal.TextMatrix(6, iGrid_Quantidade_Col) = Format(objPrevVendaMensal.dQuantidade6, "Standard")
        GridPrevMensal.TextMatrix(6, iGrid_Valor_Col) = Format(objPrevVendaMensal.dValor6, "Standard")
        GridPrevMensal.TextMatrix(6, iGrid_Total_Col) = Format(GridPrevMensal.TextMatrix(6, iGrid_Quantidade_Col) * GridPrevMensal.TextMatrix(6, iGrid_Valor_Col), "Standard")
        GridPrevMensal.TextMatrix(6, iGrid_Atualizacao_Col) = Format(objPrevVendaMensal.dtDataAtualizacao6, "dd/mm/yyyy")
    End If
    
    If objPrevVendaMensal.dtDataAtualizacao7 <> DATA_NULA Then
        GridPrevMensal.TextMatrix(7, iGrid_Quantidade_Col) = Format(objPrevVendaMensal.dQuantidade7, "Standard")
        GridPrevMensal.TextMatrix(7, iGrid_Valor_Col) = Format(objPrevVendaMensal.dValor7, "Standard")
        GridPrevMensal.TextMatrix(7, iGrid_Total_Col) = Format(GridPrevMensal.TextMatrix(7, iGrid_Quantidade_Col) * GridPrevMensal.TextMatrix(7, iGrid_Valor_Col), "Standard")
        GridPrevMensal.TextMatrix(7, iGrid_Atualizacao_Col) = Format(objPrevVendaMensal.dtDataAtualizacao7, "dd/mm/yyyy")
    End If
    
    If objPrevVendaMensal.dtDataAtualizacao8 <> DATA_NULA Then
        GridPrevMensal.TextMatrix(8, iGrid_Quantidade_Col) = Format(objPrevVendaMensal.dQuantidade8, "Standard")
        GridPrevMensal.TextMatrix(8, iGrid_Valor_Col) = Format(objPrevVendaMensal.dValor8, "Standard")
        GridPrevMensal.TextMatrix(8, iGrid_Total_Col) = Format(GridPrevMensal.TextMatrix(8, iGrid_Quantidade_Col) * GridPrevMensal.TextMatrix(8, iGrid_Valor_Col), "Standard")
        GridPrevMensal.TextMatrix(8, iGrid_Atualizacao_Col) = Format(objPrevVendaMensal.dtDataAtualizacao8, "dd/mm/yyyy")
    End If
    
    If objPrevVendaMensal.dtDataAtualizacao9 <> DATA_NULA Then
        GridPrevMensal.TextMatrix(9, iGrid_Quantidade_Col) = Format(objPrevVendaMensal.dQuantidade9, "Standard")
        GridPrevMensal.TextMatrix(9, iGrid_Valor_Col) = Format(objPrevVendaMensal.dValor9, "Standard")
        GridPrevMensal.TextMatrix(9, iGrid_Total_Col) = Format(GridPrevMensal.TextMatrix(9, iGrid_Quantidade_Col) * GridPrevMensal.TextMatrix(9, iGrid_Valor_Col), "Standard")
        GridPrevMensal.TextMatrix(9, iGrid_Atualizacao_Col) = Format(objPrevVendaMensal.dtDataAtualizacao9, "dd/mm/yyyy")
    End If
    
    If objPrevVendaMensal.dtDataAtualizacao10 <> DATA_NULA Then
        GridPrevMensal.TextMatrix(10, iGrid_Quantidade_Col) = Format(objPrevVendaMensal.dQuantidade10, "Standard")
        GridPrevMensal.TextMatrix(10, iGrid_Valor_Col) = Format(objPrevVendaMensal.dValor10, "Standard")
        GridPrevMensal.TextMatrix(10, iGrid_Total_Col) = Format(GridPrevMensal.TextMatrix(10, iGrid_Quantidade_Col) * GridPrevMensal.TextMatrix(10, iGrid_Valor_Col), "Standard")
        GridPrevMensal.TextMatrix(10, iGrid_Atualizacao_Col) = Format(objPrevVendaMensal.dtDataAtualizacao10, "dd/mm/yyyy")
    End If
    
    If objPrevVendaMensal.dtDataAtualizacao11 <> DATA_NULA Then
        GridPrevMensal.TextMatrix(11, iGrid_Quantidade_Col) = Format(objPrevVendaMensal.dQuantidade11, "Standard")
        GridPrevMensal.TextMatrix(11, iGrid_Valor_Col) = Format(objPrevVendaMensal.dValor11, "Standard")
        GridPrevMensal.TextMatrix(11, iGrid_Total_Col) = Format(GridPrevMensal.TextMatrix(11, iGrid_Quantidade_Col) * GridPrevMensal.TextMatrix(11, iGrid_Valor_Col), "Standard")
        GridPrevMensal.TextMatrix(11, iGrid_Atualizacao_Col) = Format(objPrevVendaMensal.dtDataAtualizacao11, "dd/mm/yyyy")
    End If
    
    If objPrevVendaMensal.dtDataAtualizacao12 <> DATA_NULA Then
        GridPrevMensal.TextMatrix(12, iGrid_Quantidade_Col) = Format(objPrevVendaMensal.dQuantidade12, "Standard")
        GridPrevMensal.TextMatrix(12, iGrid_Valor_Col) = Format(objPrevVendaMensal.dValor12, "Standard")
        GridPrevMensal.TextMatrix(12, iGrid_Total_Col) = Format(GridPrevMensal.TextMatrix(12, iGrid_Quantidade_Col) * GridPrevMensal.TextMatrix(12, iGrid_Valor_Col), "Standard")
        GridPrevMensal.TextMatrix(12, iGrid_Atualizacao_Col) = Format(objPrevVendaMensal.dtDataAtualizacao12, "dd/mm/yyyy")
    End If
    
    objGridPrevMensal.iLinhasExistentes = 12

    iAlterado = 0

    Traz_PrevVendaMensal_Tela = SUCESSO

    Exit Function

Erro_Traz_PrevVenda_Tela:

    Traz_PrevVendaMensal_Tela = gErr

    Select Case gErr
    
        Case 91102
        
        Case 91103
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165145)

    End Select

    Exit Function

End Function

Function PrevVendaMensal_Exclui(objPrevVendaMensal As ClassPrevVendaMensal) As Long
'Exclui a previsão de venda mensal do BD

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim sCodigo As String
Dim sProduto As String
Dim lTransacao As Long
Dim iIndice As Integer

On Error GoTo Erro_PrevVendaMensal_Exclui

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 91104
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 91105

    sCodigo = String(STRING_PREVVENDAMENSAL_CODIGO, 0)
    
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM PrevVendaMensal WHERE FilialEmpresa = ? AND Codigo = ? AND Ano = ? AND CodRegiao = ? AND Cliente = ? AND Filial = ? AND Produto = ? ", 0, sCodigo, objPrevVendaMensal.iFilialEmpresa, objPrevVendaMensal.sCodigo, objPrevVendaMensal.iAno, objPrevVendaMensal.iCodRegiao, objPrevVendaMensal.lCliente, objPrevVendaMensal.iFilial, objPrevVendaMensal.sProduto)
    If lErro <> SUCESSO Then Error 91156

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 91106
   
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then Error 91107

    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PrevVendaMensal", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 91109

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 91108

    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

   PrevVendaMensal_Exclui = SUCESSO

   Exit Function

Erro_PrevVendaMensal_Exclui:

   PrevVendaMensal_Exclui = gErr

   Select Case gErr

        Case 91104
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 91105
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 91106, 91156
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PREVVENDAMENSAL", gErr, objPrevVendaMensal.sCodigo)

        Case 91107
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PREVVENDAMENSAL", gErr)
       
        Case 91108
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 91109
             lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PREVVENDAMENSAL", gErr, objPrevVendaMensal.sCodigo)
          
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165146)

   End Select
   
   '- OK ???? Call
   Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

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

Private Sub GridPrevMensal_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridPrevMensal, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPrevMensal, iAlterado)
    End If

End Sub
Private Sub GridPrevMensal_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridPrevMensal)

End Sub

Private Sub GridItens_RowColChange()

    Call Grid_RowColChange(objGridPrevMensal)

End Sub

Private Function Inicializa_GridPrevMensal(objGridInt As AdmGrid) As Long

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Mês")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Total")
    objGridInt.colColuna.Add ("Atualização")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (PrecoUnitario.Name)
    objGridInt.colCampo.Add (PrecoTotal.Name)
    objGridInt.colCampo.Add (DataLimite.Name)
    
    'Colunas do Grid
    iGrid_Quantidade_Col = 1
    iGrid_Valor_Col = 2
    iGrid_Total_Col = 3
    iGrid_Atualizacao_Col = 4

    'Grid do GridInterno
    objGridInt.objGrid = GridPrevMensal
    
    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 13

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 6
    
    objGridInt.iLinhasExistentes = 12
    
    'Largura da primeira coluna
    GridPrevMensal.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridPrevMensal = SUCESSO

    Exit Function

End Function

Private Sub GridPrevMensal_LeaveCell()

    Call Saida_Celula(objGridPrevMensal)

End Sub

Private Sub GridPrevMensal_EnterCell()

    Call Grid_Entrada_Celula(objGridPrevMensal, iAlterado)

End Sub

Private Sub GridPrevMensal_GotFocus()

    Call Grid_Recebe_Foco(objGridPrevMensal)

End Sub

Private Sub GridPrevMensal_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

   Call Grid_Trata_Tecla(KeyAscii, objGridPrevMensal, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
   
        Call Grid_Entrada_Celula(objGridPrevMensal, iAlterado)
        
   End If

End Sub

Private Sub GridPrevMensal_Validate(Cancel As Boolean)

Dim i As Integer

    Call Grid_Libera_Foco(objGridPrevMensal)
    
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long

'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridPrevMensal)
    If lErro = SUCESSO Then

        'Verifica qual a coluna atual do Grid
        Select Case objGridInt.objGrid.Col
           
            'Quantidade
            Case iGrid_Quantidade_Col
                lErro = Saida_Celula_Quantidade(objGridInt)
                If lErro <> SUCESSO Then gError 91110
            
            'Valor
            Case iGrid_Valor_Col
                lErro = Saida_Celula_Valor(objGridPrevMensal)
                If lErro <> SUCESSO Then gError 91111
            
            'Atualizacao
            Case iGrid_Atualizacao_Col
                lErro = Saida_Celula_Atualizacao(objGridPrevMensal)
                If lErro <> SUCESSO Then gError 91113
                
        End Select
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 91114

    End If
 
    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 91110, 91111, 91112, 91113, 91114
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165147)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dQuantidade As Double
Dim dValor As Double

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    'Se quantidade estiver preenchida
    If Len(Trim(Quantidade.ClipText)) > 0 Then

        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 91116

        dQuantidade = CDbl(Quantidade.Text)
        
        'Coloca o valor Formatado na tela
        Quantidade.Text = Formata_Estoque(dQuantidade)

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 91117

    dValor = StrParaDbl(GridPrevMensal.TextMatrix(GridPrevMensal.RowSel, iGrid_Valor_Col))
    dQuantidade = StrParaDbl(GridPrevMensal.TextMatrix(GridPrevMensal.RowSel, iGrid_Quantidade_Col))

    GridPrevMensal.TextMatrix(GridPrevMensal.RowSel, iGrid_Total_Col) = IIf(dQuantidade * dValor > 0, Format(dQuantidade * dValor, "Standard"), "")

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 91116, 91117
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165148)

    End Select

    Exit Function

End Function

   
Private Function Saida_Celula_Valor(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Valor do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dValor As Double
Dim dQuantidade As Double

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridInt.objControle = PrecoUnitario

    'Se PrecoUnitario estiver preenchida
    If Len(Trim(PrecoUnitario.ClipText)) > 0 Then

        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(PrecoUnitario.Text)
        If lErro <> SUCESSO Then gError 91118

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 91119
    
    dQuantidade = StrParaDbl(GridPrevMensal.TextMatrix(GridPrevMensal.RowSel, iGrid_Quantidade_Col))
    dValor = StrParaDbl(GridPrevMensal.TextMatrix(GridPrevMensal.RowSel, iGrid_Valor_Col))

    GridPrevMensal.TextMatrix(GridPrevMensal.RowSel, iGrid_Total_Col) = IIf(dQuantidade * dValor > 0, Format(dQuantidade * dValor, "Standard"), "")
    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = gErr

    Select Case gErr

        Case 91118, 91119
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165149)

    End Select

    Exit Function

End Function


Private Function Saida_Celula_Atualizacao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Atualização do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dtAtualizacao As Date

On Error GoTo Erro_Saida_Celula_Atualizacao

    Set objGridInt.objControle = DataLimite

    'Se DataLimite estiver preenchida
    If Len(Trim(DataLimite.ClipText)) > 0 Then

        'Critica o valor
        lErro = Data_Critica(DataLimite.Text)
        If lErro <> SUCESSO Then gError 91124

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 91125

    Saida_Celula_Atualizacao = SUCESSO

    Exit Function

Erro_Saida_Celula_Atualizacao:

    Saida_Celula_Atualizacao = gErr

    Select Case gErr

        Case 91124, 91125
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165150)

    End Select

    Exit Function

End Function

'???????? APAGAR
Function Cod_Nomes_Le(sTabela As String, sCampo_Codigo As String, sCampo_Nome As String, ByVal iTamanho_Nome As Integer, colCodigoDescricao As AdmColCodigoNome) As Long
'Le todos os campos sCampo_Codigo e sCampo_Nome (de tamanho = iTamanho_Nome) da tabela sTabela e coloca na colecao
'O sCampo_Codigo deve ter tipo Inteiro.

Dim lComando As Long
Dim lErro As Long
Dim iCodigo As Integer
Dim sNome As String

On Error GoTo Erro_Cod_Nomes_Le

    lComando = 0

    sNome = String(iTamanho_Nome, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 91126

    lErro = Comando_Executar(lComando, "SELECT " & sCampo_Codigo & ", " & sCampo_Nome & " FROM " & sTabela & " ORDER BY " & sCampo_Codigo, iCodigo, sNome)
    If lErro <> AD_SQL_SUCESSO Then gError 91127

    'le o primeiro sCampo_Codigo e sCampo_Nome de sTabela
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 91128

    Do While lErro <> AD_SQL_SEM_DADOS

        'coloca o objCodigoDescricao lido na coleção
        colCodigoDescricao.Add iCodigo, sNome
       
        'le o proximo registro da tabela
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 91129
    Loop

    Call Comando_Fechar(lComando)

    Cod_Nomes_Le = SUCESSO

    Exit Function

Erro_Cod_Nomes_Le:

    Cod_Nomes_Le = gErr

    Select Case gErr

        Case 91126
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 91127, 91128, 91129
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PREVVENDAMENSAL", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165151)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Function Carrega_Filial() As Long
'Carrega a combobox Filial

Dim lErro As Long
Dim objCodigoNomeFilial As New AdmCodigoNome
Dim colCodigoNomeFilial As New AdmColCodigoNome

On Error GoTo Erro_Carrega_Filial
 
 'Leitura dos códigos e nome das Filiais
    lErro = CF("Cod_Nomes_Le", "FiliaisClientes", "CodFilial", "Nome", STRING_FILIAL_NOME, colCodigoNomeFilial)
    If lErro <> SUCESSO Then gError 91130

    'Preenche a ComboBox Filial com código e nome das filiais
    For Each objCodigoNomeFilial In colCodigoNomeFilial
        Filial.AddItem objCodigoNomeFilial.sNome
        Filial.ItemData(Filial.NewIndex) = objCodigoNomeFilial.iCodigo
    Next

    Carrega_Filial = SUCESSO

    Exit Function

Erro_Carrega_Filial:

    Carrega_Filial = gErr

    Select Case gErr

        Case 91130
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165152)

    End Select

    Exit Function

End Function

Function PrevVendaMensal_Grava(objPrevVendaMensal As ClassPrevVendaMensal) As Long

Dim lTransacao As Long
Dim alComando(0 To 2) As Long
Dim lErro As Long
Dim iIndice As Integer
Dim iFilialEmpresa As Integer
Dim sCodigo As String
Dim lCliente As Integer
Dim iAno As Integer
'Dim iCodRegiao As Integer
Dim iFilial As Integer
Dim sProduto As String
Dim iRegiao As Integer

On Error GoTo Erro_PrevVendaMensal_Grava

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 91131
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 91132

    sCodigo = String(STRING_PREVVENDAMENSAL_CODIGO, 0)
    sProduto = String(STRING_PRODUTO, 0)
    
    'se não tiver região associada a filial do cliente --> erro
    lErro = Comando_Executar(alComando(2), "SELECT Regiao FROM FiliaisClientes WHERE CodFilial = ? AND CodCliente = ? ", iRegiao, objPrevVendaMensal.iFilial, objPrevVendaMensal.lCliente)
    If lErro <> AD_SQL_SUCESSO Then gError 99331

    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 99332

    'Se Encontrou o registro
    If lErro = AD_SQL_SEM_DADOS Then gError 99333

    'correcao por tulio em 14/08/02
    'faz com que a regiao do obj seja a regiao da filial, que foi lido anteriormente...
    objPrevVendaMensal.iCodRegiao = iRegiao

    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo, Ano, Cliente, Filial, Produto FROM PrevVendaMensal WHERE FilialEmpresa = ? AND Codigo = ? AND Ano = ? AND Cliente = ? AND Filial = ? AND Produto = ? ", 0, sCodigo, iAno, lCliente, iFilial, sProduto, objPrevVendaMensal.iFilialEmpresa, objPrevVendaMensal.sCodigo, objPrevVendaMensal.iAno, objPrevVendaMensal.lCliente, objPrevVendaMensal.iFilial, objPrevVendaMensal.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 91133

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 91134

    'Se Encontrou o registro
   If lErro = AD_SQL_SUCESSO Then

        'Lock do registro
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 91135

        'Atualiza os dados da PrevVendaMensal
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE PrevVendaMensal SET CodRegiao = ?, " & _
        "Quantidade1 = ? , Valor1 = ? , DataAtualizacao1 = ? , Quantidade2 = ? , valor2 = ? , DataAtualizacao2 = ? , Quantidade3 = ? , valor3 = ? , DataAtualizacao3 = ? , " & _
        "Quantidade4 = ? , valor4 = ? , DataAtualizacao4 = ? , Quantidade5 = ? , Valor5 = ? , DataAtualizacao5 = ?, Quantidade6 = ? , valor6 = ? ,DataAtualizacao6 = ?, " & _
        "Quantidade7 = ? , valor7 = ? , DataAtualizacao7 = ? , Quantidade8 = ? , valor8 = ?, DataAtualizacao8 = ?, Quantidade9 = ? , valor9 = ?, DataAtualizacao9 = ? , " & _
        "Quantidade10 = ? , valor10 = ?, DataAtualizacao10 = ? , Quantidade11 = ? , valor11 = ? , DataAtualizacao11 = ? , Quantidade12 = ?, valor12 = ? , DataAtualizacao12 = ?", _
        alComando(0), objPrevVendaMensal.iCodRegiao, objPrevVendaMensal.dQuantidade1, objPrevVendaMensal.dValor1, objPrevVendaMensal.dtDataAtualizacao1, _
        objPrevVendaMensal.dQuantidade2, objPrevVendaMensal.dValor2, objPrevVendaMensal.dtDataAtualizacao2, _
        objPrevVendaMensal.dQuantidade3, objPrevVendaMensal.dValor3, objPrevVendaMensal.dtDataAtualizacao3, _
        objPrevVendaMensal.dQuantidade4, objPrevVendaMensal.dValor4, objPrevVendaMensal.dtDataAtualizacao4, _
        objPrevVendaMensal.dQuantidade5, objPrevVendaMensal.dValor5, objPrevVendaMensal.dtDataAtualizacao5, _
        objPrevVendaMensal.dQuantidade6, objPrevVendaMensal.dValor6, objPrevVendaMensal.dtDataAtualizacao6, _
        objPrevVendaMensal.dQuantidade7, objPrevVendaMensal.dValor7, objPrevVendaMensal.dtDataAtualizacao7, _
        objPrevVendaMensal.dQuantidade8, objPrevVendaMensal.dValor8, objPrevVendaMensal.dtDataAtualizacao8, _
        objPrevVendaMensal.dQuantidade9, objPrevVendaMensal.dValor9, objPrevVendaMensal.dtDataAtualizacao9, _
        objPrevVendaMensal.dQuantidade10, objPrevVendaMensal.dValor10, objPrevVendaMensal.dtDataAtualizacao10, _
        objPrevVendaMensal.dQuantidade11, objPrevVendaMensal.dValor11, objPrevVendaMensal.dtDataAtualizacao11, _
        objPrevVendaMensal.dQuantidade12, objPrevVendaMensal.dValor12, objPrevVendaMensal.dtDataAtualizacao12)

       If lErro <> AD_SQL_SUCESSO Then gError 91136

   Else
        'Insere os dados em PrevVendaMensal
        lErro = Comando_Executar(alComando(1), "INSERT INTO PrevVendaMensal(FilialEmpresa, Codigo, Ano, CodRegiao,Cliente, Filial, Produto, " & _
        "Quantidade1, Valor1, DataAtualizacao1, Quantidade2, valor2, DataAtualizacao2, Quantidade3, valor3, DataAtualizacao3, " & _
        "Quantidade4, valor4, DataAtualizacao4, Quantidade5, Valor5, DataAtualizacao5, Quantidade6, valor6,DataAtualizacao6, " & _
        "Quantidade7, valor7, DataAtualizacao7, Quantidade8, valor8, DataAtualizacao8, Quantidade9, valor9, DataAtualizacao9," & _
        "Quantidade10, valor10, DataAtualizacao10, Quantidade11, valor11, DataAtualizacao11, Quantidade12, valor12, DataAtualizacao12) " & _
        "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
        objPrevVendaMensal.iFilialEmpresa, objPrevVendaMensal.sCodigo, objPrevVendaMensal.iAno, objPrevVendaMensal.iCodRegiao, _
        objPrevVendaMensal.lCliente, objPrevVendaMensal.iFilial, objPrevVendaMensal.sProduto, objPrevVendaMensal.dQuantidade1, objPrevVendaMensal.dValor1, objPrevVendaMensal.dtDataAtualizacao1, _
        objPrevVendaMensal.dQuantidade2, objPrevVendaMensal.dValor2, objPrevVendaMensal.dtDataAtualizacao2, _
        objPrevVendaMensal.dQuantidade3, objPrevVendaMensal.dValor3, objPrevVendaMensal.dtDataAtualizacao3, _
        objPrevVendaMensal.dQuantidade4, objPrevVendaMensal.dValor4, objPrevVendaMensal.dtDataAtualizacao4, _
        objPrevVendaMensal.dQuantidade5, objPrevVendaMensal.dValor5, objPrevVendaMensal.dtDataAtualizacao5, _
        objPrevVendaMensal.dQuantidade6, objPrevVendaMensal.dValor6, objPrevVendaMensal.dtDataAtualizacao6, _
        objPrevVendaMensal.dQuantidade7, objPrevVendaMensal.dValor7, objPrevVendaMensal.dtDataAtualizacao7, _
        objPrevVendaMensal.dQuantidade8, objPrevVendaMensal.dValor8, objPrevVendaMensal.dtDataAtualizacao8, _
        objPrevVendaMensal.dQuantidade9, objPrevVendaMensal.dValor9, objPrevVendaMensal.dtDataAtualizacao9, _
        objPrevVendaMensal.dQuantidade10, objPrevVendaMensal.dValor10, objPrevVendaMensal.dtDataAtualizacao10, _
        objPrevVendaMensal.dQuantidade11, objPrevVendaMensal.dValor11, objPrevVendaMensal.dtDataAtualizacao11, _
        objPrevVendaMensal.dQuantidade12, objPrevVendaMensal.dValor12, objPrevVendaMensal.dtDataAtualizacao12)
        
        If lErro <> AD_SQL_SUCESSO Then gError 91137

    End If

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 91138

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PrevVendaMensal_Grava = SUCESSO

    Exit Function

Erro_PrevVendaMensal_Grava:

    PrevVendaMensal_Grava = gErr

    Select Case gErr

        Case 91131
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 91132
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 91133, 91134
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PREVVENDAMENSAL", gErr, objPrevVendaMensal.sCodigo)
        
        Case 91135
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PREVVENDAMENSAL", gErr)

        Case 91136
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PREVVENDAMENSAL", gErr)

        Case 91137
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PREVVENDAMENSAL", gErr)

        Case 91138
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 99331, 99332
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAIS", gErr)
        
        Case 99333
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGIAO_NAO_SELECIONADA2", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165153)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Quantidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridPrevMensal)

End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPrevMensal)

End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridPrevMensal.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGridPrevMensal)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub PrecoUnitario_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PrecoUnitario_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridPrevMensal)

End Sub

Private Sub PrecoUnitario_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPrevMensal)

End Sub

Private Sub PrecoUnitario_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridPrevMensal.objControle = PrecoUnitario
    lErro = Grid_Campo_Libera_Foco(objGridPrevMensal)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub PrecoTotal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PrecoTotal_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridPrevMensal)

End Sub

Private Sub PrecoTotal_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPrevMensal)

End Sub

Private Sub PrecoTotal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridPrevMensal.objControle = PrecoTotal
    lErro = Grid_Campo_Libera_Foco(objGridPrevMensal)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub DataLimite_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataLimite_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridPrevMensal)

End Sub

Private Sub DataLimite_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPrevMensal)

End Sub

Private Sub DataLimite_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridPrevMensal.objControle = DataLimite
    lErro = Grid_Campo_Libera_Foco(objGridPrevMensal)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub GridPrevMensal_RowColChange()

    Call Grid_RowColChange(objGridPrevMensal)

End Sub

Private Sub GridPrevMensal_Scroll()

    Call Grid_Scroll(objGridPrevMensal)

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1
    
    'Preenche o Produto
    lErro = Traz_Produto_Tela(objProduto)
    If lErro <> SUCESSO And lErro <> 40050 Then gError 91139
    
    If lErro = 40050 Then gError 91140
    
    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 91139
        
        Case 91140
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165154)

    End Select

    Exit Sub

End Sub

'*************************************************
'*************apagar*** já existem em rotinas CPR
'*************************************************
Function FiliaisClientes_Le_Cliente(objcliente As ClassCliente, colCodigoNome As AdmColCodigoNome) As Long
'Le na tabela FiliaisClientes todos os Codigos e Nomes de Filiais
'relacionadas ao objCliente. Retorna na colecao colCodigoNome

Dim lComando As Long
Dim iCodFilial As Integer
Dim sNome As String
Dim lErro As Long

On Error GoTo Erro_FiliaisClientes_Le_Cliente

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 91140

    sNome = String(STRING_FILIAL_CLIENTE_NOME, 0)

    lErro = Comando_Executar(lComando, "SELECT CodFilial, Nome FROM FiliaisClientes WHERE CodCliente=? ORDER BY CodFilial", iCodFilial, sNome, objcliente.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 91141

    'le a primeira filial
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 91145
    If lErro = AD_SQL_SEM_DADOS Then gError 91142

    Do While lErro <> AD_SQL_SEM_DADOS

        'coloca a filial lida na coleção
        colCodigoNome.Add iCodFilial, sNome

        'le a proxima filial
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 91143

    Loop

    lErro = Comando_Fechar(lComando)

    FiliaisClientes_Le_Cliente = SUCESSO

    Exit Function

Erro_FiliaisClientes_Le_Cliente:

    FiliaisClientes_Le_Cliente = gErr

    Select Case Err

        Case 91140
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 91141
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTES", gErr)

        Case 91142
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_SEM_FILIAL", gErr, objcliente.lCodigo)
        
        Case 91143
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165155)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function


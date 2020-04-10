VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl PrevVendaOcx 
   ClientHeight    =   4860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6315
   KeyPreview      =   -1  'True
   ScaleHeight     =   4860
   ScaleWidth      =   6315
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
      Left            =   3630
      TabIndex        =   8
      Top             =   2745
      Width           =   1215
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
      Left            =   3630
      TabIndex        =   6
      Top             =   2205
      Width           =   1215
   End
   Begin VB.TextBox Codigo 
      Height          =   315
      Left            =   975
      MaxLength       =   10
      TabIndex        =   1
      Top             =   225
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3945
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   180
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "PrevVendaUsuOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "PrevVendaUsuOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "PrevVendaUsuOcx.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "PrevVendaUsuOcx.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox Regiao 
      Enabled         =   0   'False
      Height          =   315
      Left            =   10200
      TabIndex        =   22
      Top             =   4290
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.Frame FrameData 
      Caption         =   "Datas do Período"
      Height          =   795
      Left            =   225
      TabIndex        =   17
      Top             =   1245
      Width           =   5310
      Begin MSComCtl2.UpDown UpDown3 
         Height          =   315
         Left            =   4800
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   315
         Left            =   2445
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   300
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   315
         Left            =   1245
         TabIndex        =   3
         Top             =   300
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   315
         Left            =   3600
         TabIndex        =   4
         Top             =   315
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Final:"
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
         Left            =   3090
         TabIndex        =   21
         Top             =   375
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Inicial:"
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
         Left            =   630
         TabIndex        =   20
         Top             =   360
         Width           =   585
      End
   End
   Begin VB.CommandButton ValorTabela 
      Enabled         =   0   'False
      Height          =   345
      Left            =   7665
      Picture         =   "PrevVendaUsuOcx.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4305
      Visible         =   0   'False
      Width           =   1905
   End
   Begin MSMask.MaskEdBox Cliente 
      Height          =   315
      Left            =   1425
      TabIndex        =   7
      Top             =   2745
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Quantidade 
      Height          =   315
      Left            =   1440
      TabIndex        =   10
      Top             =   3840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   315
      Left            =   3915
      TabIndex        =   12
      Top             =   4365
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   556
      _Version        =   393216
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
   Begin MSMask.MaskEdBox DataPrevisao 
      Height          =   315
      Left            =   1590
      TabIndex        =   2
      Top             =   750
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDown 
      Height          =   315
      Left            =   2790
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   750
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox Produto 
      Height          =   315
      Left            =   1425
      TabIndex        =   9
      Top             =   3300
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Vendedor 
      Height          =   315
      Left            =   1425
      TabIndex        =   5
      Top             =   2205
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox PrecoUnitario 
      Height          =   315
      Left            =   1425
      TabIndex        =   11
      Top             =   4380
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   556
      _Version        =   393216
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
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Preço Unitário:"
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
      TabIndex        =   37
      Top             =   4410
      Width           =   1290
   End
   Begin VB.Label LabelVendedor 
      AutoSize        =   -1  'True
      Caption         =   "Vendedor:"
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
      Left            =   510
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   36
      Top             =   2250
      Width           =   885
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
      Left            =   660
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   35
      Top             =   3360
      Width           =   735
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
      Left            =   3375
      TabIndex        =   34
      Top             =   3900
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descrição:"
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
      Left            =   6495
      TabIndex        =   33
      Top             =   4365
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Label Descricao 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2910
      TabIndex        =   32
      Top             =   3285
      Width           =   3150
   End
   Begin VB.Label Label5 
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   3300
      TabIndex        =   31
      Top             =   4455
      Width           =   510
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Data Previsão:"
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
      Left            =   240
      TabIndex        =   30
      Top             =   810
      Width           =   1275
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   315
      TabIndex        =   29
      Top             =   3885
      Width           =   1050
   End
   Begin VB.Label UnidMed 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3915
      TabIndex        =   28
      Top             =   3840
      Width           =   930
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
      Left            =   735
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   27
      Top             =   2805
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
      Left            =   285
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   26
      Top             =   285
      Width           =   660
   End
   Begin VB.Label RegiaoVendaLabel 
      AutoSize        =   -1  'True
      Caption         =   "Região:"
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
      Height          =   195
      Left            =   9480
      TabIndex        =   25
      Top             =   4335
      Visible         =   0   'False
      Width           =   675
   End
End
Attribute VB_Name = "PrevVendaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'ALTERAÇÕES INPAL
'Márcio 09/2000 - Incluido PrevVendas por Clientes do Vendedor

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim sProdutoAnterior As String
Dim iAlterado As Integer

Dim WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Dim WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1

'Márcio 09/2000 - Incluido PrevVendas por Clientes do Vendedor
Dim WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Dim WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1

Public Function Move_Tela_Memoria(objPrevVenda As ClassPrevVenda) As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iPreenchido As Integer
Dim iNivel As Integer
Dim objcliente As New ClassCliente
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Move_Tela_Memoria

    objPrevVenda.sCodigo = Codigo.Text
        
    'Verifica se o Produto foi preenchido
    If Len(Trim(Produto.ClipText)) > 0 Then
    
        'Passa para o formato do BD
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iPreenchido)
        If lErro <> SUCESSO Then gError 34588

        'testa se o codigo está preenchido
        If iPreenchido = PRODUTO_PREENCHIDO Then
            objPrevVenda.sProduto = sProdutoFormatado
        End If
        
    End If

    'Recolhe os demais dados
    If Len(Trim(Regiao.Text)) = 0 Then
        objPrevVenda.iCodRegiao = 0
    Else
        objPrevVenda.iCodRegiao = Codigo_Extrai(Regiao.Text)
    End If
    
    If Len(Trim(Quantidade.ClipText)) > 0 Then
        objPrevVenda.dQuantidade = CDbl(Quantidade.Text)
    Else
        objPrevVenda.dQuantidade = 0
    End If

    If Len(Trim(Valor.ClipText)) > 0 Then
        objPrevVenda.dValor = CDbl(Valor.ClipText)
    Else
        objPrevVenda.dValor = 0
    End If

    If Len(Trim(DataPrevisao.ClipText)) > 0 Then
        objPrevVenda.dtDataPrevisao = CDate(DataPrevisao.Text)
    Else
        objPrevVenda.dtDataPrevisao = DATA_NULA
    End If
        
    'Recolhe a data Inicial
    If Len(Trim(DataInicial.ClipText)) > 0 Then
        objPrevVenda.dtDataInicio = CDate(DataInicial.Text)
    Else
        objPrevVenda.dtDataInicio = DATA_NULA
    End If
            
    'Recolhe a data Final
    If Len(Trim(DataFinal.ClipText)) > 0 Then
        objPrevVenda.dtDataFim = CDate(DataFinal.Text)
    Else
        objPrevVenda.dtDataFim = DATA_NULA
    End If
    
    objPrevVenda.iFilialEmpresa = giFilialEmpresa
    
    '*** Inicio - Márcio 09/2000 - Incluido PrevVendas por Clientes do Vendedor
    
    'Se o cliente foi preenchido
    If Len(Trim(Cliente.Text)) > 0 Then
    
        'Lê o Cliente a partir do Nome Reduzido
        objcliente.sNomeReduzido = Cliente.Text
        lErro = CF("Cliente_Le_NomeReduzido", objcliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 500182
        
        'Se não econtrou o Cliente, erro
        If lErro = 12348 Then gError 500183
        
        objPrevVenda.lCliente = objcliente.lCodigo
        
    End If
    
    'Se o vendedor foi preenchido
    If Len(Trim(Vendedor.Text)) > 0 Then
    
        'Lê o Vendedor a partir do Nome Reduzido
        objVendedor.sNomeReduzido = Vendedor.Text
        lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
        If lErro <> SUCESSO And lErro <> 25008 Then gError 500184
        
        'Se não econtrou o vendedor, erro
        If lErro = 25008 Then gError 500185
        
        objPrevVenda.iVendedor = objVendedor.iCodigo
            
    End If
    
    objPrevVenda.dPrecoUnitario = StrParaDbl(PrecoUnitario.Text)
    '*** Fim - Márcio 09/2000 - Incluido PrevVendas por Clientes do Vendedor
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 34588, 500182, 500184

        Case 500183
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objcliente.sNomeReduzido)
            
        Case 500185
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO1", gErr, objcliente.sNomeReduzido)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no Banco de Dados

Dim lErro As Long
Dim objPrevVenda As New ClassPrevVenda

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "PrevVenda"

    'Lê os atributos de objPrevVenda que aparecem na Tela
    lErro = Move_Tela_Memoria(objPrevVenda)
    If lErro <> SUCESSO Then gError 34602

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do Banco de Dados), tamanho do campo
    'no Banco de Dados no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objPrevVenda.sCodigo, STRING_PREVVENDA_CODIGO, "Codigo"
    colCampoValor.Add "Produto", objPrevVenda.sProduto, STRING_PRODUTO, "Produto"
    colCampoValor.Add "DataPrevisao", objPrevVenda.dtDataPrevisao, 0, "DataPrevisao"
    colCampoValor.Add "DataInicio", objPrevVenda.dtDataInicio, 0, "DataInicio"
    colCampoValor.Add "DataFim", objPrevVenda.dtDataFim, 0, "DataFim"
    colCampoValor.Add "Cliente", objPrevVenda.lCliente, 0, "Cliente"
    colCampoValor.Add "Vendedor", objPrevVenda.iVendedor, 0, "Vendedor"
    colCampoValor.Add "CodRegiao", objPrevVenda.iCodRegiao, 0, "CodRegiao"
    colCampoValor.Add "FilialEmpresa", objPrevVenda.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "Quantidade", objPrevVenda.dQuantidade, 0, "Quantidade"
    colCampoValor.Add "PrecoUnitario", objPrevVenda.dPrecoUnitario, 0, "PrecoUnitario"
    colCampoValor.Add "Valor", objPrevVenda.dValor, 0, "Valor"

    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 34602

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do Banco de Dados

Dim lErro As Long
Dim objPrevVenda As New ClassPrevVenda

On Error GoTo Erro_Tela_Preenche

    'Passa os dados da coleção para objPrevVenda
    objPrevVenda.sCodigo = colCampoValor.Item("Codigo").vValor
    objPrevVenda.sProduto = colCampoValor.Item("Produto").vValor
    objPrevVenda.dtDataPrevisao = colCampoValor.Item("DataPrevisao").vValor
    objPrevVenda.dtDataInicio = colCampoValor.Item("DataInicio").vValor
    objPrevVenda.dtDataFim = colCampoValor.Item("DataFim").vValor
    objPrevVenda.lCliente = colCampoValor.Item("Cliente").vValor
    objPrevVenda.iVendedor = colCampoValor.Item("Vendedor").vValor
    objPrevVenda.iCodRegiao = colCampoValor.Item("CodRegiao").vValor
    objPrevVenda.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
    objPrevVenda.dQuantidade = colCampoValor.Item("Quantidade").vValor
    objPrevVenda.dPrecoUnitario = colCampoValor.Item("PrecoUnitario").vValor
    objPrevVenda.dValor = colCampoValor.Item("Valor").vValor
    
    'Preenche a tela
    lErro = Traz_PrevVenda_Tela(objPrevVenda)
    If lErro <> SUCESSO Then gError 34583

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 34583

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objPrevVenda As New ClassPrevVenda
Dim vbMsgRet As VbMsgBoxResult
Dim iPreenchida As Integer
Dim sPrevVendaFormatada As String

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o código foi informado
    If Len(Trim(Codigo.Text)) = 0 Then gError 34576
        
    '*** Inicio - Márcio 09/2000 - Incluido PrevVendas por Clientes do Vendedor
    
    'Verifica se o Cliente foi preenchido
    If Len(Trim(Cliente.Text)) = 0 Then gError 500211
    
    'Verifica se o vendedor foi preenchido
    If Len(Trim(Vendedor.Text)) = 0 Then gError 500212
    
    'Verifica se o produto foi preenchido
    If Len(Trim(Produto.ClipText)) = 0 Then gError 500213
    
    'Move dados da tela para a memória
    lErro = Move_Tela_Memoria(objPrevVenda)
    If lErro <> SUCESSO Then gError 500214
    
    '*** Fim - Márcio 09/2000 - Incluido PrevVendas por Clientes do Vendedor
    
    'Verifica se a Previsão existe
    lErro = CF("PrevVenda_Le", objPrevVenda)
    If lErro <> SUCESSO And lErro <> 34526 Then gError 34578

    'Previsão não está cadastrada
    If lErro = 34526 Then gError 34579

    'Pede Confirmação para exclusão ao usuário
    vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_PREVVENDA", Codigo.Text)

    If vbMsgRet = vbYes Then

        'Exclui a previsão de venda
        lErro = CF("PrevVenda_Exclui", objPrevVenda)
        If lErro <> SUCESSO Then gError 34580

        'Limpa a tela
        Codigo.Text = ""
        Call Limpa_Tela_PrevVenda

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 34576
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PREVVENDA_NAO_PREENCHIDO", gErr)

        Case 34578, 34580, 500214

        Case 34579
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PREVVENDA_NAO_CADASTRADA", gErr, Codigo.Text)
        
        Case 500211
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 500212
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_FORNECIDO", gErr)
        
        Case 500213
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objPrevVenda As New ClassPrevVenda
Dim sPrevVendaFormatada As String
Dim iIndice As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o código foi informado
    If Len(Trim(Codigo.Text)) = 0 Then gError 34582

    'Verifica se a data de previsão foi informada
    If Len(Trim(DataPrevisao.ClipText)) = 0 Then gError 34584

    'Verifica se o produto foi informado
    If Len(Trim(Produto.ClipText)) = 0 Then gError 34585

    'Verifica se quantidade foi informada
    If StrParaDbl(Quantidade.Text) = 0 Then gError 34586
    
    'Verifica se a data inicial foi informada
    If Len(Trim(DataInicial.ClipText)) = 0 Then gError 52914
    
    'Verifica se a data final foi informada
    If Len(Trim(DataFinal.ClipText)) = 0 Then gError 52915
    
    'data inicial não pode ser maior que a data final
    If Trim(DataInicial.ClipText) <> "" And Trim(DataFinal.ClipText) <> "" Then
    
         If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then gError 52916
    
    End If
        
    '*** Inicio - Márcio 09/2000 - Incluido PrevVendas por Clientes do Vendedor
            
    'Verifica se o Cliente foi preenchido
    If Len(Trim(Cliente.Text)) = 0 Then gError 500186
    
    'Verifica se o Vendedor foi preenchido
    If Len(Trim(Vendedor.Text)) = 0 Then gError 500187
    
    'Verifica se o Valor foi preenchido
    If StrParaDbl(Valor.Text) = 0 Then gError 500188
    
    'Verifica se o Preço Unitário foi preenchido
    If StrParaDbl(PrecoUnitario.Text) = 0 Then gError 500189
        
    'Verifica se a quantidade * preço unitário = Valor
    If Abs(StrParaDbl(PrecoUnitario.Text) * StrParaDbl(Quantidade.Text) - StrParaDbl(Valor.Text)) > 0.00001 Then gError 500199
    
    '*** Fim - Márcio 09/2000 - Incluido PrevVendas por Clientes do Vendedor
        
    'Preenche objPrevVenda
    lErro = Move_Tela_Memoria(objPrevVenda)
    If lErro <> SUCESSO Then gError 34601

    'Grava Previsao de Venda no Banco de Dados
    lErro = CF("PrevVenda_Grava", objPrevVenda)
    If lErro <> SUCESSO Then gError 34603

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 34582
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PREVVENDA_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 34601, 34603

        Case 34584
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAPREVISAO_NAO_PREENCHIDA", gErr)
            DataPrevisao.SetFocus

        Case 34585
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRODUTO_NAO_PREENCHIDO", gErr)
            Produto.SetFocus

        Case 34586
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_PREVVENDA_NAO_PREENCHIDA", gErr)
            Quantidade.SetFocus

        Case 52914
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_NAOPREENCHIDA", gErr)
            DataInicial.SetFocus
                    
        Case 52915
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAFINAL_NAOPREENCHIDA", gErr)
            DataFinal.SetFocus
            
        Case 52916
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            
        Case 500186
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
            Cliente.SetFocus
                
        Case 500187
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_FORNECIDO", gErr)
            Vendedor.SetFocus
        
        Case 500188
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_PREENCHIDO1", gErr)
            Valor.SetFocus
        
        Case 500189
            Call Rotina_Erro(vbOKOnly, "ERRO_PRECOUNITARIO_NAO_PREENCHIDO", gErr)
            PrecoUnitario.SetFocus
            
        Case 500199
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DIFERENTE_QUANTPRECOUNITARIO", gErr, StrParaDbl(Valor.Text), StrParaDbl(PrecoUnitario.Text) * StrParaDbl(Quantidade.Text))
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Public Function Limpa_Tela_PrevVenda() As Long

Dim lErro As Long
    
    'Limpa os campos da tela que não foram limpos pela função acima
    Descricao.Caption = ""
    UnidMed.Caption = ""
    Regiao.ListIndex = -1
    Regiao.Text = ""
    
    DataPrevisao.Text = Format(gdtDataAtual, "dd/mm/yy")
        
    '*** Inicio - Márcio 09/2000 - Incluido PrevVendas por Clientes do Vendedor
    If CheckFixar(0).Value = vbUnchecked Then
        Vendedor.Text = ""
    End If
    
    If CheckFixar(1).Value = vbUnchecked Then
        Cliente.Text = ""
    End If
    
    Produto.PromptInclude = False
    Produto.Text = ""
    Produto.PromptInclude = True
    
    PrecoUnitario.Text = ""
    Valor.Text = ""
    Quantidade.Text = ""
    '*** Fim - Márcio 09/2000 - Incluido PrevVendas por Clientes do Vendedor
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0
    
End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de gravação no BD
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 34581

    'Limpa a tela
    Call Limpa_Tela_PrevVenda

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 34581

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 34600
        
    'Márcio 09/2000 - Incluido PrevVendas por Clientes do Vendedor
    
    'Limpa Campos não fixos
    Codigo.Text = ""

    DataInicial.PromptInclude = False
    DataInicial.Text = ""
    DataInicial.PromptInclude = True

    DataFinal.PromptInclude = False
    DataFinal.Text = ""
    DataFinal.PromptInclude = True
    
    'Limpa restante dos campos
    Call Limpa_Tela_PrevVenda

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 34600

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    End Select

    Exit Sub

End Sub

Public Function Trata_Parametros(Optional objPrevVenda As ClassPrevVenda) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    'Se há uma previsão selecionada exibir seus dados
    If Not (objPrevVenda Is Nothing) Then

        'Verifica se a previsão existe
        lErro = CF("PrevVenda_Le", objPrevVenda)
        If lErro <> SUCESSO And lErro <> 34526 Then gError 34544

        If lErro = SUCESSO Then

            'A previsão está cadastrada
            lErro = Traz_PrevVenda_Tela(objPrevVenda)
            If lErro <> SUCESSO Then gError 52917
            
        Else

            'Previsão não está cadastrada
            Codigo.Text = objPrevVenda.sCodigo

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 34544, 52917 'Tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFinal, iAlterado)

End Sub

Private Sub DataInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicial, iAlterado)

End Sub

Private Sub DataPrevisao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataPrevisao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataPrevisao, iAlterado)

End Sub

Private Sub DataPrevisao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataPrevisao_Validate

    If Len(DataPrevisao.ClipText) = 0 Then Exit Sub

    lErro = Data_Critica(DataPrevisao.Text)
    If lErro <> SUCESSO Then gError 34591

    Exit Sub

Erro_DataPrevisao_Validate:

    Cancel = True

    Select Case gErr

        Case 34591

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

'Márcio 09/2000 - Incluido PrevVendas por Clientes do Vendedor
Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodFilial As Integer
Dim objcliente As New ClassCliente

On Error GoTo Erro_Cliente_Validate

    'Se o Cliente foi preenchido
    If Len(Trim(Cliente.Text)) > 0 Then
        
        'Busca o Cliente no BD
        lErro = TP_Cliente_Le(Cliente, objcliente, iCodFilial)
        If lErro <> SUCESSO Then gError 500195
                    
    End If
            
    Exit Sub
        
Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr
    
        Case 500195
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
    End Select
    
    Exit Sub

End Sub

'Márcio 09/2000 - Incluido PrevVendas por Vendedors do Vendedor
Private Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Vendedor_Validate

    'Se o Vendedor foi preenchido
    If Len(Trim(Vendedor.Text)) > 0 Then
        
        'Busca o Vendedor no BD
        lErro = TP_Vendedor_Le(Vendedor, objVendedor)
        If lErro <> SUCESSO Then gError 500196
                    
    End If
            
    Exit Sub
        
Erro_Vendedor_Validate:

    Cancel = True

    Select Case gErr
    
        Case 500196
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
    End Select
    
    Exit Sub

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

Dim lErro As Long

    Set objEventoCodigo = Nothing
    Set objEventoProduto = Nothing
    Set objEventoVendedor = Nothing
    Set objEventoCliente = Nothing
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Private Sub LabelCodigo_Click()

Dim lErro As Long
Dim objPrevVenda As New ClassPrevVenda
Dim colSelecao As Collection

On Error GoTo Erro_LabelCodigo_Click

    lErro = Move_Tela_Memoria(objPrevVenda)
    If lErro <> SUCESSO Then gError 34607

    'Chama a Tela que Lista as PrevVendas
    Call Chama_Tela("PrevVendaLista", colSelecao, objPrevVenda, objEventoCodigo)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case 34607

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPrevVenda As ClassPrevVenda

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objPrevVenda = obj1

    lErro = Traz_PrevVenda_Tela(objPrevVenda)
    If lErro <> SUCESSO Then gError 34608

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 34608

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1
    
    'Preenche o Produto
    lErro = Traz_Produto_Tela(objProduto)
    If lErro <> SUCESSO And lErro <> 40050 Then gError 40047
    
    If lErro = 40050 Then gError 58251
    
    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 40047
            
        Case 58251
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

'*** Inicio - Márcio 09/2000 - Incluido PrevVendas por Clientes do Vendedor

Private Sub LabelVendedor_Click()

Dim lErro As Long
Dim objVendedor As ClassVendedor
Dim colSelecao As Collection

    'Chama a Tela que Lista de Vendedores
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

    Exit Sub

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objVendedor As ClassVendedor
Dim bCancel As Boolean

On Error GoTo Erro_objEventoVendedor_evSelecao

    Set objVendedor = obj1
    
    Vendedor.Text = objVendedor.iCodigo
    Call Vendedor_Validate(bCancel)
    
    Me.Show

    Exit Sub

Erro_objEventoVendedor_evSelecao:

    Select Case gErr
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub LabelCliente_Click()

Dim lErro As Long
Dim objcliente As ClassCliente
Dim colSelecao As New Collection
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_LabelCliente_Click

    'Se o vendedor foi preenchido
    If Len(Trim(Vendedor.Text)) > 0 Then
        
        'Lê o Vendedor
        objVendedor.sNomeReduzido = Vendedor.Text
        lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
        If lErro <> SUCESSO And lErro <> 25008 Then gError 500197
        
        'Se não econtrou o vendedor, erro
        If lErro = 25008 Then gError 500198
        
        colSelecao.Add objVendedor.iCodigo
    
        'Chama a Tela que Lista de Clientes
        Call Chama_Tela("ClientesVendedorLista", colSelecao, objcliente, objEventoCliente)
        
    'Se o vendedor não foi preenchido
    Else
    
        'Chama a Tela que Lista de Clientes
        Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoCliente)
    
    End If
    

    Exit Sub

Erro_LabelCliente_Click:

    Select Case gErr
    
        Case 500197
        
        Case 500198
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO1", gErr, objVendedor.sNomeReduzido)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objcliente As ClassCliente
Dim bCancel As Boolean

On Error GoTo Erro_objEventoCliente_evSelecao

    Set objcliente = obj1
    
    Cliente.Text = objcliente.lCodigo
    Call Cliente_Validate(bCancel)
    
    Me.Show

    Exit Sub

Erro_objEventoCliente_evSelecao:

    Select Case gErr
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub Produto_GotFocus()

    sProdutoAnterior = Produto.Text

End Sub

'*** Fim - Márcio 09/2000 - Incluido PrevVendas por Clientes do Vendedor

Private Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO

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

        'Se produto não estiver preenchido --> limpa descrição e unidade de medida
        If Len(Trim(Produto.ClipText)) = 0 Then

            Descricao.Caption = ""
            UnidMed.Caption = ""
            
        Else 'Caso esteja preenchido

            lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 34589

            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
                objProduto.sCodigo = sProdutoFormatado
                
                lErro = Traz_Produto_Tela(objProduto)
                If lErro <> SUCESSO And lErro <> 40050 Then gError 40048
                
                'Caso não esteja Cadastrado o Prodduto no BD
                If lErro = 40050 Then gError 58250
                
            End If

        End If

        sProdutoAnterior = Produto.Text

    End If

    Exit Sub

Erro_Produto_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 34589, 40048
        
        Case 58250
            Descricao.Caption = ""
            UnidMed.Caption = ""

            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)

            If vbMsg = vbYes Then
                
                Call Chama_Tela("Produto", objProduto)

            End If
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

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
        If lErro <> SUCESSO Then gError 34612

        objProduto.sCodigo = sProdutoFormatado

    End If
    
    'Cahama a lista de Produtos que Podem ser Vendidos
    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoProduto)
    
    Exit Sub

Erro_ProdutoLabel_Click:

    Select Case gErr

        Case 34612

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Function Traz_Produto_Tela(objProduto As ClassProduto) As Long
'Preenche o campo do produto

Dim lErro As Long
Dim objProdutoFilial As New ClassProdutoFilial

On Error GoTo Erro_Traz_Produto_Tela
    
    'Lê o Produto que está sendo Passado
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 40049
    
    'Se ele não existir --- > ERRO
    If lErro = 28030 Then gError 40050
    
    'Se ele for Gerencial --- > ERRO
    If objProduto.iGerencial = GERENCIAL Then gError 58071
    
    'Se for Produto não vendavel ---> ERRO
    If objProduto.iFaturamento = PRODUTO_NAO_VENDAVEL Then gError 58072
    
    objProdutoFilial.sProduto = objProduto.sCodigo
    objProdutoFilial.iFilialEmpresa = giFilialEmpresa
    
    lErro = CF("ProdutoFilial_Le", objProdutoFilial)
    If lErro <> SUCESSO And lErro <> 28261 Then gError 58255
    
    'Se não encontrou
    If lErro = 28261 Then gError 58256
    
    'Preenche a Tela com o Produto e sua Descrição
    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, Produto, Descricao)
    If lErro <> SUCESSO Then gError 34593
    
    'Preenche unidade de medida
    UnidMed.Caption = objProduto.sSiglaUMVenda
        
    '*** Inicio Márcio 09/2000 - Incluido PrevVendas por Clientes do Vendedor
    lErro = Traz_ValorUnitario_Tela()
    If lErro <> SUCESSO Then gError 500190
        
    Traz_Produto_Tela = SUCESSO

    Exit Function

Erro_Traz_Produto_Tela:

    Traz_Produto_Tela = gErr

    Select Case gErr

        Case 34593, 40049, 58255, 500190 'Tratados na Rotina chamada
        
        Case 40050 'Se não encontrou em Produtos
                                
        Case 58071
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)
        
        Case 58072
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PODE_SER_VENDIDO2", gErr, objProduto.sCodigo)
        
        Case 58256 'Produto não cadastrado em ProdutoFilial
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_FILIAL_INEXISTENTE", gErr, objProduto.sCodigo, giFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

'*** Inicio Márcio 09/2000 - Incluido PrevVendas por Clientes do Vendedor

Function Traz_ValorUnitario_Tela() As Long

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim objItemPedido As New ClassItemPedido
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Traz_ValorUnitario_Tela

    'Se o produto e o cliente estão preenchidos
    If Len(Trim(Produto.ClipText)) > 0 And Len(Trim(Cliente.Text)) > 0 And sProdutoAnterior <> Produto.Text Then
            
        'Lê dados do cliente
        objcliente.sNomeReduzido = Cliente.Text
        lErro = CF("Cliente_Le_NomeReduzido", objcliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 500191
        
        'Se não econtrou o Cliente, erro
        If lErro = 12348 Then gError 500192
                
        'Formata o Produto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 500193
        
        'Lê dados do último ItemPV relacionado ao Produto e Cliente da tela
        objItemPedido.sProduto = sProdutoFormatado
        objItemPedido.iFilialEmpresa = giFilialEmpresa
        lErro = CF("Cliente_Le_UltimoItemPedido", objItemPedido, objcliente.lCodigo)
        If lErro <> SUCESSO And lErro <> 500019 Then gError 500194
        
        'Se encontrou o pedido
        If lErro <> 500019 Then
            PrecoUnitario.Text = Format(objItemPedido.dPrecoUnitario, "Standard")
            
            'Se a quantidade está preenchida
            If Len(Trim(Quantidade.Text)) > 0 Then
                Valor.Text = Format(StrParaDbl(PrecoUnitario.Text) * StrParaDbl(Quantidade.Text), "Standard")
            End If
        Else
            PrecoUnitario.Text = ""
        End If
        
    End If
    
    Traz_ValorUnitario_Tela = SUCESSO

    Exit Function
    
Erro_Traz_ValorUnitario_Tela:

    Select Case gErr
        
        Case 500191, 500193, 500194
        
        Case 500192
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objcliente.sNomeReduzido)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
    End Select
    
    Exit Function

End Function

Private Sub Regiao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objRegiaoVenda As New ClassRegiaoVenda
Dim iCodigo As Integer

On Error GoTo Erro_Regiao_Validate

    'Verifica se foi preenchida a ComboBox Regiao
    If Len(Trim(Regiao.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Regiao
    If Regiao.ListIndex >= 0 Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Regiao, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 52905

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objRegiaoVenda.iCodigo = iCodigo

        'Tentativa de leitura da Regiao de Venda com esse código no BD
        lErro = CF("RegiaoVenda_Le", objRegiaoVenda)
        If lErro <> SUCESSO And lErro <> 16137 Then gError 52906

        If lErro = 16137 Then gError 52907 'Não encontrou Regiao de Venda no BD

        'Encontrou Regiao Venda no BD, coloca no Text da Combo
        Regiao.Text = CStr(objRegiaoVenda.iCodigo) & SEPARADOR & objRegiaoVenda.sDescricao

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 52908

    Exit Sub

Erro_Regiao_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 52905, 52906
    
        Case 52907  'Não encontrou Regiao de Venda no BD
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_REGIAO")

            If vbMsgRes = vbYes Then
                'Chama a tela RegiaoVenda
                Call Chama_Tela("RegiaoVenda", objRegiaoVenda)
            End If

        Case 52908
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGIAO_VENDA_NAO_ENCONTRADA", gErr, Regiao.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Valor_Validate

    If Len(Trim(Valor.Text)) <> 0 Then
        
        'Critica Valor
        lErro = Valor_Positivo_Critica(Valor.Text)
        If lErro <> SUCESSO Then gError 34606
        
        Valor.Text = Format(Valor.Text, "Fixed")
        
    End If

    Exit Sub

Erro_Valor_Validate:

    Cancel = True

    Select Case gErr

        Case 34606

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Quantidade_Validate

    If Len(Trim(Quantidade.Text)) = 0 Then Exit Sub

    'Critica quantidade
    lErro = Valor_Positivo_Critica(Quantidade.Text)
    If lErro <> SUCESSO Then gError 34604
    
    Quantidade.Text = Formata_Estoque(CDbl(Quantidade.Text))
    
    'Márcio 09/2000 - Incluido PrevVendas por Clientes do Vendedor
    If Len(Trim(PrecoUnitario.Text)) > 0 Then
        Valor.Text = Format(StrParaDbl(PrecoUnitario.Text) * StrParaDbl(Quantidade.Text), "Standard")
    End If
    
    Exit Sub

Erro_Quantidade_Validate:

    Cancel = True

    Select Case gErr

        Case 34604
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

'Márcio 09/2000 - Incluido PrevVendas por Clientes do Vendedor

Private Sub PrecoUnitario_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PrecoUnitario_Validate
    
    'Se o preço unitário não foi preenchido, sai da rotina
    If Len(Trim(PrecoUnitario.Text)) = 0 Then Exit Sub

    'Critica PrecoUnitario
    lErro = Valor_Positivo_Critica(PrecoUnitario.Text)
    If lErro <> SUCESSO Then gError 500193
    
    PrecoUnitario.Text = Formata_Estoque(CDbl(PrecoUnitario.Text))
    
    'Se a quantidade foi preenchida
    If Len(Trim(Quantidade.Text)) > 0 Then
        Valor.Text = Format(StrParaDbl(PrecoUnitario.Text) * StrParaDbl(Quantidade.Text), "Standard")
    End If
    
    Exit Sub

Erro_PrecoUnitario_Validate:

    Cancel = True

    Select Case gErr

        Case 500193
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colCodigoNome As AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome

On Error GoTo Erro_Form_Load
    
    '*** Inicio - Márcio 09/2000 - Incluido PrevVendas por Clientes do Vendedor
    'Esconde controles que não serão usados na tela
    Label1.left = POSICAO_FORA_TELA
    ValorTabela.left = POSICAO_FORA_TELA
    ValorTabela.TabStop = False
    RegiaoVendaLabel.left = POSICAO_FORA_TELA
    Regiao.left = POSICAO_FORA_TELA
    Regiao.TabStop = False
    
    'Marca Checks Para Fixar Cliente e Vendedor
    CheckFixar(0).Value = vbChecked
    CheckFixar(1).Value = vbChecked
    
    '*** Fim - Márcio 09/2000 - Incluido PrevVendas por Clientes do Vendedor
    
    'Inicializa os Eventos de Browser
    Set objEventoCodigo = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoCliente = New AdmEvento
    Set objEventoVendedor = New AdmEvento
    
    Set colCodigoNome = New AdmColCodigoNome

    'Leitura dos códigos e descrições das Regiões de Venda
    lErro = CF("Cod_Nomes_Le", "RegioesVendas", "Codigo", "Descricao", STRING_REGIAO_VENDA_DESCRICAO, colCodigoNome)
    If lErro <> SUCESSO Then gError 16614

    'Preenche a ComboBox Região com código e descrição das Regiões de Venda
    For Each objCodigoNome In colCodigoNome
        Regiao.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        Regiao.ItemData(Regiao.NewIndex) = objCodigoNome.iCodigo
    Next

    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 34565

    'Carrega a arvore de Produtos
''    lErro = CF("Carga_Arvore_Produto_Venda",TvwProdutos.Nodes)
''    If lErro Then gError 34566

    'Carrega Data Atual
    DataPrevisao.PromptInclude = False
    DataPrevisao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataPrevisao.PromptInclude = True
        
    'Formata a quantidade
    Quantidade.Format = FORMATO_ESTOQUE
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case gErr

        Case 34564, 34565

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Traz_PrevVenda_Tela(objPrevVenda As ClassPrevVenda) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim bCancel As Boolean

On Error GoTo Erro_Traz_PrevVenda_Tela

    Call Limpa_Tela_PrevVenda
    
    'Preenche código de previsão
    Codigo.Text = CStr(objPrevVenda.sCodigo)

    'Verifica se produto retornado é válido
    objProduto.sCodigo = objPrevVenda.sProduto
    
    lErro = Traz_Produto_Tela(objProduto)
    If lErro <> SUCESSO And lErro <> 40050 Then gError 40051

    'Produto não cadastrado --> Erro
    If lErro = 40050 Then gError 34596
    
    'Coloca código da Regiao em .Text e chama Validate
    If objPrevVenda.iCodRegiao <> 0 Then
        Regiao.Text = CStr(objPrevVenda.iCodRegiao)
        Regiao_Validate (bCancel)
    Else
        Regiao.Text = ""
    End If
    
    Call StrParaMasked(Quantidade, Formata_Estoque(objPrevVenda.dQuantidade))
    
    'Preenche data Inicial
    Call DateParaMasked(DataInicial, objPrevVenda.dtDataInicio)
    
    'Preenche data Final
    Call DateParaMasked(DataFinal, objPrevVenda.dtDataFim)

    'Preenche data de previsão
    Call DateParaMasked(DataPrevisao, objPrevVenda.dtDataPrevisao)

    Call StrParaMasked(Valor, Format(objPrevVenda.dValor, "Fixed"))

    '*** Inicio - Márcio 09/2000 - Incluido PrevVendas por Clientes do Vendedor
    PrecoUnitario.Text = Format(objPrevVenda.dPrecoUnitario, "Standard")
        
    'Vendedor
    Vendedor.Text = objPrevVenda.iVendedor
    Call Vendedor_Validate(bCancel)
    
    'Cliente
    Cliente.Text = objPrevVenda.lCliente
    Call Cliente_Validate(bCancel)
    
    '*** Fim - Márcio 09/2000 - Incluido PrevVendas por Clientes do Vendedor

    iAlterado = 0

    Traz_PrevVenda_Tela = SUCESSO

    Exit Function

Erro_Traz_PrevVenda_Tela:

    Traz_PrevVenda_Tela = gErr

    Select Case gErr

        Case 34596
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 40051 'Tratados nas Rotinas chamadas
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

'*** Inicio - Márcio 09/2000 - Incluido PrevVendas por Clientes do Vendedor

Private Sub PrecoUnitario_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Cliente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Vendedor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

'*** Fim - Márcio 09/2000 - Incluido PrevVendas por Clientes do Vendedor

Private Sub Valor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDown_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataPrevisao, DIMINUI_DATA)
    If lErro Then gError 49862

    Exit Sub

Erro_UpDown_DownClick:

    Select Case gErr

        Case 49862

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub UpDown_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataPrevisao, AUMENTA_DATA)
    If lErro Then gError 49863

    Exit Sub

Erro_UpDown_UpClick:

    Select Case gErr

        Case 49863

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub Regiao_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Regiao_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub DataInicial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataInicial_Validate

    If Len(DataInicial.ClipText) = 0 Then Exit Sub
    
    'Critica Data
    lErro = Data_Critica(DataInicial.Text)
    If lErro <> SUCESSO Then gError 52909

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True


    Select Case gErr

        Case 52909

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown2_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro Then gError 52910

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case gErr

        Case 52910 'Tratado na rotina chamada

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown2_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro Then gError 52911

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case gErr

        Case 52911 'Tratado na rotina chamada

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub DataFinal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    If Len(DataFinal.ClipText) = 0 Then Exit Sub

    'Critica Data
    lErro = Data_Critica(DataFinal.Text)
    If lErro <> SUCESSO Then gError 52912

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 52912

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub UpDown3_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown3_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro Then gError 52913

    Exit Sub

Erro_UpDown3_DownClick:

    Select Case gErr

        Case 52913 'Tratado na rotina chamada

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub UpDown3_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown3_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro Then gError 52914

    Exit Sub

Erro_UpDown3_UpClick:

    Select Case gErr

        Case 52914 'Tratado na rotina chamada

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_PREV_VENDAS
    Set Form_Load_Ocx = Me
    Caption = "Previsão de Vendas"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "PrevVenda"
    
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
        
        If Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call ProdutoLabel_Click
        End If
    
    End If

End Sub

Private Sub ValorTabela_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProdutoFilial As New ClassProdutoFilial
Dim dtDataFinal As Date
Dim dPrecoTabela As Double
Dim dValorTabela As Double
Dim dQuantidade As Double

On Error GoTo Erro_ValorTabela_Click

    'Se produto não estiver preenchido --> erro
    If Len(Trim(Produto.ClipText)) = 0 Then gError 25792

    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 25793

    'Se produto não estiver preenchido --> erro
    If iProdutoPreenchido = PRODUTO_VAZIO Then gError 25794

    'Verifica se a data inicial foi informada
    If Len(Trim(DataInicial.ClipText)) = 0 Then gError 25795
    
    'Verifica se a data final foi informada
    If Len(Trim(DataFinal.ClipText)) = 0 Then gError 25796
    
    'Verifica se DataInical é anterior ou igual à DataFinal
    If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then gError 25797
        
    dtDataFinal = CDate(DataFinal.Text)
        
    'Verifica se Quantidade foi preenchida
    If Len(Trim(Quantidade.ClipText)) = 0 Then gError 25798

    dQuantidade = CDbl(Quantidade.Text)
    
    objProdutoFilial.sProduto = sProdutoFormatado
    objProdutoFilial.iFilialEmpresa = giFilialEmpresa
    
    'Pesquisa na tabela default último preço unitário vigente dentro da faixa de data
    lErro = CF("TabelaPrecoPadrao_Le", objProdutoFilial, dtDataFinal, dPrecoTabela)
    If lErro <> SUCESSO And lErro <> 25788 And lErro <> 25791 Then gError 25799
    
    'Não encontrou Tabela Padrão
    If lErro = 25788 Then gError 25800
    'Não encontrou ítem dentro da Tabela Padrão
    If lErro = 25791 Then gError 25801
    
    'Coloca valor de Tabela padrão na tela
    dValorTabela = dPrecoTabela * dQuantidade
    Valor.Text = Format(dValorTabela, "Fixed")
            
    Exit Sub

Erro_ValorTabela_Click:
    
    Select Case gErr
        
        Case 25792, 25794
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case 25793, 25799 'tratados nas rotinas chamadas
        
        Case 25795
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_NAOPREENCHIDA", gErr)
                    
        Case 25796
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAFINAL_NAOPREENCHIDA", gErr)
        
        Case 25797
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            
        Case 25798
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_PREVVENDA_NAO_PREENCHIDA", gErr)
            
        Case 25800
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_TABELA_PADRAO", gErr, objProdutoFilial.sProduto)
        
        Case 25801
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECOITEM_INEXISTENTE3", gErr, objProdutoFilial.sProduto, dtDataFinal)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub RegiaoVendaLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(RegiaoVendaLabel, Source, X, Y)
End Sub

Private Sub RegiaoVendaLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(RegiaoVendaLabel, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigo, Source, X, Y)
End Sub

Private Sub LabelCodigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigo, Button, Shift, X, Y)
End Sub

Private Sub UnidMed_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(UnidMed, Source, X, Y)
End Sub

Private Sub UnidMed_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(UnidMed, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Descricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Descricao, Source, X, Y)
End Sub

Private Sub Descricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Descricao, Button, Shift, X, Y)
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

Private Sub LabelCliente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCliente, Source, X, Y)
End Sub

Private Sub LabelCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCliente, Button, Shift, X, Y)
End Sub

Private Sub LabelVendedor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCliente, Source, X, Y)
End Sub

Private Sub LabelVendedor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCliente, Button, Shift, X, Y)
End Sub
'
'Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label8, Source, X, Y)
'End Sub
'
'Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
'End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub


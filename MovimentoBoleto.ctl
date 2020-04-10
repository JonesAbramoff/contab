VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl MovimentoBoleto 
   ClientHeight    =   6495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7110
   KeyPreview      =   -1  'True
   ScaleHeight     =   6495
   ScaleWidth      =   7110
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4890
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   195
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "MovimentoBoleto.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "F8 - Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "MovimentoBoleto.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "F7 - Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   615
         Picture         =   "MovimentoBoleto.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "F6 - Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "MovimentoBoleto.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "F5- Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame FrameIdentificacao 
      Caption         =   "Identificação"
      Height          =   870
      Left            =   75
      TabIndex        =   11
      Top             =   105
      Width           =   4575
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   1785
         Picture         =   "MovimentoBoleto.ctx":0994
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Numeração Automática"
         Top             =   345
         Width           =   300
      End
      Begin VB.CommandButton BotaoTrazer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   2505
         Picture         =   "MovimentoBoleto.ctx":0A7E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "F4 - Exibe na tela o movimento com o código informado."
         Top             =   225
         Width           =   1560
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   930
         TabIndex        =   1
         Top             =   330
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   7
         Mask            =   "#######"
         PromptChar      =   " "
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
         Left            =   225
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   12
         Top             =   390
         Width           =   660
      End
   End
   Begin VB.Frame FrameBoletosNaoDetalhados 
      Caption         =   "Boletos Não Detalhados - Cartões de Crédito"
      Height          =   840
      Left            =   75
      TabIndex        =   5
      Top             =   5580
      Width           =   6975
      Begin VB.TextBox ValorSangriaCCredito 
         Height          =   315
         Left            =   5055
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
      Begin VB.Frame FrameCartaoDebito 
         Caption         =   "Cartões de Débito"
         Height          =   1245
         Left            =   -3600
         TabIndex        =   6
         Top             =   240
         Width           =   3135
         Begin VB.TextBox ValorSangriaCDebito 
            Height          =   315
            Left            =   1470
            TabIndex        =   10
            Top             =   765
            Width           =   1215
         End
         Begin VB.Label LabelEmCaixaCDebito 
            AutoSize        =   -1  'True
            Caption         =   "Em caixa:"
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
            Left            =   480
            TabIndex        =   9
            Top             =   420
            Width           =   840
         End
         Begin VB.Label LabelEmCaixaCDebitoValor 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1455
            TabIndex        =   8
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label LabelSangriaCDebito 
            AutoSize        =   -1  'True
            Caption         =   "Sangria:"
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
            Left            =   600
            TabIndex        =   7
            Top             =   840
            Width           =   720
         End
      End
      Begin VB.Label LabelEmCaixaCCredito 
         AutoSize        =   -1  'True
         Caption         =   "Em caixa:"
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
         Left            =   1170
         TabIndex        =   25
         Top             =   420
         Width           =   840
      End
      Begin VB.Label LabelEmCaixaCCreditoValor 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2115
         TabIndex        =   24
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label LabelSangriaCCredito 
         AutoSize        =   -1  'True
         Caption         =   "Sangria:"
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
         Left            =   4200
         TabIndex        =   23
         Top             =   420
         Width           =   720
      End
   End
   Begin VB.Frame FrameBoletosDetalhados 
      Caption         =   "Boletos Detalhados"
      Height          =   4365
      Left            =   75
      TabIndex        =   0
      Top             =   1095
      Width           =   6975
      Begin VB.CommandButton BotaoValoresEmCaixa 
         Height          =   585
         Left            =   165
         Picture         =   "MovimentoBoleto.ctx":3748
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "F9 - Traz para a tela os boletos em caixa que podem ser ""sangrados""."
         Top             =   3705
         Width           =   1620
      End
      Begin VB.ComboBox Administradora 
         Height          =   315
         Left            =   510
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   255
         Width           =   1605
      End
      Begin VB.ComboBox Parcelamento 
         Height          =   315
         Left            =   2115
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   240
         Width           =   1590
      End
      Begin MSMask.MaskEdBox ValorSangria 
         Height          =   300
         Left            =   5100
         TabIndex        =   20
         Top             =   285
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
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
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ValorTotal 
         Height          =   300
         Left            =   3765
         TabIndex        =   21
         Top             =   255
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridBoletos 
         Height          =   3135
         Left            =   150
         TabIndex        =   4
         Top             =   255
         Width           =   6630
         _ExtentX        =   11695
         _ExtentY        =   5530
         _Version        =   393216
         Rows            =   5
         Cols            =   5
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         FocusRect       =   2
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
      Begin VB.Label LabelSangriaValor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5130
         TabIndex        =   29
         Top             =   3660
         Width           =   1290
      End
      Begin VB.Label LabelEmCaixaValor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3825
         TabIndex        =   28
         Top             =   3660
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Totais:"
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
         Left            =   3135
         TabIndex        =   27
         Top             =   3720
         Width           =   600
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   915
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RT1 
      Height          =   525
      Left            =   195
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   926
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"MovimentoBoleto.ctx":64A2
   End
End
Attribute VB_Name = "MovimentoBoleto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'mario
 
'###### Observações de Tela #####
'se o movimento de cartao for de débito não especificado o Terminal de Cartão usado é de Pós,
'já se o movimento de  Cartão for de crédito não especificado o Terminal Usado é o Manual
'e o código da Admnistradora recebe 0 para ambos.
'' #################################################################



'Variáveis Globais
Dim objGridBoletos As AdmGrid
Dim iAlterado As Integer
Dim iGrid_Administradora_Col As Integer
Dim iGrid_Parcelamento_Col As Integer
Dim iGrid_ValorTotal_Col As Integer
Dim iGrid_ValorSangria_Col As Integer


'Property Variables:
Dim m_Caption As String
Event Unload()

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Sangria Boleto"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "MovimentoBoleto"

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


Private Sub Codigo_GotFocus()

    Call MaskEdBox_TrataGotFocus(Codigo)

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
    'Parent.UnloadDoFilho

   RaiseEvent Unload

End Sub

Public Property Get Caption() As String

    Caption = m_Caption

End Property

Public Property Let Caption(ByVal New_Caption As String)

    Parent.Caption = New_Caption
    m_Caption = New_Caption

End Property

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

End Sub

Public Sub Form_Load()
'Função inicialização da Tela
Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Instanciar o objGridBoletos para apontar para uma posição de memória
    Set objGridBoletos = New AdmGrid

    'Inicialização de Grid Boletos
    lErro = Inicializa_GridBoletos(objGridBoletos)
    If lErro <> SUCESSO Then gError 107721

    'Função que Carrega as Admnistradoras de Meio de Pagto
    lErro = Carrega_Combo_Administradora()
    If lErro <> SUCESSO Then gError 107725

    Call BotaoValoresEmCaixa_Click

    'Indica que o carregamento da tela aconteceu com sucesso
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        'Erros tratados na rotina chamada
        Case 107721 To 107723, 107817, 107725


        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162913)

    End Select

    Exit Sub

End Sub

Function Inicializa_GridBoletos(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Inicializa_GridBoletos

   'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Administradora")
    objGridInt.colColuna.Add ("Parcelamento")
    objGridInt.colColuna.Add ("Valor em Caixa")
    objGridInt.colColuna.Add ("Valor Sangria")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Administradora.Name)
    objGridInt.colCampo.Add (Parcelamento.Name)
    objGridInt.colCampo.Add (ValorTotal.Name)
    objGridInt.colCampo.Add (ValorSangria.Name)

    'Colunas do Grid
    iGrid_Administradora_Col = 1
    iGrid_Parcelamento_Col = 2
    iGrid_ValorTotal_Col = 3
    iGrid_ValorSangria_Col = 4

    'Grid do GridInterno
    objGridInt.objGrid = GridBoletos

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_LINHAS_GRID

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 9

    'Largura da primeira coluna
    GridBoletos.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
     
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridBoletos = SUCESSO

    Exit Function

Erro_Inicializa_GridBoletos:

    Inicializa_GridBoletos = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162914)

    End Select

    Exit Function

End Function

Function Recalcula_Totais() As Long
'Calcula o Total de Boletos não Especificado e o Totas de Boletos Especificados

Dim lErro As Long
Dim iIndice As Integer
Dim dValorTotalBoletos As Double
Dim dValorSangria As Double
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_Recalcula_Totais

    'Para todos os Cheque do Grid
    For iIndice = 1 To objGridBoletos.iLinhasExistentes

        'Acumula o Valor Total dos Boletos
        dValorTotalBoletos = dValorTotalBoletos + StrParaDbl(GridBoletos.TextMatrix(iIndice, iGrid_ValorTotal_Col))

        'Acumula o Valor Total da Sangria
        dValorSangria = dValorSangria + StrParaDbl(GridBoletos.TextMatrix(iIndice, iGrid_ValorSangria_Col))

    Next

    'Exibe o valor total tando de Sangria como o valor total de boletos que existem no caixa
    LabelEmCaixaValor.Caption = Format(dValorTotalBoletos, "Standard")
    LabelSangriaValor.Caption = Format(dValorSangria, "Standard")
    
    Recalcula_Totais = SUCESSO

    Exit Function

Erro_Recalcula_Totais:

    Recalcula_Totais = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162915)

    End Select

    Exit Function

End Function

Function Carrega_Combo_Administradora() As Long
'Função que Carrega a Combo de Admnistradoras

Dim lErro As Long
Dim objAdmMeioPagto As New ClassAdmMeioPagto
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto
Dim iIndice As Integer

On Error GoTo Erro_Carrega_Combo_Administradora

    'Adiciona todas as Admnistradoras lidas na Coleção Global na Combo
    For Each objAdmMeioPagto In gcolAdmMeioPagto
        
        If objAdmMeioPagto.iTipoMeioPagto = TIPOMEIOPAGTOLOJA_CARTAO_CREDITO Then
        
            For Each objAdmMeioPagtoCondPagto In gcolCartao
        
                If objAdmMeioPagtoCondPagto.iAdmMeioPagto = objAdmMeioPagto.iCodigo And objAdmMeioPagtoCondPagto.iTipoCartao = TIPO_MANUAL And objAdmMeioPagtoCondPagto.dSaldo > 0 Then
        
                        'Senão for nenhum dois dos acima carrega na combo de Administradora
                        Administradora.AddItem objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto
                        Administradora.ItemData(Administradora.NewIndex) = objAdmMeioPagtoCondPagto.iAdmMeioPagto
                        Exit For

                End If
                
            Next
            
        End If
        
    Next
    
    Carrega_Combo_Administradora = SUCESSO

    Exit Function

Erro_Carrega_Combo_Administradora:

    Carrega_Combo_Administradora = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162916)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable

    If iCaminho = ROTINA_GRID_TRATA_TECLA_CAMPO2 Or iCaminho = ROTINA_GRID_ENTRADA_CELULA Or iCaminho = ROTINA_GRID_CLICK Then

        'Pesquisa controle da coluna em questão
        Select Case objControl.Name
    
            Case Parcelamento.Name
    
                If Len(Trim(GridBoletos.TextMatrix(iLinha, iGrid_Administradora_Col))) = 0 Then
    
                    objControl.Enabled = False
    
                Else
    
                    objControl.Enabled = True
    
                    'Carrega a Combo de Parcelamentos
                    lErro = Carrega_Combo_Parcelamento(iLinha)
                    If lErro <> SUCESSO Then gError 107726
    
                End If
            
            'Se o campo for valor sangria
            Case ValorSangria.Name
    
                'Verifca se o Campo valorTotal está Preenchido
                 If Len(Trim(GridBoletos.TextMatrix(iLinha, iGrid_ValorTotal_Col))) <> 0 Then
    
                    objControl.Enabled = True
    
                Else
    
                    'Desabilita o Campo Valor Total de Sangria
                    objControl.Enabled = False
    
                End If
    
        End Select

    End If

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 107726
            'Erro Tratado Dentro da Função

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162917)

    End Select

    Exit Sub

End Sub

Function Carrega_Combo_Parcelamento(iLinha As Integer) As Long
'Função que Carrega a Combo de Parcelamentos, Verificando se o parcelamento selecionado é igual a um dos parcelamento que é referente a admnistradora selecinada

Dim lErro As Long
Dim sNomeParcelamento As String
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto
Dim sNomeAdmnistradora As String
Dim objAdmMeioPagto As New ClassAdmMeioPagto
Dim iIndice As Integer
On Error GoTo Erro_Carrega_Combo_Parcelamento

    'Atribui a Variável o nome do parcelamento
    sNomeParcelamento = GridBoletos.TextMatrix(iLinha, iGrid_Parcelamento_Col)

    'Atribui o Nome da Admnistradora Selecionada
    sNomeAdmnistradora = GridBoletos.TextMatrix(iLinha, iGrid_Administradora_Col)
    
    'Limpa a Combo
     Parcelamento.Clear
    
    'Incluir todos os Parcelamentos cadastratrados na Coleção Glogal de Parcelamentos Referenciando a Admnistradora Referenciada
    For Each objAdmMeioPagtoCondPagto In gcolCartao

         If sNomeAdmnistradora = objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto And objAdmMeioPagtoCondPagto.iTipoCartao = TIPO_MANUAL And objAdmMeioPagtoCondPagto.dSaldo > 0 Then

            Parcelamento.AddItem objAdmMeioPagtoCondPagto.sNomeParcelamento
            Parcelamento.ItemData(Parcelamento.NewIndex) = objAdmMeioPagtoCondPagto.iParcelamento

        End If
    Next


    For iIndice = 0 To Parcelamento.ListCount - 1

        If sNomeParcelamento = Parcelamento.List(iIndice) Then
            Parcelamento.ListIndex = iIndice
            Exit For
        End If

    Next
    
    Carrega_Combo_Parcelamento = SUCESSO

    Exit Function

Erro_Carrega_Combo_Parcelamento:

    Carrega_Combo_Parcelamento = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162918)

    End Select

    Exit Function

End Function

Private Sub BotaoValoresEmCaixa_Click()
'Função que Verifica os Valores em Caixa

Dim lErro As Long
Dim colMovimentosCaixa As New Collection
Dim lNumero As Long

On Error GoTo Erro_BotaoValoresEmCaixa_Click

    'Limpa a Tela
    Call Limpa_Tela_MovimentoBoleto

    'traz todos os movimentos de caixa relacionados com boletos
    lErro = Traz_ValoresEmCaixa_Tela()
    If lErro <> SUCESSO Then gError 107801

    'Função que Recalcula os Totáis
    lErro = Recalcula_Totais()
    If lErro <> SUCESSO Then gError 107803

    Exit Sub

Erro_BotaoValoresEmCaixa_Click:

    Select Case gErr

        Case 107801 To 107803

        Case 107855
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_INEXISTENTE, gErr)
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162919)

    End Select

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'Função que Incrementa o Código Atravez da Tecla F2
Dim lErro As Long

On Error GoTo Erro_UserControl_KeyDown

    Select Case KeyCode

        Case KEYCODE_PROXIMO_NUMERO
            
            'Função que Incrementa o Código( Ultimo Movto + 1)
            If Not TrocaFoco(Me, BotaoProxNum) Then Exit Sub
            Call BotaoProxNum_Click

        Case KEYCODE_BROWSER

            Call LabelCodigo_Click

        Case vbKeyF4
            If Not TrocaFoco(Me, BotaoTrazer) Then Exit Sub
            Call BotaoTrazer_Click

        Case vbKeyF5
            If Not TrocaFoco(Me, BotaoGravar) Then Exit Sub
            Call BotaoGravar_Click
            
        Case vbKeyF6
            If Not TrocaFoco(Me, BotaoExcluir) Then Exit Sub
            Call BotaoExcluir_Click
            
        Case vbKeyF7
            If Not TrocaFoco(Me, BotaoLimpar) Then Exit Sub
            Call BotaoLimpar_Click
            
        Case vbKeyF8
            If Not TrocaFoco(Me, BotaoFechar) Then Exit Sub
            Call BotaoFechar_Click

        Case vbKeyF9
            If Not TrocaFoco(Me, BotaoValoresEmCaixa) Then Exit Sub
            Call BotaoValoresEmCaixa_Click

    End Select

    Exit Sub

Erro_UserControl_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162920)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()
'Botão que Gera um Próximo Numero para Movto

Dim lErro As Long
Dim lNumero As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Função que Gera o Próximo Código para a Tela de Sangria de Boletos
    lErro = CF_ECF("Caixa_Obtem_NumAutomatico", lNumero)
    If lErro <> SUCESSO Then gError 107730

    'Exibir o Numero na Tela
    Codigo.Text = lNumero

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 107730

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162921)

    End Select

    Exit Sub

End Sub

Private Sub BotaoTrazer_Click()
'Função que chama a função que preenche o grid

Dim lErro As Long
Dim lCodigo As String

On Error GoTo Erro_botaoTrazer_click

    'Verifica se o código não está preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 107753

    lCodigo = StrParaLong(Codigo.Text)
    
    Call Limpa_Tela_MovimentoBoleto
    
    Codigo.Text = lCodigo
    
    'Chama a função que preenche o grid
    lErro = Traz_MovimentoBoleto_Tela(StrParaLong(Codigo.Text))
    If lErro <> SUCESSO Then gError 107754

    'Anula a Alteração
    iAlterado = 0

    Exit Sub

Erro_botaoTrazer_click:

    Select Case gErr

        Case 107753
            Call Rotina_ErroECF(vbOKOnly, ERRO_CODIGO_NAO_PREENCHIDO1, gErr)

        Case 107754
            'Erro tradado Dentro da Função que Foi Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162922)

    End Select

    Exit Sub

End Sub

Function Traz_MovimentoBoleto_Tela(lNumero As Long) As Long
'Função que Trar o Movimento Boleto para a tela

Dim lErro As Long
Dim colMovimentosCaixa As New Collection
Dim objMovimentoCaixa As ClassMovimentoCaixa
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto
Dim iLinha As Integer

On Error GoTo Erro_Traz_MovimentoBoleto_Tela

    Call Limpa_Tela_MovimentoBoleto
    
    Codigo.Text = lNumero
    
    'Função que Lê os Movimentos de Caixa
    lErro = CF_ECF("Caixa_MovimentoCaixa_Le_NumMovto", colMovimentosCaixa, lNumero)
    If lErro <> SUCESSO And lErro <> 107850 Then gError 107755

    If lErro = 107850 Then gError 105700

    For Each objMovimentoCaixa In colMovimentosCaixa
        
        If objMovimentoCaixa.iTipo <> MOVIMENTOCAIXA_SANGRIA_BOLETO_CC Then gError 105738

        For Each objAdmMeioPagtoCondPagto In gcolCartao
        
            'Verifica se o Código da Administradora é igual ao Código da Adimistradora do Movto
            If objAdmMeioPagtoCondPagto.iAdmMeioPagto = objMovimentoCaixa.iAdmMeioPagto Then

                'Verifica se o codigo da admnistradora é igual a zero
                If objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto = STRING_NAO_DETALHADO_CCREDITO Then
                        
                    'Exibe o Saldo + Valor do Movto
                    LabelEmCaixaCCreditoValor.Caption = CStr(Format(objAdmMeioPagtoCondPagto.dSaldo + objMovimentoCaixa.dValor, "standard"))
                    'Coloca o Valor do Movto
                    ValorSangriaCCredito.Text = Format(objMovimentoCaixa.dValor, "Standard")
                    
                    Exit For

                'Senão Preenche o Grid, Verifica-se a Tupla Adm + Parce + Termi são iguais a do Movto
                ElseIf objAdmMeioPagtoCondPagto.iParcelamento = objMovimentoCaixa.iParcelamento And objAdmMeioPagtoCondPagto.iTipoCartao = TIPO_MANUAL Then

                    iLinha = iLinha + 1

                    'Joga no Grig
                    GridBoletos.TextMatrix(iLinha, iGrid_Administradora_Col) = objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto
                    GridBoletos.TextMatrix(iLinha, iGrid_Parcelamento_Col) = objAdmMeioPagtoCondPagto.sNomeParcelamento
                                       
                    'Exibe no Grig o Valor do saldo antes de ser executado a Sangria
                    GridBoletos.TextMatrix(iLinha, iGrid_ValorTotal_Col) = Format(objAdmMeioPagtoCondPagto.dSaldo + objMovimentoCaixa.dValor, "standard")
                    GridBoletos.TextMatrix(iLinha, iGrid_ValorSangria_Col) = Format(objMovimentoCaixa.dValor, "standard")

                    Exit For

                End If

            End If

        Next
    
    Next
    
    'atualizar Linhas Existentas
    objGridBoletos.iLinhasExistentes = iLinha

    'Função que Atualiza os Totais
    lErro = Recalcula_Totais()
    If lErro <> SUCESSO Then gError 107811

    Traz_MovimentoBoleto_Tela = SUCESSO

    Exit Function

Erro_Traz_MovimentoBoleto_Tela:

    Traz_MovimentoBoleto_Tela = gErr

    Select Case gErr

        Case 105700
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_INEXISTENTE, gErr, lNumero)

        Case 105738
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_NAO_SANGRIA_BOLETO_CC, gErr, lNumero)

        Case 107755, 107811

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162923)

    End Select

    Exit Function

End Function

Private Sub GridBoletos_Click()

    Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridBoletos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        'Variavel não definida
        Call Grid_Entrada_Celula(objGridBoletos, iAlterado)
    End If

End Sub

Private Sub GridBoletos_EnterCell()

    'Parametro não opcional
    Call Grid_Entrada_Celula(objGridBoletos, iAlterado)

End Sub

Private Sub GridBoletos_GotFocus()

    Call Grid_Recebe_Foco(objGridBoletos)

End Sub

Private Sub GridBoletos_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long

On Error GoTo Erro_GridBoletos_KeyDown

    Call Grid_Trata_Tecla1(KeyCode, objGridBoletos)

    If KeyCode = vbKeyDelete Then
        
        'Função que recalcula os totais no grid
        lErro = Recalcula_Totais()
        If lErro <> SUCESSO Then gError 105730
        
    End If

    Exit Sub
    
Erro_GridBoletos_KeyDown:

    Select Case gErr

        Case 105730

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162924)

    End Select

    Exit Sub

End Sub

Private Sub GridBoletos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridBoletos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridBoletos, iAlterado)
    End If


End Sub

Private Sub GridBoletos_LeaveCell()

    Call Saida_Celula(objGridBoletos)

End Sub

Private Sub GridBoletos_LostFocus()

    Call Grid_Libera_Foco(objGridBoletos)

End Sub

Private Sub GridBoletos_RowColChange()

    Call Grid_RowColChange(objGridBoletos)

End Sub

Private Sub GridBoletos_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridBoletos)

End Sub

Private Sub GridBoletos_Scroll()

    Call Grid_Scroll(objGridBoletos)

End Sub

Private Sub Administradora_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Administradora_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridBoletos)

End Sub

Private Sub Administradora_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBoletos)

End Sub

Private Sub Administradora_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridBoletos.objControle = Administradora
    lErro = Grid_Campo_Libera_Foco(objGridBoletos)

    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Parcelamento_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Parcelamento_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridBoletos)

End Sub

Private Sub Parcelamento_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBoletos)

End Sub

Private Sub Parcelamento_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridBoletos.objControle = Parcelamento
    lErro = Grid_Campo_Libera_Foco(objGridBoletos)

    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ValorTotal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorTotal_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridBoletos)

End Sub

Private Sub ValorTotal_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBoletos)

End Sub

Private Sub ValorTotal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridBoletos.objControle = ValorTotal
    lErro = Grid_Campo_Libera_Foco(objGridBoletos)

    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ValorSangria_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorSangria_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridBoletos)

End Sub

Private Sub ValorSangria_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBoletos)

End Sub

Private Sub ValorSangria_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridBoletos.objControle = ValorSangria
    lErro = Grid_Campo_Libera_Foco(objGridBoletos)

    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

    'Verifica qual a coluna atual do Grid
        Select Case objGridInt.objGrid.Col

            'Administradora
            Case iGrid_Administradora_Col
                lErro = Saida_Celula_Administradora(objGridInt)
                If lErro <> SUCESSO Then gError 107769

            'Parcelamento
            Case iGrid_Parcelamento_Col
                lErro = Saida_Celula_Parcelamento(objGridInt)
                If lErro <> SUCESSO Then gError 107770

            'ValorSangria
            Case iGrid_ValorSangria_Col
                lErro = Saida_Celula_ValorSangria(objGridInt)
                If lErro <> SUCESSO Then gError 107773

        End Select

    'Função que Finaliza a Saida de Celula
    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 107757

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 107757, 107769 To 107771, 107773
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162925)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Administradora(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iCodigo As Integer
Dim iIndice As Integer
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_Saida_Celula_Administradora

    Set objGridInt.objControle = Administradora

    If Administradora.Text <> GridBoletos.TextMatrix(GridBoletos.Row, iGrid_Administradora_Col) Then

        'Limpa o Grid na Linha Relacionada a Administradora em Questão na Coluna Relacionada ao Parcelamento
        GridBoletos.TextMatrix(GridBoletos.Row, iGrid_Parcelamento_Col) = ""
        
        'Limpa o Grid na Linha Relacionada a Administradora em Questão na Coluna Relacionada ao ValorTotal
        GridBoletos.TextMatrix(GridBoletos.Row, iGrid_ValorTotal_Col) = ""

        'Limpa o Grid na Linha Relacionada a Administradora em Questão na Coluna Relacionada ao Valor da Sangria
        GridBoletos.TextMatrix(GridBoletos.Row, iGrid_ValorSangria_Col) = ""

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 107758

    'Verifica se célula do Grid que identifica o campo Admnistradora esta preenchido
    If Len(Trim(GridBoletos.TextMatrix(GridBoletos.Row, iGrid_Administradora_Col))) <> 0 Then

        lErro = Recalcula_Totais()
        If lErro <> SUCESSO Then gError 107814

        'Acrescenta uma linha no Grid se for o caso
        If GridBoletos.Row - GridBoletos.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    Saida_Celula_Administradora = SUCESSO

    Exit Function

Erro_Saida_Celula_Administradora:

    Saida_Celula_Administradora = gErr

    Select Case gErr

        Case 107758

        Case 107814

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162926)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Parcelamento(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iCodigo As Integer
Dim iLinha As Integer
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto
Dim iAchou As Integer

On Error GoTo Erro_Saida_Celula_Parcelamento

    Set objGridInt.objControle = Parcelamento

    If Parcelamento.Text <> GridBoletos.TextMatrix(GridBoletos.Row, iGrid_Parcelamento_Col) Then

        'Verifica se o Parcelamento esta selecionada ou foi Alterada
        If Len(Trim(Parcelamento.Text)) = 0 Then
    
            'Limpa o Campo Valor
            GridBoletos.TextMatrix(GridBoletos.Row, iGrid_ValorTotal_Col) = ""
    
            'Limpa o Campo Relacionado a Sangria
            GridBoletos.TextMatrix(GridBoletos.Row, iGrid_ValorSangria_Col) = ""
    
        'Se o campo Parcelamento foi Alterado e se o Campo Relacionado a Terminal Já Estivar Preenchido
        Else
    
            'Para ver se existe duplicidade no Grig
            For iLinha = 1 To objGridBoletos.iLinhasExistentes
    
                If iLinha <> GridBoletos.Row Then
    
                    If GridBoletos.TextMatrix(iLinha, iGrid_Administradora_Col) = GridBoletos.TextMatrix(GridBoletos.Row, iGrid_Administradora_Col) And _
                    GridBoletos.TextMatrix(iLinha, iGrid_Parcelamento_Col) = Parcelamento.Text Then gError 117544
    
                End If
    
            Next
            
    
            'Procura na Coleção de Cartões o Tupla correspondente e preenche o Grid
            For Each objAdmMeioPagtoCondPagto In gcolCartao
    
                'Verifica se o Nome da Admnistradora é Igual e o parcelamento Selecionada é Igual ao da Tupla Admnistradora + parcelmento
                If objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto = GridBoletos.TextMatrix(GridBoletos.Row, iGrid_Administradora_Col) And objAdmMeioPagtoCondPagto.sNomeParcelamento = Parcelamento.Text And objAdmMeioPagtoCondPagto.iTipoCartao = TIPO_MANUAL Then
    
                    'Se for Boleto Manual
                    'If objAdmMeioPagtoCondPagto.iTipoCartao = BOLETO_MANUAL Then
    
                        'Preenche o Grid com o Valor Total refente aos Boletos em Questão
    
                        GridBoletos.TextMatrix(GridBoletos.Row, iGrid_ValorTotal_Col) = CStr(Format(objAdmMeioPagtoCondPagto.dSaldo, "standard"))
                        GridBoletos.TextMatrix(GridBoletos.Row, iGrid_ValorSangria_Col) = CStr(Format(objAdmMeioPagtoCondPagto.dSaldo, "standard"))
                        iAchou = 1
                        Exit For
    
                    End If
    
            Next
    
            If iAchou = 0 Then gError 117545
    
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 107759

    'Acrescenta uma linha no Grid se for o caso
    If GridBoletos.Row - GridBoletos.FixedRows = objGridInt.iLinhasExistentes Then
        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    End If

    'Função que recalcula os totais no grid
    lErro = Recalcula_Totais()
    If lErro <> SUCESSO Then gError 105728

    Saida_Celula_Parcelamento = SUCESSO

    Exit Function

Erro_Saida_Celula_Parcelamento:

    Saida_Celula_Parcelamento = gErr

    Select Case gErr

        Case 105728, 107759

        Case 117544
            Call Rotina_ErroECF(vbOKOnly, ERRO_LINHA_REPETIDA, gErr)

        Case 117545
            Call Rotina_ErroECF(vbOKOnly, ERRO_PARCELAMENTO_NAO_EXISTENTE1, gErr, Parcelamento.Text)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162927)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ValorSangria(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iCodigo As Integer
Dim iIndice As Integer
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_Saida_Celula_ValorSangria

    Set objGridInt.objControle = ValorSangria

    'Verifica se a Admnistradora esta selecionada ou foi Alterada
    If Len(Trim(ValorSangria.Text)) <> 0 Then

        'Verifica se o Valor Digitado é Valido
        lErro = Valor_NaoNegativo_Critica(ValorSangria.Text)
        If lErro <> SUCESSO Then gError 107761

        'Verifica se o valor da sangria é maior que o valor total se for Erro
        If StrParaDbl(ValorSangria.Text) > StrParaDbl(GridBoletos.TextMatrix(GridBoletos.Row, iGrid_ValorTotal_Col)) Then gError 107762


    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 107764

    'Acrescenta uma linha no Grid se for o caso
    If GridBoletos.Row - GridBoletos.FixedRows = objGridInt.iLinhasExistentes Then
        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    End If

    'Função que recalcula os totais no grid
    lErro = Recalcula_Totais()
    If lErro <> SUCESSO Then gError 107763

    Saida_Celula_ValorSangria = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorSangria:

    Saida_Celula_ValorSangria = gErr

    Select Case gErr

        Case 107761

        Case 107762
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALORSANGRIA_NAO_DISPONIVEL, gErr, ValorSangria.Text, giCodCaixa)

        Case 107763

        Case 107764

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162928)

    End Select

    Exit Function

End Function

'Private Sub ValorSangriaCDebito_Validate(Cancel As Boolean)
''Função que valida os dados no campo valor de sangria em cartões de débito na conta
'
'Dim lErro As Long
'
'On Error GoTo Erro_ValorSangriaCDebito_Validate
'
'    'Verifica se o campo esta preenchido
'    If Len(Trim(ValorSangriaCDebito.Text)) <> 0 Then
'
'        'se esta preenchido então verificar o valor
'        lErro = Valor_NaoNegativo_Critica(ValorSangriaCDebito.Text)
'        If lErro <> SUCESSO Then gError 107765
'
'        'Verifica se o valor da sangria é maior do que o valor de boletos de débitos não especificados
'        If StrParaDbl(ValorSangriaCDebito.Text) > StrParaDbl(LabelEmCaixaCDebitoValor.Caption) Then gError 107766
'
'        ValorSangriaCDebito.Text = Format(ValorSangriaCDebito.Text, "standard")
'
'    End If
'
'    Exit Sub
'
'Erro_ValorSangriaCDebito_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 107765
'
'        Case 107766
'            Call Rotina_ErroECF(vbOKOnly, ERRO_VALORSANGRIA_NAO_DISPONIVEL, gErr, ValorSangriaCDebito.Text, giCodCaixa)
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162929)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Private Sub ValorSangriaCCredito_Validate(Cancel As Boolean)
'Função que valida os dados no campo valor de sangria em cartões de débito na conta

Dim lErro As Long

On Error GoTo Erro_ValorSangriaCDebito_Validate

    'Verifica se o campo esta preenchido
    If Len(Trim(ValorSangriaCCredito.Text)) <> 0 Then

        'se esta preenchido então verificar o valor
        lErro = Valor_NaoNegativo_Critica(ValorSangriaCCredito.Text)
        If lErro <> SUCESSO Then gError 107815

        'Verifica se o valor da sangria é maior do que o valor de boletos de débitos não especificados
        If StrParaDbl(ValorSangriaCCredito.Text) > StrParaDbl(LabelEmCaixaCCreditoValor.Caption) Then gError 107816

        ValorSangriaCCredito.Text = Format(ValorSangriaCCredito.Text, "standard")

    End If

    Exit Sub

Erro_ValorSangriaCDebito_Validate:

    Cancel = True

    Select Case gErr

        Case 107815

        Case 107816
            Call Rotina_ErroECF(vbOKOnly, ERRO_SANGRIA_NAOESP_NAO_DISPONIVEL, gErr, ValorSangriaCCredito.Text, LabelEmCaixaCCreditoValor.Caption)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162930)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Função que Realiza a Gravação

Dim lErro As Long
Dim colMovimentosCaixa As New Collection
Dim vbMsgRes As VbMsgBoxResult
Dim iTipoMovimento As Integer
Dim lNumMovto As Long

On Error GoTo Erro_BotaoGravar_Click

    If Not AFRAC_ImpressoraCFe(giCodModeloECF) Then
    
        lErro = CF_ECF("Requisito_XXII")
        If lErro <> SUCESSO Then gError 207971

    End If

    'Função que efeuara a gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 107783

    'Função que Limpa a Tela
    lErro = Limpa_Tela_MovimentoBoleto
    If lErro <> SUCESSO Then gError 107784

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 107783, 107784, 207971

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162931)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long

Dim lErro As Long
Dim lNumMovto As Long
Dim iTipoMovimento As Integer
Dim colMovimentosCaixa As New Collection
Dim vbMsgRes As VbMsgBoxResult
Dim objMovCC As ClassMovimentoCaixa

On Error GoTo Erro_Gravar_Registro

    'Função que Valida a Gravação, se existem linhas repetidas no Grid
    lErro = MovimentoBoleto_Valida_Gravacao()
    If lErro <> SUCESSO Then gError 107767

    'Transforma o Codigo Texto para Long
    lNumMovto = StrParaLong(Codigo.Text)

    'Lê os Movimentos de Caixa da coleção global e carrega no coleção local a tela para o numero do Movto
    lErro = CF_ECF("Caixa_MovimentoCaixa_Le_NumMovto", colMovimentosCaixa, lNumMovto)
    If lErro <> SUCESSO And lErro <> 107850 Then gError 107768

    'Verifica se já existe um movimento para o código referido, Verifica se é Alteração
    If colMovimentosCaixa.Count > 0 Then
    
        Set objMovCC = colMovimentosCaixa(1)

        If (objMovCC.iTipo <> MOVIMENTOCAIXA_SANGRIA_BOLETO_CC) Then gError 86293
        
         iTipoMovimento = MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_B0LETO

        'Envia aviso perguntando se deseja atualizar o movimemtos
        vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_ALTERACAO_MOVIMENTOCAIXA, Codigo.Text)

        'Se a Reposta for Negativa
        If vbMsgRes = vbNo Then gError 107777

        'Função que Faz a Alteração na Sangria de Boleto Previamente Executada, adciona o iTipoMovimento
        lErro = MovimentoCaixa_Prepara_Exclusao(colMovimentosCaixa, iTipoMovimento)
        If lErro <> SUCESSO Then gError 107774

    End If

    'Move os Dados da Sangria para a Memoria
    lErro = Move_Dados_Memoria(colMovimentosCaixa)
    If lErro <> SUCESSO Then gError 107775

    'Função que Grava os movimentos em Arquivos
    lErro = Caixa_Grava_Movimento(colMovimentosCaixa)
    If lErro <> SUCESSO Then gError 107791

    'Atualiza os Dados da Memoria
    lErro = MovimentoBoleto_Atualiza_Memoria(colMovimentosCaixa)
    If lErro <> SUCESSO Then gError 107776

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

   Select Case gErr
        
        Case 86293
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_NAO_CARTAO, gErr, Codigo.Text)

        Case 86294
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_JA_TRANSMITIDO, gErr)

        Case 107767, 107768, 107774 To 107777, 107791

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162932)

    End Select

End Function

Function MovimentoBoleto_Valida_Gravacao() As Long
'Função que Valida a Gravação

Dim lErro As Long
Dim iIndice As Integer
Dim iCont As Integer
Dim dValor As Double

On Error GoTo Erro_MovimentoBoleto_Valida_Gravacao

    'Verifica se o código Foi Preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 107778

    'O valor da sangria nao pode estar zerado
    If StrParaDbl(LabelSangriaValor.Caption) = 0 And StrParaDbl(ValorSangriaCCredito.Text) = 0 Then gError 107782

    'Verifica se no Grid o Campo Referente a Sangria Esta Preenchido
    For iIndice = 1 To objGridBoletos.iLinhasExistentes
        
        If Len(GridBoletos.TextMatrix(iIndice, iGrid_Administradora_Col)) = 0 Then gError 105730
        
        If Len(GridBoletos.TextMatrix(iIndice, iGrid_Parcelamento_Col)) = 0 Then gError 105731
        
        If StrParaDbl(GridBoletos.TextMatrix(iIndice, iGrid_ValorSangria_Col)) = 0 Then gError 105729

        'Para ver se existe duplicidade no Grig
        For iCont = 1 To objGridBoletos.iLinhasExistentes

            If iCont <> iIndice Then

                If GridBoletos.TextMatrix(iCont, iGrid_Administradora_Col) = GridBoletos.TextMatrix(iIndice, iGrid_Administradora_Col) And _
                GridBoletos.TextMatrix(iCont, iGrid_Parcelamento_Col) = GridBoletos.TextMatrix(iIndice, iGrid_Parcelamento_Col) Then gError 107781

            End If

        Next

    Next

    MovimentoBoleto_Valida_Gravacao = SUCESSO

    Exit Function

Erro_MovimentoBoleto_Valida_Gravacao:

    MovimentoBoleto_Valida_Gravacao = gErr

    Select Case gErr

        Case 105729
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALORSANGRIA_NAO_INFORMADO_GRID, gErr, iIndice)

        Case 105730
            Call Rotina_ErroECF(vbOKOnly, ERRO_ADMINISTRADORA_NAO_PREENCHIDO_GRID, gErr, iIndice)

        Case 105731
            Call Rotina_ErroECF(vbOKOnly, ERRO_PARCELAMENTO_NAO_PREENCHIDO_GRID, gErr, iIndice)

        Case 107778
            Call Rotina_ErroECF(vbOKOnly, ERRO_CODIGO_NAO_PREENCHIDO1, gErr)

        Case 107781
            Call Rotina_ErroECF(vbOKOnly, ERRO_LINHA_REPETIDA, gErr)

        Case 107782, 107780
            Call Rotina_ErroECF(vbOKOnly, ERRO_SANGRIA_NAOPODE_SER_EXECUTADA, gErr)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162933)

    End Select

    Exit Function

End Function

Function MovimentoCaixa_Prepara_Exclusao(colMovimentosCaixa As Collection, Optional iTipoMovimento As Integer) As Long
'Função que Para Cada obj da Coleção adciona a esse movimento dizando q foi alterado o valor da Sangria

Dim lErro As Long
Dim objMovimentoCaixa As New ClassMovimentoCaixa

On Error GoTo Erro_MovimentoCaixa_Prepara_Exclusao

    For Each objMovimentoCaixa In colMovimentosCaixa

        'Adciona o Tipo de Movimento a Coleção de Movimentos
        objMovimentoCaixa.iTipo = iTipoMovimento

    Next

    MovimentoCaixa_Prepara_Exclusao = SUCESSO

    Exit Function

Erro_MovimentoCaixa_Prepara_Exclusao:

    MovimentoCaixa_Prepara_Exclusao = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162934)

    End Select

    Exit Function

End Function

Function Move_Dados_Memoria(colMovimentosCaixa As Collection) As Long
'Função que Move os dados para a memoria

Dim lErro As Long
Dim iIndice As Integer
Dim objMovimentosCaixa As ClassMovimentoCaixa
Dim objAdmMeioPagto As New ClassAdmMeioPagto
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto
Dim objAdmMeioPagtoCondPagtoAux As New ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_Move_Dados_Memoria

    'verifica para cada linha do grid
    For iIndice = 1 To objGridBoletos.iLinhasExistentes

        'Instancia um novo obj
        Set objMovimentosCaixa = New ClassMovimentoCaixa

        'Guarda o valor da Sangria
        objMovimentosCaixa.dValor = StrParaDbl(GridBoletos.TextMatrix(iIndice, iGrid_ValorSangria_Col))

        'Guarda em qual filial empresa que esta Trabalhando
        objMovimentosCaixa.iFilialEmpresa = giFilialEmpresa
        
        'Guardo o codigo do movimento
        objMovimentosCaixa.lNumMovto = StrParaLong(Codigo.Text)

        For Each objAdmMeioPagtoCondPagtoAux In gcolCartao

            If GridBoletos.TextMatrix(iIndice, iGrid_Administradora_Col) = objAdmMeioPagtoCondPagtoAux.sNomeAdmMeioPagto And GridBoletos.TextMatrix(iIndice, iGrid_Parcelamento_Col) = objAdmMeioPagtoCondPagtoAux.sNomeParcelamento Then

                'Guardo o Codigo da Admnistradora no Movimento Caixa
                objMovimentosCaixa.iAdmMeioPagto = objAdmMeioPagtoCondPagtoAux.iAdmMeioPagto
                'Guarda o Código do Parcelamento da Linha
                objMovimentosCaixa.iParcelamento = objAdmMeioPagtoCondPagtoAux.iParcelamento
                'Move o Tipo de Terminal para o ObjMovimentoCaixa
                objMovimentosCaixa.iTipoCartao = TIPO_MANUAL
                
                Exit For

            End If

        Next

        objMovimentosCaixa.iTipo = MOVIMENTOCAIXA_SANGRIA_BOLETO_CC
    
        'Adciona a Coleção de ColMovimentosCaixa
        colMovimentosCaixa.Add objMovimentosCaixa

    Next

    If StrParaDbl(ValorSangriaCCredito.Text) <> 0 Then

        'Instancia novo Obj
        Set objMovimentosCaixa = New ClassMovimentoCaixa

        'Guarda Zero no Código do Admnistradora não especificada
        objMovimentosCaixa.lNumMovto = StrParaLong(Codigo.Text)

        'Guarda em qual filial empresa que esta Trabalhando
        objMovimentosCaixa.iFilialEmpresa = giFilialEmpresa

        'Guarda o valor da Sangria
        objMovimentosCaixa.dValor = StrParaDbl(ValorSangriaCCredito.Text)

        'Guardo o Codigo da Admnistradora no Movimento Caixa
        objMovimentosCaixa.iAdmMeioPagto = 0

        'Guarda Zero no Código do Parcelamento não especificado
        objMovimentosCaixa.iParcelamento = PARCELAMENTO_AVISTA

        'Guarda o Tipo de Movimento
        objMovimentosCaixa.iTipo = MOVIMENTOCAIXA_SANGRIA_BOLETO_CC

        'Adciona a Coleção de ColMovimentosCaixa
        colMovimentosCaixa.Add objMovimentosCaixa

    End If

    Move_Dados_Memoria = SUCESSO

    Exit Function

Erro_Move_Dados_Memoria:

    Move_Dados_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162935)

    End Select

    Exit Function

End Function

Function MovimentoBoleto_Atualiza_Memoria(colMovimentosCaixa As Collection) As Long
'Função que Limpa a Coleção Global a Tela apos a Função de Gravação

Dim lErro As Long
Dim objMovimentoCaixa As New ClassMovimentoCaixa

On Error GoTo Erro_MovimentoBoleto_Atualiza_Memoria

    For Each objMovimentoCaixa In colMovimentosCaixa

        'Função que Atualiza os Boletos Excluidos
        lErro = CF_ECF("MovimentoBoleto_Atualiza_Memoria1", objMovimentoCaixa)
        If lErro <> SUCESSO Then gError 107787

        If objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_B0LETO Then

            'Função que Retira de memória os Movimentos Excluidos
            lErro = MovimentoCaixa_Exclui_Memoria(objMovimentoCaixa)
            If lErro <> SUCESSO Then gError 107786
            
        Else
        
            'Adcionar a Coleção Global o objMovimento Caixa
            gcolMovimentosCaixa.Add objMovimentoCaixa
    
        End If

    Next

    MovimentoBoleto_Atualiza_Memoria = SUCESSO

    Exit Function

Erro_MovimentoBoleto_Atualiza_Memoria:

    MovimentoBoleto_Atualiza_Memoria = gErr

    Select Case gErr

        Case 107786, 107787

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162936)

    End Select

    Exit Function

End Function

Function MovimentoCaixa_Exclui_Memoria(objMovimentoCaixa As ClassMovimentoCaixa) As Long
'Função que Exclui da Memória os Movimentos que Foram Alterados

Dim lErro As Long
Dim objMovimentoCaixaAux As New ClassMovimentoCaixa
Dim iIndice As Integer

On Error GoTo Erro_MovimentoCaixa_Exclui_Memoria

    For iIndice = gcolMovimentosCaixa.Count To 1 Step -1

        Set objMovimentoCaixaAux = gcolMovimentosCaixa.Item(iIndice)

        'Verifica se o movimento é o mesmo
        If objMovimentoCaixa.lNumMovto = objMovimentoCaixaAux.lNumMovto And objMovimentoCaixa.lSequencial = objMovimentoCaixaAux.lSequencial Then

            'Exclui o movimento da Coleção Global de MovimentosCaixa
            gcolMovimentosCaixa.Remove (iIndice)

            'Sai do Loop
            Exit For

        End If

    Next

    MovimentoCaixa_Exclui_Memoria = SUCESSO

    Exit Function

Erro_MovimentoCaixa_Exclui_Memoria:

    MovimentoCaixa_Exclui_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162937)

    End Select

    Exit Function

End Function


Function Limpa_Tela_MovimentoBoleto() As Long
'Função que Limpa a Tela

Dim lErro As Long
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_Limpa_Tela_MovimentoBoleto

    'Limpa os Controles básico da Tela
    Call Limpa_Tela(Me)

    'Limpa Grid
    Call Grid_Limpa(objGridBoletos)

    'Limpa os Labes da Tela
    LabelEmCaixaValor.Caption = Format(0, "standard")
    LabelSangriaValor.Caption = Format(0, "standard")
    LabelEmCaixaCCreditoValor.Caption = Format(0, "standard")

    iAlterado = 0
    
    Limpa_Tela_MovimentoBoleto = SUCESSO

    Exit Function

Erro_Limpa_Tela_MovimentoBoleto:

    Limpa_Tela_MovimentoBoleto = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162938)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()
'Botão que Exclui um Movimento de Sangria

Dim lErro As Long
Dim colMovcxExclui As New Collection
Dim vbMsgRes As VbMsgBoxResult
Dim lNumero As Long
Dim iTipoMovimento As Integer
Dim objMovCx As New ClassMovimentoCaixa

On Error GoTo Erro_BotaoExcluir_Click

    If Not AFRAC_ImpressoraCFe(giCodModeloECF) Then
    
        lErro = CF_ECF("Requisito_XXII")
        If lErro <> SUCESSO Then gError 207972

    End If

    'Verifica se o Codigo não foi Preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 107788

    lNumero = StrParaLong(Codigo.Text)

    'Verifica os Movimentos de Caixa para o Código em Questão
    lErro = CF_ECF("Caixa_MovimentoCaixa_Le_NumMovto", colMovcxExclui, lNumero)
    If lErro <> SUCESSO And lErro <> 107850 Then gError 107854
    
    'Senão encontrou ninguem
    If lErro = 107850 Then gError 107789
    
    Set objMovCx = colMovcxExclui(1)
    
    If (objMovCx.iTipo <> MOVIMENTOCAIXA_SANGRIA_BOLETO_CD) And (objMovCx.iTipo <> MOVIMENTOCAIXA_SANGRIA_BOLETO_CC) Then gError 86293
    
    'Pergunta se deseja Realmente Excluir o Movimento
    vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_EXCLUSAO_MOVIMENTOCAIXA, Codigo.Text)
    If vbMsgRes = vbNo Then gError 107811
    
    iTipoMovimento = MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_B0LETO
         
    'guardo os sequenciais para poder reconhecer o movimento na inicialização do sistema
    For Each objMovCx In colMovcxExclui
        objMovCx.lNumRefInterna = objMovCx.lSequencial
    Next
    
    'Prepara os Movimentos para a Exclusão
    lErro = MovimentoCaixa_Prepara_Exclusao(colMovcxExclui, iTipoMovimento)
    If lErro <> SUCESSO Then gError 107790

    'Função que Grava a Exclusão de Boletos
    lErro = Caixa_Grava_Movimento(colMovcxExclui)
    If lErro <> SUCESSO Then gError 107792
    
    'Atualiza os Dados na Memória
    lErro = MovimentoBoleto_Atualiza_Memoria(colMovcxExclui)
    If lErro <> SUCESSO Then gError 107836
    
    'Função Que Limpa a Tela
    Call Limpa_Tela_MovimentoBoleto

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 86293
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_NAO_CARTAO, gErr, Codigo.Text)

        Case 86294
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_JA_TRANSMITIDO, gErr)

        Case 107788
            Call Rotina_ErroECF(vbOKOnly, ERRO_CODIGO_NAO_PREENCHIDO1, gErr)

        Case 107789
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_INEXISTENTE, gErr, lNumero)

        Case 107790, 107792, 107836, 107811, 107854, 207972

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162939)

    End Select

    Exit Sub
End Sub

Function Caixa_Grava_Movimento(colMovimentosCaixa As Collection) As Long

Dim lErro As Long
Dim lSequencial As Long
Dim colRegistro As New Collection
Dim iIndice As Integer
Dim objMovimentoCaixa As New ClassMovimentoCaixa
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto
Dim sNomeTerminal As String
Dim iCont As Integer
Dim objMovimentoCaixaAux As New ClassMovimentoCaixa
Dim sMensagem As String
Dim objMovCx As ClassMovimentoCaixa
Dim vbMsgRes As VbMsgBoxResult
Dim objOperador As New ClassOperador
Dim lTamanho As Long
Dim sRetorno As String
Dim sArquivo As String

On Error GoTo Erro_Caixa_Grava_Movimento

    If giStatusSessao = SESSAO_ENCERRADA Then

        'Envia aviso perguntando se de seja Abrir sessão
        vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_DESEJA_ABRIR_SESSAO, giCodCaixa)

        If vbMsgRes = vbNo Then gError 107793

        'Função que Executa Abertura na Sessão
        lErro = CF_ECF("Operacoes_Executa_Abertura")
        If lErro <> SUCESSO Then gError 107794

    End If

    'Se for Necessário a Altorização do Gerente para abertura do Caixa
    If gobjLojaECF.iGerenteAutoriza = AUTORIZACAO_GERENTE Then

        'Chama a Tela de Senha
        Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)

        'Sai de Função se a Tela de Login não Retornar ok
        If giRetornoTela <> vbOK Then gError 107795

    End If
    
    lTamanho = 255
    sRetorno = String(lTamanho, 0)
    
    'Obtém o diretório onde deve ser armazenado o arquivo com dados do backoffice
    Call GetPrivateProfileString(APLICACAO_DADOS, "DirDadosCC", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    
    'Retira os espaços no final da string
    sRetorno = StringZ(sRetorno)
    
    'Se não encontrou
    If Len(Trim(sRetorno)) = 0 Or sRetorno = CStr(CONSTANTE_ERRO) Then gError 127097
    
    If right(sRetorno, 1) <> "\" Then sRetorno = sRetorno & "\"
    
    sArquivo = sRetorno & giCodEmpresa & "_" & giFilialEmpresa & "_" & NOME_ARQUIVOCC
    
    'Abre o arquivo de retorno
    Open sArquivo For Input Lock Read Write As #10
    
    'Função que Abre a Transação de Caixa, Identificador dentro do Caixa para um determinado MOVTO
    lErro = CF_ECF("Caixa_Transacao_Abrir", lSequencial)
    If lErro <> SUCESSO Then gError 107796

    lTamanho = 255
    sRetorno = String(lTamanho, 0)
        
    'Obtém a ultima transacao transferida
    Call GetPrivateProfileString(APLICACAO_DADOS, "UltimaTransacaoTransf", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
        
    'Retira os espaços no final da string
    sRetorno = StringZ(sRetorno)

    For Each objMovimentoCaixa In colMovimentosCaixa

        'se o numero da ultima transacao transferida ultrapassar o numero da transacao do movimento de caixa
        If objMovimentoCaixa.lSequencial <> 0 And StrParaLong(sRetorno) > objMovimentoCaixa.lSequencial Then gError 133846

        'Caso nao precise de autorizacao do gerente nesta transacao ==> objOperador.iCodigo vai estar zerado
        objMovimentoCaixa.iGerente = objOperador.iCodigo

        lErro = Caixa_Grava_MovCx(objMovimentoCaixa, lSequencial)
        If lErro <> SUCESSO Then gError 105708

    Next

    lSequencial = lSequencial - 1
    
    'Fecha a Transação , Grava o Novo Sequencial
    lErro = CF_ECF("Caixa_Transacao_Fechar", lSequencial)
    If lErro <> SUCESSO Then gError 107799

    Close #10

    Caixa_Grava_Movimento = SUCESSO

    Exit Function

Erro_Caixa_Grava_Movimento:

    Close #10

    Caixa_Grava_Movimento = gErr

    Select Case gErr

        Case 105707, 105708, 107793 To 107799, 109807, 109813
            'Erro Tratado Dentro da Função Chamada

        Case 107793, 107794, 107795
        
        Case 133846
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_JA_TRANSMITIDO, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162940)

    End Select
    
    lErro = CF_ECF("Caixa_Transacao_Rollback", glTransacaoPAFECF)
    
    Exit Function

End Function

Function Caixa_Grava_MovCx(objMovimentoCaixa As ClassMovimentoCaixa, lSequencial As Long) As Long
'grava cada movimento de caixa passado como parametro

Dim lErro As Long
Dim colRegistro As New Collection
Dim sMensagem As String
Dim objMovCx As ClassMovimentoCaixa
Dim objTela As Object

On Error GoTo Erro_Caixa_Grava_MovCx

    'se for um movimento de exclusao de sangria ==> cria um novo movimento de caixa, grava no arquivão e deixa o que foi passado
    'como parametro sem alterar para permitir posteriormente retirar o movimento da memoria pelo sequencial original
    If objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_B0LETO Then
                                 
        Set objMovCx = New ClassMovimentoCaixa
            
        lErro = CF_ECF("MovimentoCaixa_Copia", objMovimentoCaixa, objMovCx)
        If lErro <> SUCESSO Then gError 105707
    
    Else
    
        Set objMovCx = objMovimentoCaixa
    
    End If
    
    'Guarda o Sequencial no objmovimentoCaixa
    objMovCx.lSequencial = lSequencial

    lSequencial = lSequencial + 1

    'Guarda no objMovimentoCaixa os Dados que Serão Usados para a Geração do Movimento de Caixa
    lErro = CF_ECF("Move_DadosGlobais_Memoria", objMovCx)
    If lErro <> SUCESSO Then gError 107797

    'Funçao que Gera o Arquivo preparando para a gravação
    Call CF_ECF("MovimentoBoleto_Gera_Log", colRegistro, objMovCx)

    'Função que Vai Gravar as Informações no Arquivo de Caixa
    lErro = CF_ECF("MovimentoCaixaECF_Grava", colRegistro)
    If lErro <> SUCESSO Then gError 107798
    
    Set objTela = Me
    
    'para não ficar 3 movimentos com o mesmo Código(Numero de Movto) na Coleção gcolMovto
    If objMovCx.iTipo = MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_B0LETO Then
        
'        'Faz a sangria
'        lErro = CF_ECF("Sangria_AFRAC", objMovCx.dValor, sMensagem, -1)
'        If lErro <> SUCESSO Then gError 109807
'
'    Else
''        'Faz a sangria
''        lErro = CF_ECF("Sangria_AFRAC", objMovCx.dValor, sMensagem, 0)
''        If lErro <> SUCESSO Then gError 109813
'    End If

        lErro = AFRAC_AbrirRelatorioGerencial(RELGER_SANGRIA_CARTAO_CREDITO, objTela)

        lErro = AFRAC_ImprimirRelatorioGerencial("Exclusão Sangria Boleto - Valor: " & Format(objMovCx.dValor, "standard"), objTela)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Exclusao Sangria - Boleto")
        If lErro <> SUCESSO Then gError 117671

        lErro = AFRAC_FecharRelatorioGerencial(objTela)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Exclusao Sangria - Boleto")
        If lErro <> SUCESSO Then gError 117672

        
    Else
        
        lErro = AFRAC_AbrirRelatorioGerencial(RELGER_SANGRIA_CARTAO_CREDITO, objTela)

        lErro = AFRAC_ImprimirRelatorioGerencial("Sangria Boleto - Valor: " & Format(objMovCx.dValor, "standard"), objTela)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Sangria - Boleto")
        If lErro <> SUCESSO Then gError 117673

        lErro = AFRAC_FecharRelatorioGerencial(objTela)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Sangria - Boleto")
        If lErro <> SUCESSO Then gError 117674

    End If

    Caixa_Grava_MovCx = SUCESSO
    
    Exit Function

Erro_Caixa_Grava_MovCx:

    Caixa_Grava_MovCx = gErr

    Select Case gErr

        Case 105701, 107883, 107884, 109808, 109814, 117671 To 117674

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162941)

    End Select
    
    Exit Function

End Function

Function Traz_ValoresEmCaixa_Tela() As Long
'Função que Traz Todos os valores de teodos os movimentos relacionados com boletos para a Tela independente de código

Dim lErro As Long
Dim iCont1 As Integer
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_Traz_ValoresEmCaixa_Tela

    LabelEmCaixaCCreditoValor.Caption = CStr(Format(0, "standard"))
    ValorSangriaCCredito.Text = CStr(Format(0, "standard"))

    'Verifica na Coleção Global
    For Each objAdmMeioPagtoCondPagto In gcolCartao

        If objAdmMeioPagtoCondPagto.iTipoCartao = BOLETO_MANUAL Then
            
            'se for cartao especificado e com saldo
            If objAdmMeioPagtoCondPagto.iAdmMeioPagto <> 0 And objAdmMeioPagtoCondPagto.dSaldo <> 0 Then
            
                iCont1 = iCont1 + 1
        
                'Joga no Grid o nome da Adm
                GridBoletos.TextMatrix(iCont1, iGrid_Administradora_Col) = objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto
                GridBoletos.TextMatrix(iCont1, iGrid_Parcelamento_Col) = objAdmMeioPagtoCondPagto.sNomeParcelamento
                GridBoletos.TextMatrix(iCont1, iGrid_ValorTotal_Col) = Format(objAdmMeioPagtoCondPagto.dSaldo, "Standard")
                GridBoletos.TextMatrix(iCont1, iGrid_ValorSangria_Col) = Format(objAdmMeioPagtoCondPagto.dSaldo, "Standard")
            
            Else
            
                'se for cartao de credito nao especificado e com saldo
                If objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto = STRING_NAO_DETALHADO_CCREDITO And objAdmMeioPagtoCondPagto.dSaldo <> 0 Then
    
                    LabelEmCaixaCCreditoValor.Caption = CStr(Format(objAdmMeioPagtoCondPagto.dSaldo, "standard"))
                    ValorSangriaCCredito.Text = CStr(Format(objAdmMeioPagtoCondPagto.dSaldo, "standard"))
    
                End If
                
            End If

        End If
        
    Next

    objGridBoletos.iLinhasExistentes = iCont1

    Traz_ValoresEmCaixa_Tela = SUCESSO

    Exit Function

Erro_Traz_ValoresEmCaixa_Tela:

    Traz_ValoresEmCaixa_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162942)

    End Select

    Exit Function

End Function

Function Traz_ValoresEmCaixa_ValoresMovto_Tela(colMovimentosCaixa As Collection) As Long
'Função que Traz a tela os movimentos encontrados

Dim lErro As Long
Dim colCartao As New Collection
Dim iIndice As Integer
Dim objMovimentoCaixa As New ClassMovimentoCaixa
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto
Dim objAdmMeioPagtoCondPagtoAux As ClassAdmMeioPagtoCondPagto
Dim objAdmMeioPagtoParcelasAux As ClassAdmMeioPagtoParcelas
Dim objAdmMeioPagtoParcelas As ClassAdmMeioPagtoParcelas
Dim iCont As Integer
Dim iProcuraCombo As Integer
Dim sNomeAdm As String
Dim sNomeParc As String
Dim sNomeTerm As String

On Error GoTo Erro_Traz_ValoresEmCaixa_ValoresMovto_Tela

    'Copia Item a Item
    For Each objAdmMeioPagtoCondPagto In gcolCartao

        Set objAdmMeioPagtoCondPagtoAux = New ClassAdmMeioPagtoCondPagto

        objAdmMeioPagtoCondPagtoAux.dDesconto = objAdmMeioPagtoCondPagto.dDesconto
        objAdmMeioPagtoCondPagtoAux.dJuros = objAdmMeioPagtoCondPagto.dJuros
        objAdmMeioPagtoCondPagtoAux.dSaldo = objAdmMeioPagtoCondPagto.dSaldo
        objAdmMeioPagtoCondPagtoAux.dTaxa = objAdmMeioPagtoCondPagto.dTaxa
        objAdmMeioPagtoCondPagtoAux.dValorMinimo = objAdmMeioPagtoCondPagto.dValorMinimo
        objAdmMeioPagtoCondPagtoAux.iAdmMeioPagto = objAdmMeioPagtoCondPagto.iAdmMeioPagto
        objAdmMeioPagtoCondPagtoAux.iFilialEmpresa = objAdmMeioPagtoCondPagto.iFilialEmpresa
        objAdmMeioPagtoCondPagtoAux.iJurosParcelamento = objAdmMeioPagtoCondPagto.iJurosParcelamento
        objAdmMeioPagtoCondPagtoAux.iNumParcelas = objAdmMeioPagtoCondPagto.iNumParcelas
        objAdmMeioPagtoCondPagtoAux.iParcelamento = objAdmMeioPagtoCondPagto.iParcelamento
        objAdmMeioPagtoCondPagtoAux.iParcelasRecebto = objAdmMeioPagtoCondPagto.iParcelasRecebto
        objAdmMeioPagtoCondPagtoAux.iTipoCartao = objAdmMeioPagtoCondPagto.iTipoCartao
        objAdmMeioPagtoCondPagtoAux.sNomeAdmMeioPagto = objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto
        objAdmMeioPagtoCondPagtoAux.sNomeParcelamento = objAdmMeioPagtoCondPagto.sNomeParcelamento

        For Each objAdmMeioPagtoParcelas In objAdmMeioPagtoCondPagto.colParcelas

            Set objAdmMeioPagtoParcelasAux = New ClassAdmMeioPagtoParcelas

            objAdmMeioPagtoParcelasAux.dPercRecebimento = objAdmMeioPagtoParcelas.dPercRecebimento
            objAdmMeioPagtoParcelasAux.iAdmMeioPagto = objAdmMeioPagtoParcelas.iAdmMeioPagto
            objAdmMeioPagtoParcelasAux.iFilialEmpresa = objAdmMeioPagtoParcelas.iFilialEmpresa
            objAdmMeioPagtoParcelasAux.iIntervaloRecebimento = objAdmMeioPagtoParcelas.iIntervaloRecebimento
            objAdmMeioPagtoParcelasAux.iParcela = objAdmMeioPagtoParcelas.iParcela
            objAdmMeioPagtoParcelasAux.iParcelamento = objAdmMeioPagtoParcelas.iParcelamento

            objAdmMeioPagtoCondPagtoAux.colParcelas.Add objAdmMeioPagtoParcelasAux
        
        Next

    colCartao.Add objAdmMeioPagtoCondPagtoAux

    Next

    'Verifica para Cada Indice da Coleção
    For iIndice = colCartao.Count To 1 Step -1
        
        'Verifica para da Coleção de Moviementos de Caixa
        For Each objMovimentoCaixa In colMovimentosCaixa
            'Verifica se o Código da Admnistradora é Igual a Zero , se For Indica q é Cartão de Débito o de Crédito
            If objMovimentoCaixa.iAdmMeioPagto = 0 Then
                'Se for Cartão de Débito
                If objMovimentoCaixa.iTipoCartao = BOLETO_POS Then
'                    'Instancia o objAdmMeioPagtoCondPagto para a apontar para Coleção de cartões
'                    Set objAdmMeioPagtoCondPagto = colCartao.Item(iIndice)
'                    If objAdmMeioPagtoCondPagto.iTipoCartao = BOLETO_POS Then
'                        'Preenche o Valor que Representa a Soma Total do que se Tem de Cartão de débito para ser recebido
'                        LabelEmCaixaCDebitoValor.Caption = Format(objMovimentoCaixa.dValor + objAdmMeioPagtoCondPagto.dSaldo, "Standard")
'                        ValorSangriaCDebito.Text = Format(objMovimentoCaixa.dValor, "standard")
'
'                    End If
'
                'Se for Cartão de Crédito
                ElseIf objMovimentoCaixa.iTipoCartao = BOLETO_MANUAL Then
                    'Instancia o objAdmMeioPagtoCondPagto para a apontar para Coleção de cartões
                    Set objAdmMeioPagtoCondPagto = colCartao.Item(iIndice)

                    If objAdmMeioPagtoCondPagto.iTipoCartao = BOLETO_MANUAL Then

                        'Preenche o Valor que Representa a Soma Total do que se Tem de Cartão de Crédito para ser recebido
                        LabelEmCaixaCCreditoValor.Caption = Format(objMovimentoCaixa.dValor + objAdmMeioPagtoCondPagto.dSaldo, "Standard")
                        ValorSangriaCCredito.Text = Format(objMovimentoCaixa.dValor, "standard")
                    End If
                End If

            Else

                'Instancia o objAdmMeioPagtoCondPagto para a apontar para Coleção de cartões
                Set objAdmMeioPagtoCondPagto = colCartao.Item(iIndice)
                'Verfica se o Movimento de Caixa está Relacionada com o cartão na coleção Global de Cartão (Parcelamento , Terminal)
                If objMovimentoCaixa.iAdmMeioPagto = objAdmMeioPagtoCondPagto.iAdmMeioPagto And objMovimentoCaixa.iParcelamento = objAdmMeioPagtoCondPagto.iParcelamento And objMovimentoCaixa.iTipoCartao = objAdmMeioPagtoCondPagto.iTipoCartao Then
                    'Verifica na Combo se existe a Admnistradora se não exitir Seleciona
                    iCont = iCont + 1

                    'Escreve no Grid o Nome da Admnistradora
                    GridBoletos.TextMatrix(iCont, iGrid_Administradora_Col) = objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto
                    'Escreve no Grid o Nome do Parcelamento
                    GridBoletos.TextMatrix(iCont, iGrid_Parcelamento_Col) = objAdmMeioPagtoCondPagto.sNomeParcelamento

                    'Escreve no Grid o Valor Todal daquela Adm , Parcelamento , Terminal
                    GridBoletos.TextMatrix(iCont, iGrid_ValorTotal_Col) = Format(objMovimentoCaixa.dValor + objAdmMeioPagtoCondPagto.dSaldo, "standard")
                    'Escreve no Grid o Valor da Sangria de Boleto daquela Adm , Parcelamento , Terminal
                    GridBoletos.TextMatrix(iCont, iGrid_ValorSangria_Col) = Format(objMovimentoCaixa.dValor, "standard")

                End If

            End If

        Next
    
        'Remove o Boleto da Coleção de Cartão
        colCartao.Remove (iIndice)

    Next
    
    'Atualiza o Numero de Linhas existentes no Grid
    objGridBoletos.iLinhasExistentes = iCont
    
    'Função que Serve para Recalcular os Totais
    lErro = Recalcula_Totais()
    If lErro <> SUCESSO Then gError 107818

    Traz_ValoresEmCaixa_ValoresMovto_Tela = SUCESSO

    Exit Function

Erro_Traz_ValoresEmCaixa_ValoresMovto_Tela:

    Traz_ValoresEmCaixa_ValoresMovto_Tela = gErr

    Select Case gErr

        Case 107818
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162943)

    End Select

    Exit Function

End Function


Private Sub BotaoLimpar_Click()
'Função Limpa a Tela

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Botaolimpar_Click

   If iAlterado = REGISTRO_ALTERADO Then

        'Envia aviso perguntando se os pagamentos devem ser aproveitados
        vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_ALTERACAO_MOVIMENTOCAIXA1, Codigo.Text)

        If vbMsgRes = vbYes Then

             lErro = Gravar_Registro()
             If lErro <> SUCESSO Then gError 107812

        End If

    End If

    Call Limpa_Tela_MovimentoBoleto

    'Função que Traz para a Tela os Saldos Relacionados aos Cartões de Débito e Crédito não especificados
    lErro = Carrega_MovimentoBoleto_DebitoCredito()
    If lErro <> SUCESSO Then gError 107817

    Exit Sub

Erro_Botaolimpar_Click:

    Select Case gErr

        Case 107812, 107817

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162944)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
'Botão que Executa o Fechamento da Tela

    Unload Me


End Sub


Private Sub Administradora_Click()

    iAlterado = REGISTRO_ALTERADO
    

End Sub

Private Sub Parcelamento_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Function Carrega_MovimentoBoleto_DebitoCredito() As Long
'Função que Carrega os Controles Relacionados aos Cartões de Débito e Credito não especificados

Dim lErro As Long
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_Carrega_MovimentoBoleto_DebitoCredito

    For Each objAdmMeioPagtoCondPagto In gcolCartao
'
'        If objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto = STRING_NAO_DETALHADO_CDEBITO Then
'
'            LabelEmCaixaCDebitoValor.Caption = CStr(Format(objAdmMeioPagtoCondPagto.dSaldo, "standard"))
'            ValorSangriaCDebito.Text = CStr(Format(objAdmMeioPagtoCondPagto.dSaldo, "standard"))
'
        If objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto = STRING_NAO_DETALHADO_CCREDITO Then

            LabelEmCaixaCCreditoValor.Caption = CStr(Format(objAdmMeioPagtoCondPagto.dSaldo, "standard"))
            ValorSangriaCCredito.Text = IIf(objAdmMeioPagtoCondPagto.dSaldo <> 0, CStr(Format(objAdmMeioPagtoCondPagto.dSaldo, "standard")), "")

        End If

    Next


    Carrega_MovimentoBoleto_DebitoCredito = SUCESSO

    Exit Function

Erro_Carrega_MovimentoBoleto_DebitoCredito:

    Carrega_MovimentoBoleto_DebitoCredito = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162945)

    End Select

    Exit Function

End Function

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

Private Sub LabelCodigo_Click()

Dim objMovimentoCaixa As New ClassMovimentoCaixa
    
   'Chama tela de MovimentoBoletoLista
    Call Chama_TelaECF_Modal("MovimentoBoletoLista", objMovimentoCaixa)
    
    If Not (objMovimentoCaixa Is Nothing) Then
        'Verifica se o Codvendedor está preenchido e joga na coleção
        If objMovimentoCaixa.lNumMovto <> 0 Then
            Codigo.Text = objMovimentoCaixa.lNumMovto
            Call CodMovimentoBoleto_Validate(False)
            
        End If
    End If
    
    Exit Sub

End Sub

Sub CodMovimentoBoleto_Validate(Cancel As Boolean)
'Função que Verifica se o Código Passado como parâmetro Existe na Coleção Globa de MOVTOCAIXA

Dim lErro As Long
Dim objMovimentoCaixa As New ClassMovimentoCaixa

On Error GoTo Erro_CodMovimentoBoleto_Validate

    'Verifica se existe movimento com o código passado
    For Each objMovimentoCaixa In gcolMovimentosCaixa
    
        If objMovimentoCaixa.lNumMovto = StrParaLong(Codigo.Text) Then
            
            'Função que traz o MovimentoBoleto para a Tela
            lErro = Traz_MovimentoBoleto_Tela(StrParaLong(Codigo.Text))
            If lErro <> SUCESSO Then gError 107819
        
            Exit For
        
        End If
        
    
    Next
    
    'Anula a Alteração
    iAlterado = 0
   
    Exit Sub
    
Erro_CodMovimentoBoleto_Validate:
    
    Select Case gErr

        Case 107819

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162946)

    End Select

    Exit Sub

End Sub

Private Sub DesmembraMovto_Click()
'Função que Desmembra Log

Dim lErro As Long
Dim colMovimentosCaixa As New Collection
Dim colImfCompl As New Collection
Dim objMovimentosCaixa As New ClassMovimentoCaixa

On Error GoTo Erro_DesmembraMovto_Click

    lErro = CF_ECF("Desmembra_MovimentosCaixa", colMovimentosCaixa, colImfCompl, TIPOREGISTROECF_MOVIMENTOCAIXA_BOLETO)
    If lErro <> SUCESSO Then gError 111067
    
    For Each objMovimentosCaixa In colMovimentosCaixa
        If objMovimentosCaixa.iTipo <> MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_B0LETO Then
            gcolMovimentosCaixa.Add colMovimentosCaixa
        End If
    Next
        
    Exit Sub
    

Erro_DesmembraMovto_Click:
    
    Select Case gErr

        Case 111067

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162947)

    End Select

    Exit Sub


End Sub



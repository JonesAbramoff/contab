VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl MovimentoCheque 
   ClientHeight    =   6495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9705
   KeyPreview      =   -1  'True
   ScaleHeight     =   6495
   ScaleWidth      =   9705
   Begin VB.Frame Frame1 
      Caption         =   "Identificação"
      Height          =   840
      Left            =   150
      TabIndex        =   38
      Top             =   45
      Width           =   4185
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
         Left            =   2385
         Picture         =   "MovimentoCheque.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "F4 - Exibe na tela o movimento com o código informado."
         Top             =   195
         Width           =   1560
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   1890
         Picture         =   "MovimentoCheque.ctx":2CCA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "F2 - Numeração Automática"
         Top             =   360
         Width           =   300
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   1035
         TabIndex        =   2
         Top             =   345
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
         Caption         =   "&Código:"
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
         Left            =   345
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   1
         Top             =   390
         Width           =   660
      End
   End
   Begin VB.CommandButton BotaoDesmarcarTodos 
      Caption         =   "Desmarcar Todos"
      Height          =   780
      Left            =   5985
      Picture         =   "MovimentoCheque.ctx":2DB4
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "F10 - Desmarcar todos os cheques"
      Top             =   120
      Width           =   1320
   End
   Begin VB.CommandButton BotaoMarcarTodos 
      Caption         =   "Marcar Todos"
      Height          =   780
      Left            =   4545
      Picture         =   "MovimentoCheque.ctx":3F96
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "F9 - Marcar todos os cheques"
      Top             =   120
      Width           =   1320
   End
   Begin VB.PictureBox Picture2 
      Height          =   555
      Left            =   7470
      ScaleHeight     =   495
      ScaleWidth      =   2055
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   150
      Width           =   2115
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   105
         Picture         =   "MovimentoCheque.ctx":4FB0
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "F5 - Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   600
         Picture         =   "MovimentoCheque.ctx":510A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "F6 - Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1080
         Picture         =   "MovimentoCheque.ctx":5294
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "F7 - Limpar"
         Top             =   105
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1560
         Picture         =   "MovimentoCheque.ctx":57C6
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "F8 - Fechar"
         Top             =   105
         Width           =   420
      End
   End
   Begin VB.Frame FrameCheques 
      Caption         =   "Cheques"
      Height          =   3930
      Left            =   120
      TabIndex        =   11
      Top             =   885
      Width           =   9480
      Begin VB.TextBox SequencialCaixa 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   420
         TabIndex        =   39
         Top             =   645
         Width           =   585
      End
      Begin VB.TextBox Tipo 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   7470
         TabIndex        =   21
         Top             =   270
         Width           =   1200
      End
      Begin VB.TextBox Cliente 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   6120
         TabIndex        =   20
         Top             =   270
         Width           =   1275
      End
      Begin VB.TextBox NumeroCheque 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   825
         TabIndex        =   19
         Top             =   300
         Width           =   900
      End
      Begin VB.TextBox Conta 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   1920
         TabIndex        =   18
         Top             =   345
         Width           =   765
      End
      Begin VB.TextBox Agencia 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   3330
         TabIndex        =   17
         Top             =   285
         Width           =   690
      End
      Begin VB.TextBox Banco 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   2745
         TabIndex        =   16
         Top             =   330
         Width           =   525
      End
      Begin VB.CheckBox Selecionar 
         Height          =   195
         Left            =   345
         TabIndex        =   13
         Top             =   360
         Width           =   420
      End
      Begin MSMask.MaskEdBox DataDeposito 
         Height          =   240
         Left            =   4065
         TabIndex        =   14
         Top             =   270
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   240
         Left            =   5100
         TabIndex        =   15
         Top             =   240
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   423
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
      Begin MSFlexGridLib.MSFlexGrid GridCheques 
         Height          =   3225
         Left            =   75
         TabIndex        =   7
         Top             =   300
         Width           =   9180
         _ExtentX        =   16193
         _ExtentY        =   5689
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
   End
   Begin VB.Frame FrameTotais 
      Caption         =   "Totais"
      Height          =   1500
      Left            =   120
      TabIndex        =   0
      Top             =   4905
      Width           =   9480
      Begin VB.Frame FrameTotaisDetalhadosNaoDetalhados 
         Caption         =   "Detalhados + Não Detalhados"
         Height          =   1155
         Left            =   6345
         TabIndex        =   32
         Top             =   225
         Width           =   2775
         Begin VB.Label LabelTotalDetalhadoNaoDetalhado 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
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
            Left            =   735
            TabIndex        =   36
            Top             =   390
            Width           =   510
         End
         Begin VB.Label LabelTotalDetalhadoNaoDetalhadoValor 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1350
            TabIndex        =   35
            Top             =   345
            Width           =   1215
         End
         Begin VB.Label LabelSelecionadoDetalhadoNaoDetalhado 
            AutoSize        =   -1  'True
            Caption         =   "Selecionado:"
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
            TabIndex        =   34
            Top             =   795
            Width           =   1125
         End
         Begin VB.Label LabelSelecionadoDetalhadoNaoDetalhadoValor 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1350
            TabIndex        =   33
            Top             =   735
            Width           =   1215
         End
      End
      Begin VB.Frame FrameTotaisNaoDetalhados 
         Caption         =   "Cheques Não Detalhados"
         Height          =   1155
         Left            =   3225
         TabIndex        =   27
         Top             =   225
         Width           =   2775
         Begin VB.Label LabelNaoDetalhadoSelecionadoValor 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1350
            TabIndex        =   31
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label LabelSelecionadoNaoDetalhado 
            AutoSize        =   -1  'True
            Caption         =   "Selecionado:"
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
            TabIndex        =   30
            Top             =   750
            Width           =   1125
         End
         Begin VB.Label LabelTotalNaoDetalhadoValor 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1350
            TabIndex        =   29
            Top             =   330
            Width           =   1215
         End
         Begin VB.Label LabelTotalNaoDetalhado 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
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
            Left            =   735
            TabIndex        =   28
            Top             =   405
            Width           =   510
         End
      End
      Begin VB.Frame FrameTotaisDetalhados 
         Caption         =   "Cheques Detalhados"
         Height          =   1155
         Left            =   120
         TabIndex        =   22
         Top             =   225
         Width           =   2775
         Begin VB.Label LabelTotalDetalhado 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
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
            Left            =   735
            TabIndex        =   26
            Top             =   390
            Width           =   510
         End
         Begin VB.Label LabelTotalDetalhadoValor 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1335
            TabIndex        =   25
            Top             =   345
            Width           =   1215
         End
         Begin VB.Label LabelSelecionadoDetalhado 
            AutoSize        =   -1  'True
            Caption         =   "Selecionado:"
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
            TabIndex        =   24
            Top             =   765
            Width           =   1125
         End
         Begin VB.Label LabelSelecionadoDetalhadoValor 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1335
            TabIndex        =   23
            Top             =   735
            Width           =   1215
         End
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
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   926
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"MovimentoCheque.ctx":5944
   End
End
Attribute VB_Name = "MovimentoCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 
'Property Variables:

'Variáveis Globais à Tela
Dim objGridCheques As AdmGrid
Dim colCheques As New Collection

'Constantes
Const NUM_MAXIMO_LINHAS_GRID = 1000

Dim gcolImfCompl As New Collection

'Variáveis Relacionadas ao Grig

Dim iGrid_Selecionado_Col As Integer
Dim iGrid_Seq_Col As Integer
Dim iGrid_Banco_Col As Integer
Dim iGrid_Agencia_Col As Integer
Dim iGrid_Conta_Col As Integer
Dim iGrid_NumeroCheque_Col As Integer
Dim iGrid_DataDeposito_Col As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_Cliente_Col As Integer
Dim iGrid_Tipo_Col As Integer

'Varialvel Global que é Usada para guardar qual é proximo numero de movto de caixa
Dim glProxNumAuto As Long
Dim iAlterado As Integer

Dim m_Caption As String
Event Unload()

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Sangria Cheque"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "MovimentoCheque"

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

Private Sub BotaoTrazer_Click()
'Função que chama a função que preenche o grid

Dim lErro As Long
Dim lCodigo As String

On Error GoTo Erro_botaoTrazer_click

    'Verifica se o código não está preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 105732

    'Função que traz o MovimentoBoleto para a Tela
    lErro = Traz_MovimentoCheque_Tela(StrParaLong(Codigo.Text))
    If lErro <> SUCESSO Then gError 105736

    'Anula a Alteração
    iAlterado = 0

    Exit Sub

Erro_botaoTrazer_click:

    Select Case gErr

        Case 105732
            Call Rotina_ErroECF(vbOKOnly, ERRO_CODIGO_NAO_PREENCHIDO1, gErr)

        Case 105736
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162951)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_GotFocus()

    Call MaskEdBox_TrataGotFocus(Codigo)

End Sub

Private Sub GridCheques_Scroll()

    Call Grid_Scroll(objGridCheques)

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

'**** Tela Iniciada por Sergio Dia 5/08/02 ****'
Public Sub Form_Load()
'Função inicialização da Tela
Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Instanciar o objGridCheques para apontar para uma posição de memória
    Set objGridCheques = New AdmGrid
    
    'Inicialização de Grid Cheques
    lErro = Inicializa_GridCheques(objGridCheques)
    If lErro <> SUCESSO Then gError 107704

    'Função que Carrega os Cheques no GridCheques
    lErro = Carrega_GridCheques()
    If lErro <> SUCESSO Then gError 107705

    'Função que exibe o valor total dos Cheques especificados e não especificados
    lErro = Recalcula_Totais()
    If lErro <> SUCESSO Then gError 107706
    
    'Indica que o carregamento da tela aconteceu com sucesso
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        'Erros tratados na rotina chamada
        Case 107704 To 107706
        
        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162952)

    End Select

    Exit Sub

End Sub

Function Inicializa_GridCheques(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Inicializa_GridCheques

   'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Sel.")
    objGridInt.colColuna.Add ("Seq.")
    objGridInt.colColuna.Add ("Banco")
    objGridInt.colColuna.Add ("Agência")
    objGridInt.colColuna.Add ("Conta")
    objGridInt.colColuna.Add ("Cheque")
    objGridInt.colColuna.Add ("Depósito")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Tipo")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Selecionar.Name)
    objGridInt.colCampo.Add (SequencialCaixa.Name)
    objGridInt.colCampo.Add (Banco.Name)
    objGridInt.colCampo.Add (Agencia.Name)
    objGridInt.colCampo.Add (Conta.Name)
    objGridInt.colCampo.Add (NumeroCheque.Name)
    objGridInt.colCampo.Add (DataDeposito.Name)
    objGridInt.colCampo.Add (Valor.Name)
    objGridInt.colCampo.Add (Cliente.Name)
    objGridInt.colCampo.Add (Tipo.Name)
    
    'Colunas do Grid
    iGrid_Selecionado_Col = 1
    iGrid_Seq_Col = 2
    iGrid_Banco_Col = 3
    iGrid_Agencia_Col = 4
    iGrid_Conta_Col = 5
    iGrid_NumeroCheque_Col = 6
    iGrid_DataDeposito_Col = 7
    iGrid_Valor_Col = 8
    iGrid_Cliente_Col = 9
    iGrid_Tipo_Col = 10
    
    'Grid do GridInterno
    objGridInt.objGrid = GridCheques

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_LINHAS_GRID

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 12

    'Largura da primeira coluna
    GridCheques.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)
    
    Inicializa_GridCheques = SUCESSO

    Exit Function

Erro_Inicializa_GridCheques:

    Inicializa_GridCheques = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162953)

    End Select

    Exit Function

End Function

Function Carrega_GridCheques() As Long
'Função que Carrega os Cheques no Grid

Dim lErro As Long
Dim objCheque As New ClassChequePre
Dim iIndice As Integer
Dim iCont As Integer

On Error GoTo Erro_Carrega_GridCheques

    'Limpa Grid
    Call Grid_Limpa(objGridCheques)

    'Verifica na Coleção Global de Cheques se o Cheque é Especificado ou não
    For iIndice = gcolCheque.Count To 1 Step -1
        
        Set objCheque = gcolCheque.Item(iIndice)
        
        If objCheque.lNumMovtoSangria = 0 And objCheque.iStatus <> STATUS_EXCLUIDO Then
            
            iCont = iCont + 1
            
            If objCheque.iNaoEspecificado = CHEQUE_ESPECIFICADO Then
            
                GridCheques.TextMatrix(iCont, iGrid_Selecionado_Col) = objCheque.iChequeSel
                GridCheques.TextMatrix(iCont, iGrid_Seq_Col) = objCheque.lSequencialCaixa
                GridCheques.TextMatrix(iCont, iGrid_Banco_Col) = objCheque.iBanco
                GridCheques.TextMatrix(iCont, iGrid_Agencia_Col) = objCheque.sAgencia
                GridCheques.TextMatrix(iCont, iGrid_Conta_Col) = objCheque.sContaCorrente
                GridCheques.TextMatrix(iCont, iGrid_NumeroCheque_Col) = objCheque.lNumero
                GridCheques.TextMatrix(iCont, iGrid_DataDeposito_Col) = Format(objCheque.dtDataDeposito, "dd/mm/yyyy")
                GridCheques.TextMatrix(iCont, iGrid_Valor_Col) = Format(objCheque.dValor, "Standard")
                GridCheques.TextMatrix(iCont, iGrid_Cliente_Col) = objCheque.sCPFCGC
                GridCheques.TextMatrix(iCont, iGrid_Tipo_Col) = STRING_ESPECIFICADO
        
            Else
                            
                GridCheques.TextMatrix(iCont, iGrid_Seq_Col) = objCheque.lSequencialCaixa
                GridCheques.TextMatrix(iCont, iGrid_Tipo_Col) = STRING_NAO_ESPECIFICADO
                GridCheques.TextMatrix(iCont, iGrid_DataDeposito_Col) = Format(objCheque.dtDataDeposito, "dd/mm/yyyy")
                GridCheques.TextMatrix(iCont, iGrid_Valor_Col) = Format(objCheque.dValor, "Standard")
            
            End If
        
            If iCont - 1 = NUM_MAXIMO_LINHAS_GRID Then
                Call Rotina_AvisoECF(vbOK, AVISO_NUM_CHEQUES_MAIOR_NUM_MAX_GRID, NUM_MAXIMO_LINHAS_GRID, gcolCheque.Count)
                Exit For
            End If
        
        End If
    
    Next
'
'    'Faz com que a Coleção Global de Cheques aponte para a Coleção Local a Tela
'    Set gcolCheque = colCheques
    
    'Atualiza o numero de Linhas no Grig
    objGridCheques.iLinhasExistentes = iCont
    
    Carrega_GridCheques = SUCESSO
    
    Exit Function

Erro_Carrega_GridCheques:

    Carrega_GridCheques = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162954)

    End Select

    Exit Function

End Function

Function Recalcula_Totais() As Long
'Calcula o Total de Cheques não Especificado e o Totas de Cheques Especificados

Dim lErro As Long
Dim iIndice As Integer
Dim dValorEspecificados As Double
Dim dValorNaoEspecificados As Double
Dim dValorTotalCheques As Double
Dim dValorNaoEspecificadosSel As Double
Dim dValorEspecificadosSel As Double

On Error GoTo Erro_Recalcula_Totais

    'Para todos os Cheque do Grid
    For iIndice = 1 To objGridCheques.iLinhasExistentes
    
        If GridCheques.TextMatrix(iIndice, iGrid_Tipo_Col) = STRING_ESPECIFICADO Then
        
            'Guarda o Total de Cheques  especificados
            dValorEspecificados = dValorEspecificados + StrParaDbl(GridCheques.TextMatrix(iIndice, iGrid_Valor_Col))
            
            'Verifica se os Cheques estão selecionadas
            If GridCheques.TextMatrix(iIndice, iGrid_Selecionado_Col) = MARCADO Then
            
                'Acumula a Valor do Cheque Especificado na Varialvel para mais tarde jogar na Tela
                dValorEspecificadosSel = dValorEspecificadosSel + StrParaDbl(GridCheques.TextMatrix(iIndice, iGrid_Valor_Col))
                            
             End If
             
        ElseIf GridCheques.TextMatrix(iIndice, iGrid_Tipo_Col) = STRING_NAO_ESPECIFICADO Then
            
            'Guarda o Total de Cheques não especificados
            dValorNaoEspecificados = dValorNaoEspecificados + StrParaDbl(GridCheques.TextMatrix(iIndice, iGrid_Valor_Col))
            
            'Verifica se os Cheques estão selecionadas
            If GridCheques.TextMatrix(iIndice, iGrid_Selecionado_Col) = MARCADO Then
            
                'Acumula a Valor do Cheque Não Especificado na Varialvel para mais tarde jogar na Tela
                dValorNaoEspecificadosSel = dValorNaoEspecificadosSel + StrParaDbl(GridCheques.TextMatrix(iIndice, iGrid_Valor_Col))
                            
             End If
                
        End If
            
            'Acumula o Valor Total dos Cheques
            dValorTotalCheques = dValorTotalCheques + StrParaDbl(GridCheques.TextMatrix(iIndice, iGrid_Valor_Col))
    Next
            
    'Exibe na Tela o Valor Total de Cheques Especificado e o valor de Cheques Selecionados especificados
    LabelTotalDetalhadoValor.Caption = Format(dValorEspecificados, "Standard")
    LabelSelecionadoDetalhadoValor.Caption = Format(dValorEspecificadosSel, "Standard")
    
            
    'Exibe na Tela o Valor Total de Cheques não Especificado e o valor de Cheques Selecionados não especificados
    LabelTotalNaoDetalhadoValor.Caption = Format(dValorNaoEspecificados, "Standard")
    LabelNaoDetalhadoSelecionadoValor.Caption = Format(dValorNaoEspecificadosSel, "Standard")
    
    'Exibe na Tela o Numero de Cheque Não Especificado mais dos Especicados e o valor acumulado
    LabelTotalDetalhadoNaoDetalhadoValor.Caption = Format(dValorTotalCheques, "Standard")
    LabelSelecionadoDetalhadoNaoDetalhadoValor.Caption = Format(dValorEspecificadosSel + dValorNaoEspecificadosSel, "Standard")
    
    'adciona a Variavel Global quanto em Valor existe em Cheques
    gdSaldocheques = dValorTotalCheques
    
    Recalcula_Totais = SUCESSO

    Exit Function

Erro_Recalcula_Totais:

    Recalcula_Totais = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162955)

    End Select

    Exit Function

End Function

Private Sub GridCheques_Click()

    Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridCheques, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        'Variavel não definida
        Call Grid_Entrada_Celula(objGridCheques, iAlterado)
    End If

End Sub

Private Sub GridCheques_EnterCell()
    'Parametro não opcional
    Call Grid_Entrada_Celula(objGridCheques, iAlterado)

End Sub

Private Sub GridCheques_GotFocus()

    Call Grid_Recebe_Foco(objGridCheques)

End Sub

Private Sub GridCheques_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call Grid_Trata_Tecla1(KeyCode, objGridCheques)

End Sub

Private Sub GridCheques_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridCheques, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCheques, iAlterado)
    End If

End Sub

Private Sub GridCheques_LeaveCell()

    Call Saida_Celula(objGridCheques)

End Sub

Private Sub GridCheques_LostFocus()

    Call Grid_Libera_Foco(objGridCheques)

End Sub
Private Sub GridCheques_RowColChange()

    Call Grid_RowColChange(objGridCheques)

End Sub

Private Sub GridCheques_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridCheques)

End Sub

Private Sub Banco_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Banco_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCheques)


End Sub

Private Sub Banco_KeyPress(KeyAscii As Integer)


    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCheques)

End Sub

Private Sub Banco_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCheques.objControle = Banco
    lErro = Grid_Campo_Libera_Foco(objGridCheques)
    
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Agencia_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Agencia_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCheques)


End Sub

Private Sub Agencia_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCheques)

End Sub

Private Sub Agencia_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCheques.objControle = Agencia
    lErro = Grid_Campo_Libera_Foco(objGridCheques)
    
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Conta_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Conta_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCheques)

End Sub

Private Sub Conta_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCheques)

End Sub

Private Sub Conta_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCheques.objControle = Conta
    lErro = Grid_Campo_Libera_Foco(objGridCheques)
    
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub NumeroCheque_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NumeroCheque_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCheques)


End Sub

Private Sub NumeroCheque_KeyPress(KeyAscii As Integer)


    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCheques)

End Sub

Private Sub NumeroCheque_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCheques.objControle = NumeroCheque
    lErro = Grid_Campo_Libera_Foco(objGridCheques)
    
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DataDeposito_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataDeposito_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCheques)


End Sub

Private Sub DataDeposito_KeyPress(KeyAscii As Integer)


    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCheques)

End Sub

Private Sub DataDeposito_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCheques.objControle = DataDeposito
    lErro = Grid_Campo_Libera_Foco(objGridCheques)
    
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Valor_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Valor_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCheques)

End Sub

Private Sub Valor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCheques)

End Sub

Private Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCheques.objControle = Valor
    lErro = Grid_Campo_Libera_Foco(objGridCheques)
    
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Cliente_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Cliente_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCheques)


End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)


    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCheques)

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCheques.objControle = Cliente
    lErro = Grid_Campo_Libera_Foco(objGridCheques)
    
    If lErro <> SUCESSO Then Cancel = True

End Sub


Private Sub Tipo_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Tipo_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCheques)


End Sub

Private Sub Tipo_KeyPress(KeyAscii As Integer)


    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCheques)

End Sub

Private Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCheques.objControle = Tipo
    lErro = Grid_Campo_Libera_Foco(objGridCheques)
    
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Selecionar_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Selecionar_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCheques)


End Sub

Private Sub Selecionar_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCheques)

End Sub

Private Sub Selecionar_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCheques.objControle = Selecionar
    lErro = Grid_Campo_Libera_Foco(objGridCheques)
    
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Selecionar_Click()

Dim lErro As Long

On Error GoTo Erro_Selecionar_Click

    lErro = Recalcula_Totais()
    If lErro <> SUCESSO Then gError 107707

    Exit Sub
    
Erro_Selecionar_Click:

    Select Case gErr

        Case 107707
            'Erro Tratado Dentro da Função
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162956)

    End Select

    Exit Sub

End Sub

Private Sub BotaoMarcarTodos_Click()
'Função que Seleciona Todos os Cheques

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_BotaoMarcarTodos_Click

    'Seleciona Todos os Cheques
    For iIndice = 1 To objGridCheques.iLinhasExistentes
       
       GridCheques.TextMatrix(iIndice, iGrid_Selecionado_Col) = MARCADO
    
    Next
    
    'dá um refresh nas checkboxes
    lErro = Grid_Refresh_Checkbox(objGridCheques)
    If lErro <> SUCESSO Then gError 107708

    'Função Que Recalcula os Totalizadores
    lErro = Recalcula_Totais()
    If lErro <> SUCESSO Then gError 107709
     
    Exit Sub

Erro_BotaoMarcarTodos_Click:

    Select Case gErr

        Case 107708, 107709
            'Erro Tratado Dentro da Função Chamada
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162957)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoDesmarcarTodos_Click()
'Botão que Desmarca Todos os Cheques

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_BotaoDesmarcarTodos_Click

    'varre o grid de cheques
    For iIndice = 1 To objGridCheques.iLinhasExistentes

        'desmarca cada um
        GridCheques.TextMatrix(iIndice, iGrid_Selecionado_Col) = DESMARCADO

    Next

    'dá um refresh nas checkboxes
    lErro = Grid_Refresh_Checkbox(objGridCheques)
    If lErro <> SUCESSO Then gError 107710

    'Função Que Recalcula os Totalizadores
    lErro = Recalcula_Totais()
    If lErro <> SUCESSO Then gError 107711

    Exit Sub
    
Erro_BotaoDesmarcarTodos_Click:
    
    Select Case gErr

        Case 107710, 107711
            'Erro Tratado Dentro da Função Chamada
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162958)

    End Select

    Exit Sub
    
End Sub


Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 107715

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 107715
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162959)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()
    
    'Fechar a Tela
    Unload Me

End Sub


Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

Function Desmarcar_Grid() As Long
'Função que Serve para Desmarcar as Opcções marcadas no Grid

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Desmarcar_Grid
    
    
    For iIndice = 1 To objGridCheques.iLinhasExistentes
    
        GridCheques.TextMatrix(iIndice, iGrid_Selecionado_Col) = DESMARCADO
    
    Next
    
    'dá um refresh nas checkboxes
    lErro = Grid_Refresh_Checkbox(objGridCheques)
    If lErro <> SUCESSO Then gError 107718

    'Função que Carrega o Grid
    lErro = Carrega_GridCheques
    If lErro <> SUCESSO Then gError 107904

    'Função Que Recalcula os Totalizadores
    lErro = Recalcula_Totais()
    If lErro <> SUCESSO Then gError 107719

    Desmarcar_Grid = SUCESSO
    
    Exit Function

Erro_Desmarcar_Grid:

    Select Case gErr
        
        Case 107718, 107719, 107904
            'Erro Tratado Dentro da Função Chamada
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162960)

    End Select

    Exit Function
    
End Function

Private Sub BotaoProxNum_Click()
Dim lNumero As Long
Dim objCheque As New ClassChequePre
Dim iNumChequesDisponiveis As Integer
Dim lErro As Long

On Error GoTo Erro_BotaoProxNum_Click

'    'Verifica para Cada Cheque na Coleção de Globais de Cheque
'    For Each objCheque In gcolCheque
'
'        'Verifica se o Numero que identifica qual foi o Movto de Caixa que efetuou a Sangria esta em Branco
'        If objCheque.lNumMovtoSangria = 0 Then
'
'            iNumChequesDisponiveis = iNumChequesDisponiveis + 1
'
'        End If
'
'    Next
'
'    'Verifica se já existem mais Cheques para realizar a Sangria do que os que estão no Grig
'    If iNumChequesDisponiveis > objGridCheques.iLinhasExistentes Then
'
'        'Recarega-se o grid com o  numero de Cheque disponíveis para a sangria
'        lErro = Carrega_GridCheques()
'        If lErro <> SUCESSO Then gError 107861
'
'    End If
    
    'Função que Gera o Proximo Numero
    lErro = CF_ECF("Caixa_Obtem_NumAutomatico", lNumero)
    If lErro <> SUCESSO Then gError 107862
    
    'Exibir o numero na tela
    Codigo.Text = lNumero
    
    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 107861, 107862

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162961)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()
'Botão que Exclui um Movimento de Sangria

Dim lErro As Long
Dim colMovcxExclui As New Collection
Dim vbMsgRes As VbMsgBoxResult
Dim lNumero As Long
Dim iTipoMovimento As Integer
Dim colChequeExclui As New Collection
Dim objChequePre As New ClassChequePre
Dim lSequencialLoja As Long
Dim colInfoChequesExcluir As New Collection
Dim colInfoChequesExcluirEsp As New Collection
Dim colColInfoCheques As New Collection
Dim objMovCaixa As ClassMovimentoCaixa

On Error GoTo Erro_BotaoExcluir_Click

    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 207974

    'Verifica se o Codigo não foi Preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 107888

    lNumero = StrParaLong(Codigo.Text)

    'Verifica os Movimentos de Caixa para o Código em Questão
    lErro = CF_ECF("Caixa_MovimentoCaixa_Le_NumMovto", colMovcxExclui, lNumero)
    If lErro <> SUCESSO And lErro <> 107850 Then gError 107889

    'Senão encontrou ninguem
    If lErro = 107850 Then gError 107890

    Set objMovCaixa = colMovcxExclui(1)

    If objMovCaixa.iTipo <> MOVIMENTOCAIXA_SANGRIA_CHEQUE Then gError 86287
    
    'Lê os movimentos que serão carregados
    lErro = Caixa_ChequePre_Le_NumMovtoSangria(colChequeExclui, lNumero)
    If lErro <> SUCESSO Then gError 107907

    'Pergunta se deseja Realmente Excluir o Movimento
    vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_EXCLUSAO_MOVIMENTOCAIXA, Codigo.Text)
    If vbMsgRes = vbNo Then gError 107891

    iTipoMovimento = MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_CHEQUE

    'Prepara os Movimentos para a Exclusão
    lErro = MovimentoCaixa_Prepara_Exclusao(colMovcxExclui, iTipoMovimento)
    If lErro <> SUCESSO Then gError 107892
    
    'Adciona o SequencialLoja à Coleção de Cheques que terão a sangria excluida
    For Each objChequePre In colChequeExclui
    
        'Atribui o Valor que esta no objCheque para a variável
        lSequencialLoja = objChequePre.lSequencialCaixa
        
        If objChequePre.iNaoEspecificado = CHEQUE_ESPECIFICADO Then
            'Atribui o valor a colecão de cheques que serão excluidos
            colInfoChequesExcluirEsp.Add lSequencialLoja
        Else
            'Atribui o valor a colecão de cheques que serão excluidos
            colInfoChequesExcluir.Add lSequencialLoja
        End If
    Next
    
    'Função que Grava a Exclusão de Boletos
    lErro = Caixa_Grava_Movimento(colMovcxExclui, colInfoChequesExcluir, colInfoChequesExcluirEsp)
    If lErro <> SUCESSO Then gError 107893

    'Atualiza os Dados na Memória
    lErro = MovimentoCheque_Exclui_Memoria(colMovcxExclui, colChequeExclui)
    If lErro <> SUCESSO Then gError 107894
        
    'Função Que Limpa a Tela
    Call Limpa_Tela_MovimentoCheque

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 86287
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_NAO_SANGRIA_CHEQUE, gErr, Codigo.Text)
            
        Case 86290
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_JA_TRANSMITIDO, gErr)
        
        Case 107888
            Call Rotina_ErroECF(vbOKOnly, ERRO_CODIGO_NAO_PREENCHIDO1, gErr)

        Case 107890
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_INEXISTENTE, gErr, lNumero)

        Case 107889, 107891, 107892, 107893, 107894, 107907, 207974

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162962)

    End Select

    Exit Sub

End Sub

Sub CodMovimentoCheque_Validate(Cancel As Boolean)
'Função que Verifica se o Código Passado como parâmetro Existe na Coleção Globa de MOVTOCAIXA

Dim lErro As Long
Dim objMovimentoCaixa As New ClassMovimentoCaixa

On Error GoTo Erro_CodMovimentoCheque_Validate

    
    Exit Sub
    
Erro_CodMovimentoCheque_Validate:
    
    Select Case gErr

        Case 105733

        Case 107864

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162963)

    End Select

    Exit Sub

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
            If Not TrocaFoco(Me, BotaoMarcarTodos) Then Exit Sub
            Call BotaoMarcarTodos_Click
            
        Case vbKeyF10
            If Not TrocaFoco(Me, BotaoDesmarcarTodos) Then Exit Sub
            Call BotaoDesmarcarTodos_Click
            
    End Select

    Exit Sub

Erro_UserControl_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162964)

    End Select

    Exit Sub

End Sub



Private Sub LabelCodigo_Click()

Dim objMovimentoCaixa As New ClassMovimentoCaixa
Dim lErro As Long
    
On Error GoTo Erro_LabelCodigo_Click
    
    'Chama tela de MovimentoBoletoLista
    Call Chama_TelaECF_Modal("MovimentoChequeLista", objMovimentoCaixa)
    
    If Not (objMovimentoCaixa Is Nothing) Then
        'Verifica se o Codvendedor está preenchido e joga na coleção
        If objMovimentoCaixa.lNumMovto <> 0 Then
            
            Codigo.Text = objMovimentoCaixa.lNumMovto
            
            'Função que traz o MovimentoBoleto para a Tela
            lErro = Traz_MovimentoCheque_Tela(StrParaLong(Codigo.Text))
            If lErro <> SUCESSO Then gError 105737
            
        End If
    End If
    
    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr
    
        Case 105737
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162965)
    
    End Select

    Exit Sub

End Sub

Function Traz_MovimentoCheque_Tela(lNumero As Long) As Long
'Função que Traz o Movto de Dinheiro para a Tela

Dim lErro As Long
Dim colMovimentosCaixa As New Collection
Dim objMovimentoCaixa As New ClassMovimentoCaixa

On Error GoTo Erro_Traz_MovimentoCheque_Tela

    'Função que Lê os Movimentos de Caixa
    lErro = CF_ECF("Caixa_MovimentoCaixa_Le_NumMovto", colMovimentosCaixa, lNumero)
    If lErro <> SUCESSO And lErro <> 107850 Then gError 105734

    'se o movimento de caixa nao foi encontrado ==> erro
    If lErro = 107850 Then gError 105735

    'se o tipo do movimento de caixa nao for sangria de cheque ==> erro
    If colMovimentosCaixa.Item(1).iTipo <> MOVIMENTOCAIXA_SANGRIA_CHEQUE Then gError 105740

    'Limpa Grid
    Call Grid_Limpa(objGridCheques)
    
    'Função que Lê os Cheques que foram sangrados e carrega no Grid
    lErro = Traz_MovimentoCheque_Tela_NumMovto(lNumero)
    If lErro <> SUCESSO Then gError 107866
    
    'Carrega o Grid Recalcula Totais
    lErro = Recalcula_Totais()
    If lErro <> SUCESSO Then gError 107868

    Traz_MovimentoCheque_Tela = SUCESSO
    
    Exit Function

Erro_Traz_MovimentoCheque_Tela:

    Traz_MovimentoCheque_Tela = gErr
    
    Select Case gErr

        Case 105735
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_INEXISTENTE, gErr, lNumero)

        Case 105740
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_NAO_SANGRIA_CHEQUE, gErr, lNumero)

        Case 107866

        Case 107868

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162966)

    End Select

    Exit Function

End Function

Function Traz_MovimentoCheque_Tela_NumMovto(lNumero As Long) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objCheque As New ClassChequePre
Dim colChequePre As New Collection

On Error GoTo Erro_Traz_MovimentoCheque_Tela_NumMovto

    'Lê os movimentos que serão carregados
    lErro = Caixa_ChequePre_Le_NumMovtoSangria(colChequePre, lNumero)
    If lErro <> SUCESSO Then gError 107869

    'Preencher o GridCheque com os Dados dos Cheques Selecionados com o Numero do Movimento passado como parâmetro
    For Each objCheque In colChequePre
        
        iIndice = iIndice + 1
        
        'Guarda na Coleção o Sequencial Loja relacionado ao Cheque da Lina do Grid
        
        If objCheque.iNaoEspecificado = CHEQUE_ESPECIFICADO Then
            
            GridCheques.TextMatrix(iIndice, iGrid_Selecionado_Col) = MARCADO
            GridCheques.TextMatrix(iIndice, iGrid_Seq_Col) = objCheque.lSequencialCaixa
            GridCheques.TextMatrix(iIndice, iGrid_Banco_Col) = objCheque.iBanco
            GridCheques.TextMatrix(iIndice, iGrid_Agencia_Col) = objCheque.sAgencia
            GridCheques.TextMatrix(iIndice, iGrid_Conta_Col) = objCheque.sContaCorrente
            GridCheques.TextMatrix(iIndice, iGrid_NumeroCheque_Col) = objCheque.lNumero
            GridCheques.TextMatrix(iIndice, iGrid_DataDeposito_Col) = Format(objCheque.dtDataDeposito, "dd/mm/yyyy")
            GridCheques.TextMatrix(iIndice, iGrid_Valor_Col) = Format(objCheque.dValor, "Standard")
            GridCheques.TextMatrix(iIndice, iGrid_Cliente_Col) = objCheque.sCPFCGC
            GridCheques.TextMatrix(iIndice, iGrid_Tipo_Col) = STRING_ESPECIFICADO
            
        Else
            GridCheques.TextMatrix(iIndice, iGrid_Selecionado_Col) = MARCADO
            GridCheques.TextMatrix(iIndice, iGrid_Seq_Col) = objCheque.lSequencialCaixa
            GridCheques.TextMatrix(iIndice, iGrid_Valor_Col) = Format(objCheque.dValor, "Standard")
            GridCheques.TextMatrix(iIndice, iGrid_Tipo_Col) = STRING_NAO_ESPECIFICADO
            GridCheques.TextMatrix(iIndice, iGrid_DataDeposito_Col) = Format(objCheque.dtDataDeposito, "dd/mm/yyyy")
        
        End If
        
    Next
    
    objGridCheques.iLinhasExistentes = iIndice
    
    'dá um refresh nas checkboxes
    lErro = Grid_Refresh_Checkbox(objGridCheques)
    If lErro <> SUCESSO Then gError 107905
    
    
    Traz_MovimentoCheque_Tela_NumMovto = SUCESSO

    Exit Function
    
Erro_Traz_MovimentoCheque_Tela_NumMovto:

    Traz_MovimentoCheque_Tela_NumMovto = gErr
    
    Select Case gErr
    
        Case 107869, 107905
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162967)

    End Select

    Exit Function

End Function

Function Caixa_ChequePre_Le_NumMovtoSangria(colChequesPre As Collection, lNumMovtoSangria As Long) As Long

Dim lErro As Long
Dim objCheque As New ClassChequePre
Dim bAchou As Boolean

On Error GoTo Erro_Caixa_ChequePre_Le_NumMovtoSangria
    bAchou = False
    
    For Each objCheque In gcolCheque
    
        'Verifica se o numero do movimento de sangria é igual ao numero do movimento dacoleção global de Cheque
        If objCheque.lNumMovtoSangria = lNumMovtoSangria Then
        
            'Adciona na Coleção Global de Cheque
            colChequesPre.Add objCheque
            bAchou = True
            
        End If
        
    Next
    
    'Se não existir nenhum movimento para esta código passado como parâmetro
    If bAchou = False Then gError 107865
    
    Caixa_ChequePre_Le_NumMovtoSangria = SUCESSO
    
    Exit Function

Erro_Caixa_ChequePre_Le_NumMovtoSangria:

    Caixa_ChequePre_Le_NumMovtoSangria = gErr

    Select Case gErr
    
        Case 107865
            Call Rotina_ErroECF(vbOKOnly, ERRO_CHEQUE_SANGRIA_NAOENCONTRADO, gErr)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162968)

    End Select

    Exit Function

End Function

Private Sub BotaoGravar_Click()
'Botão que Realiza a Gravação do Movto
Dim lErro As Long
Dim colMovimentosCaixa As New Collection
Dim vbMsgRes As VbMsgBoxResult
Dim iTipoMovimento As Integer
Dim lNumMovto As Long

On Error GoTo Erro_BotaoGravar_Click
    
    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 207973
    
    'Verifica se já foi executa a redução z para a data de hoje
    If gdtUltimaReducao = Date Then gError 118019
    
    'Função que efeuara a gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 107902
    
    'Funçao que Limpa a Tela de Movto Cheque
    Call Limpa_Tela_MovimentoCheque
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 107900 To 107902, 118019, 207973

        Case 118019
            Call Rotina_ErroECF(vbOKOnly, ERRO_REDUCAO_JA_EXECUTADA, gErr, Format(Date, "dd/mm/yyyy"))

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162969)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long

Dim lErro As Long
Dim colMovcxExclui As New Collection
Dim colMovCxInclui As New Collection
Dim vbMsgRes As VbMsgBoxResult
Dim iTipoMovimento As Integer
Dim lSequencialCaixa As Long
Dim colInfoChequesExcluir As New Collection
Dim colInfoChequesExcluirEsp As New Collection
Dim objChequePre As New ClassChequePre
Dim iIndice As Integer
Dim colColInfoCheques As New Collection
Dim objMovimentoCaixa As New ClassMovimentoCaixa
Dim colInfoChequesInserir As New Collection
Dim colInfoChequesInserirEsp As New Collection
Dim colChequeExclui As New Collection
Dim colChequeInclui As New Collection

On Error GoTo Erro_Gravar_Registro

    'Verifica se o Código não esta preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 107870

    'Função que Lê os Movto de Caixa com o Código Passado
    lErro = CF_ECF("Caixa_MovimentoCaixa_Le_NumMovto", colMovcxExclui, StrParaLong(Codigo.Text))
    If lErro <> SUCESSO And lErro <> 107850 Then gError 107873
    
    'Verifica se a Coleção Voltou carregada se voltou significa q é alteração
    If colMovcxExclui.Count > 0 Then
        
        Set objMovimentoCaixa = colMovcxExclui(1)
        
        If objMovimentoCaixa.iTipo <> MOVIMENTOCAIXA_SANGRIA_CHEQUE Then gError 86287
        
        'Envia aviso perguntando se deseja atualizar o movimemtos
        vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_ALTERACAO_MOVIMENTOCAIXA, Codigo.Text)
        If vbMsgRes = vbNo Then gError 107874
        
        'Atribuir ao tipo de movimento de caixa que é um  Movimento de exclusao de Sangria de cheque
        iTipoMovimento = MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_CHEQUE
                
        'Prepara com um flaq o movimento que vai ser excluido
        lErro = MovimentoCaixa_Prepara_Exclusao(colMovcxExclui, iTipoMovimento)
        If lErro <> SUCESSO Then gError 107875
        
        'Lê os Cheques do Movimento e carrega em colCheques, que já foram sangrados
        lErro = Caixa_ChequePre_Le_NumMovtoSangria(colChequeExclui, StrParaLong(Codigo.Text))
        If lErro <> SUCESSO Then gError 107876
        
        'Adciona o SequencialLoja à Coleção de Cheques que terão a sangria excluida
        For Each objChequePre In colChequeExclui
        
            'Atribui o Valor que esta no objCheque para a variável
            lSequencialCaixa = objChequePre.lSequencialCaixa
            
            If objChequePre.iNaoEspecificado = CHEQUE_ESPECIFICADO Then
                'Atribui o valor a colecão de cheques que serão excluidos
                colInfoChequesExcluirEsp.Add lSequencialCaixa
            Else
                'Atribui o valor a colecão de cheques que serão excluidos
                colInfoChequesExcluir.Add lSequencialCaixa
            End If
        Next
        
    End If
       
    'Guarda os dados na memoria que serão inseridos
    lErro = Move_Dados_Memoria(colMovCxInclui, colChequeInclui, colInfoChequesInserir, colInfoChequesInserirEsp)
    If lErro <> SUCESSO Then gError 107877
    
    'Função que grava os dados no arquivao
    lErro = Caixa_Grava_Movimento(colMovcxExclui, colInfoChequesExcluir, colInfoChequesExcluirEsp, colMovCxInclui, colInfoChequesInserir, colInfoChequesInserirEsp)
    If lErro <> SUCESSO Then gError 107878
    
    'Atualiza os Dados da Memoria
    lErro = MovimentoCheque_Exclui_Memoria(colMovcxExclui, colChequeExclui)
    If lErro <> SUCESSO Then gError 107895

    'Atualiza os Dados da Memoria
    lErro = MovimentoCheque_Inclui_Memoria(colMovCxInclui, colChequeInclui, StrParaLong(Codigo.Text))
    If lErro <> SUCESSO Then gError 105704

    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    Select Case gErr
    
        Case 86287
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_NAO_CHEQUE, gErr, Codigo.Text)
    
        Case 86288
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_JA_TRANSMITIDO, gErr)
            
        Case 105704, 107873 To 107878, 107895
            
        Case 107870
            Call Rotina_ErroECF(vbOKOnly, ERRO_CODIGO_NAO_PREENCHIDO1, gErr)
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162970)

    End Select

    Exit Function

End Function

Function Caixa_Grava_Movimento(colMovcxExclui As Collection, colInfoExcluir As Collection, colInfoExcluirEsp As Collection, Optional colMovCxInclui As Collection, Optional colcolInfoComplementar As Collection, Optional colcolInfoComplementarEsp As Collection) As Long

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim lSequencial As Long
Dim colRegistro As New Collection
Dim objOperador As New ClassOperador
Dim iIndice As Integer
Dim colInfoComplementar As Collection
Dim objMovimentoCaixa As New ClassMovimentoCaixa
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto
Dim sNomeTerminal As String
Dim iCont As Integer
Dim objMovimentoCaixaAux As New ClassMovimentoCaixa
Dim objCheque As New ClassChequePre
Dim sMensagem As String
Dim objMovCx As ClassMovimentoCaixa
Dim lTamanho As Long
Dim sRetorno As String
Dim sArquivo As String

On Error GoTo Erro_Caixa_Grava_Movimento

    If giStatusSessao = SESSAO_ENCERRADA Then

        'Envia aviso perguntando se de seja Abrir sessão
        vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_DESEJA_ABRIR_SESSAO, giCodCaixa)

        If vbMsgRes = vbNo Then gError 107879

        'Função que Executa Abertura na Sessão
        lErro = CF_ECF("Operacoes_Executa_Abertura")
        If lErro <> SUCESSO Then gError 107880

    End If

    'Se for Necessário a Altorização do Gerente para abertura do Caixa
    If gobjLojaECF.iGerenteAutoriza = AUTORIZACAO_GERENTE Then

        'Chama a Tela de Senha
        Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)

        'Sai de Função se a Tela de Login não Retornar ok
        If giRetornoTela <> vbOK Then gError 107881


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
    If lErro <> SUCESSO Then gError 107882

    lTamanho = 255
    sRetorno = String(lTamanho, 0)
        
    'Obtém a ultima transacao transferida
    Call GetPrivateProfileString(APLICACAO_DADOS, "UltimaTransacaoTransf", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
        
    'Retira os espaços no final da string
    sRetorno = StringZ(sRetorno)

    'grava as exclusoes
    For Each objMovimentoCaixa In colMovcxExclui
            
        'se o numero da ultima transacao transferida ultrapassar o numero da transacao do movimento de caixa
        If objMovimentoCaixa.lSequencial <> 0 And StrParaLong(sRetorno) > objMovimentoCaixa.lSequencial Then gError 133847
            
        'Caso nao precise de autorizacao do gerente nesta transacao ==> objOperador.iCodigo vai estar zerado
        objMovimentoCaixa.iGerente = objOperador.iCodigo
            
        lErro = Caixa_Grava_MovCx(objMovimentoCaixa, lSequencial, colInfoExcluir, colInfoExcluirEsp, colcolInfoComplementar, colcolInfoComplementarEsp)
        If lErro <> SUCESSO Then gError 105705
    
    Next
        
    'se as inclusoes estiverem preenchidas
    If Not colMovCxInclui Is Nothing Then
        
        'grava as inclusoes
        For Each objMovimentoCaixa In colMovCxInclui
                
            'se o numero da ultima transacao transferida ultrapassar o numero da transacao do movimento de caixa
            If objMovimentoCaixa.lSequencial <> 0 And StrParaLong(sRetorno) > objMovimentoCaixa.lSequencial Then gError 133848
                
            'Caso nao precise de autorizacao do gerente nesta transacao ==> objOperador.iCodigo vai estar zerado
            objMovimentoCaixa.iGerente = objOperador.iCodigo
                
            lErro = Caixa_Grava_MovCx(objMovimentoCaixa, lSequencial, colInfoExcluir, colInfoExcluirEsp, colcolInfoComplementar, colcolInfoComplementarEsp)
            If lErro <> SUCESSO Then gError 105706
        
        Next
        
    End If
        
    lSequencial = lSequencial - 1
    
    'Fecha a Transação
    lErro = CF_ECF("Caixa_Transacao_Fechar", lSequencial)
    If lErro <> SUCESSO Then gError 107885

    Close #10

    Caixa_Grava_Movimento = SUCESSO

    Exit Function

Erro_Caixa_Grava_Movimento:

    Close #10

    Caixa_Grava_Movimento = gErr

    Select Case gErr

        Case 105705, 105706, 107882, 107885
            'Erro Tratado Dentro da Função Chamada

        Case 107879 To 107881
        
        Case 133847, 133848
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_JA_TRANSMITIDO, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162971)

    End Select
    
    Call CF_ECF("Caixa_Transacao_Rollback", glTransacaoPAFECF)
    
    Exit Function

End Function

Function Caixa_Grava_MovCx(objMovimentoCaixa As ClassMovimentoCaixa, lSequencial As Long, colInfoExcluir As Collection, colInfoExcluirEsp As Collection, Optional colcolInfoComplementar As Collection, Optional colcolInfoComplementarEsp As Collection) As Long
'grava cada movimento de caixa passado como parametro

Dim lErro As Long
Dim colRegistro As New Collection
Dim sMensagem As String
Dim objMovCx As ClassMovimentoCaixa
Dim objTela As Object

On Error GoTo Erro_Caixa_Grava_MovCx

    'se for um movimento de exclusao de sangria ==> cria um novo movimento de caixa, grava no arquivão e deixa o que foi passado
    'como parametro sem alterar para permitir posteriormente retirar o movimento da memoria pelo sequencial original
    If objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_CHEQUE Then

        Set objMovCx = New ClassMovimentoCaixa
            
        lErro = CF_ECF("MovimentoCaixa_Copia", objMovimentoCaixa, objMovCx)
        If lErro <> SUCESSO Then gError 105701
        
    Else
    
        'se for um movimento de inclusao de sangria ==> os dados alterados nesta rotina e gravados no arquivão
        'serão passados para fora para serem gravados em memoria
        Set objMovCx = objMovimentoCaixa
        
    End If
        
    'Guarda o Sequencial no objmovimentoCaixa
    objMovCx.lSequencial = lSequencial

    lSequencial = lSequencial + 1

    'Guarda no objMovCx os Dados que Serão Usados para a Geração do Movimento de Caixa
    lErro = CF_ECF("Move_DadosGlobais_Memoria", objMovCx)
    If lErro <> SUCESSO Then gError 107883
            
    If objMovCx.iTipo <> MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_CHEQUE Then
    
        If objMovCx.iAdmMeioPagto = 0 Then
            'Funçao que Gera o Arquivo preparando para a gravação
            Call CF_ECF("MovimentoCheque_Gera_Log", colRegistro, objMovCx, colcolInfoComplementar)
        Else
            'Funçao que Gera o Arquivo preparando para a gravação
            Call CF_ECF("MovimentoCheque_Gera_Log", colRegistro, objMovCx, colcolInfoComplementarEsp)
        End If
    Else
    
        If objMovCx.iAdmMeioPagto = 0 Then
            'Funçao que Gera o Arquivo preparando para a gravação
            Call CF_ECF("MovimentoCheque_Gera_Log", colRegistro, objMovCx, colInfoExcluir)
        Else
            'Funçao que Gera o Arquivo preparando para a gravação
            Call CF_ECF("MovimentoCheque_Gera_Log", colRegistro, objMovCx, colInfoExcluirEsp)
        End If
        
    End If
    
    'Função que Vai Gravar as Informações no Arquivo de Caixa
    lErro = CF_ECF("MovimentoCaixaECF_Grava", colRegistro)
    If lErro <> SUCESSO Then gError 107884
    
    Set objTela = Me
            
    
    If objMovCx.iTipo = MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_CHEQUE Then

        lErro = AFRAC_AbrirRelatorioGerencial(RELGER_SANGRIA_CHEQUE, objTela)

        lErro = AFRAC_ImprimirRelatorioGerencial("Exclusão Sangria Cheque - Valor: " & Format(objMovCx.dValor, "standard"), objTela)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Exclusao Sangria - Cheque")
        If lErro <> SUCESSO Then gError 117666

        lErro = AFRAC_FecharRelatorioGerencial(objTela)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Exclusao Sangria - Cheque")
        If lErro <> SUCESSO Then gError 117667

        
    Else
        
        lErro = AFRAC_AbrirRelatorioGerencial(RELGER_SANGRIA_CHEQUE, objTela)

        lErro = AFRAC_ImprimirRelatorioGerencial("Sangria Cheque - Valor: " & Format(objMovCx.dValor, "standard"), objTela)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Sangria - Cheque")
        If lErro <> SUCESSO Then gError 117669

        lErro = AFRAC_FecharRelatorioGerencial(objTela)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Sangria - Cheque")
        If lErro <> SUCESSO Then gError 117670
        
    End If

    Caixa_Grava_MovCx = SUCESSO
    
    Exit Function

Erro_Caixa_Grava_MovCx:

    Caixa_Grava_MovCx = gErr

    Select Case gErr

        Case 105701, 107883, 107884, 109808, 109814, 117665 To 117670

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162972)

    End Select
    
    Exit Function

End Function

Function MovimentoCaixa_Prepara_Exclusao(colMovimentosCaixa As Collection, iTipoMovimento As Integer) As Long
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
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162973)

    End Select

    Exit Function

End Function

Function MovimentoCheque_Exclui_Memoria(colMovcxExclui As Collection, colChequeExclui As Collection) As Long
'exclui os dados da sangria da memoria global

Dim lErro As Long
Dim objMovimentosCaixa As New ClassMovimentoCaixa
Dim objCheque As New ClassChequePre
Dim iCont As Integer
Dim objChequeAux As New ClassChequePre
Dim dvalorMovimento As Double

On Error GoTo Erro_MovimentoCheque_Exclui_Memoria

    For Each objMovimentosCaixa In colMovcxExclui

            'Função que Retira de memória os Movimentos Excluidos
            lErro = MovimentoCaixa_Exclui_Memoria(objMovimentosCaixa)
            If lErro <> SUCESSO Then gError 107886

            gdSaldocheques = gdSaldocheques + objMovimentosCaixa.dValor

    Next

    'Atualiza a Coleção de Forma a Voltar para aColeção global de Cheques
    For Each objCheque In colChequeExclui
    
        'Zera o Numero da Sangria
        objCheque.lNumMovtoSangria = 0
            
    Next
    
    MovimentoCheque_Exclui_Memoria = SUCESSO

    Exit Function

Erro_MovimentoCheque_Exclui_Memoria:

    MovimentoCheque_Exclui_Memoria = gErr

    Select Case gErr

        Case 107886

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162974)

    End Select

    Exit Function

End Function

Function MovimentoCheque_Inclui_Memoria(colMovCxInclui As Collection, colChequeInclui As Collection, ByVal lNumMovtoSangria As Long) As Long
'Inclui os dados da sangria na memoria global

Dim lErro As Long
Dim objMovCx As ClassMovimentoCaixa
Dim objCheque As New ClassChequePre

On Error GoTo Erro_MovimentoCheque_Inclui_Memoria

    For Each objMovCx In colMovCxInclui

        'Adcionar a Coleção Global o objMovimento Caixa
        gcolMovimentosCaixa.Add objMovCx
        
        gdSaldocheques = gdSaldocheques - objMovCx.dValor

    Next

    'Atualiza a Coleção de Forma a Voltar para aColeção global de Cheques
    For Each objCheque In colChequeInclui
    
        'Zera o Numero da Sangria
        objCheque.lNumMovtoSangria = lNumMovtoSangria
            
    Next
    
    MovimentoCheque_Inclui_Memoria = SUCESSO

    Exit Function

Erro_MovimentoCheque_Inclui_Memoria:

    MovimentoCheque_Inclui_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162975)

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
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162976)

    End Select

    Exit Function

End Function

'Function MovimentoCheque_Atualiza_Memoria1(objMovimentosCaixa As ClassMovimentoCaixa, iIndice As Integer) As Long
''Função que Limpa a Coleção Global a Tela apos a Função de Gravação
'
'Dim lErro As Long
'Dim objCheque As New ClassChequePre
'
'On Error GoTo Erro_MovimentoCheque_Atualiza_Memoria1
'
'    For Each objCheque In gcolCheque
'
'        'Verifica se é a Mensma administardora ,parcelamento , Cartão
'        If objCheque.lNumMovtoSangria = objMovimentosCaixa.lNumMovto Then
'
'            gdSaldocheques = gdSaldocheques + objMovimentosCaixa.dValor
'
'
'        End If
'
'    Next
'
'    MovimentoCheque_Atualiza_Memoria1 = SUCESSO
'
'    Exit Function
'
'Erro_MovimentoCheque_Atualiza_Memoria1:
'
'    MovimentoCheque_Atualiza_Memoria1 = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162977)
'
'    End Select
'
'    Exit Function
'
'End Function

Sub Limpa_Tela_MovimentoCheque()
'Função que Limpa a Tela

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Limpa_Tela_MovimentoCheque

    'Limpa os Controles básico da Tela
    Codigo.Text = ""
    
    'Desmarcar Grig
    lErro = Carrega_GridCheques
    If lErro <> SUCESSO Then gError 107899
    
    lErro = Recalcula_Totais
    If lErro <> SUCESSO Then gError 107908
    
    Exit Sub
    
Erro_Limpa_Tela_MovimentoCheque:

    Select Case gErr

        Case 107899, 107908

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 162978)

    End Select

    Exit Sub

End Sub

Function Move_Dados_Memoria(colMovCxInclui As Collection, colChequeInclui As Collection, colInfoCheques As Collection, colInfoChequesEsp As Collection) As Long
'Função que Move os dados memória

Dim lErro As Long
Dim iIndice As Integer
Dim objCheque As ClassChequePre
Dim lSequencialLoja As Long
Dim objChequeAux As New ClassChequePre
Dim lSequencialLojaAux As Long
Dim iIndiceAux As Integer
Dim dValorEsp As Double
Dim dValorNaoEsp As Double
Dim objMovimentoCaixa As New ClassMovimentoCaixa

On Error GoTo Erro_Move_Dados_Memoria
    
    'Procura no grid os cheques que serão sangrados
    For iIndice = 1 To objGridCheques.iLinhasExistentes
    
        'Verifica se o cheque correspondente aquela linha esta marcado se estiver
        If GridCheques.TextMatrix(iIndice, iGrid_Selecionado_Col) = MARCADO Then
            
            lSequencialLojaAux = StrParaLong(GridCheques.TextMatrix(iIndice, iGrid_Seq_Col))
            
            For Each objChequeAux In gcolCheque
                    
               If objChequeAux.lSequencialCaixa = lSequencialLojaAux Then
                
                    'se o cheque já estiver sangrado e nao se tratar de uma alteracao ==> erro
                    If objChequeAux.lNumMovtoSangria <> 0 And objChequeAux.lNumMovtoSangria <> StrParaLong(Codigo.Text) Then gError 105731
                
                    lSequencialLoja = objChequeAux.lSequencialCaixa
                    
                    If GridCheques.TextMatrix(iIndice, iGrid_Tipo_Col) = STRING_ESPECIFICADO Then
                        'Adcionar o SequencialLoja a Coleção de Informações Adcionais que Foi passada por parâmetro
                        colInfoChequesEsp.Add lSequencialLoja
                    Else
                        'Adcionar o SequencialLoja a Coleção de Informações Adcionais que Foi passada por parâmetro
                        colInfoCheques.Add lSequencialLoja
                    End If
                    
                    colChequeInclui.Add objChequeAux
                    
                End If
            Next
            
            If GridCheques.TextMatrix(iIndice, iGrid_Tipo_Col) = STRING_ESPECIFICADO Then
                dValorEsp = dValorEsp + StrParaDbl(GridCheques.TextMatrix(iIndice, iGrid_Valor_Col))
            Else
                dValorNaoEsp = dValorNaoEsp + StrParaDbl(GridCheques.TextMatrix(iIndice, iGrid_Valor_Col))
            End If
        
        End If
    
    Next
    
    'se tem cheques especificados
    If dValorEsp > 0 Then
    
        'Instancia um novo obj
        Set objMovimentoCaixa = New ClassMovimentoCaixa
        
        'guarda o tipo de Movto
        objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_SANGRIA_CHEQUE
            
        'Guarda em qual filial empresa que esta Trabalhando
        objMovimentoCaixa.iFilialEmpresa = giFilialEmpresa
                    
        'Preenche o obj com o Numero do Movimento
        objMovimentoCaixa.lNumMovto = StrParaLong(Codigo.Text)
            
        objMovimentoCaixa.iParcelamento = COD_A_VISTA
    
        'guarda o valor do moviento no objMovimentoCaixa
        objMovimentoCaixa.dValor = dValorEsp
    
        'Guardo o Codigo da Admnistradora no Movimento Caixa
        objMovimentoCaixa.iAdmMeioPagto = MEIO_PAGAMENTO_CHEQUE
    
        colMovCxInclui.Add objMovimentoCaixa
    
    End If
    
    'se tem cheques nao especificado
    If dValorNaoEsp > 0 Then

        'Instancia um novo obj
        Set objMovimentoCaixa = New ClassMovimentoCaixa
        
        'guarda o tipo de Movto
        objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_SANGRIA_CHEQUE
            
        'Guarda em qual filial empresa que esta Trabalhando
        objMovimentoCaixa.iFilialEmpresa = giFilialEmpresa
                    
        'Preenche o obj com o Numero do Movimento
        objMovimentoCaixa.lNumMovto = StrParaLong(Codigo.Text)
            
        objMovimentoCaixa.iParcelamento = COD_A_VISTA

        'guarda o valor do moviento no objMovimentoCaixa
        objMovimentoCaixa.dValor = dValorNaoEsp
    
        'Guardo o Codigo do meio de Pagto no Movimento Caixa
        objMovimentoCaixa.iAdmMeioPagto = 0
    
        colMovCxInclui.Add objMovimentoCaixa

    End If

    If colMovCxInclui.Count = 0 Then gError 105771

    Exit Function

Erro_Move_Dados_Memoria:

    Move_Dados_Memoria = gErr

    Select Case gErr

        Case 105731
            Call Rotina_ErroECF(vbOKOnly, ERRO_CHEQUE_SANGRADO_GRID, gErr, iIndice, objChequeAux.lNumMovtoSangria)

        Case 105771
            Call Rotina_ErroECF(vbOKOnly, ERRO_CHEQUE_NAO_SELECIONADO, gErr, iIndice, objChequeAux.lNumMovtoSangria)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162979)

    End Select

    Exit Function
    
End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo BotaoLimpar_Click

    'Função que Lima a Tela
    Call Limpa_Tela_MovimentoCheque

    Exit Sub

BotaoLimpar_Click:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162980)

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

    lErro = CF_ECF("Desmembra_MovimentosCheque", colMovimentosCaixa, colImfCompl, TIPOREGISTROECF_MOVIMENTOCAIXA_CHEQUES)
    If lErro <> SUCESSO Then gError 111064
    
    For Each objMovimentosCaixa In colMovimentosCaixa
        If objMovimentosCaixa.iTipo <> MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_CHEQUE Then
            gcolMovimentosCaixa.Add colMovimentosCaixa
            gcolImfCompl.Add colImfCompl
        End If
    Next
    
    Exit Sub

Erro_DesmembraMovto_Click:
    
    Select Case gErr

        Case 111064

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 162981)

    End Select

    Exit Sub


End Sub


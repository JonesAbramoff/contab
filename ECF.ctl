VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ECF 
   ClientHeight    =   4125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6195
   KeyPreview      =   -1  'True
   ScaleHeight     =   4125
   ScaleWidth      =   6195
   Begin VB.Frame Frame1 
      Caption         =   "Impressora"
      Height          =   1545
      Left            =   105
      TabIndex        =   17
      Top             =   1290
      Width           =   5970
      Begin MSMask.MaskEdBox CodImp 
         Height          =   315
         Left            =   1215
         TabIndex        =   4
         Top             =   345
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin VB.Label LabelFabricante 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1215
         TabIndex        =   24
         Top             =   975
         Width           =   1785
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fabricante:"
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
         TabIndex        =   23
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label LabelModelo 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3930
         TabIndex        =   22
         Top             =   1020
         Width           =   1785
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Modelo:"
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
         Left            =   3150
         TabIndex        =   21
         Top             =   1065
         Width           =   690
      End
      Begin VB.Label LabelNumSerie 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3225
         TabIndex        =   20
         Top             =   345
         Width           =   2505
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Num. Serie:"
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
         Left            =   2115
         TabIndex        =   19
         Top             =   405
         Width           =   1005
      End
      Begin VB.Label LabelCodImp 
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
         Left            =   480
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   18
         Top             =   405
         Width           =   660
      End
   End
   Begin VB.CheckBox Ativo 
      Caption         =   "Ativo"
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
      Left            =   2565
      TabIndex        =   16
      Top             =   270
      Value           =   1  'Checked
      Width           =   900
   End
   Begin VB.CommandButton Teste_Log 
      Caption         =   "Teste_Log"
      Height          =   360
      Left            =   5190
      TabIndex        =   15
      Top             =   825
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3960
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   90
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "ECF.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "ECF.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "ECF.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ECF.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CheckBox ImpressoraCheque 
      Caption         =   "Tem Impressora de Cheques"
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
      Left            =   300
      TabIndex        =   5
      Top             =   3000
      Width           =   2850
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   1770
      Picture         =   "ECF.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Numeração Automática"
      Top             =   255
      Width           =   300
   End
   Begin VB.Frame FrameHorariodeVerao 
      Caption         =   "Horário de Verão"
      Height          =   660
      Index           =   0
      Left            =   105
      TabIndex        =   11
      Top             =   3390
      Width           =   5970
      Begin VB.OptionButton HorarioVeraoInativo 
         Caption         =   "Inativo"
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
         Height          =   270
         Left            =   2790
         TabIndex        =   13
         Top             =   285
         Width           =   1020
      End
      Begin VB.OptionButton HorarioVeraoAtivo 
         Caption         =   "Ativo"
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
         Height          =   270
         Left            =   840
         TabIndex        =   12
         Top             =   300
         Width           =   1125
      End
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1155
      TabIndex        =   1
      Top             =   225
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Caixa 
      Height          =   315
      Left            =   1155
      TabIndex        =   3
      Top             =   840
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin VB.Label LabelNomeCaixa 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1800
      TabIndex        =   25
      Top             =   840
      Width           =   2505
   End
   Begin VB.Label LabelECF 
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
      Left            =   405
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   0
      Top             =   285
      Width           =   660
   End
   Begin VB.Label LabelCaixa 
      AutoSize        =   -1  'True
      Caption         =   "Caixa:"
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
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   10
      Top             =   885
      Width           =   540
   End
End
Attribute VB_Name = "ECF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Const de Inclusão de Log

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoCaixa As AdmEvento
Attribute objEventoCaixa.VB_VarHelpID = -1
Private WithEvents objEventoECF As AdmEvento
Attribute objEventoECF.VB_VarHelpID = -1
Private WithEvents objEventoImp As AdmEvento
Attribute objEventoImp.VB_VarHelpID = -1

'Declarações Globais
Dim iAlterado As Integer

Private Sub Ativo_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Emissor de Cupom Fiscal"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ECF"
    
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


Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    
        Case KEYCODE_BROWSER
        
            If Me.ActiveControl Is Codigo Then
                Call LabelECF_Click
            ElseIf Me.ActiveControl Is Caixa Then
                Call LabelCaixa_Click
            ElseIf Me.ActiveControl Is CodImp Then
                Call LabelCodImp_Click
            End If
        
        Case KEYCODE_PROXIMO_NUMERO
            Call BotaoProxNum_Click
    
    End Select

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

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
'    'Carrega a lista de ECF's cadastrados no BD
'    lErro = Carrega_ECF_Lista()
'    If lErro <> SUCESSO Then gError 79527
    
'    'Carrega a combo de Impressoras com as Impressoras cadastrados no BD
'    lErro = Carrega_Impressoras_Combo()
'    If lErro <> SUCESSO Then gError 79530
'
'    'Carrega a combo de Caixas com os Caixas cadastrados no BD
'    lErro = Carrega_Caixa_Combo()
'    If lErro <> SUCESSO Then gError 79528
'
    Set objEventoCaixa = New AdmEvento
    Set objEventoECF = New AdmEvento
    Set objEventoImp = New AdmEvento
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 79527, 79528, 79530
            'Erros Tratados Dentro da Função Chamadora
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159158)

    End Select

    iAlterado = 0

    Exit Sub

    lErro_Chama_Tela = SUCESSO

End Sub

'Private Function Carrega_ECF_Lista() As Long
''Carrega a list de Ecf's com os ECF's da filial ativa cadastrados no BD
'
'Dim lErro As Long
'Dim colECF As New Collection
'Dim objECF As ClassECF
'
'On Error GoTo Erro_Carrega_ECF_Lista
'
'    'Lê os códigos dos ECFs da filial ativa
'    lErro = CF("ECF_Le_Todos", colECF)
'    If lErro <> SUCESSO And lErro <> 79524 Then gError 79523
'
'    'Se encontrou pelo menos um ECF no BD
'    If lErro <> 79524 Then
'
'        'Preenche a listbox com os códigos dos ECFs
'        For Each objECF In colECF
'            ECFs.AddItem objECF.iCodigo
'            ECFs.ItemData(ECFs.NewIndex) = objECF.iCodigo
'        Next
'
'    End If
'
'    Carrega_ECF_Lista = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_ECF_Lista:
'
'    Carrega_ECF_Lista = gErr
'
'    Select Case gErr
'
'        Case 79523
'            'Erro Tratado Dentro da Função Chamadora
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159159)
'
'    End Select
'
'    Exit Function
'
'End Function

'Private Function Carrega_Impressoras_Combo() As Long
''Carrega a combo de Fabricantes com os Fabricantes existentes no BD
'
'Dim lErro As Long
'Dim colImpressorasECF As New Collection
'Dim objImpressorasECF As ClassFabricanteECF
'
'On Error GoTo Erro_Carrega_Impressoras_Combo
'
'    'Lê os códigos e nomes reduzidos de todos os Fabricantes de ECF cadastrados no BD
'    lErro = CF("ImpressorasECF_Le_Todos", colImpressorasECF)
'    If lErro <> SUCESSO And lErro <> 79534 Then gError 79536
'
'    'Se não Houver Registros no Banco de Dados a Tela não pode ser Inicializada.
'    If lErro = 79534 Then gError 104277 '????? ok nao tratou
'
'    'Adiciona na combobox o Código e o Nome das Impressoras encontrados
'    For Each objImpressorasECF In colImpressorasECF
'        Fabricante.AddItem objImpressorasECF.iCodigo & SEPARADOR & objImpressorasECF.sNome
'        Fabricante.ItemData(Fabricante.NewIndex) = objImpressorasECF.iCodigo
'    Next
'
'    Carrega_Impressoras_Combo = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_Impressoras_Combo:
'
'    Carrega_Impressoras_Combo = gErr
'
'    Select Case gErr
'
'        Case 79536
'            'Erro Tradado Dentro da Função
'
'        Case 104277
'            Call Rotina_Erro(vbOKOnly, "ERRO_NAO_EXISTE_IMPRESSORAS_CADASTRADAS", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159160)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Carrega_Caixa_Combo() As Long
''Carrega a combo de Caixas com os Caixas da filial ativa
'
'Dim lErro As Long
'Dim colCaixa As New Collection
'Dim objCaixa As ClassCaixa
'
'On Error GoTo Erro_Carrega_Caixa_Combo
'
'    'Lê os códigos e os nomes reduzidos dos Caixas da filial ativa
'    lErro = CF("Caixa_Le_Todos", colCaixa)
'    If lErro <> SUCESSO And lErro <> 79525 Then gError 79526
'
'    'Se não Existir Caixa não Abre a Tela
'    If lErro = 79525 Then gError 104276 '???? ok Por que é erro
'
'    'Adiciona na combobox o par Código - Nome Reduzido do Caixa
'    For Each objCaixa In colCaixa
'        Caixa.AddItem objCaixa.iCodigo & SEPARADOR & objCaixa.sNomeReduzido
'        Caixa.ItemData(Caixa.NewIndex) = objCaixa.iCodigo
'    Next
'
'    Carrega_Caixa_Combo = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_Caixa_Combo:
'
'    Carrega_Caixa_Combo = gErr
'
'    Select Case gErr
'
'        Case 79526
'            'Erro Tratado Dentro da Função Chamadora
'
'        Case 104276
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_NAO_EXISTE_CAIXA_CADASTRADA", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159161)
'
'    End Select
'
'    Exit Function
'
'End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai o ECF da tela

Dim lErro As Long
Dim objECF As New ClassECF

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "ECF"

    'le os dados da tela
    lErro = Move_Tela_Memoria(objECF)
    If lErro <> SUCESSO Then gError 79586

    'Preenche a coleção colCampoValor
    colCampoValor.Add "Codigo", objECF.iCodigo, 0, "Codigo"
    colCampoValor.Add "ImpressoraECF", objECF.iImpressoraECF, 0, "ImpressoraECF"
    colCampoValor.Add "Caixa", objECF.iCaixa, 0, "Caixa"
    colCampoValor.Add "ImpressoraCheque", objECF.iImpressoraCheque, 0, "ImpressoraCheque"
    colCampoValor.Add "HorarioVerao", objECF.iHorarioVerao, 0, "HorarioVerao"
    colCampoValor.Add "FilialEmpresa", objECF.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "Ativo", objECF.iAtivo, 0, "Ativo"
    'Faz o filtro dos dados que serão exibidos
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 79586
            'Erro Tratado Dentro da Função Chamadora
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159162)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objECF As New ClassECF

On Error GoTo Erro_Tela_Preenche

    'Carrega objECF com os dados passados em colCampoValor
    objECF.iCodigo = colCampoValor.Item("Codigo").vValor
    objECF.iImpressoraECF = colCampoValor.Item("ImpressoraECF").vValor
    objECF.iCaixa = colCampoValor.Item("Caixa").vValor
    objECF.iImpressoraCheque = colCampoValor.Item("ImpressoraCheque").vValor
    objECF.iHorarioVerao = colCampoValor.Item("HorarioVerao").vValor
    objECF.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
    objECF.iAtivo = colCampoValor.Item("Ativo").vValor
        
    'Funação que Traz os dados de ECF Para Tela
    lErro = Traz_ECF_Tela(objECF)
    If lErro <> SUCESSO Then gError 79585

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 79585
            'Erro Tratado dentro da função chamadora
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159163)

    End Select

    Exit Sub

End Sub

Function Move_Tela_Memoria(objECF As ClassECF) As Long
'Guarda no objECF os dados informados na tela

Dim objCaixa As New ClassCaixa
Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    'Guarda Codigo, FilialEmpresa
    objECF.iCodigo = StrParaInt(Codigo.Text)
    objECF.iFilialEmpresa = giFilialEmpresa
        
    'Guarda em objECF o codigo Selecionado na List Fabricantes
    objECF.iImpressoraECF = StrParaInt(CodImp.Text)
      
    objECF.iCaixa = StrParaInt(Caixa.Text)
    
    'Guarda em objECF o valor de ImpressoraCheque
    objECF.iImpressoraCheque = ImpressoraCheque.Value
    
    'Guarda se o Horário de Verão esta ativo ou não
    objECF.iHorarioVerao = NAO_SELECIONADO
  
    If Ativo.Value = vbUnchecked Then
        objECF.iAtivo = ECF_INATIVO
    Else
        objECF.iAtivo = ECF_ATIVO
    End If
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
        
        Case 118000
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXA_NAO_CADASTRADO", gErr, objCaixa.iCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159164)

    End Select

    Exit Function

End Function

Private Function Traz_ECF_Tela(objECF As ClassECF) As Long
'Exibe na tela os dados carregados em objECF
Dim iIndice As Integer

On Error GoTo Erro_Traz_ECF_Tela

    If objECF.iAtivo = ECF_ATIVO Then
        Ativo.Value = vbChecked
    Else
        Ativo.Value = vbUnchecked
    End If

    'Exibe Codigo e POS na tela
    Codigo.Text = objECF.iCodigo
    
    CodImp.Text = objECF.iImpressoraECF
    Call CodImp_Validate(False)
    
    Caixa.Text = objECF.iCaixa
    Call Caixa_Validate(False)
    
    'Se o ECF possui Impressora de Cheque
    If objECF.iImpressoraCheque = IMPRESSORA_PRESENTE Then
        'Marca a checkbox
        ImpressoraCheque.Value = vbChecked
    
    'Senão
    Else
        'Desmarca a checkbox
        ImpressoraCheque.Value = vbUnchecked
    End If
    
   'Se o ECF está configurado em horário de verão
    If objECF.iHorarioVerao = HORARIO_VERAO_ATIVO Then
        'Marca a opção Horário Ativo
        HorarioVeraoAtivo.Value = True
    
    'Senão
    Else
        'Marca a opção Horário Inativo
        HorarioVeraoInativo.Value = True
    
    End If

    'Indica que não existe nenhum campo alterado
    iAlterado = 0
    
    Traz_ECF_Tela = SUCESSO
    
    Exit Function

Erro_Traz_ECF_Tela:

    Traz_ECF_Tela = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159165)
    
    End Select

    Exit Function
    
End Function

Private Sub LabelCodImp_Click()

Dim objImp As New ClassImpressoraECF
Dim colSelecao As Collection
    
    If Len(Trim(CodImp.Text)) <> 0 Then
        objImp.iCodigo = StrParaInt(CodImp.Text)
    End If
    
    Call Chama_Tela("ImpressoraECFLista", colSelecao, objImp, objEventoImp)

    Exit Sub

End Sub

Private Sub objEventoImp_evSelecao(obj1 As Object)

Dim objImp As ClassImpressoraECF
Dim lErro As Long
Dim iIndex As Integer
Dim objModeloECF As New ClassModeloECF

On Error GoTo Erro_objEventoImp_evSelecao

    Set objImp = obj1
    
    CodImp.Text = objImp.iCodigo
    Call CodImp_Validate(False)
    
    Me.Show
        
    Exit Sub

Erro_objEventoImp_evSelecao:
    
    Select Case gErr
                
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159166)

    End Select
    
    Exit Sub

End Sub

Private Sub LabelCaixa_Click()

Dim objCaixa As New ClassCaixa
Dim colSelecao As Collection
    
    If Len(Trim(Caixa.Text)) <> 0 Then
        objCaixa.iCodigo = StrParaInt(Caixa.Text)
    End If
    
    Call Chama_Tela("CaixaLista", colSelecao, objCaixa, objEventoCaixa)

    Exit Sub

End Sub

Private Sub objEventoCaixa_evSelecao(obj1 As Object)

Dim objCaixa As ClassCaixa
Dim lErro As Long
Dim iIndex As Integer

On Error GoTo Erro_objEventoCaixa_evSelecao

    Set objCaixa = obj1
    
    Caixa.Text = objCaixa.iCodigo
    Call Caixa_Validate(False)
        
    Me.Show
        
    Exit Sub

Erro_objEventoCaixa_evSelecao:
    
    Select Case gErr
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159167)

    End Select
    
    Exit Sub

End Sub

Private Sub LabelECF_Click()

Dim objECF As New ClassECF
Dim colSelecao As Collection
    
    If Len(Trim(Codigo.Text)) <> 0 Then
        objECF.iCodigo = StrParaInt(Codigo.Text)
    End If
    
    Call Chama_Tela("ECFLojaLista", colSelecao, objECF, objEventoECF)

    Exit Sub

End Sub

Private Sub objEventoECF_evSelecao(obj1 As Object)

Dim objECF As ClassECF
Dim lErro As Long
Dim iIndex As Integer

On Error GoTo Erro_objEventoECF_evSelecao

    Set objECF = obj1
    
    objECF.iFilialEmpresa = giFilialEmpresa
            
    'Lê ECF no BD a partir do código
    lErro = CF("ECF_Le", objECF)
    If lErro <> SUCESSO And lErro <> 79573 Then gError 118003

    If lErro = SUCESSO Then
        'Exibe os dados do ECF na tela
        lErro = Traz_ECF_Tela(objECF)
        If lErro <> SUCESSO Then gError 118004
    End If
        
    Me.Show
        
    Exit Sub

Erro_objEventoECF_evSelecao:
    
    Select Case gErr
                
        Case 118003, 118004
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159168)

    End Select
    
    Exit Sub

End Sub

Private Sub CodImp_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub CodImp_Validate(Cancel As Boolean)
    
Dim objImp As New ClassImpressoraECF
Dim objModeloECF As New ClassModeloECF
Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long
    
On Error GoTo Erro_CodImp_Validate

    'se o código estiver preenchido
    If Len(Trim(CodImp.Text)) <> 0 Then
    
        objImp.iFilialEmpresa = giFilialEmpresa
        objImp.iCodigo = StrParaInt(CodImp.Text)
        
        lErro = CF("ImpressoraECF_Le", objImp)
        If lErro <> SUCESSO And lErro <> 103447 Then gError 112797
    
        If lErro <> SUCESSO Then gError 112798
    
        objModeloECF.iCodigo = objImp.iCodModelo
        
        'busca na tabela ModeloECF
        lErro = CF("ModeloECF_Le", objModeloECF)
        If lErro <> SUCESSO And lErro <> 103459 Then gError 112775
        
        If lErro = 103459 Then gError 112776
        
    End If
    LabelNumSerie.Caption = objImp.sNumSerie
    
    LabelFabricante.Caption = objModeloECF.sFabricante
    LabelModelo.Caption = objModeloECF.sNome
        
    Exit Sub

Erro_CodImp_Validate:

    Cancel = True

    Select Case gErr

        Case 112775, 112797
        
        Case 112776
            Call Rotina_Erro(vbOKOnly, "ERRO_MODELOECF_NAO_CADASTRADO", gErr)
            
        Case 112798
            
            'pergunta se deseja cadastrar
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_IMPRESSORA", objImp.iCodigo)
            
            'Se confirma
            If vbMsgRes = vbYes Then
                
                Call Chama_Tela("ImpressoraECF", objImp)
            
            End If
   
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159169)

    End Select

    Exit Sub
    
End Sub

Private Sub HorarioVeraoAtivo_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub HorarioVeraoInativo_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub
Private Sub Codigo_Change()

    'Indica que o conteúdo do campo foi alterado
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Caixa_Change()

    'Indica que o conteúdo do campo foi alterado
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Caixa_Click()

    'Indica que o conteúdo do campo foi alterado
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Caixa_Validate(Cancel As Boolean)

Dim objCaixa As New ClassCaixa
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Caixa_Validate

    'se o código estiver preenchido
    If Len(Trim(Caixa.Text)) <> 0 Then

        objCaixa.iCodigo = StrParaInt(Caixa.Text)
        objCaixa.iFilialEmpresa = giFilialEmpresa
        
        'Lê Caixa no BD a partir do nome reduzido
        lErro = CF("Caixas_Le", objCaixa)
        If lErro <> SUCESSO And lErro <> 79405 Then gError 118001
        
        If lErro = 79405 Then gError 118002
    End If
    
    LabelNomeCaixa.Caption = objCaixa.sNomeReduzido
            
    Exit Sub

Erro_Caixa_Validate:

    Cancel = True

    Select Case gErr

        Case 118001

        Case 118002
            
            'pergunta se deseja cadastrar
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CAIXA", objCaixa.iCodigo)
            
            'Se confirma
            If vbMsgRes = vbYes Then
                
                Call Chama_Tela("Caixa", objCaixa)
            
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159170)

    End Select

    Exit Sub
        
End Sub

Private Sub ImpressoraCheque_Click()

    'Indica que o conteúdo do campo foi alterado
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub BotaoProxNum_Click()
'Gera um novo número disponível para código de ECF

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click
    
    'Chama a função que gera o Código Automático para o novo ECF
    lErro = ECF_Codigo_Automatico(lCodigo)
    If lErro <> SUCESSO Then gError 79538

    'Exibe o novo código na tela
    Codigo.Text = CStr(lCodigo)
        
    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 79538
            'Erro tratado na rotina chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159171)
    
    End Select

    Exit Sub

End Sub

Function ECF_Codigo_Automatico(lCodigo As Long) As Long
'Gera o proximo codigo da Tabela de Requisitante

Dim lErro As Long

On Error GoTo Erro_ECF_Codigo_Automatico

    'Chama a rotina que gera o sequencial
    lErro = CF("Config_ObterAutomatico", "LojaConfig", "NUM_PROXIMO_ECF", "ECF", "Codigo", lCodigo)
    If lErro <> SUCESSO Then Error 104271

    ECF_Codigo_Automatico = SUCESSO

    Exit Function

Erro_ECF_Codigo_Automatico:

    ECF_Codigo_Automatico = Err

    Select Case Err

        Case 104271
            'Erro Tratado dentro da Função Chamadora
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159172)

    End Select

    Exit Function

End Function

'Private Sub ECFs_Click()
''Carrega para a tela o ECF selecionado através de clique
'
'Dim lErro As Long
'Dim objECF As New ClassECF
'
'On Error GoTo Erro_ECFs_Click
'
'    'Verifica se nenhum item está selecionado sair da Função
'    If ECFs.ListIndex = -1 Then Exit Sub
'
'    'Define o parâmetro que será passado para ECF_Le
'    objECF.iCodigo = ECFs.ItemData(ECFs.ListIndex)
'
'    'Define o parâmetro que será passado para ECF_Le
'    objECF.iFilialEmpresa = giFilialEmpresa
'
'
'    'Procura a ECF no BD através do código
'    lErro = CF("ECF_Le", objECF)
'    If lErro <> SUCESSO And lErro <> 79573 Then gError 79574
'
'    'Se não encontrou =>erro
'    If lErro = 79573 Then gError 79575
'
'    'Traz para a tela os dados do ECF selecionado
'    lErro = Traz_ECF_Tela(objECF)
'    If lErro <> SUCESSO Then gError 79576
'
'    'Fecha o comando das setas se estiver aberto
'    Call ComandoSeta_Fechar(Me.Name)
'
'    Exit Sub
'
'Erro_ECFs_Click:
'
'    Select Case gErr
'
'        Case 79574, 79576
'            'Erro Tratado Dentro da Fuinção Chamadora
'        Case 79575
'            Call Rotina_Erro(vbOKOnly, "ERRO_ECF_NAO_CADASTRADO", gErr, objECF.iCodigo)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159173)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Private Sub BotaoGravar_Click()
'Chama as rotinas que irão efetuar a gravação do ECF no BD

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama a rotina de gravação do registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 79540

    'Limpa a Tela
    Call Limpa_Tela(Me)

    'Zera o iAlterado para demostrar que não houve alteração
    iAlterado = 0
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 79540
            'Erro Tratado Dentro da Função que Chamada
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159174)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long
'Verifica se os dados obrigatórios da ECF foram preenchidos
'Grava ECF no BD
'Atualiza List

Dim lErro As Long
Dim objECF As New ClassECF
Dim iCodigo As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se os campos obrigatórios da tela foram preenchidos
    lErro = ECF_Critica_CamposPreenchidos()
    If lErro <> SUCESSO Then gError 79544

    'Passa para objECF os dados contidos na tela
    lErro = Move_Tela_Memoria(objECF)
    If lErro <> SUCESSO Then gError 79545
    
    'Verifica se Houva Alguma Alteração
    lErro = Trata_Alteracao(objECF, objECF.iCodigo, objECF.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 32328
    
    'Chama a função que insere ou atualiza o ECF no BD
    lErro = CF("ECF_Grava", objECF)
    If lErro <> SUCESSO Then gError 79546
               
'    'Retira o ECF da lista de ECF's
'    Call ListaECFs_Exclui(objECF.iCodigo)
'
'    'Recoloca o ECF na lista de ECF's
'    Call ListaECFs_Adiciona(objECF)
    
    'Limpa a Tela
    Call Limpa_Tela_ECF
    
    Gravar_Registro = SUCESSO

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 32328, 79544 To 79546
            'Erros Tratados Dentro da Função Chamados
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159175)

    End Select

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

End Function

Function ECF_Critica_CamposPreenchidos() As Long
'Verifica se os campos obrigatórios da tela foram preenchidos

Dim lErro As Long

On Error GoTo Erro_ECF_Critica_CamposPreenchidos

    'Se o código não foi preenchido => erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 79541

    'Se o Fabricante não foi selecionado => erro
    If Len(Trim(CodImp.Text)) = 0 Then gError 79542

    'Se o Caixa não foi selecionado => erro
    If Len(Trim(Caixa.Text)) = 0 Then gError 79543

    ECF_Critica_CamposPreenchidos = SUCESSO
    
    Exit Function

Erro_ECF_Critica_CamposPreenchidos:

    ECF_Critica_CamposPreenchidos = gErr
    
    Select Case gErr

        Case 79541
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 79542
            Call Rotina_Erro(vbOKOnly, "ERRO_IMPRESSORA_NAO_SELECIONADA", gErr)
        
        Case 79543
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXA_NAO_SELECIONADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159176)
        
    End Select
    
    Exit Function

End Function

'Private Sub ListaECFs_Exclui(iCodigo As Long)
''Percorre a ListBox de ECF's para remover o ECF que está sendo cadastrado, caso ele exista
'
'Dim iIndice As Integer
'
'    'Faz o loop na lista de ECF's
'    For iIndice = 0 To ECFs.ListCount - 1
'
'        'Se o Código do ECF que se deseja retirar da lista é igual ao ItemData do ECF atual
'        If ECFs.ItemData(iIndice) = iCodigo Then
'
'            'Remove esse ECF da lista
'            ECFs.RemoveItem (iIndice)
'            Exit For
'
'        End If
'
'    Next
'
'End Sub
'
'Private Sub ListaECFs_Adiciona(objECF As ClassECF)
''Adiciona na ListBox de ECF's o ECF que acabou de ser gravado
'
'Dim iIndice As Integer
'
'    'Faz o loop na lista de ECF's
'    For iIndice = 0 To ECFs.ListCount - 1
'
'        'Se o ItemData do ECF atual é maior que o código do ECF que se deseja incluir => sai do loop e mantém a posição (índice) onde deve entrar o novo ECF
'        If ECFs.ItemData(iIndice) > objECF.iCodigo Then Exit For
'    Next
'
'    'Adiciona o ECF à lista e guarda seu código no ItemData
'    ECFs.AddItem objECF.iCodigo, iIndice
'    ECFs.ItemData(ECFs.NewIndex) = objECF.iCodigo
'
'    Exit Sub
'
'End Sub

Private Sub BotaoExcluir_Click()
'Chama as funções que irão excluir uma ECF do BD

Dim lErro As Long
Dim objECF As New ClassECF
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Se o código não foi preenchido = > erro
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 79599

    'Guarda em objECF o código que será passado como parâmtero para ECF_Le
    objECF.iCodigo = StrParaInt(Codigo.Text)
    objECF.iFilialEmpresa = giFilialEmpresa
    
    'Lê no BD os dados do ECF que será excluído
    lErro = CF("ECF_Le", objECF)
    If lErro <> SUCESSO And lErro <> 79573 Then gError 79600

    'Se o ECF não estiver cadastrado => erro
    If lErro = 79573 Then gError 79601
    
    'Envia aviso perguntando se realmente deseja excluir ECF
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_ECF", objECF.iCodigo)

    If vbMsgRes = vbYes Then
    'Se sim
    
        'Chama a Função Move_Tela_Memória
        lErro = Move_Tela_Memoria(objECF)
        If lErro <> SUCESSO Then gError 104273
        
        'Chama a função que irá excluir o ECF
        lErro = CF("ECF_Exclui", objECF)
        If lErro <> SUCESSO Then gError 79618

'        'Retira o nome do ECF da lista de ECF's
'        Call ListaECFs_Exclui(objECF.iCodigo)

        'Limpa a Tela
        Call Limpa_Tela_ECF
                
    End If

    iAlterado = 0
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr
        
        Case 79599
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 79600, 79618, 104273
            'Erro Tratado dentro da função chamadora

        Case 79601
            Call Rotina_Erro(vbOKOnly, "ERRO_ECF_NAO_CADASTRADO", gErr, objECF.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159177)

    End Select

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
'chamada de Limpa_Tela_ECF

Dim lErro As Long

On Error GoTo Erro_Botaolimpar_Click

    'Faz o teste que verifica se algum campo foi alterado
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 79569

    'Limpa Tela
    Call Limpa_Tela_ECF

    Exit Sub

Erro_Botaolimpar_Click:

    Select Case gErr

        Case 79569
            'Erro Tratado Dentro da Função que Foi Chamada
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159178)

    End Select

End Sub

Sub Limpa_Tela_ECF()
'Limpa os campos da tela de ECF e seleciona as opções default

On Error GoTo Erro_Limpa_Tela_ECF

    'Limpa a tela
    Call Limpa_Tela(Me)
    
    Ativo.Value = vbChecked

    'Passa o foco para o campo código
    Codigo.SetFocus
    
    LabelFabricante.Caption = ""
    LabelModelo.Caption = ""
    LabelNumSerie.Caption = ""
    LabelNomeCaixa.Caption = ""
    
    'Desmarca o Horário de Verão
    HorarioVeraoAtivo.Value = False
    HorarioVeraoInativo.Value = False
       
    'Desmarca as Checkbox ImpressoraCheque
    ImpressoraCheque.Value = vbUnchecked
        
    'Fecha o comando de setas
    Call ComandoSeta_Fechar(Me.Name)
        
    'Indica que não existe nenhum campo alterado
    iAlterado = 0
    
    Exit Sub

Erro_Limpa_Tela_ECF:

    Select Case gErr
      
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159179)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub Form_Activate()

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
    
    Set objEventoCaixa = Nothing
    Set objEventoECF = Nothing
    Set objEventoImp = Nothing
    
    'Fecha o comando de setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Function Trata_Parametros(Optional objECF As ClassECF) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se houver ECF passado como parâmetro, exibe seus dados
    If Not (objECF Is Nothing) Then

        If objECF.iCodigo > 0 Then

            objECF.iFilialEmpresa = giFilialEmpresa
            
            'Lê ECF no BD a partir do código
            lErro = CF("ECF_Le", objECF)
            If lErro <> SUCESSO And lErro <> 79573 Then gError 79619

            If lErro = SUCESSO Then

                'Exibe os dados do ECF na tela
                lErro = Traz_ECF_Tela(objECF)
                If lErro <> SUCESSO Then gError 79620

                
            'Se não encontrou o ECF no BD
            Else
                'Exibe esse código na tela
                Codigo.Text = objECF.iCodigo
            
            End If
    
        End If

    End If
    
    'Indica que não houve nenhum campo alterado
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 79619, 79620
            'Erros Tratados Dentro das Funções que foram chamadas
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159180)

    End Select

    Exit Function

End Function

Private Sub Teste_Log_Click()
'Função de Teste
Dim lErro As Long
Dim objECF As New ClassECF
Dim objLog As New ClassLog

On Error GoTo Erro_Teste_Log_Click
 
    lErro = Log_Le(objLog)
    If lErro <> SUCESSO And lErro <> 104202 Then gError 104200
    
    lErro = Rede_Desmembra_Log(objECF, objLog)
    If lErro <> SUCESSO And lErro = 104195 Then gError 104196

    Exit Sub
    
Erro_Teste_Log_Click:
    
    Select Case gErr
                                                                    
        Case 104196
            'Erro Tratado Dentro da Função Chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159181)
         
        End Select
         
    Exit Sub
    
End Sub

Function Rede_Desmembra_Log(objECF As ClassECF, objLog As ClassLog) As Long
'Função que informações do banco de Dados e Carrega no Obj

Dim lErro As Long
Dim iPosicao3 As Integer
Dim iPosicao2 As Integer
Dim iIndice As Integer

On Error GoTo Erro_Rede_Desmembra_Log

    'Inicilalização do objRede
    Set objECF = New ClassECF
     
    'Primeira Posição
    iPosicao3 = 1
    'Procura o Primeiro Escape dentro da String sAdmMeiopagto e Armazena a Posição
    iPosicao2 = (InStr(iPosicao3, objLog.sLog, Chr(vbKeyEscape)))
    iIndice = 0
    
    Do While iPosicao2 <> 0
        
       iIndice = iIndice + 1
        'Recolhe os Dados do Banco de Dados e Coloca no objAdmMeioPagto
        Select Case iIndice
            
            Case 1: objECF.iCodigo = StrParaInt(Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3))
            Case 2: objECF.iFilialEmpresa = StrParaInt(Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3))
            Case 3: objECF.iCaixa = Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3)
            Case 4: objECF.iImpressoraECF = StrParaInt(Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3))
            Case 5: objECF.iImpressoraCheque = Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3)
            Case 6: objECF.iHorarioVerao = StrParaInt(Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3))
            
        End Select
        
        'Atualiza as Posições
        iPosicao3 = iPosicao2 + 1
        iPosicao2 = (InStr(iPosicao3, objLog.sLog, Chr(vbKeyEscape)))
    
    
    Loop
        lErro = Traz_ECF_Tela(objECF)
        If lErro <> SUCESSO Then gError 104279
        
        Rede_Desmembra_Log = SUCESSO
        
        Exit Function
        
        
Erro_Rede_Desmembra_Log:

    Select Case gErr
        Case 104279
            'Erro Tradado Dentro da Função Chamada
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159182)

        End Select
    
    Exit Function
    

End Function

Function Log_Le(ByVal objLog As ClassLog) As Long

Dim lErro As Long
Dim tLog As typeLog
Dim lComando As Long

On Error GoTo Erro_Log_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 104197

    'Inicializa o Buffer da Variáveis String
    tLog.sLog1 = String(STRING_CONCATENACAO, 0)
    tLog.sLog2 = String(STRING_CONCATENACAO, 0)
    tLog.sLog3 = String(STRING_CONCATENACAO, 0)
    tLog.sLog4 = String(STRING_CONCATENACAO, 0)

    'Seleciona código e nome dos meios de pagamentos da tabela AdmMeioPagto
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Operacao, Log1, Log2, Log3, Log4 , Data , Hora FROM Log ", tLog.lNumIntDoc, tLog.iOperacao, tLog.sLog1, tLog.sLog2, tLog.sLog3, tLog.sLog4, tLog.dtData, tLog.dHora)
    If lErro <> SUCESSO Then gError 104198

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 104199


    If lErro = AD_SQL_SUCESSO Then

        'Carrega o objLog com as Infromações de bonco de dados
        objLog.lNumIntDoc = tLog.lNumIntDoc
        objLog.iOperacao = tLog.iOperacao
        objLog.sLog = tLog.sLog1 & tLog.sLog2 & tLog.sLog3 & tLog.sLog4
        objLog.dtData = tLog.dtData
        objLog.dHora = tLog.dHora

    End If

    If lErro = AD_SQL_SEM_DADOS Then gError 104202
    
    Log_Le = SUCESSO

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

Erro_Log_Le:

    Log_Le = gErr

   Select Case gErr

    Case gErr

        Case 104198, 104199
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOG", gErr)
    
        Case 104202
            Call Rotina_Erro(vbOKOnly, "ERRO_LOG_NAO_EXISTENTE", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159183)

        End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function



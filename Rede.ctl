VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl Rede 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame Frame3 
      Caption         =   "Redes"
      Height          =   5160
      Left            =   6675
      TabIndex        =   27
      Top             =   645
      Width           =   2760
      Begin VB.ListBox Redes 
         Height          =   4740
         ItemData        =   "Rede.ctx":0000
         Left            =   120
         List            =   "Rede.ctx":0002
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   285
         Width           =   2490
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dados para Importação de Extrato"
      Height          =   2850
      Left            =   90
      TabIndex        =   22
      Top             =   2940
      Width           =   6525
      Begin VB.TextBox NomeDiretorio 
         Height          =   285
         Left            =   1530
         MaxLength       =   100
         TabIndex        =   9
         Top             =   2250
         Width           =   4350
      End
      Begin VB.CommandButton BotaoProcurar 
         Caption         =   "..."
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
         Left            =   5880
         TabIndex        =   10
         Top             =   2190
         Width           =   555
      End
      Begin VB.ComboBox Bandeira 
         Height          =   315
         ItemData        =   "Rede.ctx":0004
         Left            =   1545
         List            =   "Rede.ctx":0014
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   495
         Width           =   1860
      End
      Begin VB.ComboBox CodContaCorrente 
         Height          =   315
         Left            =   1530
         TabIndex        =   7
         Top             =   1065
         Width           =   1845
      End
      Begin MSMask.MaskEdBox Estabelecimento 
         Height          =   315
         Left            =   1530
         TabIndex        =   8
         Top             =   1635
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Diretório:"
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
         Index           =   2
         Left            =   735
         TabIndex        =   26
         Top             =   2265
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bandeira:"
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
         Index           =   2
         Left            =   705
         TabIndex        =   25
         Top             =   540
         Width           =   825
      End
      Begin VB.Label LabelCtaCorrente 
         AutoSize        =   -1  'True
         Caption         =   "Conta Corrente:"
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
         Left            =   180
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   24
         Top             =   1110
         Width           =   1350
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Estabelecimento:"
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
         Index           =   1
         Left            =   60
         TabIndex        =   23
         Top             =   1680
         Width           =   1470
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identificação"
      Height          =   2070
      Left            =   90
      TabIndex        =   17
      Top             =   645
      Width           =   6510
      Begin VB.ComboBox Filial 
         Height          =   315
         Left            =   4560
         TabIndex        =   5
         Top             =   1530
         Width           =   1860
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   2160
         Picture         =   "Rede.ctx":0040
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Numeração Automática"
         Top             =   390
         Width           =   300
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
         Left            =   2910
         TabIndex        =   2
         Top             =   405
         Value           =   1  'Checked
         Width           =   900
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   1530
         TabIndex        =   0
         Top             =   375
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   3
         Mask            =   "###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Nome 
         Height          =   315
         Left            =   1530
         TabIndex        =   3
         Top             =   945
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Cliente 
         Height          =   300
         Left            =   1530
         TabIndex        =   4
         Top             =   1530
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
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
         Index           =   1
         Left            =   855
         TabIndex        =   21
         Top             =   420
         Width           =   660
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
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
         Index           =   0
         Left            =   960
         TabIndex        =   20
         Top             =   1005
         Width           =   555
      End
      Begin VB.Label ClienteLabel 
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
         Left            =   855
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   19
         Top             =   1590
         Width           =   660
      End
      Begin VB.Label Label1 
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
         Index           =   7
         Left            =   4050
         TabIndex        =   18
         Top             =   1590
         Width           =   465
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7290
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   45
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "Rede.ctx":012A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   585
         Picture         =   "Rede.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1110
         Picture         =   "Rede.ctx":040E
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "Rede.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
End
Attribute VB_Name = "Rede"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" _
                                  (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" _
                                  (ByVal pidList As Long, _
                                  ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
                                  (ByVal lpString1 As String, ByVal _
                                  lpString2 As String) As Long

Private Type BrowseInfo
   hWndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type

'A tela de Redes mesmo no caixa central independente
'só vai operar com clientes cadastrados no backoffice
'para evitar usar um cliente que pode estar cadastrado com outro
'codigo em outra filial, mas que deveria ser o mesmo.

'Variáveis Globais a Tela Redes
Public iAlterado As Integer
Dim iClienteAlterado As Integer

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoContaCorrenteInt As AdmEvento
Attribute objEventoContaCorrenteInt.VB_VarHelpID = -1

'Property Variables:
Dim m_Caption As String
Event Unload()

Private Sub Ativo_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Redes"
    Call Form_Load
    
End Function

Public Function Name() As String
'Nome da Tela
    Name = "Rede"
    
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

Private Sub Bandeira_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Bandeira_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub BotaoProxNum_Click()
'Gera um novo número disponível para código da Rede

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click
    
    'Chama a função que gera o sequencial do Código Automático para a nova rede
    lErro = CF("Config_Obter_Inteiro_Automatico", "LojaConfig", "NUM_PROX_REDE", "Redes", "Codigo", iCodigo)
    If lErro <> SUCESSO Then gError 99453

    'Exibe o novo código na tela
    Codigo.Text = CStr(iCodigo)
        
    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 99453
            'Erro tratado na rotina chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166530)
    
    End Select

    Exit Sub

End Sub

Private Sub Cliente_Change()
    iAlterado = REGISTRO_ALTERADO
    iClienteAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim colCodigoNome As New AdmColCodigoNome
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer

On Error GoTo Erro_Cliente_Validate

    'se o cliente tiver sido alterado
    If iClienteAlterado = REGISTRO_ALTERADO Then
        
        'se cliente estiver preenchido
        If Len(Trim(Cliente.Text)) <> 0 Then
        
            'leio o cliente
            lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
            If lErro <> SUCESSO Then gError 113810
            
            'leio as filiais do cliente
            lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
            If lErro <> SUCESSO Then gError 113811
            
            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", Filial, colCodigoNome)
            
            If colCodigoNome.Count = 1 Or iCodFilial <> 0 Then
            
                If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ
                
                'Seleciona filial na Combo Filial
                Call CF("Filial_Seleciona", Filial, iCodFilial)
            
            End If

            
        End If
        
        iClienteAlterado = 0
    
    End If

    Exit Sub
    
Erro_Cliente_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 113810, 113811
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166531)
    
    End Select
    
    Exit Sub

End Sub

Private Sub ClienteLabel_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection
Dim sSelecaoSQL As String

    'Preenche NomeReduzido com o cliente da tela
    objCliente.sNomeReduzido = Cliente.Text
    
    If colSelecao Is Nothing Then Set colSelecao = New Collection

    'Preenche o parâmetro a ser passado
    colSelecao.Add CLIENTE_ATIVO
    
    sSelecaoSQL = "Ativo=?"
    
    'Chama Tela ClienteLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente, sSelecaoSQL)

End Sub

Private Sub Codigo_GotFocus()
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
End Sub

Private Sub Estabelecimento_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NomeDiretorio_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim colCodigoNome As New AdmColCodigoNome
Dim iCodFilial As Integer
Dim lErro As Long

On Error GoTo Erro_objEventoCliente_evSelecao

    Set objCliente = obj1
    
    Cliente.Text = objCliente.lCodigo
    
    Call Cliente_Validate(False)
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoCliente_evSelecao:
    
    Select Case gErr
        
        Case 113812 To 113814
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166532)
    
    End Select
    
    Exit Sub

End Sub

Public Sub Filial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Filial_Click()

Dim lErro As Long

On Error GoTo Erro_Filial_Click

    iAlterado = REGISTRO_ALTERADO
       
    If Filial.ListIndex = -1 Then Exit Sub
    
    Exit Sub
    
Erro_Filial_Click:

    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166533)
            
    End Select

    Exit Sub

End Sub

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 110106

    'Se não encontrou o CÓDIGO
    If lErro = 6730 Then

        'Verifica se o cliente foi digitado
        If Len(Trim(Cliente.Text)) = 0 Then gError 110107

        sCliente = Cliente.Text
        objFilialCliente.iCodFilial = iCodigo

        'Pesquisa se existe Filial com o código extraído
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError 110108

        If lErro = 17660 Then gError 110109

        'Coloca na tela a Filial lida
        Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome

    End If

    'Não encontrou a STRING
    If lErro = 6731 Then gError 110110

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case 110106, 110108

        Case 110107
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 110109
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)

                If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientesLoja", objFilialCliente)
            Else
            End If

        Case 110110
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166534)

    End Select

    Exit Sub

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

''***** fim do trecho a ser copiado ******
''Inicio Sergio 28/05/2002
Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    lErro = Carrega_Redes()
    If lErro <> SUCESSO Then gError 104205
    
    lErro = Carrega_CodContaCorrente()
    If lErro <> SUCESSO Then gError 18051
    
    Ativo.Value = MARCADO
    
    Set objEventoCliente = New AdmEvento
    Set objEventoContaCorrenteInt = New AdmEvento

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        'Erros tratados na rotina chamadora
        Case 104205

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166535)

    End Select
    
    Exit Sub

End Sub

Function Carrega_Redes() As Long

Dim lErro As Long
Dim objRede As New ClassRede
Dim colRedes As New Collection

On Error GoTo Erro_Carrega_Redes

    'Função que lê as Redes no Banco de Dados
    lErro = CF("Redes_Le_Todas", colRedes)
    If lErro <> SUCESSO Then gError 104207

    'Carrega a List Box Redes com os Dados aramazenados na Collection
    For Each objRede In colRedes
    
        Redes.AddItem objRede.iCodigo & SEPARADOR & objRede.sNome
        Redes.ItemData(Redes.NewIndex) = objRede.iCodigo
        
    Next
    
    Carrega_Redes = SUCESSO

    Exit Function

Erro_Carrega_Redes:

    Carrega_Redes = gErr

    Select Case gErr
    
        Case 104206
        'Erro Tratados dentro das Função Chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166536)

        End Select

    Exit Function

End Function


Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objRede As New ClassRede

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Redes"

    'Le os dados da Tela AdmMeioPagto
    lErro = Move_Tela_Memoria(objRede)
    If lErro <> SUCESSO Then gError 104213

    'Preenche a coleção colCampoValor, com nome do campo,
    colCampoValor.Add "Codigo", objRede.iCodigo, 0, "Codigo"
    colCampoValor.Add "Nome", objRede.sNome, STRING_REDE_NOME, "Nome"
    colCampoValor.Add "FilialEmpresa", objRede.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "Cliente", objRede.lCliente, 0, "Cliente"
    colCampoValor.Add "FilialCli", objRede.iFilialCli, 0, "FilialCli"
    colCampoValor.Add "Ativo", objRede.iAtivo, 0, "Ativo"
    
    'Utilizado na hora de passar o parâmetro FilialEmpresa
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr
        'Erro tratado na rotina chamadora
        Case 104213
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166537)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD
Dim lErro As Long
Dim objRede As New ClassRede

On Error GoTo Erro_Tela_Preenche

    objRede.iCodigo = colCampoValor.Item("Codigo").vValor
            
    If objRede.iCodigo > 0 Then
        
        'Carrega objAdmMeioPagto com os dados passados em colCampoValor
        objRede.sNome = colCampoValor.Item("Nome").vValor
        objRede.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
        objRede.iAtivo = colCampoValor.Item("Ativo").vValor
        
        'Traz dados de Redes para a Tela
        lErro = Traz_Rede_Tela(objRede)
        If lErro <> SUCESSO Then gError 104214
        
    End If
        
    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 104214
        'Erro tratado na rotina chamadora
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166538)

    End Select
    
    Exit Sub

End Sub

Function Move_Tela_Memoria(objRede As ClassRede) As Long
'Move os dados da tela para o objRede
Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_Move_Tela_Memoria
      
    'Move a FilialEmpresa que esta sendo Referenciada para a Memória
    objRede.iFilialEmpresa = giFilialEmpresa
    'Move o Codigo Para Memoria
    objRede.iCodigo = StrParaInt(Codigo.Text)
    
    objRede.sNome = Nome.Text
    
    If Len(Trim(Cliente.Text)) <> 0 Then
        'preenche o cliente pelo nome reduzido
        objCliente.sNomeReduzido = Trim(Cliente.Text)
        
        'lê o cliente pelo nome reduzido
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 113817
        
        'se não encontrar o cliente-> erro
        If lErro = 12348 Then gError 113818
    End If
    
    objRede.lCliente = objCliente.lCodigo
    
    If Ativo.Value = vbUnchecked Then
        objRede.iAtivo = REDE_INATIVO
    Else
        objRede.iAtivo = REDE_ATIVO
    End If
    
    objRede.iBandeira = Codigo_Extrai(Bandeira.Text)
    objRede.iCodConta = Codigo_Extrai(CodContaCorrente.Text)
    objRede.sEstabelecimento = Estabelecimento.Text
    objRede.sDirImportacaoExtrato = NomeDiretorio.Text
    
    If Filial.ListIndex <> -1 Then objRede.iFilialCli = Filial.ItemData(Filial.ListIndex)
    
    Move_Tela_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
    
    Select Case gErr
    
        Case 113817
        
        Case 113818
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigoLoja)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166539)
        
        End Select

    Exit Function
    
End Function

Function Traz_Rede_Tela(objRede As ClassRede) As Long
'Função que Traz as Informações da Admnistradoras contida no objAdmMeioPagto para Tela
Dim lErro As Long

On Error GoTo Erro_Traz_Rede_Tela

    Call Limpa_Tela_Rede
    
    'Traz o Codigo para a Tela
    Codigo.Text = objRede.iCodigo
    
    'Traz o Nome para a Tela
    Nome.Text = objRede.sNome
    
    lErro = CF("Rede_Le", objRede)
    If lErro <> SUCESSO And lErro <> 104244 Then gError 113819
    
    If lErro = 104244 Then gError 113820
    
    'Coloca o Cliente na Tela
    If objRede.lCliente > 0 Then
        Cliente.Text = objRede.lCliente
        Call Cliente_Validate(bSGECancelDummy)
    End If
    
    If objRede.iFilialCli > 0 Then
        'Coloca a Filial na Tela
        Filial.Text = objRede.iFilialCli
        Call Filial_Validate(bSGECancelDummy)
    End If
    
    If objRede.iAtivo = REDE_ATIVO Then
        Ativo.Value = vbChecked
    Else
        Ativo.Value = vbUnchecked
    End If
    
    If objRede.iCodConta <> 0 Then
        CodContaCorrente.Text = objRede.iCodConta
        Call CodContaCorrente_Validate(bSGECancelDummy)
    End If
    Estabelecimento.Text = objRede.sEstabelecimento
    NomeDiretorio.Text = objRede.sDirImportacaoExtrato
    
    Bandeira.ListIndex = objRede.iBandeira
    
    'Demonstra que não Houve Alteração na Tela
    iAlterado = 0
    
    Traz_Rede_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Rede_Tela:

    Traz_Rede_Tela = gErr

    Select Case gErr
    
        Case 113819
        
        Case 113820
            Call Rotina_Erro(vbOKOnly, "ERRO_REDE_NAO_ENCONTRADA", gErr, objRede.iCodigo)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166540)
        
        End Select
        
        Exit Function
        
End Function

Function Trata_Parametros(Optional objRede As ClassRede) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se houver POS passado como parâmetro, exibe seus dados
    If Not (objRede Is Nothing) Then

        objRede.iFilialEmpresa = giFilialEmpresa

        If objRede.iCodigo > 0 Then

            'Lê POS no BD a partir do código
            lErro = CF("Rede_Le", objRede)
            If lErro <> SUCESSO And lErro <> 104244 Then gError 104216
            
            If lErro = SUCESSO Then

                'Exibe os dados de AdmMeioPagto
                lErro = Traz_Rede_Tela(objRede)
                If lErro <> SUCESSO Then gError 104217
            Else
    
                Codigo.Text = objRede.iCodigo
                Nome.Text = objRede.sNome
                    
            End If

        End If
    
    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 104216, 104217
        'Erro tratado dentro da Função Chamadora
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166541)

    End Select

    Exit Function

End Function

Private Sub Redes_DblClick()
'Função que Traz para a Tela a rede Através de um DuploClick na List Redes

Dim lErro As Long
Dim objRede As New ClassRede

On Error GoTo Erro_Redes_DblClick
   
    'As Informações Retiradas da List e Inseridas no objRede
    objRede.iCodigo = Redes.ItemData(Redes.ListIndex)
    objRede.sNome = Nome_Extrai(Redes.List(Redes.ListIndex))
    objRede.iFilialEmpresa = giFilialEmpresa
    
    'Chama a Função que Preenche a Tela com os dados Lidos do Bd
    lErro = Traz_Rede_Tela(objRede)
    If lErro <> SUCESSO Then gError 104223
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Exit Sub
    
Erro_Redes_DblClick:

    Select Case gErr
    
        Case 104223
            'Erro Tratado dentro da Função Chamadora
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166542)

    End Select

    Exit Sub
    
End Sub

Private Sub Codigo_Change()
'Determina se Houve Mudança

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Nome_Change()
'Determina se Houve Mudança

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub BotaoGravar_Click()
'Função que Inicializa a Gravação de Novo Registro

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    'Chamada da Função Gravar Registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 104224
    
    'Limpa a Tela
     Call Limpa_Tela_Rede
     
    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoGravar_Click:
    
    Select Case gErr
            
        Case 104224
            'Erro Tratada dentro da Função Chamadora
            
        Case 112773
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REDE_NAO_PERMITE_ALTERACAO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166543)

    End Select

    Exit Sub
    
End Sub
             
Function Gravar_Registro() As Long
'Função de Gravação de Rede

Dim objRede As New ClassRede
Dim lErro As Long
Dim sNome As String

On Error GoTo Erro_Gravar_Registro

    'Verifica se o campo Código esta preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 104225
    
    'Verifica se o campo Nome esta preenchido
    If Len(Trim(Nome.Text)) = 0 Then gError 104226
    
    'se o cliente não estiver preenchido-> erro
    If Len(Trim(Cliente.Text)) = 0 Then gError 113815
    
    'se a filial do cliente não estiver preechida-> erro
    If Len(Trim(Filial.Text)) = 0 Then gError 113816

    'Move para a memória os campos da Tela
    lErro = Move_Tela_Memoria(objRede)
    If lErro <> SUCESSO Then gError 104227
    
    sNome = objRede.sNome
    
    'Utilização para incluir FilialEmpresa como parâmetro
    lErro = Trata_Alteracao(objRede, objRede.iCodigo, objRede.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 104264

    'Chama a Função que Grava rede na Tabela
    lErro = CF("Rede_Grava", objRede)
    If lErro <> SUCESSO Then gError 104228
    
    'Exclui a Rede Através do Codigo da List
    Call Rede_Exclui_List(objRede, sNome)
    
    'Inclui a Rede Utilizando o Código da List
    Call Rede_Inclui_List(objRede, sNome)
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:
   
    Gravar_Registro = gErr
        
        Select Case gErr
        
            Case 113815
                Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
            
            Case 113816
                Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_INFORMADA", gErr)
            
            Case 104225
                lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
                
            Case 104226
                lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_NAO_PREENCHIDO", gErr)
            
            Case 104228, 104264, 104227
                'Erro Tratado Dentro da Função
                    
            Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166544)

        End Select
        
    Exit Function
    
End Function


Public Function Nome_Extrai(sTexto As String) As String
'Função que retira de um texto no formato "Codigo - Nome" apenas o nome.

Dim iPosicao As Integer
Dim sString As String

    iPosicao = InStr(1, sTexto, "-")
    sString = Mid(sTexto, iPosicao + 1)
    
    Nome_Extrai = sString
    
    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objRede As New ClassRede
Dim vbMsgRes As VbMsgBoxResult
Dim sNome As String

On Error GoTo Erro_BotaoExcluir_Click
    
    'verifica se o Codigo Está Preenchido senão Erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 104240
    
    If Codigo.Text = REDE_VISANET Or Codigo.Text = REDE_REDECARD Or Codigo.Text = REDE_TECBAN Then gError 112774
    
    'Para Saber qual é a FilialEmpresa que Esta sendo Referenciada
    objRede.iFilialEmpresa = giFilialEmpresa
    
    'Passa o codigo para a leitura no banco de dados
    objRede.iCodigo = Codigo.Text
    
    'Lê a Rede no Banco e Trazer o objRede
    lErro = CF("Rede_Le", objRede)
    If lErro <> SUCESSO And lErro <> 104244 Then gError 104241
    
    'Se não for encontrado a Rede no Bd
    If lErro = 104244 Then gError 104245
    
    'Envia aviso perguntando se realmente deseja excluir Rede
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_DESEJA_EXCLUIR_REDE", objRede.iCodigo)

    If vbMsgRes = vbYes Then

        'Exclui Rede
        lErro = CF("Rede_Exclui", objRede)
        If lErro <> SUCESSO Then gError 104243
        
        'Exclui a Rede na List Redes
        Call Rede_Exclui_List(objRede, sNome)
    
    End If
    
    'Limpa a Tela
    Call Limpa_Tela_Rede
    
    'Fechar o Comando de Setas
    Call ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0
    
    Exit Sub
        
Erro_BotaoExcluir_Click:

    Select Case gErr
    
        Case 104240
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 104241, 104243, 104265
            'Erro Tratado Dentro da Função Chamadora
        
        Case 104245
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REDE_NAO_ENCONTRADA", gErr, objRede.iCodigo)
        
        Case 112774
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REDE_NAO_PERMITE_ALTERACAO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166545)

    End Select

    Exit Sub
    
End Sub


Private Sub BotaoLimpar_Click()
'Função que tem as chamadas para as Funções que limpam a tela
Dim lErro As Long

On Error GoTo Erro_Botaolimpar_Click
                    
    'Limpa todo os Contoles menos Combo e Label's
    Call Teste_Salva(Me, iAlterado)
    'Limpa Tela de Redes
    
    Call Limpa_Tela_Rede
    'Fecha o comando das setas se estiver aberto
    
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then gError 104261
    
    iAlterado = 0
    
    Exit Sub
        
Erro_Botaolimpar_Click:

    Select Case gErr
        Case 104261
            'Erro Tratado dentro da Função Chamadora
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166546)

    End Select
    
    Exit Sub
        
End Sub

Sub Limpa_Tela_Rede()
'Função que Limpa a Tela de Rede
    
    'Função que Ja Existe no Sistema Limpa a Tela
    Call Limpa_Tela(Me)
    
    Filial.Clear
    
    Redes.ListIndex = -1
    Bandeira.ListIndex = -1
    CodContaCorrente.ListIndex = -1
    
    iClienteAlterado = 0

    Ativo.Value = vbChecked
    
End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoCliente = Nothing
    Set objEventoContaCorrenteInt = Nothing
    
    'Libera a referência da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Private Sub Rede_Inclui_List(objRede As ClassRede, sNome As String)
'Adiciona na ListBox informações da Rede
Dim iIndice As Integer
    
    If objRede.iCodigo = REDE_REDECARD Or objRede.iCodigo = REDE_TECBAN Or objRede.iCodigo = REDE_VISANET Then objRede.sNome = sNome
        
    Redes.AddItem objRede.iCodigo & SEPARADOR & objRede.sNome
    Redes.ItemData(Redes.NewIndex) = objRede.iCodigo
    
    Exit Sub

End Sub

Private Sub Rede_Exclui_List(objRede As ClassRede, sNome As String)
'Percorre a ListBox de Rede para remover a informação em questão

Dim iIndice As Integer
    
    'Percorre a listBox
    For iIndice = 0 To Redes.ListCount - 1
        'se o Codigo For Igual então é Excluida da List
        If Redes.ItemData(iIndice) = objRede.iCodigo Then
            If objRede.iCodigo = REDE_REDECARD Or objRede.iCodigo = REDE_TECBAN Or objRede.iCodigo = REDE_VISANET Then sNome = Nome_Extrai(Redes.List(iIndice))
            Redes.RemoveItem (iIndice)
            Exit For
        End If
     Next

End Sub

Private Sub BotaoFechar_Click()

'Função que Fecha a Tela
    Unload Me

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Cliente Then
            Call ClienteLabel_Click
        ElseIf Me.ActiveControl Is CodContaCorrente Then
            Call LblConta_Click
        End If

    End If

End Sub

Private Sub CodContaCorrente_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CodContaCorrente_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CodContaCorrente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_CodContaCorrente_Validate
    
    If Len(Trim(CodContaCorrente.Text)) = 0 Then Exit Sub

    'verifica se esta preenchida com o item selecionado na ComboBox CodContacOrrente
    If CodContaCorrente.Text = CodContaCorrente.List(CodContaCorrente.ListIndex) Then Exit Sub

    'Tenta selecionar a conta corrente na combo
    lErro = Combo_Seleciona(CodContaCorrente, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 18120

    If lErro = 6730 Then
    
        'Pega o codigo que estana combo
        objContaCorrenteInt.iCodigo = iCodigo
        
        'Procura no BD
        lErro = CF("ContaCorrenteInt_Le", objContaCorrenteInt.iCodigo, objContaCorrenteInt)
        If lErro <> SUCESSO And lErro <> 11807 Then Error 18121
    
        'Se nao estiver cadastrada --> Erro
        If lErro = 11807 Then Error 18122
                
        'Se estiver cadastrada põe na tela
        CodContaCorrente.Text = CStr(objContaCorrenteInt.iCodigo) & SEPARADOR & objContaCorrenteInt.sNomeReduzido

    ElseIf lErro = 6731 Then
    
        Error 18119

    End If
    
    Exit Sub

Erro_CodContaCorrente_Validate:

    Cancel = True

    Select Case Err

        Case 18119
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, CodContaCorrente.Text)
            
        Case 18120, 18121
        
        Case 18122
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CODCONTACORRENTE_INEXISTENTE", objContaCorrenteInt.iCodigo)
        
            If vbMsgRes = vbYes Then
                'Lembrar de manter na tela o numero passado como parametro
                Call Chama_Tela("CtaCorrenteInt", objContaCorrenteInt)
            Else
            End If
        
        Case 43533
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_PERTENCE_FILIAL", Err, CodContaCorrente.Text, giFilialEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158922)

    End Select

    Exit Sub

End Sub

Private Sub LblConta_Click()
'chama browse de conta corrente

Dim objContasCorrentesInternas As New ClassContasCorrentesInternas
Dim colSelecao As New Collection

    If Len(Trim(CodContaCorrente.Text)) > 0 Then objContasCorrentesInternas.iCodigo = Codigo_Extrai(CodContaCorrente.Text)

    Call Chama_Tela("CtaCorrenteLista", colSelecao, objContasCorrentesInternas, objEventoContaCorrenteInt)

End Sub

Private Sub objEventoContaCorrenteInt_evSelecao(obj1 As Object)

Dim objContaCorrenteInt As ClassContasCorrentesInternas

    Set objContaCorrenteInt = obj1
    
    CodContaCorrente.Text = objContaCorrenteInt.iCodigo
    Call CodContaCorrente_Validate(bSGECancelDummy)
    
    Me.Show

End Sub

Private Function Carrega_CodContaCorrente() As Long
'Carrega as contas correntes na combo de contas correntes

Dim lErro As Long
Dim colCodigoNomeRed As New AdmColCodigoNome
Dim objCodigoNome As New AdmCodigoNome
Dim iFilialAux As Integer

On Error GoTo Erro_Carrega_CodContaCorrente

    iFilialAux = giFilialEmpresa
    giFilialEmpresa = EMPRESA_TODA

    'Le o nome e o codigo de todas a contas correntes
    lErro = CF("ContasCorrentesInternas_Le_CodigosNomesRed", colCodigoNomeRed)
    If lErro <> SUCESSO Then Error 18054

    For Each objCodigoNome In colCodigoNomeRed
    
        'Insere na combo de contas correntes
        CodContaCorrente.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        CodContaCorrente.ItemData(CodContaCorrente.NewIndex) = objCodigoNome.iCodigo

    Next
    
    giFilialEmpresa = iFilialAux

    Carrega_CodContaCorrente = SUCESSO

    Exit Function

Erro_Carrega_CodContaCorrente:

    giFilialEmpresa = iFilialAux

    Carrega_CodContaCorrente = Err

    Select Case Err

        Case 18054

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158906)

    End Select

    Exit Function

End Function

Private Sub BotaoProcurar_Click()

Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo

On Error GoTo Erro_BotaoProcurar_Click

    szTitle = "Localização física dos arquivos .html"
    With tBrowseInfo
        .hWndOwner = Me.hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With

    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
       
        NomeDiretorio.Text = sBuffer
        Call NomeDiretorio_Validate(bSGECancelDummy)
  
    End If
  
    Exit Sub

Erro_BotaoProcurar_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192856)

    End Select

    Exit Sub
  
End Sub

Private Sub NomeDiretorio_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iPOS As Integer

On Error GoTo Erro_NomeDiretorio_Validate

    If Len(Trim(NomeDiretorio.Text)) = 0 Then Exit Sub
    
    If Right(NomeDiretorio.Text, 1) <> "\" And Right(NomeDiretorio.Text, 1) <> "/" Then
        iPOS = InStr(1, NomeDiretorio.Text, "/")
        If iPOS = 0 Then
            NomeDiretorio.Text = NomeDiretorio.Text & "\"
        Else
            NomeDiretorio.Text = NomeDiretorio.Text & "/"
        End If
    End If

    If Len(Trim(Dir(NomeDiretorio.Text, vbDirectory))) = 0 Then gError 192857

    Exit Sub

Erro_NomeDiretorio_Validate:

    Cancel = True

    Select Case gErr

        Case 192857, 76
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_INVALIDO", gErr, NomeDiretorio.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192858)

    End Select

    Exit Sub

End Sub

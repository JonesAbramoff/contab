VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ChequeNEsp 
   ClientHeight    =   3930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8235
   KeyPreview      =   -1  'True
   ScaleHeight     =   3930
   ScaleWidth      =   8235
   Begin VB.Frame Frame1 
      Caption         =   "Cheque Não Especificado"
      Height          =   1335
      Left            =   105
      TabIndex        =   13
      Top             =   60
      Width           =   6345
      Begin MSMask.MaskEdBox Sequencial 
         Height          =   300
         Left            =   1185
         TabIndex        =   14
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin VB.Label LabelECF 
         AutoSize        =   -1  'True
         Caption         =   "ECF:"
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
         Left            =   720
         TabIndex        =   23
         Top             =   930
         Width           =   420
      End
      Begin VB.Label LabelCupom 
         AutoSize        =   -1  'True
         Caption         =   "Cupom Fiscal (COO):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   3330
         TabIndex        =   22
         Top             =   930
         Width           =   1770
      End
      Begin VB.Label CupomFiscal 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5160
         TabIndex        =   21
         Top             =   885
         Width           =   1110
      End
      Begin VB.Label ECF 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1185
         TabIndex        =   20
         Top             =   885
         Width           =   945
      End
      Begin VB.Label Label6 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2535
         TabIndex        =   19
         Top             =   405
         Width           =   510
      End
      Begin VB.Label LabelSequencial 
         AutoSize        =   -1  'True
         Caption         =   "Sequencial:"
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
         Left            =   135
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   18
         Top             =   405
         Width           =   1020
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Bom Para:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   4215
         TabIndex        =   17
         Top             =   405
         Width           =   885
      End
      Begin VB.Label DataDepositoNEsp 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5160
         TabIndex        =   16
         Top             =   360
         Width           =   1110
      End
      Begin VB.Label ValorNEsp 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   3105
         TabIndex        =   15
         Top             =   360
         Width           =   945
      End
   End
   Begin MSMask.MaskEdBox CPFCGC 
      Height          =   300
      Left            =   6180
      TabIndex        =   5
      Top             =   1980
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   529
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      PromptInclude   =   0   'False
      MaxLength       =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##############"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox DataDeposito 
      Height          =   300
      Left            =   3945
      TabIndex        =   6
      Top             =   2010
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   529
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Conta 
      Height          =   300
      Left            =   2130
      TabIndex        =   7
      Top             =   2040
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   529
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      PromptInclude   =   0   'False
      MaxLength       =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Numero 
      Height          =   300
      Left            =   3105
      TabIndex        =   8
      Top             =   2010
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   529
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      PromptInclude   =   0   'False
      MaxLength       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "######"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Agencia 
      Height          =   300
      Left            =   1290
      TabIndex        =   9
      Top             =   2025
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   529
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      PromptInclude   =   0   'False
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Banco 
      Height          =   300
      Left            =   495
      TabIndex        =   10
      Top             =   2025
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   529
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      PromptInclude   =   0   'False
      MaxLength       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "###"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   300
      Left            =   4950
      TabIndex        =   11
      Top             =   1995
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
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
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6510
      ScaleHeight     =   495
      ScaleWidth      =   1605
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   150
      Width           =   1665
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ChequeNEsp.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "ChequeNEsp.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1110
         Picture         =   "ChequeNEsp.ctx":068C
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridCheques 
      Height          =   2130
      Left            =   135
      TabIndex        =   4
      Top             =   1740
      Width           =   7950
      _ExtentX        =   14023
      _ExtentY        =   3757
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cheques Especificados:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   180
      TabIndex        =   12
      Top             =   1530
      Width           =   2055
   End
End
Attribute VB_Name = "ChequeNEsp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 Option Explicit

'Variáveis Globáis

Private Const NUM_CHEQUES_GRID = 20

Public objGridCheques As AdmGrid

Dim iAlterado  As Integer
Dim iSetasNaoEsp As Integer

Private WithEvents objEventoCheque As AdmEvento
Attribute objEventoCheque.VB_VarHelpID = -1

Dim iGrid_Banco_Col As Integer
Dim iGrid_Agencia_Col As Integer
Dim iGrid_Conta_Col As Integer
Dim iGrid_Numero_Col As Integer
Dim iGrid_DataDeposito_Col As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_CPFCGC_Col As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Especificação de Cheque"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ChequeNEsp"

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

Private Sub LabelSequencial_Click()

Dim objCheque As New ClassChequePre
Dim colSelecao As New Collection
Dim sSelecao As String

On Error GoTo Erro_LabelSequencial_Click

    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then

        objCheque.lSequencialLoja = StrParaLong(Sequencial.Text)

        sSelecao = "Localizacao = ?"
        colSelecao.Add CHEQUEPRE_LOCALIZACAO_LOJA
    
    Else

        objCheque.lSequencialBack = StrParaLong(Sequencial.Text)

        sSelecao = "Localizacao = ?"
        colSelecao.Add CHEQUEPRE_LOCALIZACAO_BACKOFFICE

    End If

    'Chama o Browser ChequeLojaLista
    Call Chama_Tela("ChequeLojaLista", colSelecao, objCheque, objEventoCheque, sSelecao)

    Exit Sub

Erro_LabelSequencial_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144436)

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

'***** fim do trecho a ser copiado ******

Public Sub Form_Load()
 'Inicialização da Tela de Cheque

Dim lErro As Long

    Set objEventoCheque = New AdmEvento

    Set objGridCheques = New AdmGrid

    'Inicializa o Grid de Cheques
    lErro = Inicializa_Grid_Cheques(objGridCheques)
    If lErro <> SUCESSO Then gError 105007

    'Define que não Houve Alteração
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 105007

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144437)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Public Function Inicializa_Grid_Cheques(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Cheques

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Banco")
    objGridInt.colColuna.Add ("Agência")
    objGridInt.colColuna.Add ("Conta")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Bom Para")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Cliente (CPF/CNPJ)")
    
    'Controles que participam do Grid
    objGridInt.colCampo.Add (Banco.Name)
    objGridInt.colCampo.Add (Agencia.Name)
    objGridInt.colCampo.Add (Conta.Name)
    objGridInt.colCampo.Add (Numero.Name)
    objGridInt.colCampo.Add (DataDeposito.Name)
    objGridInt.colCampo.Add (Valor.Name)
    objGridInt.colCampo.Add (CPFCGC.Name)

    

    'Colunas da Grid
    iGrid_Banco_Col = 1
    iGrid_Agencia_Col = 2
    iGrid_Conta_Col = 3
    iGrid_Numero_Col = 4
    iGrid_DataDeposito_Col = 5
    iGrid_Valor_Col = 6
    iGrid_CPFCGC_Col = 7
    
    'Grid do GridInterno
    objGridInt.objGrid = GridCheques

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_CHEQUES_GRID + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 5
    
    'Largura da primeira coluna
    GridCheques.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Cheques = SUCESSO

    Exit Function

End Function

Function Trata_Parametros(Optional objCheque As ClassChequePre) As Long
'Trata os parametros

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se há um operador preenchido
    If Not (objCheque Is Nothing) Then

        lErro = Traz_Cheque_Tela(objCheque)
        If lErro <> SUCESSO And lErro <> 105035 Then gError 104319

        If lErro <> SUCESSO Then

                'Limpa a Tela
                Call Limpa_Tela(Me)

                'Mantém o Código do operador na tela
                Sequencial.Text = objCheque.lSequencial

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 104319
            'Erros Tratados Dentro da Função Chamadas
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144438)

    End Select

    iAlterado = 0

    Exit Function

End Function

Private Sub Sequencial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Sequencial_GotFocus()

    Call MaskEdBox_TrataGotFocus(Sequencial, iAlterado)

End Sub

Private Sub Sequencial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCheque As New ClassChequePre

On Error GoTo Erro_Sequencial_Validate

    'se não estiver preenchido-> sai
    If Len(Trim(Sequencial.Text)) > 0 Then

        'critica o campo
        lErro = Long_Critica(Sequencial.Text)
        If lErro <> SUCESSO Then gError 105031

        objCheque.lSequencialLoja = StrParaLong(Sequencial.Text)
        objCheque.lSequencialBack = StrParaLong(Sequencial.Text)
        objCheque.iFilialEmpresaLoja = giFilialEmpresa
        objCheque.iFilialEmpresa = giFilialEmpresa
    
        'Se o Sequencial do Cheque nao for nulo Traz o Cheque para a tela
        lErro = Traz_Cheque_Tela(objCheque)
        If lErro <> SUCESSO And lErro <> 105035 Then gError 105772

        If lErro = 105035 Then gError 105773

    End If

    Exit Sub

Erro_Sequencial_Validate:

    Cancel = True

    Select Case gErr

        Case 105031, 105772

        Case 105773
            Call Rotina_Erro(vbOKOnly, "ERRO_CHEQUE_INEXISTENTE", gErr, Sequencial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144439)

    End Select

    Exit Sub

End Sub

Function Traz_Cheque_Tela(objCheque As ClassChequePre) As Long
'Função que le as Informações de um cheque passado como parametro e Traz estes dados para a tela.

Dim lErro As Long

On Error GoTo Erro_Traz_Cheque_Tela

    'Função que Lê no Banco de Dados Informações do Cheque Refereciado
    lErro = CF("Cheque_Le", objCheque)
    If lErro <> SUCESSO And lErro <> 104346 Then gError 105034

    'Se não for Encontrado Registro no Banco de Dados Referente ao Cheque
    If lErro = 104346 Then gError 105035

    'se o cheque for especificado ==> erro
    If objCheque.iNaoEspecificado = CHEQUE_ESPECIFICADO Then gError 105006

    Sequencial.Text = objCheque.lSequencial
    ValorNEsp.Caption = Format(objCheque.dValor, "Standard")

    DataDepositoNEsp.Caption = Format(objCheque.dtDataDeposito, "dd/mm/yyyy")

    If objCheque.lCupomFiscal <> 0 Then CupomFiscal.Caption = CStr(objCheque.lCupomFiscal)
    
    If objCheque.iECF <> 0 Then ECF.Caption = CStr(objCheque.iECF)

    Traz_Cheque_Tela = SUCESSO

    Exit Function

Erro_Traz_Cheque_Tela:

    Traz_Cheque_Tela = gErr

    Select Case gErr

        Case 105006
            Call Rotina_Erro(vbOKOnly, "ERRO_CHEQUE_ESPECIFICADO", gErr, objCheque.lSequencial)

        Case 105034, 105035

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144440)

    End Select

    Exit Function

End Function

Private Sub objEventoCheque_evSelecao(obj1 As Object)

Dim objCheque As ClassChequePre
Dim lErro As Long
Dim lCodigoMsgErro As Long

On Error GoTo Erro_objEventoCheque_evSelecao

    Set objCheque = obj1

    'Move os dados para a tela
    lErro = Traz_Cheque_Tela(objCheque)
    If lErro <> SUCESSO And lErro <> 105035 Then gError 104322

    'Cheque não Encontrado no Banco de Dados
    If lErro = 105035 Then gError 104321

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoCheque_evSelecao:

    Select Case gErr

        Case 104321
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CHEQUE_INEXISTENTE", gErr, objCheque.lSequencial)

        Case 104322

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144441)

    End Select

    Exit Sub

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da ceula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
        
        'Verifica qual a coluna atual do Grid
        Select Case objGridInt.objGrid.Col
    
            Case iGrid_Banco_Col
                lErro = Saida_Celula_Banco(objGridInt)
                If lErro <> SUCESSO Then gError 105009
    
            Case iGrid_Agencia_Col
                lErro = Saida_Celula_Agencia(objGridInt)
                If lErro <> SUCESSO Then gError 105012
    
            Case iGrid_Conta_Col
                lErro = Saida_Celula_Conta(objGridInt)
                If lErro <> SUCESSO Then gError 105014
    
            Case iGrid_Numero_Col
                lErro = Saida_Celula_Numero(objGridInt)
                If lErro <> SUCESSO Then gError 105016
    
            Case iGrid_DataDeposito_Col
                lErro = Saida_Celula_DataDeposito(objGridInt)
                If lErro <> SUCESSO Then gError 105019
    
            Case iGrid_Valor_Col
                lErro = Saida_Celula_Valor(objGridInt)
                If lErro <> SUCESSO Then gError 105022
    
            Case iGrid_CPFCGC_Col
                lErro = Saida_Celula_CPFCGC(objGridInt)
                If lErro <> SUCESSO Then gError 105029
    
        End Select
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 105030

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 105009, 105012, 105014, 105016, 105019, 105022, 105029, 105030

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144442)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Banco(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Banco que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Banco

    Set objGridInt.objControle = Banco

    'Verifica o preenchimento do Banco
    If Len(Trim(Banco.Text)) > 0 Then
        
        lErro = Inteiro_Critica(Banco.Text)
        If lErro <> SUCESSO Then gError 105010
        
        'Acrescenta uma linha no Grid se for o caso
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    End If
        
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 105011
    
    Saida_Celula_Banco = SUCESSO

    Exit Function

Erro_Saida_Celula_Banco:

    Saida_Celula_Banco = gErr

    Select Case gErr

        Case 105010, 105011
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144443)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Agencia(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Agencia que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Agencia

    Set objGridInt.objControle = Agencia

    If Len(Trim(Agencia.Text)) > 0 Then
    
        'Acrescenta uma linha no Grid se for o caso
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 105013
    
    Saida_Celula_Agencia = SUCESSO

    Exit Function

Erro_Saida_Celula_Agencia:

    Saida_Celula_Agencia = gErr

    Select Case gErr

        Case 105013
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144444)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Conta(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Conta que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Conta

    Set objGridInt.objControle = Conta

    If Len(Trim(Conta.Text)) > 0 Then
    
        'Acrescenta uma linha no Grid se for o caso
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 105015
    
    Saida_Celula_Conta = SUCESSO

    Exit Function

Erro_Saida_Celula_Conta:

    Saida_Celula_Conta = gErr

    Select Case gErr

        Case 105015
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144445)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Numero(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Numero que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Numero

    Set objGridInt.objControle = Numero

    If Len(Trim(Numero.Text)) > 0 Then

        lErro = Long_Critica(Numero.Text)
        If lErro <> SUCESSO Then gError 105017

        If CLng(Numero.Text) < 1 Then gError 105018

        'Acrescenta uma linha no Grid se for o caso
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 105016
    
    Saida_Celula_Numero = SUCESSO

    Exit Function

Erro_Saida_Celula_Numero:

    Saida_Celula_Numero = gErr

    Select Case gErr

        Case 105016
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 105017
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 105018
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_MENOR_QUE_UM", gErr, Numero.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144446)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DataDeposito(objGridInt As AdmGrid) As Long
'Faz a crítica da célula DataDeposito que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DataDeposito

    Set objGridInt.objControle = DataDeposito

    If Len(Trim(DataDeposito.ClipText)) > 0 Then

        'Verifica se a data final é válida
        lErro = Data_Critica(DataDeposito.Text)
        If lErro <> SUCESSO Then gError 105020

        'Acrescenta uma linha no Grid se for o caso
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 105021
    
    Saida_Celula_DataDeposito = SUCESSO

    Exit Function

Erro_Saida_Celula_DataDeposito:

    Saida_Celula_DataDeposito = gErr

    Select Case gErr

        Case 105020, 105021
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144447)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Valor(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Valor que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridInt.objControle = Valor

    If Len(Trim(Valor.Text)) > 0 Then

        'critica o valor
        lErro = Valor_Positivo_Critica(Valor.Text)
        If lErro <> SUCESSO Then gError 105023
            
        'Acrescenta uma linha no Grid se for o caso
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
            
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 105024
    
    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = gErr

    Select Case gErr

        Case 105023, 105024
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144448)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_CPFCGC(objGridInt As AdmGrid) As Long
'Faz a crítica da célula CPFCGC que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_CPFCGC

    Set objGridInt.objControle = CPFCGC

    'Se CGCCPF não foi preenchido -- Exit Sub
    If Len(Trim(CPFCGC.ClipText)) > 0 Then

        Select Case Len(Trim(CPFCGC.Text))
    
            Case STRING_CPF 'CPF
    
                'Critica Cpf
                lErro = Cpf_Critica(CPFCGC.Text)
                If lErro <> SUCESSO Then gError 105025
    
                'Formata e coloca na Tela
                CPFCGC.Format = FORMATO_CPF
                
    
            Case STRING_CGC 'CGC
    
                'Critica CGC
                lErro = Cgc_Critica(CPFCGC.Text)
                If lErro <> SUCESSO Then gError 105026
    
                'Formata e Coloca na Tela
                CPFCGC.Format = FORMATO_CGC
    
            Case Else
                gError 105027
    
        End Select

        'Acrescenta uma linha no Grid se for o caso
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 105028
    
    
    Saida_Celula_CPFCGC = SUCESSO

    Exit Function

Erro_Saida_Celula_CPFCGC:

    Saida_Celula_CPFCGC = gErr

    Select Case gErr

        Case 105025, 105026, 105028
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 105027
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC_CPF", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144449)

    End Select

    Exit Function

End Function

Public Sub Banco_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Banco_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCheques)

End Sub

Public Sub Banco_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCheques)

End Sub

Public Sub Banco_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCheques.objControle = Banco
    lErro = Grid_Campo_Libera_Foco(objGridCheques)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Agencia_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Agencia_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCheques)

End Sub

Public Sub Agencia_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCheques)

End Sub

Public Sub Agencia_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCheques.objControle = Agencia
    lErro = Grid_Campo_Libera_Foco(objGridCheques)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Conta_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Conta_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCheques)

End Sub

Public Sub Conta_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCheques)

End Sub

Public Sub Conta_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCheques.objControle = Conta
    lErro = Grid_Campo_Libera_Foco(objGridCheques)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Numero_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Numero_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCheques)

End Sub

Public Sub Numero_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCheques)

End Sub

Public Sub Numero_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCheques.objControle = Numero
    lErro = Grid_Campo_Libera_Foco(objGridCheques)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub DataDeposito_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataDeposito_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCheques)

End Sub

Public Sub DataDeposito_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCheques)

End Sub

Public Sub DataDeposito_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCheques.objControle = DataDeposito
    lErro = Grid_Campo_Libera_Foco(objGridCheques)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Valor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Valor_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCheques)

End Sub

Public Sub Valor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCheques)

End Sub

Public Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCheques.objControle = Valor
    lErro = Grid_Campo_Libera_Foco(objGridCheques)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub CPFCGC_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CPFCGC_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCheques)

End Sub

Public Sub CPFCGC_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCheques)

End Sub

Public Sub CPFCGC_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCheques.objControle = CPFCGC
    lErro = Grid_Campo_Libera_Foco(objGridCheques)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub GridCheques_Click()

    Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridCheques, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        'Variavel não definida
        Call Grid_Entrada_Celula(objGridCheques, iAlterado)
    End If

End Sub

Private Sub GridCheques_EnterCell()

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

Private Sub GridCheques_Scroll()

    Call Grid_Scroll(objGridCheques)

End Sub

Private Sub GridCheques_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridCheques)

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objCheque As New ClassChequePre

On Error GoTo Erro_Tela_Extrai

    sTabela = "ChequePreCupomView"

    'Armazena os dados presentes na tela em objOperador
    lErro = Move_Tela_Memoria(objCheque)
    If lErro <> SUCESSO Then gError 105032

    'Definição de Onde Está Sendo Trabalhado para setar Sistema de Setas Chaves diferentes
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
    
        'Filtro
        colSelecao.Add "Localizacao", OP_IGUAL, CHEQUEPRE_LOCALIZACAO_LOJA
        
    Else

        'Filtro
        colSelecao.Add "Localizacao", OP_IGUAL, CHEQUEPRE_LOCALIZACAO_BACKOFFICE

    End If

    'Preenche a colecao de campos-valores com os dados de objOperador
    colCampoValor.Add "SequencialBack", objCheque.lSequencialBack, 0, "SequencialBack"
    colCampoValor.Add "SequencialLoja", objCheque.lSequencialLoja, 0, "SequencialLoja"
    colCampoValor.Add "FilialEmpresaLoja", objCheque.iFilialEmpresaLoja, 0, "FilialEmpresaLoja"

    'Filtro
    colSelecao.Add "FilialEmpresaLoja", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 105032

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144450)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim objCheque As New ClassChequePre
Dim lErro As Long

On Error GoTo Erro_Tela_Preenche

    objCheque.lSequencialLoja = colCampoValor.Item("SequencialLoja").vValor
    objCheque.lSequencialBack = colCampoValor.Item("SequencialBack").vValor
    objCheque.iFilialEmpresaLoja = colCampoValor.Item("FilialEmpresaLoja").vValor
    objCheque.iFilialEmpresa = objCheque.iFilialEmpresaLoja

    'Se o Sequencial do Cheque nao for nulo Traz o Cheque para a tela
    lErro = Traz_Cheque_Tela(objCheque)
    If lErro <> SUCESSO Then gError 105033

    iAlterado = 0

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 105033

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144451)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
'Botão Limpa Tela

Dim lErro As Long

On Error GoTo Erro_Botaolimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 105037

    'Função que Limpa a Tela de Cheque
    Call Limpa_Tela_Cheque

    'Função que Fecha o Comando de Setas
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_Botaolimpar_Click:

    Select Case gErr

        Case 105037

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144452)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_Cheque()
'Função que limpa Tela

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Cheque

    lErro = Limpa_Tela(Me)
    If lErro <> SUCESSO Then gError 105036

    Call Grid_Limpa(objGridCheques)

    DataDepositoNEsp.Caption = ""
    ValorNEsp.Caption = ""
    CupomFiscal.Caption = ""
    ECF.Caption = ""
    
    Exit Sub

Erro_Limpa_Tela_Cheque:

    Select Case gErr

        Case 105036

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144453)

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

Public Sub form_unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoCheque = Nothing

    'Fecha o comando de setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Private Sub BotaoGravar_Click()
'Função que Inicializa a Gravação de Novo Registro

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chamada da Função Gravar Registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 105037

    'Limpa a Tela
     Call Limpa_Tela_Cheque

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 105037

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144454)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long
'Função que Verifica se os Campos Obrigatórios da Tela Cheque estão Preenchidos e chama a função Grava_Registro

Dim objCheque As New ClassChequePre
Dim lErro As Long
Dim colCheque As New Collection

On Error GoTo Erro_Gravar_Registro

    'Verifica se o campo Código esta preenchido
    If Len(Trim(Sequencial.Text)) = 0 Then gError 105038

    'valida o conteudo do grid
    lErro = Valida_Grid_Cheques()
    If lErro <> SUCESSO Then gError 105041

    'Move para a memória os campos da Tela
    lErro = Move_Tela_Memoria(objCheque)
    If lErro <> SUCESSO Then gError 105039
    
    lErro = Move_Tela_Memoria_Grid(colCheque)
    If lErro <> SUCESSO Then gError 105040

   'Chama a Função que Grava Cheque na Tabela
    lErro = CF("Cheque_Grava", objCheque, colCheque)
    If lErro <> SUCESSO Then gError 105051

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

        Select Case gErr

            Case 105038
              lErro = Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIAL_NAO_PREENCHIDO", gErr)

            Case 105039, 105040, 105041, 105051

            Case Else
                lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144455)

        End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(objCheque As ClassChequePre) As Long
'Lê os dados que estão na tela Cheque e coloca em objOperador

On Error GoTo Erro_Move_Tela_Memoria

    'Definição de Onde Está Sendo Trabalhado para setar Sistema de Setas Chaves diferentes
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then

        objCheque.lSequencialLoja = StrParaLong(Sequencial.Text)

    Else
    
        objCheque.lSequencialBack = StrParaLong(Sequencial.Text)

    End If

    'Diz qual é a filial empresa que está sendo Referênciada
    objCheque.iFilialEmpresaLoja = giFilialEmpresa
    objCheque.iFilialEmpresa = giFilialEmpresa

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144456)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria_Grid(colCheque As Collection) As Long
'Recolhe do Grid os dados

Dim lErro As Long
Dim iIndice As Integer
Dim objCheque As ClassChequePre
Dim sCPFCGC As String

On Error GoTo Erro_Move_Tela_Memoria_Grid

    For iIndice = 1 To objGridCheques.iLinhasExistentes

        Set objCheque = New ClassChequePre
    
        'Armazena os dados do cheque
        objCheque.iBanco = StrParaInt(GridCheques.TextMatrix(iIndice, iGrid_Banco_Col))
        objCheque.sAgencia = GridCheques.TextMatrix(iIndice, iGrid_Agencia_Col)
        objCheque.sContaCorrente = GridCheques.TextMatrix(iIndice, iGrid_Conta_Col)
        objCheque.lNumero = StrParaLong(GridCheques.TextMatrix(iIndice, iGrid_Numero_Col))
        objCheque.dtDataDeposito = CDate(GridCheques.TextMatrix(iIndice, iGrid_DataDeposito_Col))
        objCheque.dValor = StrParaDbl(GridCheques.TextMatrix(iIndice, iGrid_Valor_Col))
        sCPFCGC = GridCheques.TextMatrix(iIndice, iGrid_CPFCGC_Col)
        If Len(sCPFCGC) = 14 Then
            objCheque.sCPFCGC = Left(sCPFCGC, 3) & Mid(sCPFCGC, 5, 3) & Mid(sCPFCGC, 9, 3) & Right(sCPFCGC, 2)
        Else
            objCheque.sCPFCGC = Left(sCPFCGC, 2) & Mid(sCPFCGC, 4, 3) & Mid(sCPFCGC, 8, 3) & Mid(sCPFCGC, 12, 4) & Right(sCPFCGC, 2)
        End If
    
        colCheque.Add objCheque

    Next

    Move_Tela_Memoria_Grid = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria_Grid:

    Move_Tela_Memoria_Grid = gErr

    Select Case gErr

        Case 20767
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 27682

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144457)

    End Select

    Exit Function

End Function

Private Function Valida_Grid_Cheques() As Long
'valida o conteudo do grid

Dim iIndice As Integer
Dim lErro As Long
Dim dValorTotal As Double

On Error GoTo Erro_Valida_Grid_Cheques

    'Verifica se há itens no grid
    If objGridCheques.iLinhasExistentes = 0 Then gError 105042
    
    'para cada item do grid
    For iIndice = 1 To objGridCheques.iLinhasExistentes
        
        'se o banco do cheque nao foi preenchido
        If Len(Trim(GridCheques.TextMatrix(iIndice, iGrid_Banco_Col))) = 0 Then gError 105043
        
        'se a agencia do cheque nao foi preenchida
        If Len(Trim(GridCheques.TextMatrix(iIndice, iGrid_Agencia_Col))) = 0 Then gError 105044
        
        'se a conta do cheque nao foi preenchida
        If Len(Trim(GridCheques.TextMatrix(iIndice, iGrid_Conta_Col))) = 0 Then gError 105045
        
        'se o numero do cheque nao foi preenchido
        If Len(Trim(GridCheques.TextMatrix(iIndice, iGrid_Numero_Col))) = 0 Then gError 105046
        
        'se a data bom para do cheque nao foi preenchida
        If Len(Trim(GridCheques.TextMatrix(iIndice, iGrid_DataDeposito_Col))) = 0 Then gError 105047
        
        'se o valor do cheque nao foi preenchido
        If Len(Trim(GridCheques.TextMatrix(iIndice, iGrid_Valor_Col))) = 0 Then gError 105048
        
        dValorTotal = CDbl(GridCheques.TextMatrix(iIndice, iGrid_Valor_Col))

    Next

    'se o valor dos cheques especificados ultrapassar o valor do cheque especificado ==> erro
    If dValorTotal > StrParaDbl(ValorNEsp.Caption) Then gError 105050

    Valida_Grid_Cheques = SUCESSO

    Exit Function

Erro_Valida_Grid_Cheques:

    Valida_Grid_Cheques = gErr

    Select Case gErr

        Case 105042
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_CHEQUES_GRID", gErr)

        Case 105043
            Call Rotina_Erro(vbOKOnly, "ERRO_GRIDCHEQUE_BANCO_NAO_PREENCHIDO", gErr, iIndice)
        
        Case 105044
            Call Rotina_Erro(vbOKOnly, "ERRO_GRIDCHEQUE_AGENCIA_NAO_PREENCHIDA", gErr, iIndice)
        
        Case 105045
            Call Rotina_Erro(vbOKOnly, "ERRO_GRIDCHEQUE_CONTA_NAO_PREENCHIDA", gErr, iIndice)
        
        Case 105046
            Call Rotina_Erro(vbOKOnly, "ERRO_GRIDCHEQUE_NUMERO_NAO_PREENCHIDO", gErr, iIndice)
        
        Case 105047
            Call Rotina_Erro(vbOKOnly, "ERRO_GRIDCHEQUE_DATADEPOSITO_NAO_PREENCHIDA", gErr, iIndice)
        
        Case 105048
            Call Rotina_Erro(vbOKOnly, "ERRO_GRIDCHEQUE_VALOR_NAO_PREENCHIDO", gErr, iIndice)
        
        Case 105050
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORCHEQUES_MAIOR_NESPECIFICADO", gErr, dValorTotal, ValorNEsp.Caption)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144458)

    End Select

    Exit Function

End Function

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Sequencial Then
            Call LabelSequencial_Click

        End If

    End If


End Sub



'Private Sub BotaoLe_Click()
'
''Função de Teste
'Dim lErro As Long
'Dim objCheque As New ClassChequePre
'Dim objLog As New ClassLog
'
'On Error GoTo Erro_Teste_Log_Click
'
'    lErro = Log_Le(objLog)
'    If lErro <> SUCESSO And lErro <> 104202 Then gError 104200
'
'    lErro = ChequePre_Desmembra_Log(objCheque, objLog)
'    If lErro <> SUCESSO And lErro = 104195 Then gError 104196
'
'    Exit Sub
'
'Erro_Teste_Log_Click:
'
'    Select Case gErr
'
'        Case 104196
'            'Erro Tratado Dentro da Função Chamadora
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144459)
'
'        End Select
'
'    Exit Sub
'
'End Sub
'
'
'
'
'
'
'
'

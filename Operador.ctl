VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl Operador 
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7455
   KeyPreview      =   -1  'True
   ScaleHeight     =   4695
   ScaleWidth      =   7455
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5160
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   90
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "Operador.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "Operador.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "Operador.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "Operador.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ListBox Operadores 
      Height          =   3570
      Left            =   5160
      TabIndex        =   14
      Top             =   1005
      Width           =   2145
   End
   Begin VB.Frame Frame2 
      Caption         =   "Operador"
      Height          =   4560
      Left            =   120
      TabIndex        =   13
      Top             =   30
      Width           =   4905
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
         Left            =   3360
         TabIndex        =   23
         Top             =   330
         Value           =   1  'Checked
         Width           =   900
      End
      Begin VB.TextBox Confirmacao 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1875
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1866
         Width           =   1410
      End
      Begin VB.CheckBox Gerente 
         Caption         =   "Gerente"
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
         Left            =   1875
         TabIndex        =   5
         Top             =   2388
         Width           =   1170
      End
      Begin VB.TextBox Senha 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1875
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1344
         Width           =   1410
      End
      Begin VB.ComboBox Vendedores 
         Height          =   315
         Left            =   1875
         TabIndex        =   8
         Top             =   3840
         Width           =   2340
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   315
         Left            =   2520
         Picture         =   "Operador.ctx":0994
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Numeração Automática"
         Top             =   300
         Width           =   300
      End
      Begin VB.CheckBox Desconto 
         Caption         =   "Pode dar desconto"
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
         Left            =   1875
         TabIndex        =   6
         Top             =   2850
         Width           =   1950
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   1875
         TabIndex        =   0
         Top             =   300
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox LimiteDesconto 
         Height          =   315
         Left            =   1875
         TabIndex        =   7
         Top             =   3312
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   3
         Format          =   "0\%"
         Mask            =   "###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Nome 
         Height          =   315
         Left            =   1875
         TabIndex        =   2
         Top             =   822
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   50
         Format          =   "0\%"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         Caption         =   "Limite de Desconto:"
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
         Index           =   5
         Left            =   120
         TabIndex        =   22
         Top             =   3364
         Width           =   1710
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
         Left            =   1200
         TabIndex        =   21
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Senha:"
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
         Index           =   2
         Left            =   1215
         TabIndex        =   19
         Top             =   1404
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Confirmação:"
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
         Index           =   6
         Left            =   705
         TabIndex        =   18
         Top             =   1926
         Width           =   1125
      End
      Begin VB.Label Label1 
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
         Height          =   195
         Index           =   3
         Left            =   945
         TabIndex        =   17
         Top             =   3870
         Width           =   885
      End
      Begin VB.Label LabelNome 
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
         Left            =   1290
         TabIndex        =   20
         Top             =   885
         Width           =   555
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Operadores"
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
      Index           =   4
      Left            =   5160
      TabIndex        =   16
      Top             =   765
      Width           =   990
   End
End
Attribute VB_Name = "Operador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Declarações Globais
Dim iAlterado As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Operador"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "Operador"

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

Private Sub Ativo_Click()

    iAlterado = REGISTRO_ALTERADO

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

On Error GoTo Erro_Operadores_Form_Load

    iAlterado = 0
    
    'Carrega a listbox com os operadores da Filial Empresa.
    lErro = Operadores_Carrega()
    If lErro <> SUCESSO Then gError 81000
    
    'Carrega a combo com os vendedores da mesma FilialEmpresa
    lErro = Vendedores_Carrega()
    If lErro <> SUCESSO Then gError 104281
    
    'Define que não Houve Alteração
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Operadores_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 81000, 104281
            'Erros
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163670)

    End Select
    
    Exit Sub

End Sub

Private Function Operadores_Carrega() As Long
'Carrega a ListBox

Dim lErro As Long
Dim objOperador As ClassOperador
Dim colOperador As New Collection

On Error GoTo Erro_Operadores_Carrega

    'Le todos os Operadores da Filial Empresa
    lErro = CF("Operador_Le_Todos", colOperador)
    If lErro <> SUCESSO And lErro <> 81005 Then gError 81006

    'Se encontrou pelo menos um Operador no BD
    If lErro <> 81005 Then
        
        'Adcionar na ListBox Operadores os operadores cadastrados no Banco de Dados
        For Each objOperador In colOperador
        
            Operadores.AddItem objOperador.iCodigo & SEPARADOR & objOperador.sNome
            Operadores.ItemData(Operadores.NewIndex) = objOperador.iCodigo
        
        Next

    End If
    
    Operadores_Carrega = SUCESSO

    Exit Function

Erro_Operadores_Carrega:

    Operadores_Carrega = gErr

    Select Case gErr

        Case 81006
            'Erro Tratados Dentro da Função Chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163671)

    End Select

    Exit Function

End Function


Private Function Vendedores_Carrega() As Long
'Carrega a ListBox

Dim lErro As Long
Dim objVendedor As ClassVendedor
Dim colVendedor As New Collection

On Error GoTo Erro_Vendedores_Carrega

    'Le todos os Operadores da Filial Empresa
    lErro = CF("VendedorFilial_Le_Todos", colVendedor)
    If lErro <> SUCESSO And lErro <> 109490 Then gError 104283

    'Se encontrou pelo menos um Vendedor no BD
    If lErro = SUCESSO Then

        'Adcionar na ComboBox Vendedores os Vendedores cadastrados no Banco de Dados
        For Each objVendedor In colVendedor
        
            Vendedores.AddItem objVendedor.iCodigo & SEPARADOR & objVendedor.sNomeReduzido
            Vendedores.ItemData(Vendedores.NewIndex) = objVendedor.iCodigo
        
        Next
    
    End If
    
    Vendedores_Carrega = SUCESSO

    Exit Function

Erro_Vendedores_Carrega:

    Vendedores_Carrega = gErr

    Select Case gErr

        Case 104283
            'Erro Tratados Dentro da Função Chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163672)

    End Select

    Exit Function

End Function

Function Trata_Parametros(Optional objOperador As ClassOperador) As Long
'Trata os parametros

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Trata_Parametros

    'Se há um operador preenchido
    If Not (objOperador Is Nothing) Then

        'Se objoperador.iCodigo > 0
        If objOperador.iCodigo > 0 Then
            
            'Atribui a Filial Empresa a Qual se  está Trabalhando
            objOperador.iFilialEmpresa = giFilialEmpresa
            
            'Verifica se o operador existe, lendo no BD a partir do código
            lErro = CF("Operador_Le", objOperador)
            If lErro <> SUCESSO And lErro <> 81026 Then gError 81190

            'Se o operador existe
            If lErro = SUCESSO Then
                lErro = Traz_Operador_Tela(objOperador)
                If lErro <> SUCESSO Then gError 81191
                
            'Se o operador não existe
            Else

                'Mantém o Código do operador na tela
                Codigo.Text = CStr(objOperador.iCodigo)

            End If

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 81190, 81191
            'Erros Tratados Dentro da Função Chamadas
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163673)

    End Select

    iAlterado = 0

    Exit Function

End Function

Private Function Move_Tela_Memoria(objOperador As ClassOperador) As Long
'Lê os dados que estão na tela Operador e coloca em objOperador

On Error GoTo Erro_Move_Tela_Memoria

    'Se o codigo não estiver vazio coloca-o no objOperador
    objOperador.iCodigo = StrParaInt(Codigo.ClipText)
    objOperador.sNome = Nome.Text
    objOperador.sSenha = Senha.Text
    objOperador.iGerente = Gerente.Value
    objOperador.iDesconto = Desconto.Value
    objOperador.iLimiteDesconto = StrParaInt(LimiteDesconto.Text)
    objOperador.iCodigoVendedor = Codigo_Extrai(Vendedores.Text)
    
    'Diz qual é a filial empresa que está sendo Referênciada
    objOperador.iFilialEmpresa = giFilialEmpresa

    If Ativo.Value = vbUnchecked Then
        objOperador.iAtivo = OPERADOR_INATIVO
    Else
        objOperador.iAtivo = OPERADOR_ATIVO
    End If


    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163674)

    End Select

    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD
Dim lErro As Long
Dim objOperador As New ClassOperador

On Error GoTo Erro_Tela_Extrai

    sTabela = "Operador"

    'Armazena os dados presentes na tela em objOperador
    lErro = Move_Tela_Memoria(objOperador)
    If lErro <> SUCESSO Then gError 81015

    'Preenche a colecao de campos-valores com os dados de objOperador
    colCampoValor.Add "Codigo", objOperador.iCodigo, 0, "Codigo"
    colCampoValor.Add "FilialEmpresa", objOperador.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "Nome", objOperador.sNome, 50, "Nome"
    colCampoValor.Add "Senha", objOperador.sSenha, 10, "Senha"
    colCampoValor.Add "Desconto", objOperador.iDesconto, 0, "Desconto"
    colCampoValor.Add "LimiteDesconto", objOperador.iLimiteDesconto, 0, "LimiteDesconto"
    colCampoValor.Add "CodVendedor", objOperador.iCodigoVendedor, 0, "CodVendedor"
    colCampoValor.Add "Gerente", objOperador.iGerente, 0, "Gerente"
    colCampoValor.Add "Ativo", objOperador.iAtivo, 0, "Ativo"
    

    'Filtro
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 81015
            'Erro Tratado Dentro da Função Chamada.
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163675)

    End Select

    Exit Sub

End Sub

'Preenche os campos da tela com os correspondentes do BD
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim objOperador As New ClassOperador
Dim lErro As Long

On Error GoTo Erro_Tela_Preenche

    'Passa os dados da colecao de campos-valores para o objOperador
    objOperador.iCodigo = colCampoValor.Item("Codigo").vValor
    objOperador.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
    objOperador.iDesconto = colCampoValor.Item("Desconto").vValor
    objOperador.iLimiteDesconto = colCampoValor.Item("LimiteDesconto").vValor
    objOperador.sNome = colCampoValor.Item("Nome").vValor
    objOperador.sSenha = colCampoValor.Item("Senha").vValor
    objOperador.iCodigoVendedor = colCampoValor.Item("CodVendedor").vValor
    objOperador.iGerente = colCampoValor.Item("Gerente").vValor
    objOperador.iAtivo = colCampoValor.Item("Ativo").vValor
    
    
    If objOperador.iCodigo <> 0 Then

        'Se o Codigo do Operador nao for nulo Traz o Operador para a tela
        lErro = Traz_Operador_Tela(objOperador)
        If lErro <> SUCESSO Then gError 81016

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 81016
            'Erro Tratado Dentro da Função Chamada.
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163676)

    End Select

    Exit Sub

End Sub

Function Traz_Operador_Tela(objOperador As ClassOperador) As Long
'Traz os dados do operador para a tela

Dim iIndice As Integer
Dim lErro As Long


On Error GoTo Erro_Traz_Operador_Tela

    Call Limpa_Tela_Operador

    If objOperador.iAtivo = OPERADOR_ATIVO Then
        Ativo.Value = vbChecked
    Else
        Ativo.Value = vbUnchecked
    End If
    
    'Traz o Codigo para a Tela
    Codigo.Text = objOperador.iCodigo
    
    'Traz o Nome para a Tela
    Nome.Text = objOperador.sNome
        
    'Traz a Senha para tela com Asteriscos
    Senha.Text = objOperador.sSenha
    
    'Traz a Confirmação da Senha Para a Tela
    Confirmacao.Text = objOperador.sSenha
    
    'Marca ou não a Checkbox Gerente
    Gerente.Value = objOperador.iGerente
    
    'Marca ou não a Checkbox de Desconto
    Desconto.Value = objOperador.iDesconto
    Call Desconto_Click
    
    If Desconto.Value = vbChecked Then LimiteDesconto.Text = objOperador.iLimiteDesconto
    
    'Preencher a Combo Vendedor
    For iIndice = 0 To Vendedores.ListCount - 1
        If Vendedores.ItemData(iIndice) = objOperador.iCodigoVendedor Then
            Vendedores.ListIndex = iIndice
            Exit For
        End If
    Next
    
    'Não existe alteração
    iAlterado = 0

    Traz_Operador_Tela = SUCESSO

    Exit Function

Erro_Traz_Operador_Tela:

    Traz_Operador_Tela = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163677)

    End Select

    Exit Function

End Function

Private Sub Codigo_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Nome_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub
Private Sub Senha_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub
Private Sub Confirmacao_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Gerente_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Desconto_Click()
    
    'Se desconto está marcado permite prencher o desconto
    If Desconto.Value = vbChecked Then
        LimiteDesconto.Enabled = True
    Else
        LimiteDesconto.Text = ""
        LimiteDesconto.Enabled = False
    End If
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub LimiteDesconto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Vendedores_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub form_unload(Cancel As Integer)
    
Dim lErro As Long

    lErro = ComandoSeta_Liberar(Me.Name)
    
End Sub

Private Sub Operadores_DblClick()

Dim lErro As Long
Dim objOperador As New ClassOperador

On Error GoTo Erro_Operadores_DblClick

    'Carrego o obj com o Código
    objOperador.iCodigo = Codigo_Extrai(Operadores.List(Operadores.ListIndex))
   
   'Atribui a Filial Empresa a Qual se  está Trabalhando
    objOperador.iFilialEmpresa = giFilialEmpresa
    
    'Carregar o objOperador com os dados do banco de dados
    lErro = CF("Operador_Le", objOperador)
    If lErro <> SUCESSO Then gError 104312
    
    'se não encontrou o operador cadastrado no bando de dados
    If lErro = 81026 Then gError 104315
    
    'Traz o operador para tela
    lErro = Traz_Operador_Tela(objOperador)
    If lErro <> SUCESSO Then gError 81022

    Exit Sub

Erro_Operadores_DblClick:

    Select Case gErr

        Case 81022, 104312
            'Erro Tratado Dentro da Função Chamada
            
        Case 104315
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OPERADOR_NAO_CADASTRADO", gErr, objOperador.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163678)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Gera Código da proximo Operador
    lErro = Operador_Codigo_Automatico(lCodigo)
    If lErro <> SUCESSO Then gError 81030

    Codigo.Text = lCodigo

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 81030

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163679)

    End Select

    Exit Sub

End Sub

Function Operador_Codigo_Automatico(lCodigo As Long) As Long
'Gera o proximo codigo da Tabela de Requisitante

Dim lErro As Long

On Error GoTo Erro_Operador_Codigo_Automatico

    'Chama a rotina que gera o sequencial
    lErro = CF("Config_ObterAutomatico", "LojaConfig", "NUM_PROXIMO_OPERADOR", "Operador", "Codigo", lCodigo)
    If lErro <> SUCESSO Then gError 104288

    Operador_Codigo_Automatico = SUCESSO

    Exit Function

Erro_Operador_Codigo_Automatico:

    Operador_Codigo_Automatico = Err

    Select Case Err

        Case 104288
            'Erro Tratado dentro da Função Chamadora
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 163680)

    End Select

    Exit Function

End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim objOperador As New ClassOperador

On Error GoTo Erro_BotaoGravar_Click

    'Grava os registros na tabela
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 81038

    Call Limpa_Tela_Operador

    iAlterado = 0

    Exit Sub
Erro_BotaoGravar_Click:

    Select Case gErr

        Case 81038
            'Erro Tratado Dentro da Função Chamada
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163681)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long
'Grava um registro no bd

Dim lErro As Long
Dim objOperador As New ClassOperador

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se os campos Obrigatórios estão preenchidos
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 81033
    
    If Len(Trim(Nome.Text)) = 0 Then gError 104289
    
    If Len(Trim(Senha.Text)) = 0 Then gError 104290

    If Len(Trim(Confirmacao.Text)) = 0 Then gError 104291
    
    'Verifica se a Senha é Diferente da Confirmação se for Erro
    If Senha.Text <> Confirmacao.Text Then gError 104292
    
    'Se estiver marcando como sendo dado desconto --> limite tem q estar preenchido
    If Desconto.Value = vbChecked And Len(Trim(LimiteDesconto.Text)) = 0 Then gError 109690
    
    lErro = Move_Tela_Memoria(objOperador)
    If lErro <> SUCESSO Then gError 104293
    
    lErro = Trata_Alteracao(objOperador, objOperador.iFilialEmpresa, objOperador.iCodigo)
    If lErro <> SUCESSO Then gError 32331
    
    lErro = CF("Operador_Grava", objOperador)
    If lErro <> SUCESSO Then gError 104294
    
    'Retirar da Lista para que não Haja Duplicata
    Call Retira_Lista_Operador(objOperador)
    
   'Adiciona na listbox se necessário
    Call Adiciona_Lista_Operador(objOperador)

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 32331, 104294, 104293
            'Erro Tratado Dentro da Função Chamada
            
        Case 81033
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 104289
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_NAO_PREENCHIDO", gErr)

        Case 104290
            Call Rotina_Erro(vbOKOnly, "ERRO_SENHA_NAO_PREENCHIDO", gErr)

        Case 104291
            Call Rotina_Erro(vbOKOnly, "ERRO_CONFIRMACAO_SENHA_NAO_PREENCHIDO", gErr)

        Case 104292
            Call Rotina_Erro(vbOKOnly, "ERRO_CONFIRMACAO_SENHA_INVALIDA", gErr)
        
        Case 109690
            Call Rotina_Erro(vbOKOnly, "ERRO_LIMITEDESCONTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163682)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Function

End Function

Private Sub Adiciona_Lista_Operador(objOperador As ClassOperador)
'Adiciona na ListBox informações do Operador
Dim iIndice As Integer
Dim iInd As Integer
    
    For iInd = 0 To Operadores.ListCount - 1
        If Operadores.ItemData(iInd) > objOperador.iCodigo Then
            Exit For
        End If
    Next
    
    Operadores.AddItem objOperador.iCodigo & SEPARADOR & objOperador.sNome, iInd
    Operadores.ItemData(Operadores.NewIndex) = objOperador.iCodigo
    
    Exit Sub

End Sub

Private Sub Retira_Lista_Operador(objOperador As ClassOperador)
'Percorre a ListBox de OPerador para remover a informação em questão

Dim iIndice As Integer
    'Percorre a listBox
    For iIndice = 0 To Operadores.ListCount - 1
        'se o Codigo For Igual então é Excluida da List
        If Operadores.ItemData(iIndice) = objOperador.iCodigo Then
            Operadores.RemoveItem (iIndice)
            Exit For
        End If
     Next

End Sub

    
Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objOperador As New ClassOperador
Dim objOperador1 As ClassOperador
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o Operador está preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 81053

    objOperador.iCodigo = StrParaInt(Codigo.Text)
    
   'passa para a função de leitura qual é a filial que se Está Trabalhando
    objOperador.iFilialEmpresa = giFilialEmpresa

    'Verifica se o usuário tem Operador
    lErro = CF("Operador_Le", objOperador)
    If lErro <> SUCESSO And lErro <> 81026 Then gError 81054
    
    'Se não Foi encontrado no BD erro
    If lErro = 81026 Then gError 81055

    'Pede a confirmação da exclusão do operador do usuário
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_OPERADOR", objOperador.iCodigo)
    If vbMsgRes = vbYes Then

        lErro = Move_Tela_Memoria(objOperador)
        If lErro <> SUCESSO Then gError 81056

        lErro = CF("Operador_Exclui", objOperador)
        If lErro <> SUCESSO Then gError 81057

        Call Limpa_Tela_Operador
        
        'Atualizar a LisBox Operadores
        Call Retira_Lista_Operador(objOperador)

        iAlterado = 0

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 81053
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 81054, 81056, 81057

        Case 81055
            Call Rotina_Erro(vbOKOnly, "ERRO_OPERADOR_NAO_CADASTRADO", gErr, objOperador.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163683)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

End Sub
    
        
Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_Botaolimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 81067

    Call Limpa_Tela_Operador

    iAlterado = 0

    Exit Sub

Erro_Botaolimpar_Click:

    Select Case gErr

        Case 81067
            'Erro Tratado Dentro da Função Chamada
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163684)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_Operador()
'Limpa a tela

Dim lErro As Long

    Call Limpa_Tela(Me)

    Ativo.Value = vbChecked

    Gerente.Value = vbUnchecked
    Desconto.Value = vbUnchecked
    LimiteDesconto.Enabled = False
    Vendedores.ListIndex = -1
    
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

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

Private Sub LimiteDesconto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dLimiteDesconto As Double

On Error GoTo Erro_LimiteDesconto_Validate
    
    If Len(Trim(LimiteDesconto.Text)) = 0 Then Exit Sub
    
    'Verifica se o Desconto é Zero se for Erro
    lErro = Valor_Positivo_Critica(Trim(LimiteDesconto.Text))
    If lErro <> SUCESSO Then gError 111311

    
    'Verifica se é porcentagem
    lErro = Porcentagem_Critica(LimiteDesconto.Text)
    If lErro <> SUCESSO Then gError 81071

    Exit Sub

Erro_LimiteDesconto_Validate:

    Cancel = True

    Select Case gErr

        Case 81071

        Case 111311
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163685)

    End Select

    Exit Sub

End Sub

Private Sub LimiteDesconto_GotFocus()

    Call MaskEdBox_TrataGotFocus(LimiteDesconto)

End Sub


Private Sub Vendedores_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objVendedor As New ClassVendedor
Dim iFilialEmpresa As Integer

On Error GoTo Erro_Vendedores_Validate

    'se não estiver preenchida sai
    If Len(Trim(Vendedores.Text)) = 0 Then Exit Sub
    
    'se a combo foi selecionada com clique, sai
    If Vendedores.ListIndex <> -1 Then Exit Sub
    
    lErro = Combo_Seleciona(Vendedores, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 113506
    
    'se não achou pelo código
    If lErro = 6730 Then
    
        'preencho a chave de vendedor
        objVendedor.iCodigo = Codigo_Extrai(Vendedores.Text)
        
        iFilialEmpresa = giFilialEmpresa
        
        lErro = CF("VendedorFilial_Le", objVendedor, iFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 109498 Then gError 113504
        
        'se não encontrar-> erro
        If lErro = 109498 Then gError 113505
        
        Vendedores.Text = objVendedor.iCodigo & SEPARADOR & objVendedor.sNomeReduzido

    End If
    
    'se não achou pelo nome
    If lErro = 6731 Then
        
        objVendedor.sNomeReduzido = Nome_Extrai(Vendedores.Text)
        
        iFilialEmpresa = giFilialEmpresa
        
        lErro = CF("VendedorFilial_Le_NomeReduzido", objVendedor, iFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 113503 Then gError 113507
        
        If lErro = 113503 Then gError 113508
        
        Vendedores.Text = objVendedor.iCodigo & SEPARADOR & objVendedor.sNomeReduzido

    End If

    Exit Sub
    
Erro_Vendedores_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 113504, 113506, 113507
        
        Case 113505
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", gErr, objVendedor.iCodigo)
        
        Case 113508
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO1", gErr, objVendedor.sNomeReduzido)
    
    End Select
    
    Exit Sub

End Sub

Private Function Nome_Extrai(sTexto As String) As String
'Função que retira de um texto no formato "Codigo - Nome" apenas o nome.

Dim iPosicao As Integer
Dim sString As String

    iPosicao = InStr(1, sTexto, "-")
    sString = Trim(Mid(sTexto, iPosicao + 1))

    Nome_Extrai = sString

    Exit Function

End Function


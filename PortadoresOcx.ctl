VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl PortadoresOcx 
   ClientHeight    =   4350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5520
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4350
   ScaleWidth      =   5520
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   1425
      Picture         =   "PortadoresOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Numeração Automática"
      Top             =   270
      Width           =   300
   End
   Begin VB.PictureBox Picture4 
      Height          =   555
      Left            =   3660
      ScaleHeight     =   495
      ScaleWidth      =   1635
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   135
      Width           =   1695
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1125
         Picture         =   "PortadoresOcx.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "PortadoresOcx.ctx":0268
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "PortadoresOcx.ctx":079A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ListBox PortadoresList 
      Height          =   1620
      Left            =   210
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   2565
      Width           =   5130
   End
   Begin VB.CheckBox Inativo 
      Caption         =   "Inativo"
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
      Left            =   2370
      TabIndex        =   2
      Top             =   300
      Width           =   915
   End
   Begin VB.ComboBox ComboBanco 
      Height          =   315
      Left            =   915
      TabIndex        =   3
      Top             =   780
      Width           =   2655
   End
   Begin MSMask.MaskEdBox Nome 
      Height          =   300
      Left            =   870
      TabIndex        =   5
      Top             =   1845
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   50
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   285
      Left            =   915
      TabIndex        =   0
      Top             =   270
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox NomeReduzido 
      Height          =   300
      Left            =   1710
      TabIndex        =   4
      Top             =   1305
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   15
      PromptChar      =   "_"
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
      Left            =   180
      TabIndex        =   11
      Top             =   285
      Width           =   660
   End
   Begin VB.Label Label2 
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
      Left            =   240
      TabIndex        =   12
      Top             =   1890
      Width           =   555
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Portadores"
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
      Left            =   240
      TabIndex        =   13
      Top             =   2340
      Width           =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Nome Reduzido:"
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
      TabIndex        =   14
      Top             =   1335
      Width           =   1410
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Banco:"
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
      Left            =   225
      TabIndex        =   15
      Top             =   825
      Width           =   615
   End
End
Attribute VB_Name = "PortadoresOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'DECLARACAO DE VARIAVEIS GLOBAIS
Dim iAlterado As Integer

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim sCodigo As String

On Error GoTo Erro_BotaoProxNum_Click

    'Gera codigo automatico do proximo Portador
    lErro = CF("Portador_Automatico",sCodigo)
    If lErro <> SUCESSO Then Error 57553

    Codigo.Text = sCodigo

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57553
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165064)
    
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
    If lErro <> SUCESSO Then Error 41785

    Call Limpa_Tela_Portador

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 41785

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165065)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objPortador As New ClassPortador

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica preenchimento do código
    If Len(Trim(Codigo.Text)) = 0 Then gError 16504

    'Verifica preenchimento do nome reduzido
    If Len(Trim(NomeReduzido.Text)) = 0 Then gError 16505

    'Verifica preenchimento do nome
    If Len(Trim(Nome.Text)) = 0 Then gError 16506
    
    lErro = Move_Tela_Memoria(objPortador)
    If lErro <> SUCESSO Then gError 16488
    
    lErro = Trata_Alteracao(objPortador, objPortador.iCodigo)
    If lErro <> SUCESSO Then gError 80454

    'Chama função de gravação
    lErro = CF("Portador_Grava",objPortador)
    If lErro <> SUCESSO Then gError 16507

    'Remove o Portador da ListBox PortadoresList
    Call PortadoresList_Exclui(objPortador.iCodigo)

    'Insere o Portador na ListBox PortadoresList
    Call PortadoresList_Adiciona(objPortador)

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 16504
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 16505
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_PREENCHIDO", gErr)

        Case 16506
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_NAO_PREENCHIDO", gErr)

        Case 16488, 16507, 80454
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165066)

    End Select

    Exit Function

End Function

Function Move_Tela_Memoria(objPortador As ClassPortador) As Long
'Move os dados da tela para memória.

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    'Preenche objPortador
    objPortador.iCodigo = StrParaInt(Codigo.Text)
    objPortador.sNomeReduzido = NomeReduzido.Text
    objPortador.sNome = Nome.Text
    objPortador.iInativo = Inativo.Value
    
    If ComboBanco.ListIndex <> -1 Then
        objPortador.iBanco = ComboBanco.ItemData(ComboBanco.ListIndex)
    Else
        objPortador.iBanco = 0
    End If
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165067)

    End Select

    Exit Function

End Function

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Verifica preenchimento do codigo
    If Len(Trim(Codigo.Text)) > 0 Then

        lErro = Inteiro_Critica(Codigo.Text)
        If lErro <> SUCESSO Then Error 16501

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True


    Select Case Err

        Case 16501

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165068)

    End Select

    Exit Sub

End Sub

Private Sub PortadoresList_Adiciona(objPortador As ClassPortador)

    PortadoresList.AddItem objPortador.sNomeReduzido
    PortadoresList.ItemData(PortadoresList.NewIndex) = objPortador.iCodigo

End Sub

Private Sub PortadoresList_Exclui(iCodigo As Integer)

Dim iIndice As Integer

    For iIndice = 0 To PortadoresList.ListCount - 1

        If PortadoresList.ItemData(iIndice) = iCodigo Then

            PortadoresList.RemoveItem iIndice
            Exit For

        End If

    Next

End Sub

Private Sub ComboBanco_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ComboBanco_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ComboBanco_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_ComboBanco_Validate

    'Verifica se o Banco foi preenchido
    If Len(Trim(ComboBanco.Text)) = 0 Then Exit Sub
    
    'Tenta selecionar na combo
    lErro = Combo_Seleciona(ComboBanco, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 41659
    
    'Se não encontra valor que contém CÓDIGO, mas extrai o código
    If lErro = 6730 Then Error 41782
        
    'Não encontrou o valor que era STRING
    If lErro = 6731 Then Error 41783

    Exit Sub
    
Erro_ComboBanco_Validate:

    Cancel = True


    Select Case Err
       
       Case 41659
        
        Case 41782, 41783
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BANCO_NAO_CADASTRADO", Err, ComboBanco.Text)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165069)
    
    End Select
    
    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colPortadoresCodigoNome As AdmColCodigoNome
Dim objPortadorCodigoNome As AdmCodigoNome
Dim colBancosCodigoNome As AdmColCodigoNome
Dim objBancoCodigoNome As AdmCodigoNome

On Error GoTo Erro_Form_Load

    Set colPortadoresCodigoNome = New AdmColCodigoNome

    'leitura dos codigos e nomes reduzidos dos portadores
    lErro = CF("Cod_Nomes_Le","Portador", "Codigo", "NomeReduzido", STRING_PORTADOR_NOME_REDUZIDO, colPortadoresCodigoNome)
    If lErro <> SUCESSO Then Error 16479

    'preenche ListBox PortadoresList com nome reduzido de Portadores do BD
    For Each objPortadorCodigoNome In colPortadoresCodigoNome

        PortadoresList.AddItem objPortadorCodigoNome.sNome
        PortadoresList.ItemData(PortadoresList.NewIndex) = objPortadorCodigoNome.iCodigo

    Next
    
    Set colBancosCodigoNome = New AdmColCodigoNome
    
    lErro = CF("Cod_Nomes_Le","Bancos", "CodBanco", "NomeReduzido", STRING_BANCO_NOME_REDUZIDO, colBancosCodigoNome)
    If lErro <> SUCESSO Then Error 41649
    
    For Each objBancoCodigoNome In colBancosCodigoNome
        
        ComboBanco.AddItem CStr(objBancoCodigoNome.iCodigo) & SEPARADOR & objBancoCodigoNome.sNome
        ComboBanco.ItemData(ComboBanco.NewIndex) = objBancoCodigoNome.iCodigo
        
    Next
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 16479
        
        Case 41649

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165070)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objPortador As New ClassPortador

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Portador"

    'Realiza conversões necessárias de campos da tela para campos do BD
    'A tipagem dos valores DEVE SER A MESMA DO BD
    lErro = Move_Tela_Memoria(objPortador)
    If lErro <> SUCESSO Then Error 33984

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objPortador.iCodigo, 0, "Codigo"
    colCampoValor.Add "Nome", objPortador.sNome, STRING_PORTADOR_NOME, "Nome"
    colCampoValor.Add "NomeReduzido", objPortador.sNomeReduzido, STRING_PORTADOR_NOME_REDUZIDO, "NomeReduzido"
    colCampoValor.Add "Inativo", objPortador.iInativo, 0, "Inativo"
    colCampoValor.Add "Banco", objPortador.iBanco, 0, "Banco"

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 33984

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165071)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objPortador As New ClassPortador

On Error GoTo Erro_Tela_Preenche

    objPortador.iCodigo = colCampoValor.Item("Codigo").vValor

    If objPortador.iCodigo <> 0 Then

        objPortador.sNome = colCampoValor.Item("Nome").vValor
        objPortador.sNomeReduzido = colCampoValor.Item("NomeReduzido").vValor
        objPortador.iInativo = colCampoValor.Item("Inativo").vValor
        objPortador.iBanco = colCampoValor.Item("Banco").vValor

        lErro = Preenche_Tela_Portador(objPortador)
        If lErro <> SUCESSO Then Error 16480
        
        'Desseleciona a List de Portadores
        PortadoresList.ListIndex = -1

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 16480

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165072)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Unload(Cancel As Integer)

 Dim lErro As Long

    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
   
End Sub

Function Preenche_Tela_Portador(objPortador As ClassPortador) As Long
'Exibe os dados do Portador especificada em objPortador

Dim iIndice As Integer

    Codigo.Text = CStr(objPortador.iCodigo)
    Inativo.Value = objPortador.iInativo
    
    If objPortador.iBanco <> 0 Then
        For iIndice = 0 To ComboBanco.ListCount - 1
            If ComboBanco.ItemData(iIndice) = objPortador.iBanco Then
                ComboBanco.ListIndex = iIndice
                Exit For
            End If
        Next
    Else
        ComboBanco.Text = ""
    End If
    
    Nome.Text = objPortador.sNome
    NomeReduzido.Text = objPortador.sNomeReduzido

    iAlterado = 0

    Preenche_Tela_Portador = SUCESSO

End Function

Function Trata_Parametros(Optional objPortador As ClassPortador) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se existir um Portador passado como parametro, exibir seus dados
    If Not (objPortador Is Nothing) Then

        lErro = CF("Portador_Le",objPortador)
        If lErro <> SUCESSO And lErro <> 15971 Then Error 16481

        If lErro = SUCESSO Then

            'Exibe dados do Portador na tela
            lErro = Preenche_Tela_Portador(objPortador)
            If lErro <> SUCESSO Then Error 16482

        Else

            'Limpa a tela
            Call Limpa_Tela_Portador

            'Exibe apenas o código
            Codigo.Text = CStr(objPortador.iCodigo)

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 16481, 16482

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165073)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 16520

    Call Limpa_Tela_Portador
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 16520

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165074)

    End Select

    Exit Sub

End Sub

Function Limpa_Tela_Portador()
'limpa todos os campos de input da tela Portadores

Dim iIndice As Integer
Dim lErro As Long

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Call Limpa_Tela(Me)

    Codigo.Text = ""
    
    'Desmarca combo de bancos
    ComboBanco.ListIndex = -1
    
    'Limpa os campos da tela que não foram limpos pela rotina Limpa_Tela
    Inativo.Value = vbUnchecked

    iAlterado = 0

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

 Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
   
End Sub

Private Sub Inativo_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Nome_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NomeReduzido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NomeReduzido_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NomeReduzido_Validate
    
    'Se está preenchido, testa se começa por letra
    If Len(Trim(NomeReduzido.Text)) > 0 Then

        If Not IniciaLetra(NomeReduzido.Text) Then Error 57820

    End If
    
    Exit Sub

Erro_NomeReduzido_Validate:

    Cancel = True

    
    Select Case Err
    
        Case 57820
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_COMECA_LETRA", Err, NomeReduzido.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165075)
    
    End Select
    
    Exit Sub

End Sub

Private Sub PortadoresList_DblClick()

Dim lErro As Long
Dim iIndice As Integer
Dim objPortador As New ClassPortador

On Error GoTo Erro_PortadoresList_DblClick

    objPortador.iCodigo = PortadoresList.ItemData(PortadoresList.ListIndex)

    'Lê Portador no BD
    lErro = CF("Portador_Le",objPortador)
    If lErro <> SUCESSO And lErro <> 16487 Then Error 16498
    
    'Se não encontrou o Portador --> Erro
    If lErro <> SUCESSO Then Error 16500

    Call Limpa_Tela_Portador
    
    'Mostra dados na tela
    lErro = Preenche_Tela_Portador(objPortador)
    If lErro <> SUCESSO Then Error 16499

    Exit Sub

Erro_PortadoresList_DblClick:

    Select Case Err

        Case 16498, 16499

        Case 16500
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PORTADOR_NAO_CADASTRADO1", Err, objPortador.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165076)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_PORTADORES
    Set Form_Load_Ocx = Me
    Caption = "Portadores"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Portadores"
    
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
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
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

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub


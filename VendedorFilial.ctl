VERSION 5.00
Begin VB.UserControl VendedorFilialOcx 
   ClientHeight    =   3270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4005
   ScaleHeight     =   3270
   ScaleWidth      =   4005
   Begin VB.CommandButton Teste_Log_Click 
      Caption         =   "Teste_Log_Click"
      Height          =   270
      Left            =   285
      TabIndex        =   7
      Top             =   255
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   2115
      ScaleHeight     =   495
      ScaleWidth      =   1575
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   90
      Width           =   1635
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1110
         Picture         =   "VendedorFilial.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   585
         Picture         =   "VendedorFilial.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Limpar"
         Top             =   105
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "VendedorFilial.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ListBox FiliaisEmpresa 
      Height          =   1860
      ItemData        =   "VendedorFilial.ctx":080A
      Left            =   165
      List            =   "VendedorFilial.ctx":080C
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   1245
      Width           =   3585
   End
   Begin VB.ComboBox Vendedores 
      Height          =   315
      Left            =   1200
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   840
      Width           =   2550
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
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   3
      Left            =   225
      TabIndex        =   1
      Top             =   870
      Width           =   885
   End
End
Attribute VB_Name = "VendedorFilialOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Dim m_Caption As String
Event Unload()


'Determina se Houve Alteração na Tela
Dim iAlterado As Integer

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Vendedor Filial"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "VendedorFilial"

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

'Inicio Tela Dia 4/07/02 Sergio Ricardo
Public Sub Form_Load()
'Form  Carrega a List Administradoras e as Combos
Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Carrega a combo de administradoras
    lErro = Carrega_Vendedores()
    If lErro <> SUCESSO Then gError 104495

    'Carrega a combo de Meios de Pagamento
    lErro = Carrega_Filiais()
    If lErro <> SUCESSO Then gError 104496

    'Zera o flag de alterações indicando que não houve nenhuma ainda
    iAlterado = 0

    'Indica que o carregamento da tela aconteceu com sucesso
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        'Erros tratados na rotina chamadora
        Case 104495 To 104496
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175721)

    End Select

    Exit Sub

End Sub

Function Carrega_Vendedores() As Long
'Função que Carrega a Combo de Venderes
Dim lErro As Long
Dim colVendedores As New Collection
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Carrega_Vendedores

    lErro = CF("Vendedor_Le_Todos", colVendedores)
    If lErro <> SUCESSO And lErro <> 104287 Then gError 107500

    'Adcionar os Vendedores cadastrados
    For Each objVendedor In colVendedores
        
        Vendedores.AddItem objVendedor.sNomeReduzido
        Vendedores.ItemData(Vendedores.NewIndex) = objVendedor.iCodigo
    
    Next
    
    Carrega_Vendedores = SUCESSO
    
    Exit Function
    
Erro_Carrega_Vendedores:

    Carrega_Vendedores = gErr

    Select Case gErr

        'Erros tratados na rotina chamadora
        Case 107500
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175722)

    End Select

    Exit Function

End Function

Function Carrega_Filiais() As Long
'Função que Carrega a List de FilialEmpresa
Dim lErro As Long
Dim colFiliais As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Carrega_Filiais

    lErro = CF("FiliaisEmpresas_Le_Empresa", glEmpresa, colFiliais)
    If lErro <> SUCESSO Then gError 107501

    'Adcionar Filiais da Empresa
    For Each objFilialEmpresa In colFiliais
        
        If objFilialEmpresa.iCodFilial <> EMPRESA_TODA Then
            FiliaisEmpresa.AddItem objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome
            FiliaisEmpresa.ItemData(FiliaisEmpresa.NewIndex) = objFilialEmpresa.iCodFilial
        End If
    
    Next
    
    Carrega_Filiais = SUCESSO
    
    Exit Function
    
Erro_Carrega_Filiais:

    Carrega_Filiais = gErr

    Select Case gErr

        'Erros tratados na rotina chamadora
        Case 107501
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175723)

    End Select

    Exit Function

End Function

Private Sub Vendedores_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Vendedores_Click()
'Função Que Traz os Dados do Vendedor para a Tela
Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Vendedores_Click
    
    'Se não Existir nada Selecionado sai da Função
    If Vendedores.ListIndex = -1 Then Exit Sub
    
    'Seleciona o Código do Vendedor na Combo e passa para o objVendedor
    objVendedor.iCodigo = Codigo_Extrai(Vendedores.ItemData(Vendedores.ListIndex))
    
    'Função que Traz os Dados do Vendedor Cadastrado no Banco da Dados
    lErro = Traz_VendedorFilial_Tela(objVendedor)
    If lErro <> SUCESSO Then gError 104497
    
    'Não Houve Alteração
    iAlterado = REGISTRO_ALTERADO
    
    Exit Sub
    
Erro_Vendedores_Click:

    Select Case gErr

        'Erro tratados na rotina chamadora
        Case 104497
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175724)

    End Select

    Exit Sub

End Sub

Private Sub FiliaisEmpresa_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Function Traz_VendedorFilial_Tela(objVendedor As ClassVendedor) As Long
'Traz os Vendedores para Tela Relacionando-os com a FilialEmpresa que estão vinculados

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais
Dim colFilial As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Traz_VendedorFilial_Tela
    
    'Limpa Todas as Marcações na ListBox
    For iIndice = 0 To FiliaisEmpresa.ListCount - 1
    
        FiliaisEmpresa.Selected(iIndice) = False
    
    Next
    
    'Função que lê Vendedor Vinculado a um FilialEmpresa ou a Várias
    lErro = CF("VendedorLoja_Le_FilalEmpresa", objVendedor)
    If lErro <> SUCESSO And lErro <> 107503 Then gError 107502
    
    'Se não Encontar
    If lErro = 107503 Then gError 107504
    
    
    
    'Preenchimento da List FiliaisEmpresas, com as Marcações na CheckBox
    For iIndice = 0 To FiliaisEmpresa.ListCount - 1
        'Varre a Coleção
        For Each objFilialEmpresa In objVendedor.colFiliaisLoja
            
            'Se for Igual a CheckBox Recebe True
            If objFilialEmpresa.iCodFilial = FiliaisEmpresa.ItemData(iIndice) Then
                
                FiliaisEmpresa.Selected(iIndice) = True
            
            End If
        
        Next
        
    Next
    
    Traz_VendedorFilial_Tela = SUCESSO
    
    Exit Function
        
Erro_Traz_VendedorFilial_Tela:

    Traz_VendedorFilial_Tela = gErr

    Select Case gErr
        
        Case 107502
            'Erro Tratado Dentro da Função Chamadora
        
        Case 107504
            'Não emite MSG
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175725)

    End Select

    Exit Function

End Function


Public Function Trata_Parametros()

    Trata_Parametros = SUCESSO

End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 107510

    Call Limpa_Tela_VendedorFilial

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 107510
            'Erro tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175726)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long
'Função onde começa a Gravação

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Gravar_Registro

    'Verifica se o vendedor esta preenchido
    If Len(Trim(Vendedores.Text)) = 0 Then gError 107511

    'Função que Move o que está na tela para dentro do objVendedores
    lErro = Move_VendedorFilial_Memoria(objVendedor)
    If lErro <> SUCESSO Then gError 107512

    'Função que Grava os Vendedores Relacionados a Filial Empresa Correspondente.
    lErro = CF("VendedorLoja_Filial_Grava", objVendedor)
    If lErro <> SUCESSO Then gError 107513

    Gravar_Registro = SUCESSO

    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    Select Case gErr

        Case 107511
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDORLOJA_NAO_SELECIONADO", gErr)

        Case 107512, 107513, 107530
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175727)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()
'Botão que Limpa a Tela e Chama a Função de Grvação

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Função que Limpa a Tela
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 107529
    
    lErro = Limpa_Tela_VendedorFilial
    If lErro <> SUCESSO Then gError 107530
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr
        
        Case 107529, 107530
            'Erros Tratados dentro da função chamadora
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175728)
    
    End Select
    
    Exit Sub
    
End Sub

Function Limpa_Tela_VendedorFilial() As Long
'Função que Limpa a Tela de VendedoresFilial

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Limpa_Tela_VendedorFilial

    'Limpa a Combo Vendedores
    Vendedores.ListIndex = -1
    
    For iIndice = 0 To FiliaisEmpresa.ListCount - 1
    
        'Limpa as CheckBox Marcadas na List
        FiliaisEmpresa.Selected(iIndice) = False
        
    Next
    
    'Fecha o comando de setas
    Call ComandoSeta_Fechar(Me.Name)
        
    'Indica que não existe nenhum campo alterado
    iAlterado = 0
    
     Limpa_Tela_VendedorFilial = SUCESSO
    
    Exit Function
    
Erro_Limpa_Tela_VendedorFilial:
    
    Limpa_Tela_VendedorFilial = gErr
    
        Select Case gErr
            
            Case Else
                Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175729)
        
        End Select
        
        Exit Function
        
End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Libera a referência da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Function Move_VendedorFilial_Memoria(objVendedor As ClassVendedor) As Long
'Função que move o que Tem na Tela para Memória

Dim lErro As Long
Dim iIndice As Integer
Dim objFilialEmpresa As AdmFiliais

On Error GoTo Erro_Move_VendedorFilial_Memoria

    'Move o Codigo do Vendedor Selecionado na Combo
    objVendedor.iCodigo = Codigo_Extrai(Vendedores.ItemData(Vendedores.ListIndex))
    
    For iIndice = 0 To FiliaisEmpresa.ListCount - 1
        'Instanciar o objVendedor
        Set objFilialEmpresa = New AdmFiliais
    
        'Verifica se a Chek esta Selecionada
        If FiliaisEmpresa.Selected(iIndice) = True Then
                
            objFilialEmpresa.iCodFilial = FiliaisEmpresa.ItemData(iIndice)
        
            'Adcionar na Coleção de Filiais Empresa
            objVendedor.colFiliaisLoja.Add objFilialEmpresa

            
        End If
        
    Next
    
    
    Move_VendedorFilial_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_VendedorFilial_Memoria:

    Move_VendedorFilial_Memoria = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175730)
    
    End Select
    
    Exit Function

End Function

Function Log_Le(objLog As ClassLog) As Long

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
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Operacao, Log1, Log2, Log3, Log4 , Data , Hora FROM Log ", tLog.lNumIntDoc, tLog.iOperacao, tLog.sLog1, tLog.sLog2, tLog.sLog3, tLog.sLog4, tLog.dtData, tLog.dtData)
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175731)

        End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function VendedorFilial_Desmembra_Log(objVendedor As ClassVendedor, objLog As ClassLog) As Long
'Função que informações do banco de Dados e Carrega no Obj

Dim lErro As Long
Dim iPosicao1 As Integer
Dim iPosicao2 As Integer
Dim iPosicao3 As Integer
Dim iPosicao4 As Integer
Dim iPosicao5 As Integer
Dim iParcelas As Integer
Dim sVendedorFilial As String
Dim objFilialEmpresa As AdmFiliais
Dim iIndice As Integer

On Error GoTo Erro_VendedorFilial_Desmembra_Log

    'iPosicao1 Guarda a posição do Primeiro Control
    iPosicao1 = InStr(1, objLog.sLog, Chr(vbKeyControl))
    
    'String que Guarda as Propriedades do objVendedor
    sVendedorFilial = Mid(objLog.sLog, 1, iPosicao1 - 1)
    
    'Inicilalização do objVendedor
    Set objVendedor = New ClassVendedor
     
    'Primeira Posição
    iPosicao3 = 1
    'Procura o Primeiro Escape dentro da String sAdmMeiopagto e Armazena a Posição
    iPosicao2 = (InStr(iPosicao3, sVendedorFilial, Chr(vbKeyEscape)))
    iIndice = 0
    
    Do While iPosicao2 <> 0
        
       iIndice = iIndice + 1
        'Recolhe os Dados do Banco de Dados e Coloca no objAdmMeioPagto
        Select Case iIndice
            
            Case 1: objVendedor.iCodigo = StrParaInt(Mid(sVendedorFilial, iPosicao3, iPosicao2 - iPosicao3))
            Case 2: Exit Do
        
        End Select
        
        'Atualiza as Posições
        iPosicao3 = iPosicao2 + 1
        iPosicao2 = (InStr(iPosicao3, sVendedorFilial, Chr(vbKeyEscape)))
    
    
    Loop
    
    'Atualiza as Posições
    iPosicao3 = iPosicao1 + 1
    iPosicao2 = (InStr(iPosicao3, objLog.sLog, Chr(vbKeyEscape)))
    
    Do While iPosicao2 <> 0
    
        Set objFilialEmpresa = New AdmFiliais
       
        objFilialEmpresa.iCodFilial = StrParaInt(Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3))
            
        'Atualiza as Posições
        iPosicao3 = iPosicao2 + 2
        iPosicao2 = (InStr(iPosicao3, objLog.sLog, Chr(vbKeyEscape)))
        
        objVendedor.colFiliaisLoja.Add objFilialEmpresa
    
    Loop
      
    VendedorFilial_Desmembra_Log = SUCESSO

    Exit Function

Erro_VendedorFilial_Desmembra_Log:

    VendedorFilial_Desmembra_Log = gErr

   Select Case gErr

        Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175732)

        End Select

    
    Exit Function

End Function

Private Sub Teste_Log_Click_Click()

'Função de Teste
Dim lErro As Long
Dim objVendedor As New ClassVendedor
Dim objLog As New ClassLog

On Error GoTo Erro_Teste_Log_Click
 
    lErro = Log_Le(objLog)
    If lErro <> SUCESSO And lErro <> 104202 Then gError 104200
    
    lErro = VendedorFilial_Desmembra_Log(objVendedor, objLog)
    If lErro <> SUCESSO And lErro = 104195 Then gError 104196

    'Só para teste
    lErro = Traz_VendedorFilial_Tela(objVendedor)
    If lErro <> SUCESSO Then gError 107532

    Exit Sub
    
Erro_Teste_Log_Click:
    
    Select Case gErr
                                                                    
        Case 104196, 107532
            'Erro Tratado Dentro da Função Chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175733)
         
        End Select
         
    Exit Sub
    
End Sub

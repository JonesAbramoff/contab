VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl Teclado 
   ClientHeight    =   1470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4515
   LockControls    =   -1  'True
   ScaleHeight     =   1470
   ScaleWidth      =   4515
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   2235
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   60
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "Teclado.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "Teclado.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "Teclado.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "Teclado.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   1740
      Picture         =   "Teclado.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Numeração Automática"
      Top             =   270
      Width           =   300
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1110
      TabIndex        =   1
      Top             =   255
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      PromptInclude   =   0   'False
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Descricao 
      Height          =   315
      Left            =   1110
      TabIndex        =   4
      Top             =   780
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   30
      PromptChar      =   " "
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descrição:"
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
      Left            =   150
      TabIndex        =   3
      Top             =   840
      Width           =   915
   End
   Begin VB.Label LabelCodigo 
      Alignment       =   1  'Right Justify
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
      Top             =   315
      Width           =   660
   End
End
Attribute VB_Name = "Teclado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Variáveis Globais a Tela Teclados
Public iAlterado As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoTeclado As AdmEvento
Attribute objEventoTeclado.VB_VarHelpID = -1
Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    'seta o admEvento
    Set objEventoTeclado = New AdmEvento
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174541)

    End Select
    
    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD
Dim lErro As Long
Dim objTeclado As New ClassTeclado

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Teclado"

    'Le os dados da Tela AdmMeioPagto
    lErro = Move_Tela_Memoria(objTeclado)
    If lErro <> SUCESSO Then gError 99451

    'Preenche a coleção colCampoValor, com descricao do campo,
    colCampoValor.Add "Codigo", objTeclado.iCodigo, 0, "Codigo"
    colCampoValor.Add "descricao", objTeclado.sDescricao, STRING_TECLADO_DESCRICAO, "Descricao"
    colCampoValor.Add "FilialEmpresa", objTeclado.iFilialEmpresa, 0, "FilialEmpresa"
    
    'Utilizado na hora de passar o parâmetro FilialEmpresa
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr
        'Erro tratado na rotina chamadora
        Case 99451
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174542)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD
Dim lErro As Long
Dim objTeclado As New ClassTeclado

On Error GoTo Erro_Tela_Preenche

    objTeclado.iCodigo = colCampoValor.Item("Codigo").vValor
            
    If objTeclado.iCodigo > 0 Then
        
        'Carrega objAdmMeioPagto com os dados passados em colCampoValor
        objTeclado.sDescricao = colCampoValor.Item("Descricao").vValor
        objTeclado.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
        
        'Traz dados de Teclados para a Tela
        lErro = Traz_Teclado_Tela(objTeclado)
        If lErro <> SUCESSO Then gError 99452
        
    End If
        
    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 99452
        'Erro tratado na rotina chamadora
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174543)

    End Select
    
    Exit Sub

End Sub

Private Sub Codigo_GotFocus()
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
End Sub

Private Sub objEventoteclado_evSelecao(obj1 As Object)

Dim objTeclado As New ClassTeclado
Dim lErro As Long

On Error GoTo Erro_objEventoteclado_evSelecao

    Set objTeclado = obj1
    
    If objTeclado.iCodigo > 0 Then
    
        Codigo.Text = objTeclado.iCodigo
        
        'Lê no BD a partir do código
        lErro = CF("Teclado_Le", objTeclado)
        If lErro <> SUCESSO And lErro <> 99459 Then gError 99455
        
        If lErro = SUCESSO Then
                
            'Exibe os dados do teclado
            lErro = Traz_Teclado_Tela(objTeclado)
            If lErro <> SUCESSO Then gError 99454
                
        End If
        
    End If
    
    Me.Show
        
    Exit Sub

Erro_objEventoteclado_evSelecao:
    
    Select Case gErr
        
        Case 99454, 99455
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174544)

    End Select
    
    Exit Sub

End Sub


Function Move_Tela_Memoria(objTeclado As ClassTeclado) As Long
'Move os dados da tela para o objTeclado
Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria
         
    'Move a FilialEmpresa que esta sendo Referenciada para a Memória
    objTeclado.iFilialEmpresa = giFilialEmpresa
    'Move o Codigo Para Memoria
    objTeclado.iCodigo = StrParaInt(Codigo.Text)
    'Move o descricao Para Memoria
    objTeclado.sDescricao = Descricao.Text
    
    Move_Tela_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
    
    Select Case gErr
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174545)
        
        End Select

    Exit Function
    
End Function

Function Traz_Teclado_Tela(objTeclado As ClassTeclado) As Long
'Função que Traz as Informações do BD para o teclado
Dim lErro As Long

On Error GoTo Erro_Traz_Teclado_Tela

    Call Limpa_Tela(Me)
    
    'Traz o Codigo para a Tela
    Codigo.Text = objTeclado.iCodigo
    
    'Traz o descricao para a Tela
    Descricao.Text = objTeclado.sDescricao
        
    'Demonstra que não Houve Alteração na Tela
    iAlterado = 0
    
    Traz_Teclado_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Teclado_Tela:

    Traz_Teclado_Tela = gErr

    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174546)
        
        End Select
        
        Exit Function
        
End Function

Function Trata_Parametros(Optional objTeclado As ClassTeclado) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se houver POS passado como parâmetro, exibe seus dados
    If Not (objTeclado Is Nothing) Then

        objTeclado.iFilialEmpresa = giFilialEmpresa

        If objTeclado.iCodigo > 0 Then

            'Lê no BD a partir do código
            lErro = CF("Teclado_Le", objTeclado)
            If lErro <> SUCESSO And lErro <> 99459 Then gError 99438
            
            If lErro = SUCESSO Then

                'Exibe os dados do teclado
                lErro = Traz_Teclado_Tela(objTeclado)
                If lErro <> SUCESSO Then gError 99439
                
            Else
    
                Codigo.Text = objTeclado.iCodigo
                Descricao.Text = objTeclado.sDescricao
                    
            End If

        End If
    
    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 99438, 99439
        'Erro tratado dentro da Função Chamadora
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174547)

    End Select

    Exit Function

End Function

Private Sub Codigo_Change()
'Determina se Houve Mudança

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Descricao_Change()
'Determina se Houve Mudança

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub BotaoProxNum_Click()
'Gera um novo número disponível para código do Teclado

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click
    
    'Chama a função que gera o sequencial do Código Automático para o novo teclado
    lErro = CF("Config_Obter_Inteiro_Automatico", "LojaConfig", "NUM_PROXIMO_TECLADO", "Teclado", "Codigo", iCodigo)
    
    If lErro <> SUCESSO Then gError 99453

    'Exibe o novo código na tela
    Codigo.Text = CStr(iCodigo)
        
    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 99453
            'Erro tratado na rotina chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174548)
    
    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Função que Inicializa a Gravação de Novo Registro

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chamada da Função Gravar Registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 99440
    
    'Limpa a Tela
     Call Limpa_Tela(Me)
     
    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoGravar_Click:
    
    Select Case gErr
            
        Case 99440
            'Erro Tratada dentro da Função Chamadora
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174549)

    End Select

    Exit Sub
    
End Sub
             
Function Gravar_Registro() As Long
'Função de Gravação de Teclado

Dim objTeclado As New ClassTeclado
Dim lErro As Long

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o campo Código esta preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 99441
    
    'Verifica se o campo descricao esta preenchido
    If Len(Trim(Descricao.Text)) = 0 Then gError 99442
        
    'Move para a memória os campos da Tela
    lErro = Move_Tela_Memoria(objTeclado)
    If lErro <> SUCESSO Then gError 99443
    
    'Utilização para incluir FilialEmpresa como parâmetro
    lErro = Trata_Alteracao(objTeclado, objTeclado.iCodigo, objTeclado.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 99444

    'Chama a Função que Grava Teclado na Tabela
    lErro = CF("Teclado_Grava", objTeclado)
    If lErro <> SUCESSO Then gError 99445
        
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:
   
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = gErr
        
        Select Case gErr
            
            Case 99441
                lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
                
            Case 99442
                lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", gErr)
            
            Case 99443, 99444, 99445
                'Erro Tratado Dentro da Função
                    
            Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174550)

        End Select
        
    Exit Function
    
End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objTeclado As New ClassTeclado
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se o Codigo Está Preenchido senão Erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 99446
    
    'Para Saber qual é a FilialEmpresa que Esta sendo Referenciada
    objTeclado.iFilialEmpresa = giFilialEmpresa
    
    'Passa o codigo para a leitura no banco de dados
    objTeclado.iCodigo = Codigo.Text
    
    'Lê a Teclado no Banco e Trazer o objTeclado
    lErro = CF("Teclado_Le", objTeclado)
    If lErro <> SUCESSO And lErro <> 99459 Then gError 99447
    
    'Se não for encontrado a Teclado no Bd
    If lErro = 99459 Then gError 99448
    
    'Envia aviso perguntando se realmente deseja excluir Teclado
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_TECLADO", objTeclado.iCodigo)

    If vbMsgRes = vbYes Then

        'Exclui Teclado
        lErro = CF("Teclado_Exclui", objTeclado)
        If lErro <> SUCESSO Then gError 99449
        
        'Limpa a Tela
        Call Limpa_Tela(Me)
        
        'Fechar o Comando de Setas
        Call ComandoSeta_Fechar(Me.Name)
        
        iAlterado = 0
        
    End If
    
    Exit Sub
        
Erro_BotaoExcluir_Click:

    Select Case gErr
    
        Case 99446
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 99447, 99449
            'Erro Tratado Dentro da Função Chamadora
        
        Case 99448
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TECLADO_NAO_ENCONTRADO", gErr, objTeclado.iCodigo)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174551)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoLimpar_Click()
'Função que tem as chamadas para as Funções que limpam a tela
Dim lErro As Long

On Error GoTo Erro_Botaolimpar_Click
                    
    'Verifica se existe algo para ser salvo antes de limpar a tela
    Call Teste_Salva(Me, iAlterado)
    
    'Limpa Tela de Teclados
    Call Limpa_Tela(Me)
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then gError 99450
    
    iAlterado = 0
    
    Exit Sub
        
Erro_Botaolimpar_Click:

    Select Case gErr
    
        Case 99450
            'Erro Tratado dentro da Função Chamadora
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174552)

    End Select
    
    Exit Sub
        
End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub form_unload(Cancel As Integer)

Dim lErro As Long

    'Libera a referência da tela e fecha o comando das setas se estiver aberto
    
    Set objEventoTeclado = Nothing
    
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Private Sub BotaoFechar_Click()

'Função que Fecha a Tela
    Unload Me

End Sub

Sub Teclado_Desmembra_Log(objTeclado As ClassTeclado, objLog As ClassLog)
'Função que informações do banco de Dados e Carrega no Obj

Dim lErro As Long
Dim iPosicao3 As Integer
Dim iPosicao2 As Integer
Dim iPosicao4 As Integer
Dim iIndice As Integer

On Error GoTo Erro_Teclado_Desmembra_Log

    'Inicilalização do objteclado
    Set objTeclado = New ClassTeclado
     
    'Primeira Posição
    iPosicao3 = 1
    'Procura o Primeiro Escape dentro da String
    iPosicao2 = (InStr(iPosicao3, objLog.sLog, Chr(vbKeyEscape)))
    
    iPosicao4 = (InStr(iPosicao3, objLog.sLog, Chr(vbKeyEnd)))
    
    iIndice = 0
    
    Do While iPosicao2 <> 0
        
       iIndice = iIndice + 1
        'Recolhe os Dados do Banco de Dados e Coloca no objAdmMeioPagto
        Select Case iIndice
            
            Case 1: objTeclado.iCodigo = StrParaInt(Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3))
            Case 2: objTeclado.iFilialEmpresa = StrParaInt(Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3))
            Case 3: objTeclado.sDescricao = Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3)
                        
        End Select
        
        'Atualiza as Posições
        iPosicao3 = iPosicao2 + 1
        iPosicao2 = (InStr(iPosicao3, objLog.sLog, Chr(vbKeyEscape)))
        
        If iPosicao3 < iPosicao4 And iPosicao2 = 0 Then iPosicao2 = iPosicao4
                 
    Loop

    Exit Sub
        
Erro_Teclado_Desmembra_Log:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174553)

        End Select
    
    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
    If KeyCode = KEYCODE_BROWSER Then
        Call LabelCodigo_Click
    End If

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Teclado"
    Call Form_Load
    
End Function

Public Function Name() As String
'descricao da Tela
    Name = "Teclado"
    
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

''***** fim do trecho a ser copiado ******
Private Sub LabelCodigo_Click()

Dim objTeclado As New ClassTeclado
Dim colSelecao As Collection
    
    If Len(Trim(Codigo.Text)) > 0 Then objTeclado.iCodigo = StrParaInt(Codigo.Text)
    
    Call Chama_Tela("TecladoLista", colSelecao, objTeclado, objEventoTeclado)

    Exit Sub

End Sub


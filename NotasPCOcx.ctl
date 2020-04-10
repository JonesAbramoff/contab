VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl NotasPCOcx 
   ClientHeight    =   4335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7845
   ScaleHeight     =   4335
   ScaleWidth      =   7845
   Begin VB.TextBox Nota 
      Height          =   1095
      Left            =   990
      MaxLength       =   150
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   990
      Width           =   6735
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   1545
      Picture         =   "NotasPCOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Numeração Automática"
      Top             =   375
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5655
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "NotasPCOcx.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "NotasPCOcx.ctx":0268
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "NotasPCOcx.ctx":079A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "NotasPCOcx.ctx":0924
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ListBox NotasPC 
      Height          =   1815
      ItemData        =   "NotasPCOcx.ctx":0A7E
      Left            =   105
      List            =   "NotasPCOcx.ctx":0A80
      TabIndex        =   5
      Top             =   2400
      Width           =   7635
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   990
      TabIndex        =   1
      Top             =   360
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
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
      Left            =   210
      TabIndex        =   0
      Top             =   420
      Width           =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nota:"
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
      Left            =   375
      TabIndex        =   3
      Top             =   915
      Width           =   480
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Notas"
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
      Left            =   195
      TabIndex        =   4
      Top             =   2130
      Width           =   510
   End
End
Attribute VB_Name = "NotasPCOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private Sub ListaNotasPC_Inclui(objNotasPC As ClassNotaPC)
'Adiciona na ListBox de Tipo de NotasPC

Dim iIndice As Integer

    For iIndice = 0 To NotasPC.ListCount - 1
        If NotasPC.ItemData(iIndice) > objNotasPC.lCodigo Then Exit For
    Next

    NotasPC.AddItem objNotasPC.lCodigo & SEPARADOR & objNotasPC.sNota, iIndice
    NotasPC.ItemData(NotasPC.NewIndex) = objNotasPC.lCodigo

    Exit Sub

End Sub

Private Sub ListaNotasPC_Exclui(objNotasPC As ClassNotaPC)
'Percorre a ListBox de NotasPC para remover o tipo caso ele exista

Dim iIndice As Integer

    For iIndice = 0 To NotasPC.ListCount - 1

        If NotasPC.ItemData(iIndice) = objNotasPC.lCodigo Then
            NotasPC.RemoveItem (iIndice)
            Exit For
        End If
    Next

    Exit Sub

End Sub

Function Move_Tela_Memoria(objNotasPC As ClassNotaPC) As Long

    'Move os dados da tela para objNotasPC
    objNotasPC.lCodigo = StrParaLong(Codigo.Text)
    objNotasPC.sNota = Trim(Nota.Text)

    Move_Tela_Memoria = SUCESSO

    Exit Function

End Function


Public Function Gravar_Registro() As Long
'Verifica se dados de NotasPC necessários foram preenchidos

Dim lErro As Long
Dim objNotasPC As New ClassNotaPC

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 103299

    'Verifica se a Nota foi preenchida
    If Len(Trim(Nota.Text)) = 0 Then gError 103300

    lErro = Move_Tela_Memoria(objNotasPC)
    If lErro <> SUCESSO Then gError 103301

    If objNotasPC.lCodigo = 0 Then gError 103302

    lErro = Trata_Alteracao(objNotasPC, objNotasPC.lCodigo)
    If lErro <> SUCESSO Then gError 103303

    lErro = CF("NotasPC_Grava", objNotasPC)
    If lErro <> SUCESSO Then gError 103304

    'Exclui da ListBox
    Call ListaNotasPC_Exclui(objNotasPC)

    'Inclui na ListBox
    Call ListaNotasPC_Inclui(objNotasPC)

    Gravar_Registro = SUCESSO

    GL_objMDIForm.MousePointer = vbDefault

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 103303

        Case 103302
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_INVALIDO1", gErr)

        Case 103299
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 103300
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", gErr)

        Case 103301, 103304

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163612)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Function

End Function

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objNotaPC As New ClassNotaPC

On Error GoTo Erro_Tela_Preenche

    'Coloca colCampoValor na Tela
    'Conversão de tipagem para a tipagem da tela se necessário
    objNotaPC.lCodigo = CStr(colCampoValor.Item("Codigo").vValor)
    objNotaPC.sNota = colCampoValor.Item("Nota").vValor
    
    Call Traz_NotasPC_Tela(objNotaPC)
    If lErro <> SUCESSO Then gError 103304

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 103304

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163613)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim objCampoValor As AdmCampoValor
Dim iCodigo As Integer
Dim lErro As Long
Dim objNotaPC As New ClassNotaPC

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada a tela
    sTabela = "NotasPC"

    lErro = Move_Tela_Memoria(objNotaPC)
    If lErro <> SUCESSO Then gError 103305

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objNotaPC.lCodigo, 0, "Codigo"
    colCampoValor.Add "Nota", objNotaPC.sNota, STRING_NOTASPC_NOTA, "Nota"
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 103305

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163614)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objNotasPC As ClassNotaPC) As Long
'Trata os parametros que podem ser passados quando ocorre a chamada da tela de NotasPC

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se houve passagem de parametro
    If Not (objNotasPC Is Nothing) Then

        lErro = CF("NotasPC_Le", objNotasPC)
        If lErro <> SUCESSO And lErro <> 103293 Then gError 103289

        If lErro = SUCESSO Then

            Call Traz_NotasPC_Tela(objNotasPC)
                        
        Else
            Codigo.Text = CStr(objNotasPC.lCodigo)
        End If

    End If

   Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 103289

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163615)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objNotaPC As New ClassNotaPC

On Error GoTo Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o codigo foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 103306

    objNotaPC.lCodigo = StrParaDbl(Codigo.Text)

    lErro = CF("NotasPC_Le", objNotaPC)
    If lErro <> SUCESSO And lErro <> 103293 Then gError 103307

    'Verifica se a Nota nao esta cadastrada
    If lErro <> SUCESSO Then gError 103308

    'Envia aviso perguntando se realmente deseja excluir a Nota
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_NOTAPC", objNotaPC.lCodigo)

    If vbMsgRes = vbYes Then

        'Exclui a Nota
        lErro = CF("NotasPC_Exclui", objNotaPC)
        If lErro <> SUCESSO Then gError 103309

        'Exclui da List
        Call ListaNotasPC_Exclui(objNotaPC)

        'Limpa a tela
        Call Limpa_Tela_NotasPC

        'Fecha o comando das setas se estiver aberto
        lErro = ComandoSeta_Fechar(Me.Name)

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 103306
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 103307, 103309

        Case 103308
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTASPCPC_NAO_CADASTRADO", gErr, objNotaPC.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163616)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 103310

    'Limpa a tela
    Call Limpa_Tela_NotasPC

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 103310

        Case Else

            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163617)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 103311

    'Limpa a tela
    Call Limpa_Tela_NotasPC

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 103311

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163618)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Obtém o próximo código disponível para TipoBloqueioPC
    lErro = CF("Config_ObterAutomatico", "ComprasConfig", "NUM_PROXIMO_NOTA_PC", "NotasPC", "Codigo", lCodigo)
    If lErro <> SUCESSO Then gError 103297

    'Coloca o Código obtido na tela
    Codigo.Text = lCodigo

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 103297
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163619)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Codigo_GotFocus()

    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Private Sub Nota_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_Activate()

   Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome

On Error GoTo Erro_Form_Load

    'Preenche a listbox NotaPC
    'Le cada codigo e Nota da tabela NotaPC
    lErro = CF("Cod_Nomes_Le", "NotasPC", "Codigo", "Nota", STRING_NOTASPC_NOTA, colCodigoNome)
    If lErro <> SUCESSO Then gError 103298

    'preenche a listbox NotasPC com os objetos da colecao colCodigoNota
    For Each objCodigoNome In colCodigoNome

        NotasPC.AddItem objCodigoNome.iCodigo & SEPARADOR & objCodigoNome.sNome
        NotasPC.ItemData(NotasPC.NewIndex) = objCodigoNome.iCodigo

    Next

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 103298

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163620)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Private Sub NomeReduzido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NotasPC_DblClick()

Dim lErro As Long
Dim objNotaPC As New ClassNotaPC

On Error GoTo Erro_NotasPC_DblClick

    objNotaPC.lCodigo = NotasPC.ItemData(NotasPC.ListIndex)

    'Le a NotaPC
    lErro = CF("NotasPC_Le", objNotaPC)
    If lErro <> SUCESSO And lErro <> 103293 Then gError 103294

    'Verifica se a NotaPC nao esta cadastrada
    If lErro <> SUCESSO Then gError 103295

    Call Traz_NotasPC_Tela(objNotaPC)
    If lErro <> SUCESSO Then gError 103296

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_NotasPC_DblClick:

    Select Case gErr

        Case 103294, 103296

        Case 103295
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODEBLOQUEIOPC_NAO_CADASTRADO", Err, objNotaPC.lCodigo)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 163621)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Notas Pedidos de Compra"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "NotasPC"

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

'**** fim do trecho a ser copiado *****


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

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Sub Traz_NotasPC_Tela(objNotasPC As ClassNotaPC)

Dim iIndex As Integer

On Error GoTo Erro_Traz_NotasPC_Tela

    'Limpa a tela
    Call Limpa_Tela_NotasPC

    'Mostra os dados na tela
    Codigo.Text = CStr(objNotasPC.lCodigo)
    Nota.Text = objNotasPC.sNota
        
    iAlterado = 0

    Exit Sub

Erro_Traz_NotasPC_Tela:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163622)

    End Select

End Sub

Sub Limpa_Tela_NotasPC()

    Call Limpa_Tela(Me)
    
    NotasPC.ListIndex = -1
        
End Sub



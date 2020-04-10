VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl MnemonicoGlobalOcx 
   ClientHeight    =   4800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9495
   ScaleHeight     =   4800
   ScaleWidth      =   9495
   Begin MSComctlLib.TreeView TvwContas 
      Height          =   3555
      Left            =   7110
      TabIndex        =   4
      Top             =   1050
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   6271
      _Version        =   393217
      Indentation     =   453
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
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
   Begin MSMask.MaskEdBox Valor 
      Height          =   480
      Left            =   4680
      TabIndex        =   2
      Top             =   1110
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   847
      _Version        =   393216
      BorderStyle     =   0
      AllowPrompt     =   -1  'True
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
      PromptChar      =   " "
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   8220
      ScaleHeight     =   495
      ScaleWidth      =   1125
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   1185
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   600
         Picture         =   "MnemonicoGlobalOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   105
         Picture         =   "MnemonicoGlobalOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
   End
   Begin VB.TextBox Mnemonico 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   390
      TabIndex        =   0
      Text            =   "Mnemonico"
      Top             =   1020
      Width           =   1500
   End
   Begin VB.TextBox Descricao 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   480
      Left            =   1950
      TabIndex        =   1
      Text            =   "Descricao"
      Top             =   1020
      Width           =   3135
   End
   Begin MSFlexGridLib.MSFlexGrid GridCampos 
      Height          =   3555
      Left            =   30
      TabIndex        =   3
      Top             =   1050
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   6271
      _Version        =   393216
      Rows            =   15
      Cols            =   4
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   2
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Plano de Contas"
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
      Left            =   7110
      TabIndex        =   8
      Top             =   810
      Width           =   1410
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Campos"
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
      Left            =   30
      TabIndex        =   9
      Top             =   810
      Width           =   675
   End
End
Attribute VB_Name = "MnemonicoGlobalOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Dim objGridCampos As AdmGrid

'Colunas do Grid
Const GRID_MNEMONICO_COL = 1
Const GRID_DESCRICAO_COL = 2
Const GRID_VALOR_COL = 3

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim lTotalReg As Long
Dim colMnemonicoGlobal As New Collection
Dim sMascaraConta As String

On Error GoTo Erro_Form_Load

    Set objGridCampos = New AdmGrid

    'tela em questão
    Set objGridCampos.objForm = Me
    
    'Inicializa a máscara de Conta
    sMascaraConta = String(STRING_CONTA, 0)
    
    'le a mascara das contas
    lErro = MascaraConta(sMascaraConta)
    If lErro <> SUCESSO Then Error 58200
    
    Valor.Mask = sMascaraConta

    'Inicializa a Arvore de Plano de Contas
    lErro = CF("Carga_Arvore_Conta",TvwContas.Nodes)
    If lErro <> SUCESSO Then Error 58201
    
    lErro = Inicializa_Grid_Campos(objGridCampos)
    If lErro <> SUCESSO Then Error 39764

    lErro = CF("MnemonicoCTBValor_Le_Globais",colMnemonicoGlobal)
    If lErro <> SUCESSO Then Error 39771
    
    lErro = Preenche_GridCampos(colMnemonicoGlobal)
    If lErro <> SUCESSO Then Error 39776
    
    lErro_Chama_Tela = SUCESSO

    iAlterado = 0

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 39764, 39771, 39776, 58200, 58201 'Tratados nas Rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162778)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Inicializa_Grid_Campos(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Inicializa_Grid_Campos

    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    'titulos do Grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Mnemônico")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Valor")

    'controles que participam do Grid
    objGridInt.colCampo.Add (Mnemonico.Name)
    objGridInt.colCampo.Add (Descricao.Name)
    objGridInt.colCampo.Add (Valor.Name)

    'Relaciona com o Grid correspondente na Tela
    objGridInt.objGrid = GridCampos
  
    'linhas visíveis do Grid
    objGridInt.iLinhasVisiveis = 6

    GridCampos.ColWidth(0) = 400

    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    Inicializa_Grid_Campos = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Campos:

    Inicializa_Grid_Campos = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162779)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
 
End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objGridCampos = Nothing
        
End Sub

Private Sub GridCampos_LeaveCell()

    Call Saida_Celula(objGridCampos)

End Sub

Private Sub GridCampos_EnterCell()

    Call Grid_Entrada_Celula(objGridCampos, iAlterado)

End Sub

Private Sub GridCampos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridCampos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCampos, iAlterado)
    End If

End Sub

Private Sub GridCampos_GotFocus()

    Call Grid_Recebe_Foco(objGridCampos)

End Sub

Private Sub GridCampos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridCampos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCampos, iAlterado)
    End If

End Sub

Private Sub GridCampos_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridCampos)

End Sub

Private Sub GridCampos_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridCampos)

End Sub

Private Sub GridCampos_RowColChange()

    Call Grid_RowColChange(objGridCampos)

End Sub

Private Sub GridCampos_Scroll()

    Call Grid_Scroll(objGridCampos)
    
End Sub

Public Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO
    
End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_Botao_Gravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 39779

    iAlterado = 0
    
    Unload Me
    
    Exit Sub

Erro_Botao_Gravar_Click:

    Select Case Err
        
        Case 39779
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162780)
        
    End Select
    
    Exit Sub
    
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim colMnemonico As Collection

On Error GoTo Erro_Gravar_Registro
        
    GL_objMDIForm.MousePointer = vbHourglass
    
    lErro = Move_Tela_Memoria(colMnemonico)
    If lErro <> SUCESSO Then Error 39785
        
    lErro = CF("MnemonicoCTBValor_Grava",colMnemonico)
    If lErro <> SUCESSO Then Error 39784
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
        
        Case 39784, 39785
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162781)

    End Select

    Exit Function

End Function

Function Preenche_GridCampos(colMnemonicoGlobal As Collection) As Long
'preenche o grid com os Mnemônicos passados na coleção colMnemonicoGlobal

Dim lErro As Long
Dim iIndice As Integer
Dim objMnemonico As ClassMnemonicoCTBValor

On Error GoTo Erro_Preenche_GridCampos

    'Limpa o grid
    Call Grid_Limpa(objGridCampos)

    If colMnemonicoGlobal.Count < 13 Then
        objGridCampos.objGrid.Rows = 14
    Else
        objGridCampos.objGrid.Rows = colMnemonicoGlobal.Count + 1
    End If

    objGridCampos.iLinhasExistentes = colMnemonicoGlobal.Count

    'preenche o grid com os dados retornados na coleção colMnemonicoGlobal
    For iIndice = 1 To colMnemonicoGlobal.Count

        Set objMnemonico = colMnemonicoGlobal.Item(iIndice)
        
        GridCampos.TextMatrix(iIndice, GRID_MNEMONICO_COL) = objMnemonico.sMnemonico
        GridCampos.TextMatrix(iIndice, GRID_DESCRICAO_COL) = objMnemonico.sDescricao
        GridCampos.TextMatrix(iIndice, GRID_VALOR_COL) = objMnemonico.sValor

    Next

    Call Grid_Inicializa(objGridCampos)

    Preenche_GridCampos = SUCESSO

    Exit Function

Erro_Preenche_GridCampos:

    Preenche_GridCampos = Err

    Select Case Err

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162782)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(colMnemonico As Collection) As Long
'Move os dados o GridCampos para a coleção colMnemonico

Dim lErro As Long
Dim iLinha As Integer
Dim objMnemonico As ClassMnemonicoCTBValor

On Error GoTo Erro_Move_Tela_Memoria
    
    Set colMnemonico = New Collection

    For iLinha = 1 To objGridCampos.iLinhasExistentes
        
        Set objMnemonico = New ClassMnemonicoCTBValor
        
        objMnemonico.sMnemonico = GridCampos.TextMatrix(iLinha, GRID_MNEMONICO_COL)
        
        objMnemonico.sDescricao = GridCampos.TextMatrix(iLinha, GRID_DESCRICAO_COL)
        
        objMnemonico.sValor = GridCampos.TextMatrix(iLinha, GRID_VALOR_COL)
        
        colMnemonico.Add objMnemonico
        
    Next

    Move_Tela_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err
    
    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162783)
            
    End Select
    
    Exit Function
    
End Function

Private Sub Valor_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Valor_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridCampos)

End Sub

Private Sub Valor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCampos)
    
End Sub

Private Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCampos.objControle = Valor
    lErro = Grid_Campo_Libera_Foco(objGridCampos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub TvwContas_Expand(ByVal objNode As MSComctlLib.Node)

Dim lErro As Long

On Error GoTo Erro_TvwContas_Expand

    If objNode.Tag <> NETOS_NA_ARVORE Then
    
        'move os dados do plano de contas do banco de dados para a arvore colNodes.
        lErro = CF("Carga_Arvore_Conta1",objNode, TvwContas.Nodes)
        If lErro <> SUCESSO Then Error 36771
        
    End If
    
    Exit Sub
    
Erro_TvwContas_Expand:

    Select Case Err
    
        Case 36771
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162784)
        
    End Select
        
    Exit Sub
    
End Sub

Private Sub TvwContas_NodeClick(ByVal Node As MSComctlLib.Node)

Dim sConta As String
Dim lErro As Long
Dim sContaEnxuta As String
Dim sCaracterInicial As String

On Error GoTo Erro_TvwContas_NodeClick
    
    sCaracterInicial = Left(Node.Key, 1)

    If sCaracterInicial = "A" Then
    
        If GridCampos.Row <= 0 Then Error 58202
    
        sConta = Right(Node.Key, Len(Node.Key) - 1)
                
        sContaEnxuta = String(STRING_CONTA, 0)
    
        'Retira a Mascara da Conta
        lErro = Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
        If lErro <> SUCESSO Then Error 58203
    
        'Preenche o campo valor
        Valor.PromptInclude = False
        Valor.Text = sContaEnxuta
        Valor.PromptInclude = True
    
        'Move para o grid
        GridCampos.TextMatrix(GridCampos.Row, GRID_VALOR_COL) = Valor.Text
 
    End If
    
    Exit Sub

Erro_TvwContas_NodeClick:

    Select Case Err
    
        Case 58202
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", Err)
        
        Case 58203 'Tratado na Rotina chamada
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162785)
            
    End Select
        
    Exit Sub
    
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        Select Case GridCampos.Col

            Case GRID_VALOR_COL

                lErro = Saida_Celula_Valor(objGridInt)
                If lErro <> SUCESSO Then Error 58208

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 58209

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 58208

        Case 58209
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162786)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Valor(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sContaFormatada As String
Dim sContaMascarada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim iContaPreenchida As Integer

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridInt.objControle = Valor

    If Len(Trim(Valor.ClipText)) > 0 Then

        'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
        lErro = CF("ContaSimples_Critica_Modulo",sContaFormatada, Valor.ClipText, objPlanoConta, MODULO_CONTABILIDADE)
        If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then Error 58203

        If lErro = SUCESSO Then

            sContaFormatada = objPlanoConta.sConta

            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)

            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
            If lErro <> SUCESSO Then Error 58204

            Valor.PromptInclude = False
            Valor.Text = sContaMascarada
            Valor.PromptInclude = True

        'se não encontrou a conta simples
        ElseIf lErro = 44096 Or lErro = 44098 Then

            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica",Valor.Text, sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
            If lErro <> SUCESSO And lErro <> 5700 Then Error 58205

            'conta não cadastrada
            If lErro = 5700 Then Error 58206

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 58207

    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = Err

    Select Case Err

        Case 58203, 58205, 58207
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 58204
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 58206
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", Valor.Text)

            If vbMsgRes = vbYes Then
                objPlanoConta.sConta = sContaFormatada

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("PlanoConta", objPlanoConta)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162787)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CONFIG_CAMPOS_GLOBAIS_UTILIZ_CONTABILIZACAO
    Set Form_Load_Ocx = Me
    Caption = "Configuração dos Campos Globais Utilizados na Contabilização"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "MnemonicoGlobal"
    
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

'***** fim do trecho a ser copiado ******



Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub


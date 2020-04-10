VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ContratosMgz 
   ClientHeight    =   4845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10635
   ScaleHeight     =   4845
   ScaleWidth      =   10635
   Begin VB.CommandButton BotaoContratos 
      Caption         =   "Contratos Cadastrados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   135
      TabIndex        =   19
      Top             =   4290
      Width           =   2685
   End
   Begin VB.TextBox ProcObs 
      Height          =   300
      Left            =   6420
      TabIndex        =   18
      Top             =   2370
      Width           =   2985
   End
   Begin MSMask.MaskEdBox ProcDataCobr 
      Height          =   225
      Left            =   4905
      TabIndex        =   17
      Top             =   2070
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox ProcValor 
      Height          =   300
      Left            =   3390
      TabIndex        =   16
      Top             =   2055
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   529
      _Version        =   393216
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
   Begin VB.TextBox ProcDescricao 
      Height          =   300
      Left            =   6030
      TabIndex        =   15
      Top             =   2010
      Width           =   2985
   End
   Begin VB.ComboBox ProcTipo 
      Height          =   315
      ItemData        =   "ContratosMiguez.ctx":0000
      Left            =   2100
      List            =   "ContratosMiguez.ctx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   2040
      Width           =   1305
   End
   Begin VB.TextBox ProcID 
      Height          =   300
      Left            =   345
      TabIndex        =   13
      Top             =   2010
      Width           =   1665
   End
   Begin MSFlexGridLib.MSFlexGrid GridProcessos 
      Height          =   2985
      Left            =   135
      TabIndex        =   9
      Top             =   1200
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   5265
      _Version        =   393216
   End
   Begin VB.TextBox Contrato 
      Height          =   300
      Left            =   4365
      TabIndex        =   8
      Top             =   165
      Width           =   2145
   End
   Begin VB.PictureBox Picture3 
      Height          =   510
      Left            =   8520
      ScaleHeight     =   450
      ScaleWidth      =   1935
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   60
      Width           =   1995
      Begin VB.CommandButton BotaoGravar 
         Height          =   330
         Left            =   75
         Picture         =   "ContratosMiguez.ctx":002C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   390
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   330
         Left            =   1005
         Picture         =   "ContratosMiguez.ctx":0186
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   390
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   330
         Left            =   1470
         Picture         =   "ContratosMiguez.ctx":06B8
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   390
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   330
         Left            =   540
         Picture         =   "ContratosMiguez.ctx":0836
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Excluir"
         Top             =   60
         Width           =   390
      End
   End
   Begin MSMask.MaskEdBox Cliente 
      Height          =   300
      Left            =   1035
      TabIndex        =   0
      Top             =   165
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Vendedor 
      Height          =   300
      Left            =   4365
      TabIndex        =   10
      Top             =   690
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      Caption         =   "Processos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   180
      TabIndex        =   12
      Top             =   960
      Width           =   1110
   End
   Begin VB.Label LabelVendedor 
      AutoSize        =   -1  'True
      Caption         =   "Responsável:"
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
      Left            =   3135
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   11
      Top             =   750
      Width           =   1170
   End
   Begin VB.Label LabelContrato 
      AutoSize        =   -1  'True
      Caption         =   "Contrato:"
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
      Left            =   3510
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   7
      Top             =   210
      Width           =   795
   End
   Begin VB.Label LabelCliente 
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
      Left            =   315
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   1
      Top             =   195
      Width           =   660
   End
End
Attribute VB_Name = "ContratosMgz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'pendencias:
    'browse no label de contratos, browse pelo botao, tab order, setas, numeracao de erros, validacao de campos do grid, cadastrar registro na tabela objetos do dic
    'usar CF e transferir rotinas, constantes, type

'Property Variables:
Dim m_Caption As String
Event Unload()

'VARIAVEIS GLOBAIS DA TELA
Dim iAlterado As Integer

Const NUM_MAX_PROCESSOSCONTRATO = 100

Dim iGrid_ProcID_Col As Integer
Dim iGrid_ProcTipo_Col As Integer
Dim iGrid_ProcDescricao_Col As Integer
Dim iGrid_ProcValor_Col As Integer
Dim iGrid_ProcDataCobr_Col As Integer
Dim iGrid_ProcObservacao_Col As Integer

Dim objGridProcessos As AdmGrid

'EVENTOS DE BROWSER
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1

Sub Limpa_Tela_Contratos()

    Call Limpa_Tela(Me)

    'Limpa o Grid
    Call Grid_Limpa(objGridProcessos)

    iAlterado = 0

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    Set objEventoCliente = New AdmEvento
    Set objEventoVendedor = New AdmEvento
        
    lErro = Inicializa_GridProcessos
    If lErro <> SUCESSO Then gError 124019
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoVendedor = Nothing
    Set objEventoCliente = Nothing
        
    Set objGridProcessos = Nothing
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_COMISSOES
    Set Form_Load_Ocx = Me
    Caption = "Contratos"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Contratos"
    
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

Private Sub BotaoExcluir_Click()

Dim lErro As Long, objContrato As New ClassContratoMgz
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    If Len(Trim(Cliente.Text)) = 0 Then gError 99999
    If Len(Trim(Contrato.Text)) = 0 Then gError 99999
    
    lErro = Move_Tela_Memoria(objContrato)
    If lErro <> SUCESSO Then gError 99999
    
    'Pede a confirmação da exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_CONTRATO")
    If vbMsgRes = vbNo Then Exit Sub
    
    'Faz a exclusão
    lErro = CF("ContratoMgz_Exclui", objContrato)
    If lErro <> SUCESSO Then gError 99999

    'Limpa a Tela
    Call Limpa_Tela_Contratos
    
    Exit Sub
     
Erro_BotaoExcluir_Click:

    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 99999

    'Limpa a Tela
    Call Limpa_Tela_Contratos

    Exit Sub
     
Erro_BotaoGravar_Click:

    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 99999

    'Limpa a Tela
    Call Limpa_Tela_Contratos

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

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

Private Function Inicializa_GridProcessos() As Long

Dim iIndice As Integer

    Set objGridProcessos = New AdmGrid

    'tela em questão
    Set objGridProcessos.objForm = Me

    'titulos do grid
    objGridProcessos.colColuna.Add ("")
    objGridProcessos.colColuna.Add ("Processo")
    objGridProcessos.colColuna.Add ("Tipo")
    objGridProcessos.colColuna.Add ("Valor")
    objGridProcessos.colColuna.Add ("Cobrança")
    objGridProcessos.colColuna.Add ("Descrição")
    objGridProcessos.colColuna.Add ("Observação")

   'campos de edição do grid
    objGridProcessos.colCampo.Add (ProcID.Name)
    objGridProcessos.colCampo.Add (ProcTipo.Name)
    objGridProcessos.colCampo.Add (ProcValor.Name)
    objGridProcessos.colCampo.Add (ProcDataCobr.Name)
    objGridProcessos.colCampo.Add (ProcDescricao.Name)
    objGridProcessos.colCampo.Add (ProcObs.Name)
    
    'Colunas do Grid
    iGrid_ProcID_Col = 1
    iGrid_ProcTipo_Col = 2
    iGrid_ProcValor_Col = 3
    iGrid_ProcDataCobr_Col = 4
    iGrid_ProcDescricao_Col = 5
    iGrid_ProcObservacao_Col = 6

    objGridProcessos.objGrid = GridProcessos
    
    'tulio 9/5/02
    GridProcessos.Rows = NUM_MAX_PROCESSOSCONTRATO
    
    'linhas visiveis do grid sem contar com as linhas fixas
    objGridProcessos.iLinhasVisiveis = 7

    GridProcessos.ColWidth(0) = 300

    objGridProcessos.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridProcessos.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    Call Grid_Inicializa(objGridProcessos)
    
    Inicializa_GridProcessos = SUCESSO

End Function

Function Trata_Parametros(Optional ByVal objContrato As ClassContratoMgz) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objContrato Is Nothing) Then
    
        lErro = Traz_ContratoMgz_Tela(objContrato)
        If lErro <> SUCESSO Then gError 99999
    
    End If
    
    Trata_Parametros = SUCESSO
     
    Exit Function
    
Erro_Trata_Parametros:

    Trata_Parametros = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long, objContrato As New ClassContratoMgz

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    lErro = Move_Tela_Memoria(objContrato)
    If lErro <> SUCESSO Then gError 99999
    
    lErro = Trata_Alteracao(objContrato, objContrato.sContrato, objContrato.lCliente)
    If lErro <> SUCESSO Then gError 99999

    lErro = CF("ContratoMgz_Grava", objContrato)
    If lErro <> SUCESSO Then gError 99999
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
     
    Exit Function
    
Erro_Gravar_Registro:

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Public Function Move_Tela_Memoria(ByVal objContrato As ClassContratoMgz) As Long

Dim lErro As Long
Dim objVendedor As New ClassVendedor
Dim objCliente As New ClassCliente

On Error GoTo Erro_Move_Tela_Memoria

    objContrato.sContrato = Trim(Contrato.Text)
            
    If Len(Trim(Cliente.ClipText)) > 0 Then

        objCliente.sNomeReduzido = Cliente.Text
        
        'Lê o Cliente
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 99999
        
        'Não encontrou p Cliente --> erro
        If lErro = 12348 Then gError 99999

        objContrato.lCliente = objCliente.lCodigo
        
    End If
    
    If Len(Trim(Vendedor.Text)) > 0 Then objVendedor.sNomeReduzido = Vendedor.Text

    'Verifica se vendedor existe
    If objVendedor.sNomeReduzido <> "" Then
        lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
        If lErro <> SUCESSO And lErro <> 25008 Then gError 99999

        'Não encontrou o vendedor ==> erro
        If lErro = 25008 Then gError 99999

        objContrato.iVendedor = objVendedor.iCodigo

    End If
    
    lErro = Move_Processos_Memoria(objContrato)
    If lErro <> SUCESSO Then gError 99999

    Move_Tela_Memoria = SUCESSO
     
    Exit Function
    
Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Private Function Move_Processos_Memoria(ByVal objContrato As ClassContratoMgz) As Long

Dim lErro As Long, objProc As ClassProcContratoMgz, iIndice As Integer, sTipo As String, iAux As Integer

On Error GoTo Erro_Move_Processos_Memoria

    For iIndice = 1 To objGridProcessos.iLinhasExistentes
    
        Set objProc = New ClassProcContratoMgz
    
        With objProc
    
            .lCliente = objProc.lCliente
            .sContrato = objProc.sContrato
            
            .sProcesso = GridProcessos.TextMatrix(iIndice, iGrid_ProcID_Col)
            
            If Len(Trim(GridProcessos.TextMatrix(iIndice, iGrid_ProcDataCobr_Col))) > 0 Then
                .dtDataCobranca = CDate(GridProcessos.TextMatrix(iIndice, iGrid_ProcDataCobr_Col))
            Else
                .dtDataCobranca = DATA_NULA
            End If
            
            .dValor = StrParaDbl(GridProcessos.TextMatrix(iIndice, iGrid_ProcValor_Col))
            .iSeq = iIndice
            
            sTipo = GridProcessos.TextMatrix(iIndice, iGrid_ProcTipo_Col)
            
            For iAux = 0 To ProcTipo.ListCount - 1
                If ProcTipo.List(iAux) = sTipo Then
                    .iTipo = ProcTipo.ItemData(iAux)
                    Exit For
                End If
            Next
            
            .sDescricao = GridProcessos.TextMatrix(iIndice, iGrid_ProcDescricao_Col)
            .sObservacao = GridProcessos.TextMatrix(iIndice, iGrid_ProcObservacao_Col)
        
        End With
        
        objContrato.colProcessos.Add objProc
        
    Next
    
    Move_Processos_Memoria = SUCESSO
     
    Exit Function
    
Erro_Move_Processos_Memoria:

    Move_Processos_Memoria = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Function Traz_ContratoMgz_Tela(ByVal objContrato As ClassContratoMgz) As Long

Dim lErro As Long, objProc As ClassProcContratoMgz, iAux As Integer, iIndice As Integer

On Error GoTo Erro_Traz_ContratoMgz_Tela

    lErro = CF("ContratoMgz_Le", objContrato)
    If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError 99999
    
    Contrato.Text = objContrato.sContrato
    
    Cliente.Text = CStr(objContrato.lCliente)
    Call Cliente_Validate(bSGECancelDummy)
    
    Vendedor.Text = CStr(objContrato.iVendedor)
    Call Vendedor_Validate(bSGECancelDummy)
    
    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridProcessos)

    For Each objProc In objContrato.colProcessos
    
        iIndice = iIndice + 1
        
        GridProcessos.TextMatrix(iIndice, iGrid_ProcID_Col) = objProc.sProcesso
        
        If objProc.dtDataCobranca <> DATA_NULA Then GridProcessos.TextMatrix(iIndice, iGrid_ProcDataCobr_Col) = Format(objProc.dtDataCobranca, "dd/mm/yyyy")
        
        GridProcessos.TextMatrix(iIndice, iGrid_ProcValor_Col) = Format(objProc.dValor, "Standard")
        
        For iAux = 0 To ProcTipo.ListCount - 1
            If ProcTipo.ItemData(iAux) = objProc.iTipo Then
                GridProcessos.TextMatrix(iIndice, iGrid_ProcTipo_Col) = ProcTipo.List(iAux)
                Exit For
            End If
        Next
        
        GridProcessos.TextMatrix(iIndice, iGrid_ProcDescricao_Col) = objProc.sDescricao
        GridProcessos.TextMatrix(iIndice, iGrid_ProcObservacao_Col) = objProc.sObservacao
    
    Next
    
    objGridProcessos.iLinhasExistentes = objContrato.colProcessos.Count
    
    iAlterado = 0
    
    Traz_ContratoMgz_Tela = SUCESSO
     
    Exit Function
    
Erro_Traz_ContratoMgz_Tela:

    Traz_ContratoMgz_Tela = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Private Sub GridProcessos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridProcessos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridProcessos, iAlterado)
    End If

End Sub

Private Sub GridProcessos_GotFocus()

    Call Grid_Recebe_Foco(objGridProcessos)

End Sub

Private Sub GridProcessos_EnterCell()

    Call Grid_Entrada_Celula(objGridProcessos, iAlterado)

End Sub

Private Sub GridProcessos_LeaveCell()

    Call Saida_Celula(objGridProcessos)

End Sub

Private Sub GridProcessos_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridProcessos)

End Sub

Private Sub GridProcessos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridProcessos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridProcessos, iAlterado)
    End If

End Sub

Private Sub GridProcessos_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridProcessos)
    
End Sub

Private Sub GridProcessos_RowColChange()

    Call Grid_RowColChange(objGridProcessos)

End Sub

Private Sub GridProcessos_Scroll()

    Call Grid_Scroll(objGridProcessos)

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        Select Case GridProcessos.Col
    
            Case iGrid_ProcID_Col

                lErro = Saida_Celula_ProcID(objGridInt)
                If lErro <> SUCESSO Then gError 99999

            Case iGrid_ProcTipo_Col

                lErro = Saida_Celula_ProcTipo(objGridInt)
                If lErro <> SUCESSO Then gError 99999

            Case iGrid_ProcDescricao_Col

                lErro = Saida_Celula_ProcDescricao(objGridInt)
                If lErro <> SUCESSO Then gError 99999

            Case iGrid_ProcValor_Col

                lErro = Saida_Celula_ProcValor(objGridInt)
                If lErro <> SUCESSO Then gError 99999

            Case iGrid_ProcDataCobr_Col

                lErro = Saida_Celula_ProcDataCobr(objGridInt)
                If lErro <> SUCESSO Then gError 99999

            Case iGrid_ProcObservacao_Col

                lErro = Saida_Celula_ProcObs(objGridInt)
                If lErro <> SUCESSO Then gError 99999

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 124021

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 124021
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Sub ProcDescricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProcDescricao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProcessos)

End Sub

Private Sub ProcDescricao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProcessos)

End Sub

Private Sub ProcDescricao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProcessos.objControle = ProcDescricao
    lErro = Grid_Campo_Libera_Foco(objGridProcessos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ProcObs_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProcObs_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProcessos)

End Sub

Private Sub ProcObs_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProcessos)

End Sub

Private Sub ProcObs_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProcessos.objControle = ProcObs
    lErro = Grid_Campo_Libera_Foco(objGridProcessos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ProcValor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProcValor_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProcessos)

End Sub

Private Sub ProcValor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProcessos)

End Sub

Private Sub ProcValor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProcessos.objControle = ProcValor
    lErro = Grid_Campo_Libera_Foco(objGridProcessos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ProcDataCobr_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProcDataCobr_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProcessos)

End Sub

Private Sub ProcDataCobr_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProcessos)

End Sub

Private Sub ProcDataCobr_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProcessos.objControle = ProcDataCobr
    lErro = Grid_Campo_Libera_Foco(objGridProcessos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ProcID_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProcID_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProcessos)

End Sub

Private Sub ProcID_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProcessos)

End Sub

Private Sub ProcID_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProcessos.objControle = ProcID
    lErro = Grid_Campo_Libera_Foco(objGridProcessos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ProcTipo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProcTipo_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridProcessos)

End Sub

Private Sub ProcTipo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridProcessos)

End Sub

Private Sub ProcTipo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridProcessos.objControle = ProcTipo
    lErro = Grid_Campo_Libera_Foco(objGridProcessos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_ProcDescricao(objGridInt As AdmGrid) As Long
'faz a critica da celula conta do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ProcDescricao

    Set objGridInt.objControle = ProcDescricao

    If GridProcessos.Row - GridProcessos.FixedRows = objGridInt.iLinhasExistentes Then
    
        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 22348
    
    Saida_Celula_ProcDescricao = SUCESSO

    Exit Function

Erro_Saida_Celula_ProcDescricao:

    Saida_Celula_ProcDescricao = Err

    Select Case Err

        Case 22348
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ProcObs(objGridInt As AdmGrid) As Long
'faz a critica da celula conta do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ProcObs

    Set objGridInt.objControle = ProcObs

    If GridProcessos.Row - GridProcessos.FixedRows = objGridInt.iLinhasExistentes Then
    
        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 22348
    
    Saida_Celula_ProcObs = SUCESSO

    Exit Function

Erro_Saida_Celula_ProcObs:

    Saida_Celula_ProcObs = Err

    Select Case Err

        Case 22348
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ProcValor(objGridInt As AdmGrid) As Long
'faz a critica da celula conta do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ProcValor

    Set objGridInt.objControle = ProcValor

    If GridProcessos.Row - GridProcessos.FixedRows = objGridInt.iLinhasExistentes Then
    
        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 22348
    
    Saida_Celula_ProcValor = SUCESSO

    Exit Function

Erro_Saida_Celula_ProcValor:

    Saida_Celula_ProcValor = Err

    Select Case Err

        Case 22348
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ProcDataCobr(objGridInt As AdmGrid) As Long
'faz a critica da celula conta do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ProcDataCobr

    Set objGridInt.objControle = ProcDataCobr

    If GridProcessos.Row - GridProcessos.FixedRows = objGridInt.iLinhasExistentes Then
    
        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 22348
    
    Saida_Celula_ProcDataCobr = SUCESSO

    Exit Function

Erro_Saida_Celula_ProcDataCobr:

    Saida_Celula_ProcDataCobr = Err

    Select Case Err

        Case 22348
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ProcID(objGridInt As AdmGrid) As Long
'faz a critica da celula conta do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ProcID

    Set objGridInt.objControle = ProcID

    If GridProcessos.Row - GridProcessos.FixedRows = objGridInt.iLinhasExistentes Then
    
        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 22348
    
    Saida_Celula_ProcID = SUCESSO

    Exit Function

Erro_Saida_Celula_ProcID:

    Saida_Celula_ProcID = Err

    Select Case Err

        Case 22348
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ProcTipo(objGridInt As AdmGrid) As Long
'faz a critica da celula conta do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ProcTipo

    Set objGridInt.objControle = ProcTipo

    If GridProcessos.Row - GridProcessos.FixedRows = objGridInt.iLinhasExistentes Then
    
        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 22348
    
    Saida_Celula_ProcTipo = SUCESSO

    Exit Function

Erro_Saida_Celula_ProcTipo:

    Saida_Celula_ProcTipo = Err

    Select Case Err

        Case 22348
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

    Exit Function

End Function

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long, iFilial As Integer
Dim objCliente As New ClassCliente

On Error GoTo Erro_Cliente_Validate

    If Len(Trim(Cliente.Text)) > 0 Then
   
        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le(Cliente, objCliente, iFilial)
        If lErro <> SUCESSO Then Error 37793

    End If
    
    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case Err

        Case 37793
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO_2", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error)

    End Select

End Sub

Private Sub LabelCliente_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    If Len(Trim(Cliente.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(Cliente.Text)
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1
    
    'Preenche campo Cliente
    Cliente.Text = CStr(objCliente.lCodigo)
    Call Cliente_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub LabelVendedor_Click()
'sub que chama o browser de vendedores

Dim objVendedor As New ClassVendedor
Dim colSelecao As New Collection
    
    'se o vendedor estiver preenchido
    If Len(Trim(Vendedor.Text)) <> 0 Then
        'carrega o obj c/ o nomereduzido do vendedor
        objVendedor.sNomeReduzido = Vendedor.Text
    End If
    
    'Chama tela que lista todos os vendores
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)
'evento de inclusão de um item selecionado no browser vendedor

Dim objVendedor As ClassVendedor

On Error GoTo Erro_objEventoVendedor_evSelecao

    Set objVendedor = obj1
    
    'Preenche o Vendedor c/ o nomereduzido
    Vendedor.Text = objVendedor.iCodigo

    Call Vendedor_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

Erro_objEventoVendedor_evSelecao:

    Select Case gErr
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub Vendedor_Validate(Cancel As Boolean)
'verifica se o vendedor existe

Dim lErro As Long
Dim objVendedor As New ClassVendedor
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Vendedor_Validate

    'Verifica se vendedor está preenchido
    If Len(Trim(Vendedor.Text)) = 0 Then Exit Sub

    'Verifica se Vendedor está cadastrado no bd
    lErro = TP_Vendedor_Le(Vendedor, objVendedor)
    If lErro <> SUCESSO Then gError 119549
    
    Exit Sub

Erro_Vendedor_Validate:

    Cancel = True

    Select Case gErr

        Case 119549
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objContrato As New ClassContratoMgz

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Contratos"

    'Le os dados da Tela Almoxarifado
    lErro = Move_Tela_Memoria(objContrato)
    If lErro <> SUCESSO Then Error 22314

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Contrato", objContrato.sContrato, STRING_PROCESSO_CONTRATO_ID, "Contrato"
    colCampoValor.Add "Cliente", objContrato.lCliente, 0, "Cliente"
    colCampoValor.Add "Vendedor", objContrato.iVendedor, 0, "Vendedor"

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 22314

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objContrato As New ClassContratoMgz

On Error GoTo Erro_Tela_Preenche

    objContrato.sContrato = colCampoValor.Item("Contrato").vValor
    objContrato.lCliente = colCampoValor.Item("Cliente").vValor

    If Len(Trim(objContrato.sContrato)) > 0 Then

        'Traz dados para a Tela
        lErro = Traz_ContratoMgz_Tela(objContrato)
        If lErro <> SUCESSO Then Error 22315

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 22315

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub



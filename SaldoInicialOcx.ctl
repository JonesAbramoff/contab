VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl SaldoInicialOcx 
   ClientHeight    =   5610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9180
   ScaleHeight     =   5610
   ScaleWidth      =   9180
   Begin VB.TextBox DescCcl 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   2235
      MaxLength       =   150
      TabIndex        =   8
      Top             =   1170
      Width           =   2130
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7395
      ScaleHeight     =   495
      ScaleWidth      =   1605
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   1665
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1125
         Picture         =   "SaldoInicialOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   615
         Picture         =   "SaldoInicialOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   105
         Picture         =   "SaldoInicialOcx.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
   End
   Begin VB.ListBox ListaConta 
      Height          =   4350
      Left            =   6075
      TabIndex        =   3
      Top             =   1050
      Width           =   2985
   End
   Begin MSMask.MaskEdBox Saldo 
      Height          =   225
      Left            =   4380
      TabIndex        =   1
      Top             =   1170
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
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
   Begin MSMask.MaskEdBox Ccl 
      Height          =   225
      Left            =   840
      TabIndex        =   0
      Top             =   1155
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      AllowPrompt     =   -1  'True
      Enabled         =   0   'False
      MaxLength       =   10
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
   Begin MSFlexGridLib.MSFlexGrid GridSaldos 
      Height          =   4470
      Left            =   120
      TabIndex        =   2
      Top             =   1050
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   7885
      _Version        =   393216
      Rows            =   50
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
   End
   Begin VB.Label DescConta 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   2625
      TabIndex        =   9
      Top             =   285
      Width           =   2685
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Contas com Centro de Custo"
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
      Left            =   6105
      TabIndex        =   10
      Top             =   825
      Width           =   2430
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Conta:"
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
      Left            =   300
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   11
      Top             =   315
      Width           =   585
   End
   Begin VB.Label Conta 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   960
      TabIndex        =   12
      Top             =   285
      Width           =   1590
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Saldos Iniciais"
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
      Left            =   135
      TabIndex        =   13
      Top             =   825
      Width           =   1245
   End
End
Attribute VB_Name = "SaldoInicialOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim objGrid1 As AdmGrid
Dim iAlterado As Integer
Const GRID_CCL_COL = 1
Const GRID_DESCCCL_COL = 2
Const GRID_SALDO_INICIAL_COL = 3
Private WithEvents objEventoSaldoInicialContaCcl As AdmEvento
Attribute objEventoSaldoInicialContaCcl.VB_VarHelpID = -1

Private Sub Label1_Click()

Dim lErro As Long
Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As Collection

On Error GoTo Erro_Label1_Click

    'Recolhe o número da Conta
    If Len(Conta.Caption) > 0 Then
        objPlanoConta.sConta = Conta.Caption
    Else
        objPlanoConta.sConta = ""
    End If
    
    'Chama tela de Saldos
    Call Chama_Tela("PlanoContaLista", colSelecao, objPlanoConta, objEventoSaldoInicialContaCcl)

    Exit Sub

Erro_Label1_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174319)

    End Select

    Exit Sub

End Sub

Private Sub objEventoSaldoInicialContaCcl_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta

On Error GoTo Erro_objEventoSaldoInicialContaCcl_evSelecao

    Set objPlanoConta = obj1

    lErro = Traz_ContaCcl_Tela(objPlanoConta.sConta)
    If lErro <> SUCESSO Then Error 34664

    'Fecha Comando de Seta
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoSaldoInicialContaCcl_evSelecao:

    Select Case Err

        Case 34664

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174320)

    End Select

    Exit Sub

End Sub

Private Sub ListaConta_DblClick()

Dim lErro As Long
Dim sContaFormatada As String
Dim iContaPreenchida As Integer
Dim sConta As String

On Error GoTo Erro_ListaConta_DblClick

    sConta = Left(ListaConta.Text, InStr(ListaConta.Text, SEPARADOR) - 2)

    'coloca a conta no formato do bd
    lErro = CF("Conta_Formata", sConta, sContaFormatada, iContaPreenchida)
    If lErro <> SUCESSO Then Error 10035

    lErro = Traz_ContaCcl_Tela(sContaFormatada)
    If lErro <> SUCESSO Then Error 10036

    iAlterado = 0

    Exit Sub

Erro_ListaConta_DblClick:

    Select Case Err

        Case 10035, 10036

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174321)

    End Select

    Exit Sub

End Sub

Private Sub Saldo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Saldo_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub Saldo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub Saldo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Saldo
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

 Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Private Sub GridSaldos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid1, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If

End Sub

Private Sub GridSaldos_GotFocus()

    Call Grid_Recebe_Foco(objGrid1)

End Sub

Private Sub GridSaldos_EnterCell()

    Call Grid_Entrada_Celula(objGrid1, iAlterado)

End Sub

Private Sub GridSaldos_LeaveCell()

    Call Saida_Celula(objGrid1)

End Sub

Private Sub GridSaldos_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGrid1)

End Sub

Private Sub GridSaldos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid1, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If

End Sub

Private Sub GridSaldos_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGrid1)

End Sub

Private Sub GridSaldos_RowColChange()

    Call Grid_RowColChange(objGrid1)

End Sub

Private Sub GridSaldos_Scroll()

    Call Grid_Scroll(objGrid1)

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objGrid1 = New AdmGrid
    Set objEventoSaldoInicialContaCcl = New AdmEvento

    'tela em questão
    Set objGrid1.objForm = Me

    lErro = Inicializa_Grid_Saldo(objGrid1)
    If lErro <> SUCESSO Then Error 9937

    'Inicializa a Lista de Contas com associação com centro de custo/lucro
    lErro = Carga_ListBox_Conta()
    If lErro <> SUCESSO Then Error 9938

    Conta.Caption = ""
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 9937, 9938

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174322)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objPlanoConta As ClassPlanoConta) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se foi passado uma conta como parametro
    If Not (objPlanoConta Is Nothing) Then

        lErro = Traz_ContaCcl_Tela(objPlanoConta.sConta)
        If lErro <> SUCESSO Then Error 9939

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 9939

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174323)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Function Carga_ListBox_Conta() As Long
'le do bd as contas que estão associadas a algum centro de custo e coloca na listbox

Dim colConta As New Collection
Dim lErro As Long
Dim sContaMascarada As String
Dim sConta As String
Dim objPlanoConta As ClassPlanoConta

On Error GoTo Erro_Carga_ListBox_Conta

    'Le as contas que possuem centro de custo/lucro associados
    lErro = CF("ContaCcl_Le_Todas_Contas_Distintas", colConta)
    If lErro <> SUCESSO Then Error 9955

    'se não houver nenhuma conta com centro de custo/lucro associado ==> erro
    If colConta.Count = 0 Then Error 9956

    ListaConta.Clear

    'Coloca cada conta encontrada na listbox
    For Each objPlanoConta In colConta

        sContaMascarada = String(STRING_CONTA, 0)

        'coloca a conta no formato que é exibida na tela
        lErro = Mascara_MascararConta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then Error 9957

        'adiciona a conta na listbox
        ListaConta.AddItem sContaMascarada & " " & SEPARADOR & " " & objPlanoConta.sDescConta

    Next

    Carga_ListBox_Conta = SUCESSO

    Exit Function

Erro_Carga_ListBox_Conta:

    Carga_ListBox_Conta = Err

    Select Case Err

        Case 9955

        Case 9956
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACCL_VAZIO", Err)

        Case 9957
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararConta", Err, sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174324)

    End Select

    Exit Function

End Function

Private Function Traz_ContaCcl_Tela(sConta As String) As Long
'traz os centros de custo associados a conta em questão

Dim lErro As Long
Dim colSaldoInicialContaCcl As New Collection
Dim objContaCcl As ClassContaCcl
Dim sCclMascarado As String
Dim iLinha As Integer
Dim sContaMascarada As String
Dim objSaldoInicialContaCcl As New ClassSaldoInicialContaCcl
Dim objPlanoConta As New ClassPlanoConta

On Error GoTo Erro_Traz_ContaCcl_Tela

    If Len(sConta) > 0 Then
    
        objSaldoInicialContaCcl.iFilialEmpresa = giFilialEmpresa
        objSaldoInicialContaCcl.sConta = sConta

        'le todas as associações da conta com centros de custo/lucro
        lErro = CF("SaldoInicialContaCcl_Le_Todos_Conta", objSaldoInicialContaCcl, colSaldoInicialContaCcl)
        If lErro <> SUCESSO Then Error 9959

        'se não houver nenhuma associação cadastrada ==> erro
        If colSaldoInicialContaCcl.Count = 0 Then Error 9960

        sContaMascarada = String(STRING_CONTA, 0)

        'coloca a conta no formato que é exibida na tela
        lErro = Mascara_MascararConta(sConta, sContaMascarada)
        If lErro <> SUCESSO Then Error 9958

        lErro = CF("PlanoConta_Le_Conta1", sConta, objPlanoConta)
        If lErro <> SUCESSO Then Error 60786

        'coloca a conta mascarada na tela
        Conta.Caption = sContaMascarada
        DescConta.Caption = objPlanoConta.sDescConta

        'limpa o conteudo do grid
        Call Grid_Limpa(objGrid1)

        iLinha = 0
        
        '####################################
        'Inserido por Wagner 11/01/2006
        If colSaldoInicialContaCcl.Count >= objGrid1.objGrid.Rows Then
            Call Refaz_Grid(objGrid1, colSaldoInicialContaCcl.Count)
        End If
        '####################################

        'Para cada centro de custo/lucro associado à conta
        For Each objSaldoInicialContaCcl In colSaldoInicialContaCcl

            iLinha = iLinha + 1

            'mascara o centro de custo
            sCclMascarado = String(STRING_CCL, 0)

            lErro = Mascara_MascararCcl(objSaldoInicialContaCcl.sCcl, sCclMascarado)
            If lErro <> SUCESSO Then Error 9961

            'coloca o centro de custo no grid
            GridSaldos.TextMatrix(iLinha, GRID_CCL_COL) = sCclMascarado
            GridSaldos.TextMatrix(iLinha, GRID_DESCCCL_COL) = objSaldoInicialContaCcl.sDescCcl

            If objSaldoInicialContaCcl.dSldIni < 0 Then objSaldoInicialContaCcl.dSldIni = -objSaldoInicialContaCcl.dSldIni

            'coloca o saldo inicial no grid
            GridSaldos.TextMatrix(iLinha, GRID_SALDO_INICIAL_COL) = Format(objSaldoInicialContaCcl.dSldIni, "Standard")

            objGrid1.iLinhasExistentes = objGrid1.iLinhasExistentes + 1

        Next
    
    End If

    iAlterado = 0
    
    Traz_ContaCcl_Tela = SUCESSO

    Exit Function

Erro_Traz_ContaCcl_Tela:

    Traz_ContaCcl_Tela = Err

    Select Case Err

        Case 9958
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararConta", Err, sConta)

        Case 9959, 60786

        Case 9960
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_SEM_CONTACCL", Err, sConta)

        Case 9961
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, objContaCcl.sCcl)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174325)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Saldo(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Inicializa_Grid_Saldo

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Centro de Custo")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Saldo Inicial")

   'campos de edição do grid
    objGridInt.colCampo.Add (Ccl.Name)
    objGridInt.colCampo.Add (DescCcl.Name)
    objGridInt.colCampo.Add (Saldo.Name)

    lErro = Inicializa_Mascaras()
    If lErro <> SUCESSO Then Error 9948

    objGridInt.objGrid = GridSaldos

    'todas as linhas do grid
    objGridInt.objGrid.Rows = 51

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 17

    GridSaldos.ColWidth(0) = 585

    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Saldo = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Saldo:

    Inicializa_Grid_Saldo = Err

    Select Case Err

        Case 9948

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174326)

    End Select

    Exit Function

End Function

Private Function Inicializa_Mascaras() As Long
'inicializa as mascaras de conta e centro de custo

Dim sMascaraConta As String
Dim sMascaraCcl As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_Mascaras

    sMascaraCcl = String(STRING_CCL, 0)

    'le a mascara dos centros de custo/lucro
    lErro = MascaraCcl(sMascaraCcl)
    If lErro <> SUCESSO Then Error 9940

    Ccl.Mask = sMascaraCcl

    Inicializa_Mascaras = SUCESSO

    Exit Function

Erro_Inicializa_Mascaras:

    Inicializa_Mascaras = Err

    Select Case Err

        Case 9940

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174327)

    End Select

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        Select Case GridSaldos.Col

            Case GRID_SALDO_INICIAL_COL

                lErro = Saida_Celula_Saldo(objGridInt)
                If lErro <> SUCESSO Then Error 9962


        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 9963

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 9962

        Case 9963
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174328)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Saldo(objGridInt As AdmGrid) As Long
'faz a critica da celula saldo inicial do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Saldo

    Set objGridInt.objControle = Saldo
    If Len(Saldo.Text) > 0 Then
        lErro = Valor_NaoNegativo_Critica(Saldo.Text)
        If lErro <> SUCESSO Then Error 9964
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 9965

    Saida_Celula_Saldo = SUCESSO

    Exit Function

Erro_Saida_Celula_Saldo:

    Saida_Celula_Saldo = Err

    Select Case Err

        Case 9964, 9965
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174329)

    End Select

    Exit Function

End Function

Private Sub BotaoGravar_Click()

    Call Gravar_Registro

    Call Grid_Limpa(objGrid1)

    Conta.Caption = ""

    iAlterado = 0

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim colSaldoInicialContaCcl As New Collection

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    lErro = Mover_Tela_Memoria(colSaldoInicialContaCcl)
    If lErro <> SUCESSO Then Error 9972

    If giSetupUsoCcl = CCL_USA_CONTABIL Then

        lErro = CF("ContaCcl_Atualiza_Saldo_Contabil", colSaldoInicialContaCcl)
        If lErro <> SUCESSO Then Error 9973

    Else

        lErro = CF("ContaCcl_Atualiza_Saldo_Extra", colSaldoInicialContaCcl)
        If lErro <> SUCESSO Then Error 9995

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 9972, 9973, 9995

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174330)

    End Select

    Exit Function

End Function

Function Mover_Tela_Memoria(colSaldoInicialContaCcl As Collection) As Long
'move os dados da tela para a coleção colContaCcl

Dim iLinha As Integer
Dim sContaFormatada As String
Dim sCcl As String
Dim sCclFormatada As String
Dim dSaldo As Double
Dim iContaPreenchida As Integer
Dim iCclPreenchida As Integer
Dim lErro As Long
Dim objSaldoInicialContaCcl As New ClassSaldoInicialContaCcl
Dim colSaldoInicialContaCcl1 As New Collection

On Error GoTo Erro_Mover_Tela_Memoria

    'coloca a conta no formato do bd
    lErro = CF("Conta_Formata", Conta.Caption, sContaFormatada, iContaPreenchida)
    If lErro <> SUCESSO Then Error 9969

    'se a conta não tiver sido preenchida ==> erro
    If iContaPreenchida = CONTA_VAZIA Then Error 9970

    objSaldoInicialContaCcl.iFilialEmpresa = giFilialEmpresa
    objSaldoInicialContaCcl.sConta = sContaFormatada

    'le todas as associações da conta com centros de custo/lucro
    lErro = CF("SaldoInicialContaCcl_Le_Todos_Conta", objSaldoInicialContaCcl, colSaldoInicialContaCcl1)
    If lErro <> SUCESSO Then Error 55706

    'se não houver nenhuma associação cadastrada ==> erro
    If colSaldoInicialContaCcl1.Count = 0 Then Error 55707
    
    'Para cada centro de custo/lucro presente no grid
    For iLinha = 1 To objGrid1.iLinhasExistentes

        Set objSaldoInicialContaCcl = New ClassSaldoInicialContaCcl

        objSaldoInicialContaCcl.iFilialEmpresa = giFilialEmpresa

        objSaldoInicialContaCcl.sConta = sContaFormatada

        sCcl = GridSaldos.TextMatrix(iLinha, GRID_CCL_COL)

        'coloca o centro de custo/lucro no formato do bd
        lErro = CF("Ccl_Formata", sCcl, sCclFormatada, iCclPreenchida)
        If lErro <> SUCESSO Then Error 9971

        objSaldoInicialContaCcl.sCcl = sCclFormatada

        If Len(GridSaldos.TextMatrix(iLinha, GRID_SALDO_INICIAL_COL)) > 0 Then
            objSaldoInicialContaCcl.dSldIni = CDbl(GridSaldos.TextMatrix(iLinha, GRID_SALDO_INICIAL_COL))
        Else
            objSaldoInicialContaCcl.dSldIni = 0
        End If

        colSaldoInicialContaCcl.Add objSaldoInicialContaCcl

    Next

    Mover_Tela_Memoria = SUCESSO

    Exit Function

Erro_Mover_Tela_Memoria:

    Mover_Tela_Memoria = Err

    Select Case Err

        Case 9969, 9971, 55706

        Case 9970
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_INFORMADA", Err)

        Case 55707
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_SEM_CONTACCL", Err, Conta.Caption)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174331)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 9973

    Call Grid_Limpa(objGrid1)

    Conta.Caption = ""
    DescConta.Caption = ""

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 9973

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174332)

    End Select

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub ListaConta_KeyPress(KeyAscii As Integer)

    If ListaConta.ListIndex <> -1 Then

        If KeyAscii = ENTER_KEY Then

            Call ListaConta_DblClick

        End If

    End If

End Sub

Public Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objGrid1 = Nothing
        
    Set objEventoSaldoInicialContaCcl = Nothing

   'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
 
End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objSaldoInicialContaCcl As New ClassSaldoInicialContaCcl
Dim colSaldoInicialContaCcl As New Collection

On Error GoTo Erro_Tela_Extrai

    'Informa a tabela associada à tela
    sTabela = "SaldoInicialContaCcl"

    If Len(Trim(Conta.Caption)) > 0 Then
        objSaldoInicialContaCcl.sConta = Conta.Caption
    Else
        objSaldoInicialContaCcl.sConta = 0
    End If

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Conta", objSaldoInicialContaCcl.sConta, STRING_CONTA, "Conta"

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 34662

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174333)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objSaldoInicialContaCcl As New ClassSaldoInicialContaCcl

On Error GoTo Erro_Tela_Preenche

    objSaldoInicialContaCcl.sConta = colCampoValor.Item("Conta").vValor

    'Preenche a tela com os dados retornados
    lErro = Traz_ContaCcl_Tela(objSaldoInicialContaCcl.sConta)
    If lErro <> SUCESSO Then Error 34663

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 34663

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174334)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_SALDOS_INICIAIS_CENTRO_CUSTO_LUCROS
    Set Form_Load_Ocx = Me
    Caption = "Saldos Iniciais - Centros de Custo/Lucro"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "SaldoInicial"
    
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



Private Sub DescConta_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescConta, Source, X, Y)
End Sub

Private Sub DescConta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescConta, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Conta_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Conta, Source, X, Y)
End Sub

Private Sub Conta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Conta, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

'#########################################################
'Inserido por Wagner
Sub Refaz_Grid(ByVal objGridInt As AdmGrid, ByVal iNumLinhas As Integer)
    objGridInt.objGrid.Rows = iNumLinhas + 1

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)
End Sub
'#########################################################

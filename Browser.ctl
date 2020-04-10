VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl Browser 
   ClientHeight    =   5490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7650
   DefaultCancel   =   -1  'True
   ScaleHeight     =   5490
   ScaleWidth      =   7650
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Busca"
      Height          =   480
      Left            =   3750
      TabIndex        =   20
      Top             =   1065
      Width           =   3690
      Begin VB.OptionButton OptFiltrar 
         Caption         =   "&Filtrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2760
         TabIndex        =   3
         Top             =   225
         Width           =   825
      End
      Begin VB.OptionButton OptPosicionar 
         Caption         =   "Posicionar por ordenação"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   90
         TabIndex        =   2
         Top             =   225
         Value           =   -1  'True
         Width           =   2520
      End
   End
   Begin VB.Timer Timer1 
      Left            =   4200
      Top             =   225
   End
   Begin MSMask.MaskEdBox Pesquisa 
      Height          =   315
      Left            =   1125
      TabIndex        =   22
      Top             =   1215
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   " "
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "Browser.ctx":0000
      Left            =   1140
      List            =   "Browser.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   18
      Top             =   225
      Width           =   2475
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6270
      ScaleHeight     =   495
      ScaleWidth      =   1095
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   105
      Width           =   1155
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "Browser.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "Browser.ctx":015E
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.VScrollBar BarraScroll 
      Height          =   2745
      LargeChange     =   10
      Left            =   7200
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1620
      Width           =   255
   End
   Begin VB.ComboBox Ordenacao 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   735
      Width           =   6300
   End
   Begin VB.PictureBox Picture2 
      Height          =   975
      Left            =   75
      ScaleHeight     =   915
      ScaleWidth      =   7350
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4425
      Width           =   7410
      Begin VB.CommandButton BotaoPlanilha 
         Caption         =   "Planilha"
         Height          =   780
         Left            =   90
         Picture         =   "Browser.ctx":02E8
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   75
         Width           =   930
      End
      Begin VB.CommandButton BotaoConsulta 
         Caption         =   "Consultar"
         Height          =   780
         Left            =   3885
         Picture         =   "Browser.ctx":1162
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   75
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.CommandButton BotaoAtualiza 
         Caption         =   "Atualizar"
         Height          =   780
         Left            =   5325
         Picture         =   "Browser.ctx":1E24
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   75
         Width           =   915
      End
      Begin VB.CommandButton BotaoPesquisa 
         Caption         =   "Pesquisar"
         Height          =   780
         Left            =   4290
         Picture         =   "Browser.ctx":2146
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   75
         Width           =   915
      End
      Begin VB.CommandButton BotaoConfigura 
         Caption         =   "Configurar"
         Height          =   780
         Left            =   3255
         Picture         =   "Browser.ctx":23A8
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   75
         Width           =   930
      End
      Begin VB.CommandButton BotaoEdita 
         Caption         =   "Editar"
         Height          =   780
         Left            =   2205
         Picture         =   "Browser.ctx":27EA
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   75
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.CommandButton BotaoSeleciona 
         Caption         =   "Selecionar"
         Default         =   -1  'True
         Height          =   780
         Left            =   1155
         Picture         =   "Browser.ctx":2C2C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   75
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.CommandButton BotaoFecha 
         Cancel          =   -1  'True
         Caption         =   "Fechar"
         Height          =   780
         Left            =   6345
         Picture         =   "Browser.ctx":306E
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   75
         Width           =   930
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridBrowse 
      Height          =   2745
      Left            =   75
      TabIndex        =   1
      Top             =   1605
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   4842
      _Version        =   393216
      Rows            =   50
      Cols            =   4
      FixedCols       =   0
      ForeColorSel    =   16777215
      AllowBigSelection=   0   'False
      Enabled         =   -1  'True
      FocusRect       =   0
      AllowUserResizing=   1
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "&Procurar:"
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
      Left            =   285
      TabIndex        =   21
      Top             =   1260
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Opção:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   450
      TabIndex        =   19
      Top             =   285
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "O&rdenação:"
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
      Left            =   75
      TabIndex        =   13
      Top             =   765
      Width           =   990
   End
End
Attribute VB_Name = "Browser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Property Variables:
Dim m_Caption As String
Dim sName As String
Event Unload()
Dim iDistanciaBorda As Integer
Dim iNumeroBotoes As Integer
Dim gsOrdenacao As String
Public gsSiglaModuloChamador As String
Dim giPesquisaChange As Integer
Dim gsSelecaoSQLOrig As String
Dim gcolSelecaoOrig As New Collection
Dim gPrecisaRemoverFiltro As Boolean

Dim gbCarregandoTela As Boolean

Private objEvento As AdmEvento
Public objBrowse1 As AdmBrowse

Const ALTURA_LINHA_GRID = 250
Const DISTANCIA_ENTRE_GRID_PICTURE = 150

Private bCargaInicialECF As Boolean

Private Sub BotaoAtualiza_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoAtualiza_Click

    lErro = objBrowse1.Browse_BarraScroll_Change(objBrowse1)
    If lErro <> SUCESSO Then gError 89952

    Exit Sub
    
Erro_BotaoAtualiza_Click:

    Select Case gErr

        Case 89952
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144009)

    End Select

    Exit Sub

End Sub

Private Sub BotaoConfigura_Click()

Dim lErro As Long
Dim objBrowseUsuarioCampo As AdmBrowseUsuarioCampo

On Error GoTo Erro_BotaoConfigura_Click

    objBrowse1.iPesquisa = ADM_CONFIGURA_NORMAL

    lErro = objBrowse1.Browse_BotaoConfigura_Click(objBrowse1)
    If lErro <> SUCESSO Then gError 89953

    Exit Sub
    
Erro_BotaoConfigura_Click:

    Select Case gErr

        Case 89953
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144010)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoEdita_Click()

Dim lErro As Long
Dim objObjeto As Object

On Error GoTo Erro_BotaoEdita_Click

    lErro = objBrowse1.Browse_BotaoEdita_Click(objBrowse1)
    If lErro <> SUCESSO And lErro <> 55693 Then gError 89954

    If lErro = SUCESSO Then
        
        Set objObjeto = CreateObject(objBrowse1.objBrowseArquivo.sProjetoObjeto & "." & objBrowse1.objBrowseArquivo.sClasseObjeto)

        lErro = Move_BrowseValorCampo(objObjeto)
        If lErro <> SUCESSO Then gError 89975

    End If

    If Len(objBrowse1.objBrowseArquivo.sRotinaBotaoEdita) > 0 Then

        lErro = CallByName(CreateObject(objBrowse1.objBrowseArquivo.sProjeto & "." & objBrowse1.objBrowseArquivo.sClasseBrowser), objBrowse1.objBrowseArquivo.sRotinaBotaoEdita, VbMethod, objObjeto, lErro)
        If lErro <> SUCESSO Then gError 89955
        
    ElseIf Len(gsSiglaModuloChamador) > 0 Then
        
        lErro = Chama_Tela(objBrowse1.objBrowseArquivo.sNomeTelaEdita, objObjeto, gsSiglaModuloChamador)
        If lErro <> SUCESSO Then gError 89955
        
    Else
    
        lErro = Chama_Tela(objBrowse1.objBrowseArquivo.sNomeTelaEdita, objObjeto)
        If lErro <> SUCESSO Then gError 89955
    
    End If
    
    If gobjFAT.iBrowseFecha = MARCADO Then
        Call BotaoFecha_Click
    End If
    
    Exit Sub

Erro_BotaoEdita_Click:

    Select Case gErr

        Case 89954, 89955, 89975

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144011)

    End Select

    Exit Sub

End Sub

Private Sub BotaoConsulta_Click()

Dim lErro As Long
Dim objObjeto As Object

On Error GoTo Erro_BotaoConsulta_Click

    lErro = objBrowse1.Browse_BotaoEdita_Click(objBrowse1)
    If lErro <> SUCESSO And lErro <> 55693 Then gError 89967

    If lErro = SUCESSO Then
        
        Set objObjeto = CreateObject(objBrowse1.objBrowseArquivo.sProjetoObjeto & "." & objBrowse1.objBrowseArquivo.sClasseObjeto)

        lErro = Move_BrowseValorCampo(objObjeto)
        If lErro <> SUCESSO Then gError 89976

    End If

    If Len(objBrowse1.objBrowseArquivo.sRotinaBotaoConsulta) > 0 Then

        lErro = CallByName(CreateObject(objBrowse1.objBrowseArquivo.sProjeto & "." & objBrowse1.objBrowseArquivo.sClasseBrowser), objBrowse1.objBrowseArquivo.sRotinaBotaoConsulta, VbMethod, objObjeto, lErro)
        If lErro <> SUCESSO Then gError 89968
        
    Else
        
        Call Chama_Tela(objBrowse1.objBrowseArquivo.sNomeTelaConsulta, objObjeto)
    
    End If
        
    Exit Sub

Erro_BotaoConsulta_Click:

    Select Case gErr

        Case 89967, 89968, 89976

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144012)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objBrowseUsuarioOrdenacao As New AdmBrowseUsuarioOrdenacao
Dim sOpcao As String
Dim sNomeTela As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click
 
    If Len(Trim(ComboOpcoes.Text)) > 0 Then
 
        sOpcao = ComboOpcoes.Text
        sNomeTela = objBrowse1.objForm.Name
 
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_OPCAO_BROWSER_CONFIRMA_EXCLUSAO", ComboOpcoes.Text)
        
        If vbMsgRes = vbYes Then
 
           lErro = CF("BrowseOpcaoOrdenacao_Le", sOpcao, sNomeTela, objBrowseUsuarioOrdenacao)
           If lErro <> SUCESSO And lErro <> 178341 Then gError 178342
    
           If lErro <> SUCESSO Then gError 178343
    
    
           lErro = CF("BrowseOpcao_Exclui", sOpcao, sNomeTela)
           If lErro <> SUCESSO Then gError 178324
    
           lErro = Carrega_Combobox_Opcoes()
           If lErro <> SUCESSO Then gError 178325
    
           Call Rotina_Aviso(vbOKOnly, "AVISO_OPCAO_BROWSER_EXCLUIDA_SUCESSO", ComboOpcoes.Text)
    
        End If
    
    End If
    
    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 178324, 178325, 178342

        Case 178343
            Call Rotina_Erro(vbOKOnly, "ERRO_OPCAO_BROWSER_NAO_CADASTRADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178326)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim objBrowseUsuarioCampo As AdmBrowseUsuarioCampo
Dim objBrowseUsuarioOrdenacao As New AdmBrowseUsuarioOrdenacao
Dim objBrowseUsuario As New AdmBrowseUsuario
Dim iAchou As Integer
Dim iIndice As Integer

On Error GoTo Erro_BotaoGravar_Click
 
    If Len(Trim(ComboOpcoes.Text)) = 0 Then gError 178364
 
        For Each objBrowseUsuarioCampo In objBrowse1.colBrowseUsuarioCampo
    
            If objBrowse1.objGrid.ColWidth(objBrowseUsuarioCampo.iPosicaoTela - 1) <> objBrowseUsuarioCampo.lLargura Then
                objBrowseUsuarioCampo.lLargura = objBrowse1.objGrid.ColWidth(objBrowseUsuarioCampo.iPosicaoTela - 1)
                objBrowse1.iAlterado = 1
            End If
        Next
     
        objBrowseUsuarioOrdenacao.sNomeTela = objBrowse1.objForm.Name
        objBrowseUsuarioOrdenacao.sCodUsuario = objBrowse1.sCodUsuario
        objBrowseUsuarioOrdenacao.iIndice = objBrowse1.objOrdenacao.ItemData(objBrowse1.objOrdenacao.ListIndex)
        objBrowseUsuarioOrdenacao.sNomeIndice = objBrowse1.objOrdenacao.Text
        objBrowseUsuarioOrdenacao.sSelecaoSQL1Usuario = objBrowse1.sSelecaoSQL1Usuario
        objBrowseUsuarioOrdenacao.sSelecaoSQL1 = objBrowse1.sSelecaoSQL1
            
        If Not objBrowse1.objPrincMDIChild Is Nothing Then
        
            'se a janela está com tamanho normal, isto é, nao se está maximizada ou minimizada
            If objBrowse1.objPrincMDIChild.WindowState = 0 Then
            
                objBrowseUsuario.sNomeTela = objBrowse1.objForm.Name
                objBrowseUsuario.sCodUsuario = objBrowse1.sCodUsuario
                
                objBrowseUsuario.lEsquerda = objBrowse1.objPrincMDIChild.left
                objBrowseUsuario.lTopo = objBrowse1.objPrincMDIChild.top
                objBrowseUsuario.lLargura = objBrowse1.objPrincMDIChild.Width
                objBrowseUsuario.lAltura = objBrowse1.objPrincMDIChild.Height
                
            End If
        
        End If
            
        lErro = CF("BrowseOpcao_Grava", ComboOpcoes.Text, objBrowse1, objBrowseUsuarioOrdenacao, objBrowseUsuario)
        If lErro <> SUCESSO Then gError 178310
        
        For iIndice = 0 To ComboOpcoes.ListCount - 1
            If ComboOpcoes.List(iIndice) = ComboOpcoes.Text Then
                iAchou = 1
                Exit For
            End If
        Next
        
        If iAchou = 0 Then ComboOpcoes.AddItem ComboOpcoes.Text
        
        Call Rotina_Aviso(vbOKOnly, "AVISO_OPCAO_BROWSER_GRAVADA_SUCESSO", ComboOpcoes.Text)
    
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 178310

        Case 178364
            Call Rotina_Erro(vbOKOnly, "ERRO_OPCAO_BROWSER_NAO_PREENCHIDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178311)

    End Select

    Exit Sub

End Sub

Private Sub BotaoPesquisa_Click()

Dim lErro As Long
Dim objBrowseUsuarioCampo As AdmBrowseUsuarioCampo

On Error GoTo Erro_BotaoPesquisa_Click

    objBrowse1.iPesquisa = ADM_CONFIGURA_PESQUISA

    lErro = objBrowse1.Browse_BotaoConfigura_Click(objBrowse1)
    If lErro <> SUCESSO Then gError 89956

    Exit Sub
    
Erro_BotaoPesquisa_Click:

    Select Case gErr

        Case 89956
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144013)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFecha_Click()
    Unload Me
End Sub

Private Sub BotaoPlanilha_Click()
'Exporta os dados do browser para uma planilha em Excel

Dim lErro As Long
Dim lNumRegistros As Long
Dim vbResp As VbMsgBoxResult
Dim iModoImpressao As Integer
Dim objBrowseExcelAux As AdmBrowseExcelAux

On Error GoTo Erro_BotaoPlanilha_Click

    'Transforma o ponteiro do mouse em ampulheta
    MousePointer = vbHourglass
    
    'Instancia a coleção que retornará os registros a serem exibidos no excel
    Set objBrowse1.colRegistros = New Collection
    
    'Descobre o número de registros que serão exportados para o Excel
    lErro = CF("Browse_Le_NumRegistros2", objBrowse1, lNumRegistros)
    If lErro <> SUCESSO And lErro <> 102102 Then gError 102103
    
    'Se nã encontrou registros
    If lErro = 102102 Then
    
        'Avisa ao usuário que nenhum registro foi encontrado
        Call Rotina_Aviso(vbOKOnly, "AVISO_REGISTRO_NAO_ENCONTRADO")
        
        'Exibe o ponteiro padrão
        MousePointer = vbDefault
        
        'Sai da função, pois não faz sentido seguir com a exportação se não existem registros a serem exportados
        Exit Sub
    
    End If
    
    'Se o número de registros é superior ao limite de registros que se permite exportar => erro
    'If lNumRegistros > BROWSE_MAX_REGISTROS_EXPORTAR Then gError 102104
    
    'Se o número de registros é superior ao número sugerido de registros a serem exportados
    If lNumRegistros > BROWSE_SUGESTAO_REGISTROS_EXPORTAR Then
    
        'Pergunta ao usuário se deseja realmente exportar os registros
        vbResp = Rotina_Aviso(vbYesNo, "AVISO_REGISTROS_EM_EXCESSO", lNumRegistros, BROWSE_SUGESTAO_REGISTROS_EXPORTAR)
        
        'Transforma o ponteiro do mouse em ampulheta
        MousePointer = vbHourglass
        
        'Se o usuário respondeu não =>
        If vbResp = vbNo Then
        
            'Exibe o ponteiro padrão
            MousePointer = vbDefault
            
            'sai da função
            Exit Sub
        
        End If
    
    End If
        
'    'Lê no BD os registros que serão exportados para o Excel
'    lErro = CF("Browse_Executa_SQL", objBrowse1)
'    If lErro <> SUCESSO And lErro <> 102092 Then gError 102095
'
'
'    'Move os dados lidos do BD para o objPlanilha que será utilizado para exportar os dados para o Excel
'    lErro = Browse_Move_Dados_Memoria_Formato_Excel(objBrowse1, objPlanilha)
'    If lErro <> SUCESSO Then gError 102096
        
    'MsgBox("Deseja exibir o gráfico na tela?", vbYesNoCancel, objPlanilha.sNomePlanilha)
    'Pergunta se o usuário deseja imprimir a planilha
    vbResp = Rotina_Aviso(vbYesNo, "AVISO_IMPRIMIR_PLANILHA")
    If vbResp = vbYes Then iModoImpressao = EXCEL_MODO_IMPRESSAO
        
    'Lê no BD os registros que serão exportados para o Excel
    lErro = CF("Browse_Executa_SQL_Excel", objBrowse1, lNumRegistros, iModoImpressao)
    If lErro <> SUCESSO And lErro <> 102092 Then gError 102095
    
    'Transforma o ponteiro do mouse em ampulheta
    MousePointer = vbHourglass
        
'    'Exporta os dados do objPlanilha para o Excel
'    lErro = CF("Excel_Gera_Planilha", objPlanilha)
'    If lErro <> SUCESSO Then gError 102097
     
    'Exibe o ponteiro padrão
    MousePointer = vbDefault
    
    Exit Sub
    
Erro_BotaoPlanilha_Click:

    Select Case gErr

        Case 102104
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCEL_REGISTROS_EM_EXCESSO", gErr, BROWSE_MAX_REGISTROS_EXPORTAR)
        
        Case 102095 To 102097, 102103
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144014)

    End Select

    'Exibe o ponteiro padrão
    MousePointer = vbDefault
    
    Exit Sub

End Sub

Private Function Browse_Move_Dados_Memoria_Formato_Excel(ByVal objBrowse As AdmBrowse, ByVal objPlanilha) As Long
'Transfere os dados de objBrowse para objPlanilha no formato esperado pelo Excel para geração da planilha
'objBrowse RECEBE(Input) os dados que serão transferidos
'objPlanilha RETORNA(Output) os dados no formato esperado pelo Excel

Dim iColuna As Integer
Dim objColunas As ClassColunasExcel
Dim objCelulas As ClassCelulasExcel
Dim dLarguraColuna As Double
Dim lErro As Long
Dim iIndiceCol As Integer
Dim iIndiceLin As Integer
'##################################
'Inserido por Wagner 26/01/2006
Dim iIndiceColdeFato As Integer
Dim iIndice As Integer
Dim objBrowseUsuarioCampo As AdmBrowseUsuarioCampo
Dim vValorCampo As Variant
Dim iExercicio As Integer
Dim iIndiceColAux As Integer
'##################################

On Error GoTo Erro_Browse_Move_Dados_Memoria_Formato_Excel

    'Guarda o nome que será exibido para a planilha
    objPlanilha.sNomePlanilha = Me.Caption
    
    'Para cada coluna que será exibida no Excel
    For iColuna = 1 To objBrowse.objGrid.Cols
    
        'Instancia um novo obj para armazenar os dados da coluna
        Set objColunas = New ClassColunasExcel
        
        'Obtém a largura da coluna no padrão do Excel
        lErro = CF("Excel_Obtem_Largura_Coluna", objBrowse.objGrid.ColWidth(iColuna - 1), dLarguraColuna)
        If lErro <> SUCESSO Then gError 102098
        
        'Guarda no objeto a largura dessa coluna
        objColunas.dLarguraColuna = dLarguraColuna
        
        'Adiciona a coluna à coleção de colunas
        objPlanilha.colColunas.Add objColunas
    
        'Instancia um novo objeto para armazenar os dados da primeira célula da coluna em questão
        Set objCelulas = New ClassCelulasExcel

        'Guarda o título da coluna em questão que será exibido na primeira célula
        objCelulas.vValor = objBrowse.objGrid.TextMatrix(0, iColuna - 1)
        objCelulas.bFonteNegrito = True
         
        'Adiciona a célula à coleção de células
        objColunas.colCelulas.Add objCelulas

    Next
    
    ' *** Transfere os dados retornados pela função de leitura acima para o formato que
    ' será passado para o Excel ***
    
    '***** Trecho adicionado por Rafael Menezes em 03/09/2002
    For iIndiceCol = 1 To objBrowse.colBrowseUsuarioCampo.Count
               
        '################################################################
        'Inserido por Wagner 26/01/2006
        'Após uma modificação na posição dos campos a coleção colregistros
        'não tem sua ordem alterada, por isso é necessário confrontar a ordem inicial
        'com a ordem atual para fazer a correlação correta entre os campos
        iIndice = 0
        For Each objBrowseUsuarioCampo In objBrowse.colBrowseUsuarioCampo
            iIndice = iIndice + 1
            If objBrowseUsuarioCampo.iPosicaoTela = iIndiceCol Then
                iIndiceColdeFato = iIndice
                Exit For
            End If
        Next
        '################################################################
               
        For iIndiceLin = 1 To objBrowse.colRegistros.Count
        
            'aponta para a coluna em questão
            Set objColunas = objPlanilha.colColunas(iIndiceCol)
            
            'Instancia um novo objeto para armazenar os dados da célula
            Set objCelulas = New ClassCelulasExcel
            
            'se não for data_Nula
            'If objBrowse.colRegistros.Item(iIndiceLin)(iIndiceCol) <> DATA_NULA Then
            If objBrowse.colRegistros.Item(iIndiceLin)(iIndiceColdeFato) <> DATA_NULA Then
                
                'guarda o valor do registro dakela linha, dakela coluna
                'objCelulas.vValor = objBrowse.colRegistros.Item(iIndiceLin)(iIndiceCol)
                objCelulas.vValor = objBrowse.colRegistros.Item(iIndiceLin)(iIndiceColdeFato)
                               
                'se for uma data, formata para "dd/mm/yyyy", caso contrário, o Excel inverte o mês com o dia e o dado perde a semântica
                'passando a ser tratado como um texto qualquer
                Select Case TypeName(objCelulas.vValor)
                    Case "Date"
                        objCelulas.vValor = StrParaDate(objCelulas.vValor)
                    
                    Case "Double"
                        objCelulas.vValor = StrParaDbl(objCelulas.vValor)
                        
                        Select Case objBrowse.colValorCampo.Item(iIndiceColdeFato).iSubTipo
                            
                            Case ADM_SUBTIPO_PERCENTUAL
                                objCelulas.sNumberFormat = "0.00%"
                                
                            Case ADM_SUBTIPO_HORA
                                objCelulas.vValor = CDate(objCelulas.vValor)
                                objCelulas.sNumberFormat = "hh:mm:ss"
                            
                            Case Else
                                objCelulas.sNumberFormat = "#,##0.00##"
                            
                        End Select
                    
                    Case Else
                                        
                        '##################################################
                        'Inserido por Wagner 17/07/2006

                        If left(objBrowseUsuarioCampo.sNome, 7) = "Periodo" And left(objBrowseUsuarioCampo.sNome, 9) <> "PeriodoDe" And left(objBrowseUsuarioCampo.sNome, 10) <> "PeriodoAte" Then
                            For iIndiceColAux = 1 To objBrowse.colBrowseUsuarioCampo.Count
                                If UCase(objBrowse.colBrowseUsuarioCampo.Item(iIndiceColAux).sNome) = "EXERCICIO" Or UCase(objBrowse.colBrowseUsuarioCampo.Item(iIndiceColAux).sNome) = "EXERCÍCIO" Then
                                    iExercicio = objBrowse.colRegistros.Item(iIndiceLin)(iIndiceColAux)
                                    Exit For
                                End If
                            Next
                        End If

                        lErro = CF("Exclusao_Valida_Mascara", objBrowse.colValorCampo.Item(iIndiceColdeFato).iTipo, objBrowse.colValorCampo.Item(iIndiceColdeFato).iSubTipo, objCelulas.vValor, vValorCampo, True, objBrowse.lComando2, objBrowse.lComando3, iExercicio)
                        If lErro <> SUCESSO Then gError 181223

                        objCelulas.vValor = vValorCampo
                        '##################################################
                
                End Select
            
            End If
            
            'adiciona a célula à coleção de células da coluna em questão
            objColunas.colCelulas.Add objCelulas
        
        Next
        
    Next
    '***** fim do Trecho adicionado por Rafael Menezes em 03/09/2002

    Browse_Move_Dados_Memoria_Formato_Excel = SUCESSO
    
    Exit Function
    
Erro_Browse_Move_Dados_Memoria_Formato_Excel:
    
    Browse_Move_Dados_Memoria_Formato_Excel = gErr
    
    Select Case gErr

        Case 102098, 181223
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144015)

    End Select

    Exit Function
    
End Function

Private Sub BotaoSeleciona_Click()

Dim lErro As Long
Dim objObjeto As Object

On Error GoTo Erro_BotaoSeleciona_Click

    If Not (objEvento Is Nothing) Then

        lErro = objBrowse1.Browse_BotaoEdita_Click(objBrowse1)
        If lErro <> SUCESSO Then gError 89957

        Set objObjeto = CreateObject(objBrowse1.objBrowseArquivo.sProjetoObjeto & "." & objBrowse1.objBrowseArquivo.sClasseObjeto)

        lErro = Move_BrowseValorCampo(objObjeto)
        If lErro <> SUCESSO Then gError 89958

        Call objEvento.ChamaEventoSelecao(objObjeto)

        If giLocalOperacao <> LOCALOPERACAO_ECF Then
        
            If gobjFAT.iBrowseFecha = MARCADO Then
                Call BotaoFecha_Click
            End If
        
        Else
            Call BotaoFecha_Click
        
        End If
        
    End If
    
    Exit Sub

Erro_BotaoSeleciona_Click:

    Select Case gErr

        Case 89957, 89958, 999999
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144016)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim objBrowseUsuarioCampo As AdmBrowseUsuarioCampo

On Error GoTo Erro_Form_Load

    bCargaInicialECF = False
    
    Set objBrowse1 = New AdmBrowse

    objBrowse1.iAlterado = -1
    Set objBrowse1.objForm = Me
    GridBrowse.Rows = 100
    Set objBrowse1.objGrid = GridBrowse
    Set objBrowse1.objOrdenacao = Ordenacao
    Set objBrowse1.objBarraScroll = BarraScroll
    objBrowse1.iGrid_LinhaInicial = 12
    objBrowse1.iGrid_LinhaFinal = 21
    objBrowse1.iGrid_LinhasExibidas = 10
    Set objBrowse1.objPrincMDIChild = Parent
    gsOrdenacao = ""
    gsSiglaModuloChamador = ""
    Set objBrowse1.objComboOpcoes = ComboOpcoes
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144017)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
 
    lErro = objBrowse1.Browse_Form_Unload(objBrowse1)
 
    Set objEvento = Nothing
    Set objBrowse1 = Nothing

End Sub

Private Sub ComboOpcoes_Click()

Dim lErro As Long
Dim objBrowseOpcao As New AdmBrowseOpcao
Dim objBrowseOpcaoCampo As AdmBrowseOpcaoCampo
Dim objBrowseUsuarioOrdenacao As New AdmBrowseUsuarioOrdenacao
Dim colBrowseUsuarioCampo As New Collection
Dim iAchou As Integer
Dim iIndice As Integer
Dim sOpcao As String
Dim sNomeTela As String

On Error GoTo Erro_ComboOpcoes_Click

    If ComboOpcoes.ListIndex = -1 Then Exit Sub

    sNomeTela = objBrowse1.objForm.Name
    sOpcao = ComboOpcoes.Text

    lErro = CF("BrowseOpcaoOrdenacao_Le", sOpcao, sNomeTela, objBrowseUsuarioOrdenacao)
    If lErro <> SUCESSO And lErro <> 178341 Then gError 178356

    If lErro <> SUCESSO Then gError 178363

    Set objBrowse1.objBrowseUsuarioOrdenacao = objBrowseUsuarioOrdenacao

    lErro = CF("BrowseOpcaoCampo_Le", sOpcao, sNomeTela, colBrowseUsuarioCampo)
    If lErro <> SUCESSO Then gError 178355
    
    Set objBrowse1.objBrowseExcel = New AdmBrowseExcel
    
    lErro = CF("BrowseExcel_Le", objBrowse1, sOpcao, objBrowse1.objBrowseExcel)
    If lErro <> SUCESSO Then gError 178355

    Set objBrowse1.colBrowseUsuarioCampo = colBrowseUsuarioCampo

    iAchou = 0
    
    'preenche gsOrdenacao para indicar que Ordenacao_Click nao deve fazer nada
    gsOrdenacao = "x"

    For iIndice = 0 To Ordenacao.ListCount - 1
        If Ordenacao.List(iIndice) = objBrowseUsuarioOrdenacao.sNomeIndice Then
            Ordenacao.ListIndex = iIndice
            iAchou = 1
            Exit For
        End If
    Next

    If iAchou = 0 Then Ordenacao.ListIndex = 0
        
    'retorna a expressao de selecao criada pelo usuario, possivelmente alterada
    objBrowse1.sSelecaoSQL1Usuario = objBrowseUsuarioOrdenacao.sSelecaoSQL1Usuario
    objBrowse1.sSelecaoSQL1 = objBrowseUsuarioOrdenacao.sSelecaoSQL1

    lErro = objBrowse1.Browse_Inicializa_Campos(objBrowse1)
    If lErro <> SUCESSO Then gError 178357

    lErro = objBrowse1.Browse_Inicializa_Comando_SQL(objBrowse1)
    If lErro <> SUCESSO Then gError 178358
    
    'inicializa o comando SQL que será usado para contar o número de registros da tabela
    lErro = objBrowse1.Browse_Inicializa_Comando_SQL_Count(objBrowse1)
    If lErro <> SUCESSO Then gError 178359
    
    objBrowse1.iBind = 0
    
    lErro = objBrowse1.Browse_Inicializa_Grid_Lote(objBrowse1)
    If lErro <> SUCESSO Then gError 178360
    
    lErro = objBrowse1.Browse_Inicializa_ScrollBar_Lote(objBrowse1)
    If lErro <> SUCESSO Then gError 178361

    objBrowse1.objGrid.Row = 1
    
    objBrowse1.objBarraScroll.Max = objBrowse1.objBarraScroll.Min
    
    Call objBrowse1.Browse_BarraScroll_Change(objBrowse1)

    Exit Sub

Erro_ComboOpcoes_Click:

    Select Case gErr

        Case 178355 To 178361

        Case 178363
            Call Rotina_Erro(vbOKOnly, "ERRO_OPCAO_BROWSER_NAO_CADASTRADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178382)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

Dim iIndice As Integer

    If Len(Trim(ComboOpcoes.Text)) > 0 Then
        
        For iIndice = 0 To ComboOpcoes.ListCount - 1
            If ComboOpcoes.Text = ComboOpcoes.List(iIndice) Then
                If ComboOpcoes.ListIndex <> iIndice Then
                    ComboOpcoes.ListIndex = iIndice
                    Exit For
                End If
            End If
        Next
    
    End If

End Sub

Private Sub GridBrowse_Click()

Dim objBrowseUsuarioCampo As AdmBrowseUsuarioCampo
Dim iIndice As Integer
Dim iAchou As Integer
Dim objBrowseIndice As New AdmBrowseIndice
Dim iIndice1 As Integer
Dim iIndice2 As Integer

On Error GoTo Erro_GridBrowse_Click

    If GridBrowse.MouseRow = 0 Then

        For Each objBrowseUsuarioCampo In objBrowse1.colBrowseUsuarioCampo
            If objBrowseUsuarioCampo.iPosicaoTela = GridBrowse.MouseCol + 1 Then
                
                iIndice2 = 0
                
                For iIndice = 0 To Ordenacao.ListCount - 1
                    If Ordenacao.ItemData(iIndice) > iIndice2 Then iIndice2 = Ordenacao.ItemData(iIndice)
                Next
                                        
                For iIndice = 0 To Ordenacao.ListCount - 1
                                        
                    If Ordenacao.List(iIndice) = objBrowseUsuarioCampo.sNome Then
                        Ordenacao.ListIndex = iIndice
                        iAchou = 1
                        Exit For
                    End If
                Next
                
                If iAchou <> 1 Then
                
                    iIndice1 = 0
                
                    For Each objBrowseIndice In objBrowse1.colBrowseIndiceUsuario
                        If objBrowseIndice.iIndice > iIndice1 Then iIndice1 = objBrowseIndice.iIndice
                    Next
                
                
                    objBrowseIndice.iIndice = iIndice1 + 1
                    objBrowseIndice.sNomeTela = objBrowse1.objForm.Name
                    objBrowseIndice.sNomeIndice = objBrowseUsuarioCampo.sNome
                    objBrowseIndice.sOrdenacaoSQL = objBrowseUsuarioCampo.sNome
                    objBrowseIndice.sSelecaoSQL = "(" & objBrowseUsuarioCampo.sNome & "<? OR " & objBrowseUsuarioCampo.sNome & " Is NULL)"

                    objBrowse1.iAlterado = -1

                    objBrowse1.colBrowseIndiceUsuario.Add objBrowseIndice
                    Ordenacao.AddItem objBrowseIndice.sNomeIndice
                    Ordenacao.ItemData(Ordenacao.NewIndex) = iIndice2 + 1
                    
                    objBrowse1.iAlteradoOrdenacao = 1
                    objBrowse1.iAlterado = 1
                
                    Ordenacao.ListIndex = Ordenacao.NewIndex
                
                End If
                
                Exit For
                
            End If
            
        Next
        
    End If
    
    Exit Sub

Erro_GridBrowse_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178435)

    End Select

    Exit Sub
    
End Sub

Private Sub GridBrowse_DblClick()
    Call BotaoSeleciona_Click
End Sub

Private Sub GridBrowse_RowColChange()

    Call objBrowse1.Browse_GridBrowse_RowColChange(objBrowse1)

End Sub

Private Sub GridBrowse_Scroll()

    Call objBrowse1.Browse_GridBrowse_Scroll(objBrowse1)
    
End Sub

Private Sub Ordenacao_Click()

Dim lErro As Long
Dim iPos As Integer
Dim sMascara As String

On Error GoTo Erro_Ordenacao_Click

    If Len(gsOrdenacao) = 0 Then

        giPesquisaChange = 0
        
        lErro = objBrowse1.Browse_Ordenacao_Click(objBrowse1)
        If lErro <> SUCESSO Then gError 89959
            
        Pesquisa.Mask = ""
        Pesquisa.PromptInclude = False
        Pesquisa.Text = ""
        Pesquisa.PromptInclude = True

        iPos = InStr(Ordenacao.Text, ",")

        If giLocalOperacao <> LOCALOPERACAO_ECF Then

            If iPos = 0 Then
                If Trim(Ordenacao.Text) = "Produto" Or ((Trim(Ordenacao.Text) = "Codigo" Or Trim(Ordenacao.Text) = "Código") And UCase(left(objBrowse1.objBrowseArquivo.sNomeTela, 7)) = "PRODUTO") Then
                    'Inicializa Máscara de Produto
                    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Pesquisa)
                    If lErro <> SUCESSO Then Error 23809
                End If
            Else
                If Trim(Mid(Ordenacao.Text, 1, iPos - 1)) = "Produto" Or ((Trim(Ordenacao.Text) = "Codigo" Or Trim(Ordenacao.Text) = "Código") And UCase(left(objBrowse1.objBrowseArquivo.sNomeTela, 7)) = "PRODUTO") Then
                    'Inicializa Máscara de Produto
                    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Pesquisa)
                    If lErro <> SUCESSO Then Error 23809
                End If
            End If

        End If

        giPesquisaChange = 1
        Call Pesquisa_Change

    Else
    
        gsOrdenacao = ""
        
    End If

    Exit Sub

Erro_Ordenacao_Click:

    Select Case gErr

        Case 89959

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144018)

    End Select

    Exit Sub

End Sub

Private Sub BarraScroll_Change()

Dim lErro As Long

On Error GoTo Erro_BarraScroll_Change

    lErro = objBrowse1.Browse_BarraScroll_Change(objBrowse1)
    If lErro <> SUCESSO Then gError 89960

    Exit Sub

Erro_BarraScroll_Change:

    Select Case gErr

        Case 89960

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144019)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(sNomeTela As String, ParamArray avParams()) As Long
'Optional colSelecao As Collection, Optional objMovEstoque As ClassMovEstoque, Optional objEvento1 As AdmEvento) As Long

Dim lErro As Long
Dim objObjeto As Object
Dim objTela As Object
Dim iIndice As Integer
Dim objLote As ClassLote
Dim iPos As Integer, vSel As Variant

On Error GoTo Erro_Trata_Parametros

    gPrecisaRemoverFiltro = False
    gbCarregandoTela = True
    
    'trata a opcao de browser
    If UBound(avParams) >= 5 Then
        
        If Not IsMissing(avParams(5)) Then
        
            objBrowse1.sOpcao = avParams(5)
        
        End If
        
    End If
    
    If Len(objBrowse1.objForm.Name) = 0 Then
    
        'guarda o Nome da Tela
        sName = sNomeTela
        
        lErro = objBrowse1.Browse_Inicializa(objBrowse1)
        If lErro <> SUCESSO Then gError 89961
        
        objBrowse1.iAlterado = 0
        
        Caption = objBrowse1.objBrowseArquivo.sTituloBrowser
        
        iNumeroBotoes = BROWSE_NUMERO_BOTOES + objBrowse1.objBrowseArquivo.iBotaoSeleciona + objBrowse1.objBrowseArquivo.iBotaoEdita + objBrowse1.objBrowseArquivo.iBotaoConsulta
        If gsNomePrinc = "SGEECF" Then iNumeroBotoes = iNumeroBotoes - 1 'sumir com o botao de planilha
        
    End If
    
    If UBound(avParams) < 0 Or IsMissing(avParams(0)) Then
    
        Set objBrowse1.colSelecao = New Collection
        
        'carrega os valores dos parametros de selecao que ainda faltam
        lErro = CF("BrowseParamSelecao_Le", sNomeTela, objBrowse1.colSelecao)
        If lErro <> SUCESSO Then gError 92000
    
    End If
    
    If UBound(avParams) >= 0 Then
    
        If Not IsMissing(avParams(0)) Then
    
            Set objBrowse1.colSelecao = avParams(0)
            
            If objBrowse1.colSelecao Is Nothing Then Set objBrowse1.colSelecao = New Collection
            
            lErro = CF("Browser_Trata_Parametros_Customizado", sNomeTela, objBrowse1)
            If lErro <> SUCESSO Then gError 126713

        End If

        
    End If
    
    
    'vSelecaoSQL é uma possivel expressão de selecao SQL passada pelo programador
    If UBound(avParams) >= 3 Then
        If Not IsMissing(avParams(3)) Then
            objBrowse1.sSelecaoSQL2 = avParams(3)
        End If
    End If
        
    If Len(objBrowse1.objBrowseArquivo.sTrataParametros) > 0 Then
    
        Set objTela = Me
    
        lErro = CallByName(CreateObject(objBrowse1.objBrowseArquivo.sProjeto & "." & objBrowse1.objBrowseArquivo.sClasseBrowser), objBrowse1.objBrowseArquivo.sTrataParametros, VbMethod, objTela, objBrowse1.colSelecao)
        If lErro <> SUCESSO Then gError 89962
    
    End If
        
    If UBound(avParams) >= 1 Then
        If Not IsMissing(avParams(1)) Then
            Set objObjeto = avParams(1)
            If sName = "LotePendenteLista" Then
                Set objLote = objObjeto
                gsSiglaModuloChamador = StringZ(objLote.sOrigem)
            End If
            
        End If
    
    End If
        
    If UBound(avParams) >= 2 Then
        If Not IsMissing(avParams(2)) Then
            Set objEvento = avParams(2)
        End If
    End If
        
    If UBound(avParams) >= 4 Then
    
        If Not IsMissing(avParams(4)) Then
    
            gsOrdenacao = avParams(4)
            
            If Len(gsOrdenacao) > 0 Then
        
                For iIndice = 0 To Ordenacao.ListCount - 1
            
                    If Ordenacao.List(iIndice) = gsOrdenacao Then
                        
                        lErro = objBrowse1.Browse_Inicializa_Campos(objBrowse1)
                        If lErro <> SUCESSO Then gError 117547
                        
                        Ordenacao.ListIndex = iIndice
                
                        Pesquisa.Mask = ""
                        Pesquisa.PromptInclude = False
                        Pesquisa.Text = ""
                        Pesquisa.PromptInclude = True
                
                        iPos = InStr(Ordenacao.Text, ",")
                
                        If iPos = 0 Then
                            If Trim(Ordenacao.Text) = "Produto" Or ((Trim(Ordenacao.Text) = "Codigo" Or Trim(Ordenacao.Text) = "Código") And UCase(left(objBrowse1.objBrowseArquivo.sNomeTela, 7)) = "PRODUTO") Then
                                'Inicializa Máscara de Produto
                                lErro = CF("Inicializa_Mascara_Produto_MaskEd", Pesquisa)
                                If lErro <> SUCESSO Then Error 23809
                            End If
                        Else
                            If Trim(Mid(Ordenacao.Text, 1, iPos - 1)) = "Produto" Or ((Trim(Ordenacao.Text) = "Codigo" Or Trim(Ordenacao.Text) = "Código") And UCase(left(objBrowse1.objBrowseArquivo.sNomeTela, 7)) = "PRODUTO") Then
                                'Inicializa Máscara de Produto
                                lErro = CF("Inicializa_Mascara_Produto_MaskEd", Pesquisa)
                                If lErro <> SUCESSO Then Error 23809
                            End If
                        End If
                
                        giPesquisaChange = 1
                        
                        Exit For
                    
                    End If
            
                Next
            
            End If
        
        End If
        
    End If
        
        
    'Se não está passando parametros
    If objObjeto Is Nothing Then
    
        'novo mario
        objBrowse1.iSelecaoExterna = 1
    
        lErro = objBrowse1.Browse_Trata_Parametros1(objBrowse1)
        If lErro <> SUCESSO Then gError 89963
    
        Set objObjeto = CreateObject(objBrowse1.objBrowseArquivo.sProjetoObjeto & "." & objBrowse1.objBrowseArquivo.sClasseObjeto)
    
        lErro = Inicializa_colBrowseValorCampo(objObjeto)
        If lErro <> SUCESSO Then gError 89964
    
        'exibe os dados a partir do inicio
        If lErro_Chama_Tela = SUCESSO Then
            BarraScroll.Max = BarraScroll.Min
        Else
            Call BarraScroll_Change
            GridBrowse.Row = 1
        End If

        objBrowse1.iSelecaoExterna = 0
    

    Else
        
        lErro = Inicializa_colBrowseValorCampo(objObjeto)
        If lErro <> SUCESSO Then gError 89965
        
        lErro = objBrowse1.Browse_Trata_Parametros(objBrowse1)
        If lErro <> SUCESSO Then gError 89966
    
    End If
    
    lErro = Carrega_Combobox_Opcoes()
    If lErro <> SUCESSO Then gError 178328
    
    
    If Len(Trim(objBrowse1.sOpcao)) > 0 Then
        For iIndice = 0 To ComboOpcoes.ListCount - 1
            If ComboOpcoes.List(iIndice) = objBrowse1.sOpcao Then
                ComboOpcoes.ListIndex = iIndice
                Exit For
            End If
        Next
    End If
    
    For Each vSel In objBrowse1.colSelecao
        gcolSelecaoOrig.Add vSel
    Next
    gsSelecaoSQLOrig = objBrowse1.sSelecaoSQL
    
    gbCarregandoTela = False
    
    If gsNomePrinc = "SGEECF" And sName = "ProdutosLista" Then
        Pesquisa.PromptInclude = False
        Pesquisa.Text = objObjeto.sNomeReduzido
        OptFiltrar.Value = True
        bCargaInicialECF = True
    End If
            
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    gbCarregandoTela = False

    Trata_Parametros = gErr
    
    Select Case gErr

        Case 89961, 89962, 89963, 89964, 89965, 89966, 92000, 178328
        
        Case Else
        
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144020)

    End Select

    Exit Function
    
End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BROWSE
    Set Form_Load_Ocx = Me
    Call Form_Load
    
End Function

Public Property Get Name() As String
    Name = sName
End Property

Public Property Let Name(ByVal New_Name As String)
    sName = New_Name
End Property

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

Private Sub Pesquisa_Change()
            
            
On Error GoTo Erro_Pesquisa_Change
            
    If giPesquisaChange = 1 And Not gbCarregandoTela Then
            
        Timer1.Tag = 0
        Timer1.Interval = 10
        

    End If

    Exit Sub
    
Erro_Pesquisa_Change:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178421)

    End Select

    Exit Sub

End Sub

'Private Sub Timer1_Timer()
'
'Dim objObjeto As Object
'Dim objGrupoBrowseCampo As AdmGrupoBrowseCampo
'Dim objBrowseValorCampo As AdmBrowseValorCampo
'Dim lErro As Long
'Dim vOrdenacao As Variant
'Dim iPosFim As Integer
'Dim iPosInicio As Integer
'Dim objBrowseIndice1 As AdmBrowseIndice
'Dim lNumReg As Long
'Dim lPosAux As Long
'Dim objCampo As New AdmCampos
'Dim lComando3 As Long
'Dim sCodGrupo As String
'Dim colGrupoBrowseCampo As New Collection
'Dim colOrdenacao As New Collection
'Dim iPrimeiroCampo As Integer
'Dim sOrdenacao As String
'Dim sProduto As String
'Dim iPreenchido As Integer
'Dim iAchou As Integer, vSel As Variant
'Dim iIndice As Integer
'
'On Error GoTo Erro_Timer1_Timer
'
'    Timer1.Tag = StrParaLong(Timer1.Tag) + Timer1.Interval
'
'    If CLng(Timer1.Tag) > 800 Then
'
'        Timer1.Tag = 0
'        Timer1.Interval = 0
'        objBrowse1.sSelecaoSQL = gsSelecaoSQLOrig
'        objBrowse1.sSelecaoSQL2 = ""
'        For iIndice = objBrowse1.colSelecao.Count To 1 Step -1
'            objBrowse1.colSelecao.Remove iIndice
'        Next
'        For Each vSel In gcolSelecaoOrig
'            objBrowse1.colSelecao.Add vSel
'        Next
'
'        lComando3 = Comando_AbrirExt(GL_lConexaoDic)
'        If lComando3 = 0 Then gError 178427
'
'        Set objObjeto = CreateObject(objBrowse1.objBrowseArquivo.sProjetoObjeto & "." & objBrowse1.objBrowseArquivo.sClasseObjeto)
'
'        'Set objBrowse1.colBrowseValorCampo = New Collection
'
'        For Each objBrowseIndice1 In objBrowse1.colBrowseIndice
'            If objBrowseIndice1.iIndice = Ordenacao.ItemData(Ordenacao.ListIndex) Then
'                sOrdenacao = objBrowseIndice1.sOrdenacaoSQL
'                Exit For
'            End If
'        Next
'
'        If Len(sOrdenacao) = 0 Then
'
'            For Each objBrowseIndice1 In objBrowse1.colBrowseIndiceUsuario
'                If objBrowseIndice1.iIndice + objBrowse1.colBrowseIndice.Count = Ordenacao.ItemData(Ordenacao.ListIndex) Then
'                    sOrdenacao = objBrowseIndice1.sOrdenacaoSQL
'                    Exit For
'                End If
'            Next
'
'        End If
'
'        iPosInicio = 1
'        iPosFim = InStr(sOrdenacao, ",")
'        If iPosFim > 0 Then sOrdenacao = Trim(Mid(sOrdenacao, iPosInicio, iPosFim - 1))
'
'        If Len(Trim(sOrdenacao)) > 0 Then
'
'            sCodGrupo = String(STRING_GRUPO, 0)
'
'            'obtem o codigo do grupo
'            lErro = Obter_Grupo(sCodGrupo)
'            If lErro <> SUCESSO Then gError 178428
'
'            'le os campos disponiveis para a tela x grupo em questão
'            lErro = CF("GrupoBrowseCampo_Le", sCodGrupo, Me.Name, colGrupoBrowseCampo)
'            If lErro <> SUCESSO Then gError 178429
'
'            For Each objGrupoBrowseCampo In colGrupoBrowseCampo
'
'                If objGrupoBrowseCampo.sNome = sOrdenacao Then
'
'                    Set objBrowseValorCampo = New AdmBrowseValorCampo
'
'                    objBrowseValorCampo.sNomeCampo = objGrupoBrowseCampo.sNome
'
'                    objCampo.sNomeArq = objBrowse1.sNomeTabela
'                    objCampo.sNome = objBrowseValorCampo.sNomeCampo
'
'                    lErro = CF("Campos_Le2", objCampo, lComando3)
'                    If lErro <> SUCESSO And lErro <> 9184 Then gError 178430
'
'                    'se o campo não estiver cadastrado
'                    If lErro = 9184 Then gError 178431
'
'                    Select Case objCampo.iTipo
'
'                        Case ADM_TIPO_SMALLINT
'                            objBrowseValorCampo.vValorCampo = StrParaIntErr(Pesquisa.Text)
'
'                        Case ADM_TIPO_INTEGER
'                            objBrowseValorCampo.vValorCampo = StrParaLongErr(Pesquisa.Text)
'
'                        Case ADM_TIPO_DOUBLE
'                            objBrowseValorCampo.vValorCampo = StrParaDblErr(Pesquisa.Text)
'
'                        Case ADM_TIPO_VARCHAR
'                                objBrowseValorCampo.vValorCampo = "%" & Trim(Pesquisa.ClipText) & "%"
'
'                        Case ADM_TIPO_DATE
'                            objBrowseValorCampo.vValorCampo = StrParaDateErr(Pesquisa.Text)
'
'                        Case Else
'                            gError 178433
'
'                    End Select
'
'                    'objBrowse1.colBrowseValorCampo.Add objBrowseValorCampo
'
'                    objBrowse1.sSelecaoSQL2 = sOrdenacao & " LIKE ? "
'
'                    objBrowse1.colSelecao.Add objBrowseValorCampo.vValorCampo
'
'                    Exit For
'
'                End If
'
'            Next
'
'        End If
'
'        lErro = objBrowse1.Browse_Trata_Parametros(objBrowse1)
'        If lErro <> SUCESSO Then gError 178435
'
'        lErro = Inicializa_colBrowseValorCampo(objObjeto)
'        If lErro <> SUCESSO Then gError 178436
'
'        Call Comando_Fechar(lComando3)
'
'    End If
'
'    Exit Sub
'
'Erro_Timer1_Timer:
'
'    Select Case gErr
'
'        Case 178422, 178428 To 178430, 178432, 178435, 178436
'
'        Case 178427
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 178431
'            Call Rotina_Erro(vbOKOnly, "ERRO_CAMPO_NAO_CADASTRADO", gErr, objCampo.sNome, objCampo.sNomeArq)
'
'        Case 178433, 178434
'            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_CAMPO_INVALIDO", gErr, objCampo.iTipo)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197898)
'
'    End Select
'
'    Call Comando_Fechar(lComando3)
'
'    Exit Sub
'
'End Sub
'

Private Sub Timer1_Timer()

Dim objObjeto As Object
Dim objGrupoBrowseCampo As AdmGrupoBrowseCampo
Dim objBrowseValorCampo As AdmBrowseValorCampo
Dim lErro As Long
Dim vOrdenacao As Variant
Dim iPosFim As Integer
Dim iPosInicio As Integer
Dim objBrowseIndice1 As AdmBrowseIndice
Dim lNumReg As Long
Dim lPosAux As Long
Dim objCampo As New AdmCampos
Dim lComando3 As Long
Dim sCodGrupo As String
Dim colGrupoBrowseCampo As New Collection
Dim colOrdenacao As New Collection
Dim iPrimeiroCampo As Integer
Dim sOrdenacao As String
Dim sProduto As String
Dim iPreenchido As Integer
Dim iAchou As Integer, vSel As Variant
Dim iIndice As Integer

On Error GoTo Erro_Timer1_Timer

    Timer1.Tag = StrParaLong(Timer1.Tag) + Timer1.Interval

    If CLng(Timer1.Tag) > 800 Then

        Timer1.Tag = 0
        Timer1.Interval = 0
        
        If OptFiltrar.Value Or gPrecisaRemoverFiltro Then
            objBrowse1.sSelecaoSQL = gsSelecaoSQLOrig
            objBrowse1.sSelecaoSQL2 = ""
            For iIndice = objBrowse1.colSelecao.Count To 1 Step -1
                objBrowse1.colSelecao.Remove iIndice
            Next
            For Each vSel In gcolSelecaoOrig
                objBrowse1.colSelecao.Add vSel
            Next
        End If

        lComando3 = Comando_AbrirExt(GL_lConexaoDic)
        If lComando3 = 0 Then gError 178427

        Set objObjeto = CreateObject(objBrowse1.objBrowseArquivo.sProjetoObjeto & "." & objBrowse1.objBrowseArquivo.sClasseObjeto)

        Set objBrowse1.colBrowseValorCampo = New Collection

        For Each objBrowseIndice1 In objBrowse1.colBrowseIndice
            If objBrowseIndice1.iIndice = Ordenacao.ItemData(Ordenacao.ListIndex) Then
                sOrdenacao = objBrowseIndice1.sOrdenacaoSQL
                Exit For
            End If
        Next

        If Len(sOrdenacao) = 0 Then

            For Each objBrowseIndice1 In objBrowse1.colBrowseIndiceUsuario
                If objBrowseIndice1.iIndice + objBrowse1.colBrowseIndice.Count = Ordenacao.ItemData(Ordenacao.ListIndex) Then
                    sOrdenacao = objBrowseIndice1.sOrdenacaoSQL
                    Exit For
                End If
            Next

        End If

        iPosInicio = 1
        iPosFim = InStr(sOrdenacao, ",")

        Do While iPosFim > 0

          colOrdenacao.Add Trim(Mid(sOrdenacao, iPosInicio, iPosFim - 1))

          iPosInicio = iPosFim + 1

          iPosFim = InStr(iPosInicio, sOrdenacao, ",")

        Loop

        colOrdenacao.Add Trim(Mid(sOrdenacao, iPosInicio))

        sCodGrupo = String(STRING_GRUPO, 0)

        If giLocalOperacao <> LOCALOPERACAO_ECF Then

            'obtem o codigo do grupo
            lErro = Obter_Grupo(sCodGrupo)
            If lErro <> SUCESSO Then gError 178428

        Else
        
            sCodGrupo = "supervisor"
            
        End If
            

        'le os campos disponiveis para a tela x grupo em questão
        lErro = CF("GrupoBrowseCampo_Le", sCodGrupo, Me.Name, colGrupoBrowseCampo)
        If lErro <> SUCESSO Then gError 178429

        For Each vOrdenacao In colOrdenacao

    '        iAchou = 0
    '
    ''        For Each objGrupoBrowseCampo In colGrupoBrowseCampo
    '         For Each objBrowseValorCampo In objBrowse1.colBrowseValorCampo
    '
    ''            If objGrupoBrowseCampo.sNome = vOrdenacao Then
    '             If objBrowseValorCampo.sNomeCampo = vOrdenacao Then
    '
    '                iAchou = 1
    '
    ''                Set objBrowseValorCampo = New AdmBrowseValorCampo
    '
    ''                objBrowseValorCampo.sNomeCampo = objGrupoBrowseCampo.sNome
    '
    '                objCampo.sNomeArq = objBrowse1.sNomeTabela
    '                objCampo.sNome = objBrowseValorCampo.sNomeCampo
    '
    '                lErro = CF("Campos_Le2", objCampo, lComando3)
    '                If lErro <> SUCESSO And lErro <> 9184 Then gError 178430
    '
    '                'se o campo não estiver cadastrado
    '                If lErro = 9184 Then gError 178431
    '
    '                If iPrimeiroCampo = 0 Then
    '
    '                    iPrimeiroCampo = 1
    '
    '                    Select Case objCampo.iTipo
    '
    '                        Case ADM_TIPO_SMALLINT
    '                            objBrowseValorCampo.vValorCampo = StrParaIntErr(Pesquisa.Text)
    '
    '                        Case ADM_TIPO_INTEGER
    '                            objBrowseValorCampo.vValorCampo = StrParaLongErr(Pesquisa.Text)
    '
    '                        Case ADM_TIPO_DOUBLE
    '                            objBrowseValorCampo.vValorCampo = StrParaDblErr(Pesquisa.Text)
    '
    '                        Case ADM_TIPO_VARCHAR
    '
    '                            If objCampo.sNome = "Produto" Then
    '
    '                                lErro = CF("Produto_Formata", Pesquisa.Text, sProduto, iPreenchido)
    '                                If lErro <> SUCESSO Then gError 178432
    '
    '                                If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""
    '
    '                                objBrowseValorCampo.vValorCampo = sProduto
    '
    '                            Else
    '                                objBrowseValorCampo.vValorCampo = Pesquisa.Text
    '                            End If
    '
    '                        Case ADM_TIPO_DATE
    '                            objBrowseValorCampo.vValorCampo = StrParaDateErr(Pesquisa.Text)
    '
    '                        Case Else
    '                            gError 178433
    '
    '                    End Select
    '
    '                Else
    '
    '                    Select Case objCampo.iTipo
    '
    '                        Case ADM_TIPO_SMALLINT
    '                            objBrowseValorCampo.vValorCampo = CInt(0)
    '
    '                        Case ADM_TIPO_INTEGER
    '                            objBrowseValorCampo.vValorCampo = CLng(0)
    '
    '                        Case ADM_TIPO_DOUBLE
    '                            objBrowseValorCampo.vValorCampo = CDbl(0)
    '
    '                        Case ADM_TIPO_VARCHAR
    '                            If objCampo.iTamanho < 500 Then
    '                                objBrowseValorCampo.vValorCampo = String(500, 0)
    '                            Else
    '                                objBrowseValorCampo.vValorCampo = String(objCampo.iTamanho, 0)
    '                            End If
    '
    '                        Case ADM_TIPO_DATE
    '                            objBrowseValorCampo.vValorCampo = CDate("1/1/1997")
    '
    '                        Case Else
    '                            gError 178434
    '
    '                    End Select
    '
    '                End If
    '
    '                Exit For
    '
    ' '               objBrowse1.colBrowseValorCampo.Add objBrowseValorCampo
    '
    '            End If
    '
    '        Next
    '
    '        If iAchou = 0 Then

                For Each objGrupoBrowseCampo In colGrupoBrowseCampo

                    If objGrupoBrowseCampo.sNome = vOrdenacao Then

                        Set objBrowseValorCampo = New AdmBrowseValorCampo

                        objBrowseValorCampo.sNomeCampo = objGrupoBrowseCampo.sNome

                        objCampo.sNomeArq = objBrowse1.sNomeTabela
                        objCampo.sNome = objBrowseValorCampo.sNomeCampo

                        lErro = CF("Campos_Le2", objCampo, lComando3)
                        If lErro <> SUCESSO And lErro <> 9184 Then gError 178430

                        'se o campo não estiver cadastrado
                        If lErro = 9184 Then gError 178431

                        If iPrimeiroCampo = 0 Then

                            iPrimeiroCampo = 1

                            Select Case objCampo.iTipo

                                Case ADM_TIPO_SMALLINT
                                    objBrowseValorCampo.vValorCampo = StrParaIntErr(Pesquisa.Text)

                                Case ADM_TIPO_INTEGER
                                    objBrowseValorCampo.vValorCampo = StrParaLongErr(Pesquisa.Text)

                                Case ADM_TIPO_DOUBLE
                                    objBrowseValorCampo.vValorCampo = StrParaDblErr(Pesquisa.Text)

                                Case ADM_TIPO_VARCHAR
                                    
                                    If OptFiltrar.Value Then
                                        objBrowseValorCampo.vValorCampo = "%" & Trim(Pesquisa.ClipText) & "%"
                                    Else
                                        If objCampo.iSubTipo = 6 Then
    
                                            lErro = CF("Produto_Formata", Pesquisa.Text, sProduto, iPreenchido)
                                            If lErro <> SUCESSO Then gError 178432
    
                                            If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""
    
                                            objBrowseValorCampo.vValorCampo = sProduto
    
                                        Else
                                            objBrowseValorCampo.vValorCampo = Pesquisa.Text
                                        End If
                                    
                                    End If
                                    
                                Case ADM_TIPO_DATE
                                    objBrowseValorCampo.vValorCampo = StrParaDateErr(Pesquisa.Text)

                                Case Else
                                    gError 178433

                            End Select
                            
                            If OptFiltrar.Value And Len(Trim(Pesquisa.ClipText)) > 0 Then
                                objBrowse1.sSelecaoSQL2 = vOrdenacao & " LIKE ? "
                                For iIndice = objBrowse1.colSelecao.Count To 1 Step -1
                                    objBrowse1.colSelecao.Remove iIndice
                                Next
                                objBrowse1.colSelecao.Add objBrowseValorCampo.vValorCampo
                                For Each vSel In gcolSelecaoOrig
                                    objBrowse1.colSelecao.Add vSel
                                Next
                                gPrecisaRemoverFiltro = True
                            End If

                        Else

                            Select Case objCampo.iTipo

                                Case ADM_TIPO_SMALLINT
                                    objBrowseValorCampo.vValorCampo = CInt(0)

                                Case ADM_TIPO_INTEGER
                                    objBrowseValorCampo.vValorCampo = CLng(0)

                                Case ADM_TIPO_DOUBLE
                                    objBrowseValorCampo.vValorCampo = CDbl(0)

                                Case ADM_TIPO_VARCHAR
                                    If objCampo.iTamanho < 500 Then
                                        objBrowseValorCampo.vValorCampo = String(500, 0)
                                    Else
                                        objBrowseValorCampo.vValorCampo = String(objCampo.iTamanho, 0)
                                    End If

                                Case ADM_TIPO_DATE
                                    objBrowseValorCampo.vValorCampo = CDate("1/1/1997")

                                Case Else
                                    gError 178434

                            End Select

                        End If

                        objBrowse1.colBrowseValorCampo.Add objBrowseValorCampo

                        Exit For

                    End If

                Next

    '        End If
    
            If OptFiltrar.Value Then Exit For

        Next

        If OptFiltrar.Value Then
            lErro = objBrowse1.Browse_Trata_Parametros(objBrowse1)
        ElseIf gPrecisaRemoverFiltro Then
            lErro = objBrowse1.Browse_Trata_Parametros(objBrowse1)
            gPrecisaRemoverFiltro = False
        Else
            lErro = objBrowse1.Browse_Trata_Parametros(objBrowse1, 1)
        End If
        If lErro <> SUCESSO Then gError 178435

        lErro = Inicializa_colBrowseValorCampo(objObjeto)
        If lErro <> SUCESSO Then gError 178436


        Call Comando_Fechar(lComando3)

    End If

    If bCargaInicialECF Then
    
        bCargaInicialECF = False
        Pesquisa.SetFocus
    
    End If
    
    Exit Sub

Erro_Timer1_Timer:

    Select Case gErr

        Case 178422, 178428 To 178430, 178432, 178435, 178436

        Case 178427
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 178431
            Call Rotina_Erro(vbOKOnly, "ERRO_CAMPO_NAO_CADASTRADO", gErr, objCampo.sNome, objCampo.sNomeArq)

        Case 178433, 178434
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_CAMPO_INVALIDO", gErr, objCampo.iTipo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197898)

    End Select

    Call Comando_Fechar(lComando3)

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
   RaiseEvent Unload
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

Public Property Let MousePointer(ByVal iTipo As Integer)
    Parent.MousePointer = iTipo
End Property

Public Property Get MousePointer() As Integer
    MousePointer = Parent.MousePointer
End Property

Public Property Get Picture2() As Object
    Set Picture2 = Me.Controls("Picture2")
End Property

Public Property Get BotaoPlanilha() As Object
    Set BotaoPlanilha = Me.Controls("BotaoPlanilha")
End Property

Public Property Get BotaoSeleciona() As Object
    Set BotaoSeleciona = Me.Controls("BotaoSeleciona")
End Property

Public Property Get BotaoEdita() As Object
    Set BotaoEdita = Me.Controls("BotaoEdita")
End Property

Public Property Get BotaoConsulta() As Object
    Set BotaoConsulta = Me.Controls("BotaoConsulta")
End Property

Public Property Get BotaoConfigura() As Object
    Set BotaoConfigura = Me.Controls("BotaoConfigura")
End Property

Public Property Get BotaoPesquisa() As Object
    Set BotaoPesquisa = Me.Controls("BotaoPesquisa")
End Property

Public Property Get BotaoAtualiza() As Object
    Set BotaoAtualiza = Me.Controls("BotaoAtualiza")
End Property

Public Property Get BotaoFecha() As Object
    Set BotaoFecha = Me.Controls("BotaoFecha")
End Property

'**** fim do trecho a ser copiado *****

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Function Inicializa_colBrowseValorCampo(objObjeto As Object) As Long

Dim objBrowseValorCampo As AdmBrowseValorCampo
Dim objBrowseCampo As AdmBrowseCampo
Dim lErro As Long

On Error GoTo Erro_Inicializa_colBrowseValorCampo

    Set objBrowse1.colBrowseValorCampo = New Collection

    For Each objBrowseCampo In objBrowse1.colBrowseCampo

        Set objBrowseValorCampo = New AdmBrowseValorCampo
            
        If Len(Trim(objBrowseCampo.sNome)) <> 0 Then
            objBrowseValorCampo.vValorCampo = CallByName(objObjeto, objBrowseCampo.sNome, VbGet)
        End If
        
        objBrowseValorCampo.sNomeCampo = objBrowseCampo.sNomeCampo
            
        objBrowse1.colBrowseValorCampo.Add objBrowseValorCampo
    
    Next
    
    Inicializa_colBrowseValorCampo = SUCESSO
    
    Exit Function
    
Erro_Inicializa_colBrowseValorCampo:

    Inicializa_colBrowseValorCampo = gErr
    
    Select Case gErr
    
        Case Else
        
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144021)

    End Select

    Exit Function
    
End Function

Private Function Move_BrowseValorCampo(objObjeto As Object) As Long

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Move_BrowseValorCampo

    For iIndice = 1 To objBrowse1.colBrowseValorCampo.Count
    
        lErro = CallByName(objObjeto, objBrowse1.colBrowseCampo.Item(iIndice).sNome, VbLet, objBrowse1.colBrowseValorCampo.Item(iIndice).vValorCampo)

    Next
    
    Move_BrowseValorCampo = SUCESSO
    
    Exit Function
    
Erro_Move_BrowseValorCampo:

    Move_BrowseValorCampo = gErr

    Select Case gErr
    
        Case Else
        
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144022)

    End Select

    Exit Function
    
End Function

Public Sub Tamanho(Largura As Integer, Altura As Integer)
    
Dim iAlturaGrid As Integer

    UserControl.Size Largura, Altura
    GridBrowse.Width = Largura - 200
    BarraScroll.left = GridBrowse.Width - 200
    Picture2.top = Altura - 1100
    Picture2.Width = Largura - 400
    Picture1.left = Largura - 1155 - 400
    Call Reposiciona_Botoes
    Ordenacao.Width = Largura - 1355
    iAlturaGrid = Altura - GridBrowse.top - Picture2.Height - DISTANCIA_ENTRE_GRID_PICTURE - 100
    objBrowse1.iGrid_LinhasExibidas = iAlturaGrid \ 240 '250
    GridBrowse.Height = (objBrowse1.iGrid_LinhasExibidas + 1) * 240 '250
    BarraScroll.Height = GridBrowse.Height - 100
    Picture2.top = Altura - 1000
    objBrowse1.iGrid_LinhaFinal = objBrowse1.iGrid_LinhaInicial + objBrowse1.iGrid_LinhasExibidas - 1
    objBrowse1.objBarraScroll.LargeChange = objBrowse1.iGrid_LinhasExibidas - 1
    Call BotaoAtualiza_Click
        
End Sub

Private Sub Reposiciona_Botoes()

Dim iEspacoEntreBotoes As Integer
Dim iLarguraBotao As Integer
Dim iPosicaoBotao As Integer

    If Picture2.Width > iNumeroBotoes * (ESPACO_ENTRE_BOTOES + LARGURA_BOTOES) + ESPACO_ENTRE_BOTOES Then
        iEspacoEntreBotoes = (Picture2.Width - iNumeroBotoes * LARGURA_BOTOES) / (iNumeroBotoes + 1)
        iLarguraBotao = LARGURA_BOTOES
    Else
        iEspacoEntreBotoes = ESPACO_ENTRE_BOTOES
        iLarguraBotao = (Picture2.Width - (iNumeroBotoes + 1) * ESPACO_ENTRE_BOTOES) / iNumeroBotoes
        If iLarguraBotao < MINIMO_LARGURA_BOTAO Then iLarguraBotao = MINIMO_LARGURA_BOTAO
    End If

    iPosicaoBotao = iEspacoEntreBotoes
    
    If objBrowse1.objBrowseArquivo.iBotaoSeleciona = 1 Then
    
        objBrowse1.objForm.BotaoSeleciona.left = iPosicaoBotao
        objBrowse1.objForm.BotaoSeleciona.Width = iLarguraBotao
        iPosicaoBotao = iPosicaoBotao + iLarguraBotao + iEspacoEntreBotoes
        
    End If
    
    If objBrowse1.objBrowseArquivo.iBotaoEdita = 1 Then
    
        objBrowse1.objForm.BotaoEdita.left = iPosicaoBotao
        objBrowse1.objForm.BotaoEdita.Width = iLarguraBotao
        iPosicaoBotao = iPosicaoBotao + iLarguraBotao + iEspacoEntreBotoes
        
    End If
    
    If objBrowse1.objBrowseArquivo.iBotaoConsulta = 1 Then
    
        objBrowse1.objForm.BotaoConsulta.left = iPosicaoBotao
        objBrowse1.objForm.BotaoConsulta.Width = iLarguraBotao
        iPosicaoBotao = iPosicaoBotao + iLarguraBotao + iEspacoEntreBotoes
        
    End If
    
    objBrowse1.objForm.BotaoConfigura.left = iPosicaoBotao
    objBrowse1.objForm.BotaoConfigura.Width = iLarguraBotao
    iPosicaoBotao = iPosicaoBotao + iLarguraBotao + iEspacoEntreBotoes
    
    objBrowse1.objForm.BotaoPesquisa.left = iPosicaoBotao
    objBrowse1.objForm.BotaoPesquisa.Width = iLarguraBotao
    iPosicaoBotao = iPosicaoBotao + iLarguraBotao + iEspacoEntreBotoes
    
    objBrowse1.objForm.BotaoAtualiza.left = iPosicaoBotao
    objBrowse1.objForm.BotaoAtualiza.Width = iLarguraBotao
    iPosicaoBotao = iPosicaoBotao + iLarguraBotao + iEspacoEntreBotoes
    
    If gsNomePrinc <> "SGEECF" Then
        objBrowse1.objForm.BotaoPlanilha.left = iPosicaoBotao
        objBrowse1.objForm.BotaoPlanilha.Width = iLarguraBotao
        objBrowse1.objForm.BotaoPlanilha.Visible = True
        iPosicaoBotao = iPosicaoBotao + iLarguraBotao + iEspacoEntreBotoes
    Else
        objBrowse1.objForm.BotaoPlanilha.Visible = False
    End If
    
    objBrowse1.objForm.BotaoFecha.left = iPosicaoBotao
    objBrowse1.objForm.BotaoFecha.Width = iLarguraBotao
    iPosicaoBotao = iPosicaoBotao + iLarguraBotao + iEspacoEntreBotoes
    
End Sub

Function Carrega_Combobox_Opcoes() As Long

Dim colOpcoes As New Collection
Dim vOpcao As Variant
Dim lErro As Long

On Error GoTo Erro_Carrega_Combobox_Opcoes

    lErro = CF("BrowseOpcaoOrdenacao_Le_Opcoes", objBrowse1.objForm.Name, colOpcoes)
    If lErro <> SUCESSO Then gError 178329

    ComboOpcoes.Clear

    For Each vOpcao In colOpcoes
        ComboOpcoes.AddItem vOpcao
    Next

    Carrega_Combobox_Opcoes = SUCESSO
    
    Exit Function
    
Erro_Carrega_Combobox_Opcoes:

    Carrega_Combobox_Opcoes = gErr
    
    Select Case gErr
    
        Case 178329
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178330)

    End Select

    Exit Function

End Function

Private Sub OptPosicionar_Click()
    Call Pesquisa_Change
End Sub

Private Sub OptFiltrar_Click()
    Call Pesquisa_Change
End Sub

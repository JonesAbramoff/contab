VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl BancosOcx 
   ClientHeight    =   3180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8295
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3180
   ScaleWidth      =   8295
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   2370
      Picture         =   "BancosOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Numeração Automática"
      Top             =   240
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5985
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   165
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "BancosOcx.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "BancosOcx.ctx":0244
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "BancosOcx.ctx":03CE
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "BancosOcx.ctx":0900
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.TextBox LayoutCheque 
      Height          =   315
      Left            =   1860
      MaxLength       =   80
      TabIndex        =   4
      Top             =   2025
      Width           =   3900
   End
   Begin VB.TextBox LayoutBoleto 
      Height          =   315
      Left            =   1860
      MaxLength       =   80
      TabIndex        =   5
      Top             =   2655
      Width           =   3900
   End
   Begin VB.ListBox BancosList 
      Height          =   1950
      IntegralHeight  =   0   'False
      Left            =   5970
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   1020
      Width           =   2160
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   300
      Left            =   1860
      TabIndex        =   0
      Top             =   240
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "9999"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox RazaoSocial 
      Height          =   315
      Left            =   1875
      TabIndex        =   2
      Top             =   840
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   50
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox NomeReduzido 
      Height          =   315
      Left            =   1860
      TabIndex        =   3
      Top             =   1410
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   15
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Left            =   1125
      TabIndex        =   12
      Top             =   270
      Width           =   615
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Left            =   1170
      TabIndex        =   13
      Top             =   870
      Width           =   570
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Left            =   315
      TabIndex        =   14
      Top             =   1470
      Width           =   1425
   End
   Begin VB.Label Label4 
      Caption         =   "Layout de Cheque:"
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
      Left            =   120
      TabIndex        =   15
      Top             =   2055
      Width           =   1620
   End
   Begin VB.Label Label5 
      Caption         =   "Layout de Boleto:"
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
      Left            =   225
      TabIndex        =   16
      Top             =   2670
      Width           =   1515
   End
   Begin VB.Label Label6 
      Caption         =   "Bancos"
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
      Left            =   5970
      TabIndex        =   17
      Top             =   825
      Width           =   675
   End
End
Attribute VB_Name = "BancosOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private Sub Traz_Banco_Tela(objBanco As ClassBanco)

Dim iIndice As Integer

    'mostra dados do banco na tela
    Codigo.Text = objBanco.iCodBanco
    RazaoSocial.Text = objBanco.sNome
    NomeReduzido.Text = objBanco.sNomeReduzido
    LayoutCheque.Text = objBanco.sLayoutCheque
    LayoutBoleto.Text = objBanco.sLayoutBoleto

    'Seleciona Nome Reduzido na ListBox
    For iIndice = 0 To BancosList.ListCount - 1

        If BancosList.List(iIndice) = NomeReduzido.Text Then
            BancosList.ListIndex = iIndice
            Exit For
        End If
    Next

    iAlterado = 0

End Sub

Private Sub BancosList_DblClick()

Dim lErro As Long
Dim objBanco As New ClassBanco
On Error GoTo Erro_BancosList_DblClick

    objBanco.iCodBanco = BancosList.ItemData(BancosList.ListIndex)

    lErro = CF("Banco_Le", objBanco)
    If lErro <> SUCESSO Then Error 16092

    Call Traz_Banco_Tela(objBanco)
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_BancosList_DblClick:

    Select Case Err

        Case 16092

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143521)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim iIndice As Integer
Dim objBanco As New ClassBanco

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'verifica preenchimento do codigo
    If Len(Trim(Codigo.Text)) = 0 Then Error 16121

    objBanco.iCodBanco = CInt(Codigo.Text)

    lErro = CF("Banco_Le", objBanco)

    If lErro = 16091 Then Error 16122

    If lErro <> SUCESSO Then Error 16130

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_BANCO", objBanco.iCodBanco)

    If vbMsgRes = vbYes Then

        'Exclui Banco
        lErro = CF("Banco_Exclui", objBanco)
        If lErro <> SUCESSO Then Error 16124

        'procura indice do banco no ListBox
        For iIndice = 0 To BancosList.ListCount - 1

            If BancosList.ItemData(iIndice) = objBanco.iCodBanco Then
                'remove banco do ListBox
                BancosList.RemoveItem (iIndice)
                Exit For
            End If

        Next

        Call Limpa_Tela_Bancos
        
    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 16121
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)

        Case 16122
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BANCO_NAO_CADASTRADO", Err, objBanco.iCodBanco)

        Case 16124

        Case 16130
 
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143522)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click

    'Gera número automático para o código de Banco
    lErro = CF("Banco_Automatico", iCodigo)
    If lErro <> SUCESSO Then Error 57700

    Codigo.Text = CStr(iCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57700
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143523)
    
    End Select

    Exit Sub
    
End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Verifica preenchimento do codigo
    If Len(Trim(Codigo.Text)) > 0 Then

        'Verifica se codigo é numérico
        If Not IsNumeric(Codigo.Text) Then Error 16386

        'Verifica se codigo é menor que um
        If CInt(Codigo.Text) < 1 Then Error 16385

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True


    Select Case Err

        Case 16385
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_MENOR_QUE_UM", Err, Codigo.Text)

        Case 16386
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_NUMERICO", Err, Codigo.Text)


        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143524)

    End Select

    Exit Sub
    
End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 16209

    Call Limpa_Tela_Bancos
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 16209
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143525)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_Bancos()

Dim lErro As Long
    
    Call Limpa_Tela(Me)

    BancosList.ListIndex = -1

    iAlterado = 0
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 80455

    Call Limpa_Tela_Bancos
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 80455

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143526)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodNome As AdmCodigoNome

On Error GoTo Erro_Form_Load

    'leitura dos bancos no BD
    lErro = CF("Cod_Nomes_Le", "Bancos", "CodBanco", "NomeReduzido", STRING_BANCO_NOME_REDUZIDO, colCodigoNome)
    If lErro <> SUCESSO Then Error 16087

    'preenche listbox com nomes reduzidos dos bancos
    For Each objCodNome In colCodigoNome
        BancosList.AddItem objCodNome.sNome
        BancosList.ItemData(BancosList.NewIndex) = objCodNome.iCodigo
    Next
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 16087

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143527)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objBanco As ClassBanco) As Long
'Trata os parametros que podem ser passados quando ocorre a chamada da tela de Bancos

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se houve passagem de parametro
    If Not (objBanco Is Nothing) Then

        lErro = CF("Banco_Le", objBanco)
        If lErro <> SUCESSO And lErro <> 16091 Then Error 16129

        If lErro = SUCESSO Then
        
            Call Traz_Banco_Tela(objBanco)

        Else
        
            Codigo.Text = objBanco.iCodBanco

        End If
        
    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 16129

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143528)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Sub Move_Tela_Memoria(objBanco As ClassBanco)

    objBanco.iCodBanco = CInt(Codigo.Text)
    objBanco.sNome = RazaoSocial.Text
    objBanco.sNomeReduzido = NomeReduzido.Text
    objBanco.sLayoutCheque = LayoutCheque.Text
    objBanco.sLayoutBoleto = LayoutBoleto.Text

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objBanco As New ClassBanco

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'verifica preenchimento do codigo
    If Len(Trim(Codigo.Text)) = 0 Then Error 16105

    'verifica preenchimento do nome
    If Len(Trim(RazaoSocial.Text)) = 0 Then Error 16106

    'verifica preenchimento do nome reduzido
    If Len(Trim(NomeReduzido.Text)) = 0 Then Error 16107

    'preenche objBanco
    Call Move_Tela_Memoria(objBanco)
    
    lErro = Trata_Alteracao(objBanco, objBanco.iCodBanco)
    If lErro <> SUCESSO Then Error 16125
       
    lErro = CF("Banco_Grava", objBanco)
    If lErro <> SUCESSO Then Error 16108

    'Remove e adiciona na ListBox
    Call BancosList_Remove(objBanco)
    Call BancosList_Adiciona(objBanco)
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 16105
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)

        Case 16106
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_NAO_PREENCHIDO", Err)

        Case 16107
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_PREENCHIDO", Err)

        Case 16108, 16125

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143529)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

 Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
    
End Sub

Private Sub LayoutBoleto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LayoutBoleto_Validate(Cancel As Boolean)

Dim objArquivo As Object

On Error GoTo Erro_LayoutBoleto_Validate

    If Len(LayoutBoleto.Text) > 0 Then
    
        Set objArquivo = CreateObject("Scripting.FileSystemObject")
        
        'verifica a existencia do arquivo
        If Not objArquivo.FileExists(LayoutBoleto.Text & NOME_EXTENSAO_RELATORIO) Then Error 55945
        
        'verifica se o nome do relatorio não ultrapassa 8 caracteres
        If Len(objArquivo.GetBaseName(LayoutBoleto.Text)) > NUM_MAX_CARACTERES_NOME_RELATORIO Then Error 55946
        
    End If

    Exit Sub

Erro_LayoutBoleto_Validate:

    Cancel = True
    
    Select Case Err
    
        Case 55945
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_NAO_ENCONTRADO", Err, LayoutBoleto.Text & NOME_EXTENSAO_RELATORIO)
        
        Case 55946
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_ARQUIVO_MAIOR_PERMITIDO", Err, objArquivo.GetBaseName(LayoutBoleto.Text), NUM_MAX_CARACTERES_NOME_RELATORIO)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143530)
    
    End Select
    
    Exit Sub

End Sub

Private Sub LayoutCheque_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LayoutCheque_Validate(Cancel As Boolean)

Dim objArquivo As Object

On Error GoTo Erro_LayoutCheque_Validate

    If Len(LayoutCheque.Text) > 0 Then
    
        Set objArquivo = CreateObject("Scripting.FileSystemObject")
        
        'verifica a existencia do arquivo
        If Not objArquivo.FileExists(LayoutCheque.Text & NOME_EXTENSAO_RELATORIO) Then Error 55948
        
        'verifica se o nome do relatorio não ultrapassa 8 caracteres
        If Len(objArquivo.GetBaseName(LayoutCheque.Text)) > NUM_MAX_CARACTERES_NOME_RELATORIO Then Error 55949
        
    End If

    Exit Sub

Erro_LayoutCheque_Validate:

    Cancel = True
    
    Select Case Err
    
        Case 55948
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_NAO_ENCONTRADO", Err, LayoutCheque.Text & NOME_EXTENSAO_RELATORIO)
        
        Case 55949
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_ARQUIVO_MAIOR_PERMITIDO", Err, objArquivo.GetBaseName(LayoutCheque.Text), NUM_MAX_CARACTERES_NOME_RELATORIO)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143531)
    
    End Select
    
    Exit Sub

End Sub

Private Sub NomeReduzido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NomeReduzido_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NomeReduzido_Validate
    
    'Se está preenchido, testa se começa por letra
    If Len(Trim(NomeReduzido.Text)) > 0 Then

        If Not IniciaLetra(NomeReduzido.Text) Then Error 57819

    End If
    
    Exit Sub

Erro_NomeReduzido_Validate:

    Cancel = True

    
    Select Case Err
    
        Case 57819
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_COMECA_LETRA", Err, NomeReduzido.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143532)
    
    End Select
    
    Exit Sub

End Sub

Private Sub RazaoSocial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

'Preenche os campos da tela com os correspondentes do BD
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim iIndice As Integer

    'Coloca colCampoValor na Tela
    'Conversão de tipagem para a tipagem da tela se necessário
    Codigo.Text = CStr(colCampoValor.Item("CodBanco").vValor)
    RazaoSocial.Text = colCampoValor.Item("Nome").vValor
    NomeReduzido.Text = colCampoValor.Item("NomeReduzido").vValor
    LayoutCheque.Text = colCampoValor.Item("LayoutCheque").vValor
    LayoutBoleto.Text = colCampoValor.Item("LayoutBoleto").vValor

    'Seleciona Nome Reduzido na ListBox
    For iIndice = 0 To BancosList.ListCount - 1

        If BancosList.List(iIndice) = NomeReduzido.Text Then
            BancosList.ListIndex = iIndice
            Exit For
        End If

    Next

    iAlterado = 0
    
End Sub

'Extrai os campos da tela que correspondem aos campos no BD
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim objCampoValor As AdmCampoValor
Dim iCodigo As Integer

    'Informa tabela associada à Tela
    sTabela = "Bancos"

    'Realiza conversões necessárias de campos da tela para campos do BD
    'A tipagem dos valores DEVE SER A MESMA DO BD
    If Codigo.Text = "" Then
        iCodigo = 0
    Else
        iCodigo = CInt(Codigo.Text)
    End If
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "CodBanco", iCodigo, 0, "CodBanco"
    colCampoValor.Add "Nome", RazaoSocial.Text, STRING_BANCO_NOME, "Nome"
    colCampoValor.Add "NomeReduzido", NomeReduzido.Text, STRING_BANCO_NOME_REDUZIDO, "NomeReduzido"
    colCampoValor.Add "LayoutCheque", LayoutCheque.Text, STRING_BANCO_LAYOUT_CHEQUE, "LayoutCheque"
    colCampoValor.Add "LayoutBoleto", LayoutBoleto.Text, STRING_BANCO_LAYOUT_BOLETO, "LayoutBoleto"

End Sub

Private Sub BancosList_Adiciona(objBanco As ClassBanco)
'Adiciona na List

    'Insere Banco na ListBox
    BancosList.AddItem objBanco.sNomeReduzido
    BancosList.ItemData(BancosList.NewIndex) = objBanco.iCodBanco

End Sub

Private Sub BancosList_Remove(objBanco As ClassBanco)
'Percorre a ListBox de Bancos para remover o tipo caso ele exista

Dim iIndice As Integer

For iIndice = 0 To BancosList.ListCount - 1

    If BancosList.ItemData(iIndice) = objBanco.iCodBanco Then

        BancosList.RemoveItem iIndice
        Exit For

    End If

Next

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BANCOS
    Set Form_Load_Ocx = Me
    Caption = "Bancos"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Bancos"
    
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

''Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
''
''    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
''        Call BotaoProxNum_Click
''    End If
''
''End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(label1, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub


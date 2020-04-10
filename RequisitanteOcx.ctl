VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RequisitanteOcx 
   ClientHeight    =   4110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8970
   KeyPreview      =   -1  'True
   ScaleHeight     =   4110
   ScaleWidth      =   8970
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6630
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   105
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RequisitanteOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RequisitanteOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RequisitanteOcx.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RequisitanteOcx.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ListBox Requisitantes 
      Height          =   2985
      Left            =   6555
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   960
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Caption         =   "Requisitante"
      Height          =   2985
      Left            =   105
      TabIndex        =   0
      Top             =   960
      Width           =   6315
      Begin VB.TextBox Email 
         Height          =   345
         Left            =   2190
         TabIndex        =   10
         Top             =   2355
         Width           =   3870
      End
      Begin VB.ComboBox CodUsuario 
         Height          =   315
         Left            =   4410
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   270
         Width           =   1725
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   315
         Left            =   2970
         Picture         =   "RequisitanteOcx.ctx":0994
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Numeração Automática"
         Top             =   270
         Width           =   300
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   2190
         TabIndex        =   2
         Top             =   270
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Nome 
         Height          =   315
         Left            =   2190
         TabIndex        =   5
         Top             =   810
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NomeReduzido 
         Height          =   315
         Left            =   2190
         TabIndex        =   7
         Top             =   1335
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Ccl 
         Height          =   315
         Left            =   2190
         TabIndex        =   9
         Top             =   1860
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         PromptChar      =   " "
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Usuário:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3600
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   20
         Top             =   330
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "E-mail:"
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
         Index           =   13
         Left            =   1455
         TabIndex        =   18
         Top             =   2430
         Width           =   585
      End
      Begin VB.Label CclLabel 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Custo/Lucro:"
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
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   8
         Top             =   1950
         Width           =   2010
      End
      Begin VB.Label Label4 
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
         Left            =   1575
         TabIndex        =   4
         Top             =   870
         Width           =   555
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
         Left            =   720
         TabIndex        =   6
         Top             =   1395
         Width           =   1410
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
         Index           =   0
         Left            =   1470
         TabIndex        =   1
         Top             =   330
         Width           =   660
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Requisitantes"
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
      Left            =   6555
      TabIndex        =   11
      Top             =   735
      Width           =   1170
   End
End
Attribute VB_Name = "RequisitanteOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1

Private Function Inicializa_MascaraCcl() As Long
'Inicializa a mascara do centro de custo

Dim sMascaraCcl As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_mascaraccl
   
    'le a mascara dos centros de custo/lucro
    lErro = MascaraCcl(sMascaraCcl)
    If lErro <> SUCESSO Then gError 49096
    
    'Atribui a máscara ao controle
    Ccl.Mask = sMascaraCcl

    Inicializa_MascaraCcl = SUCESSO

    Exit Function

Erro_Inicializa_mascaraccl:

    Inicializa_MascaraCcl = gErr
    
    Select Case gErr
    
        Case 49096
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174073)
        
    End Select

    Exit Function

End Function

Private Sub ListaRequisitante_Exclui(lCodigo As Long)
'Percorre a ListBox de Requisitante para remover o tipo caso ele exista

Dim iIndice As Integer
    
    'Para cada Item da Listbox de Requisitante
    For iIndice = 0 To Requisitantes.ListCount - 1
        'Se o código guardado no Itemdata é igual ao código passado
        If Requisitantes.ItemData(iIndice) = lCodigo Then
            'Remove esse item da listBox
            Requisitantes.RemoveItem (iIndice)
            'Sai do Loop
            Exit For
        End If
    Next

    Exit Sub

End Sub

Private Sub ListaRequisitante_Adiciona(objRequisitante As ClassRequisitante)
'Adiciona na ListBox de Requisitante
    
    'Adiciona na ListBox o Requisitante passado
    Requisitantes.AddItem objRequisitante.sNomeReduzido
    'Guarda no ItemData do Requisitante inserido o valor do seu código
    Requisitantes.ItemData(Requisitantes.NewIndex) = objRequisitante.lCodigo

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objRequisitante As New ClassRequisitante

On Error GoTo Erro_Tela_Preenche

    'Carrega objRequisitante com os dados passados em colCampoValor
    objRequisitante.lCodigo = colCampoValor.Item("Codigo").vValor
    objRequisitante.sNome = colCampoValor.Item("Nome").vValor
    objRequisitante.sNomeReduzido = colCampoValor.Item("NomeReduzido").vValor
    objRequisitante.sCcl = colCampoValor.Item("Ccl").vValor
    
    'le o requisitante
    lErro = CF("Requisitante_Le", objRequisitante)
    If lErro <> SUCESSO And lErro <> 49084 Then gError 49056
    
    lErro = Traz_Requisitante_Tela(objRequisitante)
    If lErro <> SUCESSO Then gError 49056

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 49056

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174074)

    End Select

    Exit Sub

End Sub

Function Move_Tela_Memoria(objRequisitante As ClassRequisitante) As Long
'Recolhe os dados da tela e armazena em objRequisitante

Dim lErro As Long
Dim lCodigo As Long
Dim sCclFormatada As String
Dim iCclPreenchida As Integer

On Error GoTo Erro_Move_Tela_Memoria

    'Move os dados da tela para objRequisitante
    objRequisitante.lCodigo = StrParaLong(Codigo.Text)
    objRequisitante.sNome = Nome.Text
    objRequisitante.sNomeReduzido = NomeReduzido.Text
    
    'Se o Centro de Custo foi informado
    If Len(Trim(Ccl.ClipText)) > 0 Then
        'Passo o CCL para o formato do Banco de Dados
        lErro = CF("Ccl_Formata", Ccl.Text, sCclFormatada, iCclPreenchida)
        If lErro <> SUCESSO Then gError 49100
        'GUarda no objRequisitante o Centro de Custo
        If iCclPreenchida = CCL_PREENCHIDA Then objRequisitante.sCcl = sCclFormatada
    
    End If
    
    objRequisitante.sEmail = Email.Text
    objRequisitante.sCodUsuario = CodUsuario.Text
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 49100
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174075)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long
'Verifica se dados de Requisitante necessários foram preenchidos
'Grava Requisitante no BD
'Atualiza List

Dim lErro As Long
Dim objRequisitante As New ClassRequisitante

On Error GoTo Erro_Gravar_Registro
    
    'Coloca o MouseIcon de Ampulheta durante a gravação
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se foi preenchido o Código
    If Len(Trim(Codigo.Text)) = 0 Then gError 49051

    'Verifica se foi preenchido o Nome
    If Len(Trim(Nome.Text)) = 0 Then gError 49052

    'Verifica se foi preenchido o Nome Reduzido
    If Len(Trim(NomeReduzido.Text)) = 0 Then gError 49053

    'Recolhe os dados da tela para o objRequisitante
    lErro = Move_Tela_Memoria(objRequisitante)
    If lErro <> SUCESSO Then gError 49054

    'Se tentou alterar dados do Requisitante automático, Erro
    If objRequisitante.lCodigo = REQUISITANTE_AUTOMATICO_CODIGO Then gError 67303

    lErro = Trata_Alteracao(objRequisitante, objRequisitante.lCodigo)
    If lErro <> SUCESSO Then Error 32291

    'Guarda/Altera os dados do requisitante no BD
    lErro = CF("Requisitante_Grava", objRequisitante)
    If lErro <> SUCESSO Then gError 49055

    'Exclui o Requisitante da lista de requisitantes
    Call ListaRequisitante_Exclui(objRequisitante.lCodigo)

    'Adiciona o Requisitante na lista de requisitantes
    Call ListaRequisitante_Adiciona(objRequisitante)
    
    'Limpa a Tela
    Call Limpa_Tela_Requisitante_FechaSeta
    
    'Coloca o MouseIcon de setinha
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 32291

        Case 49051
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 49052
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_NAO_PREENCHIDO", gErr)

        Case 49053
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_PREENCHIDO", gErr)

        Case 49054, 49055

        Case 67303
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_REQUISITANTE_AUTOMATICO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174076)

    End Select

    'Coloca o MouseIcon de setinha
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

End Function

Function Traz_Requisitante_Tela(objRequisitante As ClassRequisitante) As Long
'Traz o Requisitante para tela

Dim lErro As Long, iIndice As Integer
Dim sCclMascarado As String

On Error GoTo Erro_Traz_Requisitante_Tela

    'Limpa as imformações da tela
    Call Limpa_Tela_Requisitante

    'Coloca o código na tela
    Codigo.Text = CStr(objRequisitante.lCodigo)
    'Coloca o Nome na tela
    Nome.Text = objRequisitante.sNome
    'Coloca o Nome Reduzido na tela
    NomeReduzido.Text = objRequisitante.sNomeReduzido
    
    'Se o centro de custo está preenchido
    If Len(Trim(objRequisitante.sCcl)) > 0 Then
        'COloca o Ccl no Formato de exibição
        lErro = Mascara_MascararCcl(objRequisitante.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then gError 49144
        
        'Coloca o Ccl na tela
        Ccl.PromptInclude = False
        Ccl.Text = sCclMascarado
        Ccl.PromptInclude = True
    
    End If
    
    If objRequisitante.sCodUsuario <> "" Then
        For iIndice = 0 To CodUsuario.ListCount - 1
            If objRequisitante.sCodUsuario = CodUsuario.List(iIndice) Then
                CodUsuario.ListIndex = iIndice
                Exit For
            End If
        Next
    Else
        CodUsuario.ListIndex = -1
    End If
    
    Email.Text = objRequisitante.sEmail
    
    iAlterado = 0

    Exit Function

Erro_Traz_Requisitante_Tela:

    Traz_Requisitante_Tela = gErr

    Select Case gErr

        Case 49144
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174077)

    End Select

    Exit Function

End Function

Function Trata_Parametros(Optional objRequisitante As ClassRequisitante) As Long

Dim lErro As Long
Dim lCodigo As Long
Dim bEncontrou As Boolean

On Error GoTo Erro_Trata_Parametros
        
    'Se houver Requisitante passado como parâmetro, exibe seus dados
    If Not (objRequisitante Is Nothing) Then
        
        'Até agora não tem os dados do Requisitante
        bEncontrou = False
        
        'Se o código do requisitante foi informmado
        If objRequisitante.lCodigo > 0 Then

            'Lê Requisitante no BD a partir do código
            lErro = CF("Requisitante_Le", objRequisitante)
            If lErro <> SUCESSO And lErro <> 49084 Then gError 49044
            'Se encontrou, guarda em bEncontrou
            If lErro = SUCESSO Then bEncontrou = True
        
        'Se o NomeReduzido foi informado
        ElseIf Len(Trim(objRequisitante.sNomeReduzido)) > 0 Then
            
            'Lê Requisitante no BD a partir do Nome Reduzido
            lErro = CF("Requisitante_Le_NomeReduzido", objRequisitante)
            If lErro <> SUCESSO And lErro <> 51152 Then gError 49043
            'Se encontrou, guarda em bEncontrou
            If lErro = SUCESSO Then bEncontrou = True
            
        End If

        'Se o requisitante passado foi encontrado no BD
        If bEncontrou Then
            'Exibe os dados do Requisitante
            lErro = Traz_Requisitante_Tela(objRequisitante)
            If lErro <> SUCESSO Then gError 49045
        Else
            'Coloca na tela as informações passadas
            If objRequisitante.lCodigo > 0 Then Codigo.Text = objRequisitante.lCodigo
            NomeReduzido.Text = objRequisitante.sNomeReduzido
            
        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 49044, 49045
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174078)

    End Select

    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai o Requisitante da tela

Dim lErro As Long
Dim objRequisitante As New ClassRequisitante

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Requisitante"

    'le os dados da tela
    lErro = Move_Tela_Memoria(objRequisitante)
    If lErro <> SUCESSO Then gError 49047

    'Preenche a coleção colCampoValor
    colCampoValor.Add "Codigo", objRequisitante.lCodigo, 0, "Codigo"
    colCampoValor.Add "Nome", objRequisitante.sNome, STRING_REQUISITANTE_NOME, "Nome"
    colCampoValor.Add "NomeReduzido", objRequisitante.sNomeReduzido, STRING_REQUISITANTE_NOMERED, "NomeReduzido"
    colCampoValor.Add "Ccl", objRequisitante.sCcl, STRING_CCL, "Ccl"

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 49047

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174079)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objRequisitante As New ClassRequisitante
Dim vbMsgRes As VbMsgBoxResult
Dim lCodigo As Long

On Error GoTo Erro_BotaoExcluir_Click

    'Coloca o MouseIcon de Ampulheta durante a Exclusão
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o codigo foi preenchido
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 49058

    'Se é o Requisitante de código 1 (Requisitante automático), Erro
    If CLng(Codigo.Text) = REQUISITANTE_AUTOMATICO_CODIGO Then gError 67302
    
    objRequisitante.lCodigo = StrParaLong(Codigo.Text)

    lErro = CF("Requisitante_Le", objRequisitante)
    If lErro <> SUCESSO And lErro <> 49084 Then gError 49059

    'Verifica se requisitante não está cadastrado
    If lErro = 49084 Then gError 49060
    
    'Envia aviso perguntando se realmente deseja excluir requisitante
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_REQUISITANTE", objRequisitante.lCodigo)

    If vbMsgRes = vbYes Then

        'Exclui Requisitante
        lErro = CF("Requisitante_Exclui", objRequisitante)
        If lErro <> SUCESSO Then gError 49061

        'Exclui da ListBox
        Call ListaRequisitante_Exclui(objRequisitante.lCodigo)

        'Limpa a Tela
        Call Limpa_Tela_Requisitante_FechaSeta
    
        iAlterado = 0
    
    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr
        
        Case 49058
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 49059, 49061

        Case 49060
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_NAO_CADASTRADO", gErr, objRequisitante.lCodigo)

        Case 67302
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_REQUISITANTE_AUTOMATICO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174080)

    End Select

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    'Fecha a Tela
    Unload Me
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava o Requisitante
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 49050

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 49050

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174081)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
'chamada de Limpa_Tela_Requisitante

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 49042

    'Limpa Tela
    Call Limpa_Tela_Requisitante_FechaSeta

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 49042

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174082)

    End Select

    Exit Sub
    
End Sub

Sub Limpa_Tela_Requisitante()

    'Limpa a tela
    Call Limpa_Tela(Me)
    
    CodUsuario.ListIndex = -1

    iAlterado = 0

End Sub

Sub Limpa_Tela_Requisitante_FechaSeta()

    'Limpa a tela
    Call Limpa_Tela(Me)
    
    'Fecha o comando de Setas
    Call ComandoSeta_Fechar(Me.Name)
    
    CodUsuario.ListIndex = -1
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Busca o próximo código de Requisitante Disponível
    lErro = CF("Requisitante_Automatico", lCodigo)
    If lErro <> SUCESSO Then gError 63823

    'Coloca o código na tela
    Codigo.Text = CStr(lCodigo)
    
    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 63823
            'Erro tratado na rotina chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174083)
    
    End Select

    Exit Sub

End Sub

Private Sub Ccl_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Ccl_Validate(Cancel As Boolean)

Dim objCcl As New ClassCcl
Dim sCcl As String
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim sCclFormatada As String

On Error GoTo Erro_Ccl_Validate

    If Len(Trim(Ccl.ClipText)) = 0 Then Exit Sub

    'Critica o Centro de Custo Informado
    lErro = CF("Ccl_Critica", Ccl.Text, sCclFormatada, objCcl)
    If lErro <> SUCESSO And lErro <> 5703 Then gError 49094

    'Não encontrou o Ccl no BD
    If lErro <> SUCESSO Then gError 49049

    Exit Sub

Erro_Ccl_Validate:

    'Cancela a saída desse campo
    Cancel = True
        
    Select Case gErr

        Case 49049
            'pergunta de deseja criar esse CCL
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", Ccl.Text)
            'Se a resposta for sim chama a tela de cadastro de CCL
            If vbMsgRes = vbYes Then Call Chama_Tela("CclTela", objCcl)
           
        Case 49094

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174084)

    End Select

    Exit Sub

End Sub

Private Sub CclLabel_Click()
'Rotinas das telas de browse
'Browse Ccl

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim colSelecao As New Collection
Dim sCclFormatada As String
Dim iCclPreenchida As Integer

On Error GoTo Erro_CclLabel_Click
    
    'Se o CCL está preenchido
    If Len(Trim(Ccl.ClipText)) > 0 Then
        'Formata o Ccl
        lErro = CF("Ccl_Formata", Ccl.Text, sCclFormatada, iCclPreenchida)
        If lErro <> SUCESSO Then gError 49072
        'Guarda o CCL no objCcl
        If iCclPreenchida = CCL_PREENCHIDA Then objCcl.sCcl = sCclFormatada

    End If

    'Chama a tela CclLista
    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCcl)

    Exit Sub

Erro_CclLabel_Click:

    Select Case gErr

        Case 49072

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174085)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Change()

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
Dim colCodigoNome As New AdmCollCodigoNome
Dim objCodigoNome As AdmlCodigoNome

On Error GoTo Erro_Form_Load

    Set objEventoCcl = New AdmEvento

    'Le o codigo e o nome reduzido de todos os requisitantes
    lErro = CF("LCod_Nomes_Le", "Requisitante", "Codigo", "NomeReduzido", STRING_REQUISITANTE_NOMERED, colCodigoNome)
    If lErro <> SUCESSO Then gError 49040

    For Each objCodigoNome In colCodigoNome

        'Insere na listbox de requisitantes
        Requisitantes.AddItem objCodigoNome.sNome
        Requisitantes.ItemData(Requisitantes.NewIndex) = objCodigoNome.lCodigo

    Next
    
    'Carrega a combobox todos os usuários
    lErro = Carrega_Usuarios()
    If lErro <> SUCESSO Then gError 49097
    
    lErro = Inicializa_MascaraCcl()
    If lErro <> SUCESSO Then gError 49097
    
    Email.MaxLength = STRING_EMAIL

    iAlterado = 0
   
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 49040, 49097

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174086)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoCcl = Nothing

    'Fecha o comando de setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Private Sub Codigo_GotFocus()
    
    'Faz o cursor ir para o início do campo
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
    
End Sub

Private Sub Nome_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NomeReduzido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoCcl_evSelecao(obj1 As Object)
'traz o ccl selecionado para a tela

Dim lErro As Long
Dim objCcl As ClassCcl
Dim sCclMascarado As String

On Error GoTo Erro_objEventoCcl_evSelecao

    Set objCcl = obj1
    
    'Coloca o CCl no formato de exibição
    lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
    If lErro <> SUCESSO Then gError 49073
    
    'Coloca o CCL na tela
    Ccl.PromptInclude = False
    Ccl.Text = sCclMascarado
    Ccl.PromptInclude = True

    'COloca esse forma acima do de browse
    Me.Show

    Exit Sub

Erro_objEventoCcl_evSelecao:

    Select Case gErr

        Case 49073

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174087)

    End Select

    Exit Sub

End Sub

Private Sub Requisitantes_DblClick()

Dim lErro As Long
Dim objRequisitante As New ClassRequisitante

On Error GoTo Erro_Requisitantes_DblClick

    'Pega no ItemData o código do Requisitante Selecionado
    objRequisitante.lCodigo = Requisitantes.ItemData(Requisitantes.ListIndex)

    'le o requisitante
    lErro = CF("Requisitante_Le", objRequisitante)
    If lErro <> SUCESSO And lErro <> 49084 Then gError 49095
    If lErro = 49084 Then gError 49145 'Não está cadastrado
    
    'TRaz p\ a tela os dados do requisitante
    lErro = Traz_Requisitante_Tela(objRequisitante)
    If lErro <> SUCESSO Then gError 49057

    'fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Exit Sub

Erro_Requisitantes_DblClick:

    Select Case gErr

        Case 49057, 49095
        
        Case 49145
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_NAO_CADASTRADO", gErr, objRequisitante.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174088)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'Quando uma tecla for pressionada
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    ElseIf KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Ccl Then
            Call CclLabel_Click
        End If
    End If
    
End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Requisitante"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Requisitante"
    
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

Private Sub CclLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CclLabel, Source, X, Y)
End Sub

Private Sub CclLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CclLabel, Button, Shift, X, Y)
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

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub CodUsuario_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CodUsuario_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Email_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Function Carrega_Usuarios() As Long
'Carrega a Combo CodUsuarios com todos os usuários do BD

Dim lErro As Long, colUsuarios As New Collection
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Carrega_Usuarios

    CodUsuario.AddItem ""
    
    lErro = CF("UsuariosFilialEmpresa_Le_Todos", colUsuarios)
    If lErro <> SUCESSO Then Error 48100

    For Each objUsuarios In colUsuarios
        CodUsuario.AddItem objUsuarios.sCodUsuario
    Next

    Carrega_Usuarios = SUCESSO

    Exit Function

Erro_Carrega_Usuarios:

    Carrega_Usuarios = Err

    Select Case Err

        Case 48100

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142643)

    End Select

    Exit Function

End Function

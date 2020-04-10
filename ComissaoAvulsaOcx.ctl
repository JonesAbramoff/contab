VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ComissaoAvulsaOcx 
   ClientHeight    =   5595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5385
   LockControls    =   -1  'True
   ScaleHeight     =   5595
   ScaleWidth      =   5385
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3165
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   135
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ComissaoAvulsaOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "ComissaoAvulsaOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "ComissaoAvulsaOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "ComissaoAvulsaOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame ComissaoAvulsaOcxComissaoAvulsaOcx 
      Caption         =   "Identificação"
      Height          =   2340
      Left            =   120
      TabIndex        =   17
      Top             =   735
      Width           =   5190
      Begin VB.CommandButton ComissoesAvulsas 
         Caption         =   "Comissões Avulsas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3045
         TabIndex        =   12
         Top             =   1815
         Width           =   2040
      End
      Begin VB.TextBox Referencia 
         Height          =   315
         Left            =   1185
         TabIndex        =   3
         Top             =   1350
         Width           =   3900
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   330
         Left            =   2355
         TabIndex        =   18
         Top             =   870
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   885
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Vendedor 
         Height          =   330
         Left            =   1200
         TabIndex        =   1
         Top             =   420
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   582
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Status 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   3915
         TabIndex        =   24
         Top             =   405
         Width           =   1170
      End
      Begin VB.Label Label1 
         Caption         =   "Status:"
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
         Left            =   3285
         TabIndex        =   23
         Top             =   450
         Width           =   675
      End
      Begin VB.Label LabelVendedor 
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
         Height          =   255
         Left            =   225
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   21
         Top             =   480
         Width           =   945
      End
      Begin VB.Label Label2 
         Caption         =   "Data:"
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
         Left            =   645
         TabIndex        =   20
         Top             =   960
         Width           =   555
      End
      Begin VB.Label Label3 
         Caption         =   "Referência:"
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
         Height          =   300
         Left            =   120
         TabIndex        =   19
         Top             =   1380
         Width           =   1020
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comissão"
      Height          =   2370
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   5175
      Begin VB.ComboBox Motivo 
         Height          =   315
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   405
         Width           =   2295
      End
      Begin MSMask.MaskEdBox ValorComissao 
         Height          =   345
         Left            =   1635
         TabIndex        =   7
         Top             =   1845
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   609
         _Version        =   393216
         Format          =   "Standard"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Aliquota 
         Height          =   345
         Left            =   1635
         TabIndex        =   6
         Top             =   1350
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   609
         _Version        =   393216
         Format          =   "#0.#0\%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox BaseCalculo 
         Height          =   330
         Left            =   1635
         TabIndex        =   5
         Top             =   870
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   582
         _Version        =   393216
         Format          =   "Standard"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Caption         =   "Motivo:"
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
         Height          =   270
         Left            =   900
         TabIndex        =   16
         Top             =   465
         Width           =   675
      End
      Begin VB.Label Label5 
         Caption         =   "Base de Cálculo:"
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
         Height          =   270
         Left            =   105
         TabIndex        =   15
         Top             =   930
         Width           =   1500
      End
      Begin VB.Label Label6 
         Caption         =   "Alíquota:"
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
         Height          =   285
         Left            =   705
         TabIndex        =   14
         Top             =   1395
         Width           =   885
      End
      Begin VB.Label Label7 
         Caption         =   "Valor Comissão:"
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
         Left            =   180
         TabIndex        =   13
         Top             =   1875
         Width           =   1455
      End
   End
End
Attribute VB_Name = "ComissaoAvulsaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

'Variaveis para sistema de Browser
Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1
Private WithEvents objEventoComissoesAvulsas As AdmEvento
Attribute objEventoComissoesAvulsas.VB_VarHelpID = -1

'?? Subir Constantes
Const MOTIVOSCOMISSAO_DESCRICAO = 50
Const STRING_COMISSAOAVULSA_REFERENCIA = 50

'?? Subir Type (s)
Private Type TypeComissaAvulsa
    lNumIntDoc As Long
    iVendedor As Integer
    dtData As Date
    sReferencia As String
    iMotivo As Integer
    dBaseCalculo As Double
    dAliquota As Double
    dValorComissao As Double
End Type

Private Type typeMotivoComissoes
    iCodigo As Integer
    sDescricao As String
    iDeducao As Integer
End Type

'??Subir Tipo de Comissão avulsa que será inserida na tabela de comissões
Const TIPO_COMISSAO_AVULSA = 4

'??Subir Constante cadastrada na tabela CPRConfig
Const NUM_PROX_COMISSAO_AVULSA = 0

Private Sub Aliquota_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dValorTotal As Double, dBaseCalculo As Double, dAliquota As Double

On Error GoTo Erro_Aliquota_Validate

    'Verifica se aliquota foi preenchida
    If Len(Trim(Aliquota.Text)) = 0 Then Exit Sub

    'Critica valor da aliquota
    lErro = Porcentagem_Critica(Aliquota.Text)
    If lErro <> SUCESSO Then gError 87596
    
    'Recolhe valores da tela
    dValorTotal = StrParaDbl(ValorComissao.Text)
    dBaseCalculo = StrParaDbl(BaseCalculo.Text)
    dAliquota = StrParaDbl(Aliquota.Text) / 100

    'Verifica se a base de cálculo foi preenchida
    If Len(Trim(BaseCalculo.Text)) > 0 Then
            
        'Calcula o Valor total da Comissão com base na Alíquota e Base de Cálculo
        lErro = ComissaoAvulsa_CalculaTotal(dBaseCalculo, dAliquota, dValorTotal)
        If lErro <> SUCESSO Then gError 87673
        
        'Se a variavel não for nula ---> preenche ValorComissao
        If dValorTotal > 0 Then
            ValorComissao.Text = Format(dValorTotal, "Standard")
        End If
    
    Else
        
        'Calcula a BaseCalculo com base na Alíquota e Valor total
        lErro = ComissaoAvulsa_CalculaBase(dBaseCalculo, dAliquota, dValorTotal)
        If lErro <> SUCESSO Then gError 87914
        
        'Se a variável dBaseCalculo não for nula ---> Preenche BaseCalculo
        If dBaseCalculo > 0 Then
            BaseCalculo.Text = Format(dBaseCalculo, "Standard")
        End If
            
    End If
    
    Exit Sub

Erro_Aliquota_Validate:

    Cancel = True

    Select Case gErr

        Case 87596, 87673, 87914
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154338)

    End Select

    Exit Sub

End Sub

Private Sub BaseCalculo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dValorTotal As Double, dBaseCalculo As Double, dAliquota As Double

On Error GoTo Erro_BaseCalculo_Validate

    'Verifica se BaseCalculo foi preenchido
    If Len(Trim(BaseCalculo.Text)) = 0 Then Exit Sub

    'Critica valor de BaseCalculo
    lErro = Valor_Positivo_Critica(BaseCalculo.Text)
    If lErro <> SUCESSO Then gError 87597
    
    'Verifica se a Alíquota foi preenchida
    If Len(Trim(Aliquota.Text)) > 0 Then
    
        'Recolhe valores da tela
        dValorTotal = StrParaDbl(ValorComissao.Text)
        dBaseCalculo = StrParaDbl(BaseCalculo.Text)
        dAliquota = StrParaDbl(Aliquota.Text) / 100
    
        'Calcula o Valor total da Comissão com base na Alíquota e Base de Cálculo
        lErro = ComissaoAvulsa_CalculaTotal(dBaseCalculo, dAliquota, dValorTotal)
        If lErro <> SUCESSO Then gError 87674
        
        'Se vaiável dValorTotal não for nula ---> Preenche ValorComissao
        If dValorTotal > 0 Then
            ValorComissao.Text = Format(dValorTotal, "Standard")
        End If
    
    End If

    Exit Sub

Erro_BaseCalculo_Validate:

    Cancel = True

    Select Case gErr

        Case 87597, 87674
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154339)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objComissoesAvulsas As New ClassComissoesAvulsas

On Error GoTo Erro_BotaoExcluir_Click

    'Verifica se chaves secundárias da tabela estão preenchidas
    If Len(Trim(Vendedor.Text)) = 0 Then gError 87612
    If Len(Trim(Data.ClipText)) = 0 Then gError 87613
    If Len(Trim(Referencia.Text)) = 0 Then gError 87614
                        
    'Atribui os valores da tela para o objComissoesAvulsas
    objComissoesAvulsas.iVendedor = Codigo_Extrai(Vendedor.Text)
    objComissoesAvulsas.dtData = StrParaDate(Data.Text)
    objComissoesAvulsas.sReferencia = Referencia.Text
                        
    'Verifica se a comissão avulsa existe
    lErro = ComissoesAvulsas_Le(objComissoesAvulsas)
    If lErro <> SUCESSO And lErro <> 87632 Then gError 87599
    
    'Se não ---> erro
    If lErro = 87632 Then gError 87600
                        
    'Pede confirmação ao usuário para a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_COMISSAOAVULSA")
    If vbMsgRes = vbYes Then
    
        'Exclui a comissão
        lErro = ComissoesAvulsas_Exclui(objComissoesAvulsas)
        If lErro <> SUCESSO Then gError 87601

        'Limpa a tela de ComissoesAvulsas
        Call Limpa_Tela(Me)
        Call LimpaTela_ComissoesAvulsas
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:
    
    Select Case gErr

        Case 87599, 87601
            'Tratado na Rotina
        
        Case 87600
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMISSAOAVULSA_INEXISTENTE", gErr)

        Case 87612
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_FORNECIDO", gErr)

        Case 87613
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)

        Case 87614
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REFERENCIA_NAO_PREENCHIDA", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154340)

    End Select

    Exit Sub

End Sub


Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Critica e grava Comissão
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 87639
    
    'Limpa a tela
    Call Limpa_Tela(Me)
    Call LimpaTela_ComissoesAvulsas
    
    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 87639 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154341)

    End Select

    Exit Sub
  
End Sub

Function Gravar_Registro()

Dim lErro As Long
Dim objComissoesAvulsas As New ClassComissoesAvulsas

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se campos obrigatórios da tela estão preenhidos
    lErro = Valida_Campos()
    If lErro <> SUCESSO Then gError 87602
    
    'Recolhe os campos da tela para o objComissoesAvulsas
    lErro = Move_Tela_Memoria(objComissoesAvulsas)
    If lErro <> SUCESSO Then gError 87603
    
    'Verifica se a Comissão Avulsa já existe
    lErro = Trata_Alteracao(objComissoesAvulsas, objComissoesAvulsas.iVendedor, objComissoesAvulsas.dtData, objComissoesAvulsas.sReferencia)
    If lErro <> SUCESSO Then gError 87677
        
    'Grava a Comissão
    lErro = CF("ComissoesAvulsas_Grava", objComissoesAvulsas)
    If lErro <> SUCESSO Then gError 87604
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    Select Case gErr

        Case 87602, 87603, 87604, 87677
            'Tratados na Rotina
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154342)

    End Select
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

End Function

Function Move_Tela_Memoria(objComissoesAvulsas As ClassComissoesAvulsas) As Long

Dim lErro As Long
Dim iCodigo As Integer
Dim iDeducao As Integer

On Error GoTo Erro_Move_Tela_Memoria

    'Atribui os valores da tela para o objComissoesAvulsas
    objComissoesAvulsas.iVendedor = Codigo_Extrai(Vendedor.Text)
    objComissoesAvulsas.dtData = StrParaDate(Data.Text)
    objComissoesAvulsas.sReferencia = Referencia.Text
    objComissoesAvulsas.iCodigoMotivo = Codigo_Extrai(Motivo.Text)
    objComissoesAvulsas.dBaseCalculo = StrParaDbl(BaseCalculo.Text)
    objComissoesAvulsas.dAliquota = (StrParaDbl(Aliquota.Text)) / 100
    
    'Verifica se existe dedução
        '---> Se existir então a Comissão será gravada com valor negativo
    iCodigo = objComissoesAvulsas.iCodigoMotivo
    lErro = Verifica_ExistenciaDeducao(iCodigo, iDeducao)
    If lErro <> SUCESSO Then gError 87913
    
    If iDeducao = 1 Then
        objComissoesAvulsas.dValorComissao = StrParaDbl(ValorComissao.Text) * (-1)
    Else
        objComissoesAvulsas.dValorComissao = StrParaDbl(ValorComissao.Text)
    End If
    
    Move_Tela_Memoria = SUCESSO

    Exit Function
    
Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
    
    Select Case gErr
    
        Case 87913
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154343)
            
    End Select

    Exit Function
    
End Function

Function Valida_Campos() As Long

Dim lErro As Long

On Error GoTo Erro_Valida_Campos

    '---> Frame Identificacao
    
    'Verifica se vendedor está preenchido
    If Len(Trim(Vendedor.Text)) = 0 Then gError 87605
    
    'Verifica se data está preenchida
    If Len(Trim(Data.ClipText)) = 0 Then gError 87606
    
    'Verifica se referência está preenchida
    If Len(Trim(Referencia.Text)) = 0 Then gError 87607


    '---> Frame Comissão
    
    'Verifica se motivo foi preenchido
    If Motivo.ListIndex = -1 Then gError 87608
    
    'Verifica se BaseCalculo foi preenchida
    If Len(Trim(BaseCalculo.Text)) = 0 Then gError 87609
    
    'Verifica se alíquota foi preenchida
    If Len(Trim(Aliquota.Text)) = 0 Then gError 87610
    
    'Verifica se ValorComissao foi preenchido
    If Len(Trim(ValorComissao.Text)) = 0 Then gError 87611

    Valida_Campos = SUCESSO
    
    Exit Function
    
Erro_Valida_Campos:

    Valida_Campos = gErr
    
    Select Case gErr
    
        Case 87605
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_FORNECIDO", gErr)

        Case 87606
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)

        Case 87607
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REFERENCIA_NAO_PREENCHIDA", gErr)

        Case 87608
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOTIVO_NAO_INFORMADO", gErr)

        Case 87609
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BASECALCULO_NAO_INFORMADA", gErr)

        Case 87610
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALIQUOTA_NAO_PREENCHIDA", gErr)

        Case 87611
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORCOMISSAO_NAO_PREENCHIDA", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154344)

    End Select
    
    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 87615
    
    'Limpa a tela de Comissões Avulsas
    Call LimpaTela_ComissoesAvulsas
    Call Limpa_Tela(Me)

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 87615
            'Erro tratado na rotina chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154345)

    End Select

    Exit Sub

End Sub

Private Sub ComissoesAvulsas_Click()

Dim lErro As Long
Dim objComissaoAvulsa As New ClassComissoesAvulsas
Dim colSelecao As New Collection
Dim sSelecao As String
Dim iPreenchido As Integer

On Error GoTo Erro_ComissoesAvulsas_Click
    
    'Recolhe os dados da tela para o objComissoesAvulsas
    lErro = Move_Tela_Memoria(objComissaoAvulsa)
    If lErro <> SUCESSO Then gError 87882

    'Monta a instrução SQL dinâmicamente
    If objComissaoAvulsa.iVendedor <> 0 Then
        sSelecao = "Vendedor = ?"
        iPreenchido = 1
        colSelecao.Add (objComissaoAvulsa.iVendedor)
    End If
    
    If objComissaoAvulsa.dtData <> DATA_NULA Then
        If iPreenchido = 1 Then
            sSelecao = sSelecao & " AND Data = ?"
        Else
            iPreenchido = 1
            sSelecao = "Data = ?"
        End If
        colSelecao.Add (objComissaoAvulsa.dtData)
    End If
    
    If Len(Trim(objComissaoAvulsa.sReferencia)) <> 0 Then
        If iPreenchido = 1 Then
            sSelecao = sSelecao & " AND Referencia = ?"
        Else
            iPreenchido = 1
            sSelecao = "Referencia = ?"
        End If
        colSelecao.Add (objComissaoAvulsa.sReferencia)
    End If

    'Chama a tela de Browser
    Call Chama_Tela("ComissoesAvulsasLista", colSelecao, objComissaoAvulsa, objEventoComissoesAvulsas, sSelecao)
    
    Exit Sub
    
Erro_ComissoesAvulsas_Click:

    Select Case gErr
                
        Case 87882
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154346)
    
    End Select
    
    Exit Sub


End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    'Verifica se a data foi preenchida
    If Len(Trim(Data.ClipText)) = 0 Then Exit Sub
    
    'Verifica se é uma data válida
    lErro = Data_Critica(Data.Text)
    If lErro <> SUCESSO Then gError 87616
    
    'Verifica se a data informada é maior que a data atual
    If StrParaDate(Data.Text) > gdtDataAtual Then gError 87617

    Exit Sub
    
Erro_Data_Validate:

    Cancel = True

    Select Case gErr

        Case 87616

        Case 87617
             lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INFORMADA_MAIOR_DATA_HOJE", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154347)

    End Select

    Exit Sub

End Sub


Private Sub LabelVendedor_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection

    If Len(Trim(Vendedor.Text)) > 0 Then
        'Preenche com o Vendedor da tela
        objVendedor.iCodigo = Codigo_Extrai(Vendedor.Text)
    End If
    
    'Chama o Browser de vendedores
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub


Private Sub objEventoComissoesAvulsas_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objComissoesAvulsas As ClassComissoesAvulsas
Dim lNumIntDoc As Long
Dim colComissao As New colComissao
Dim iStatus As Integer

On Error GoTo Erro_objEventoComissoesAvulsas_evSelecao

    Set objComissoesAvulsas = obj1

    If Not (objComissoesAvulsas Is Nothing) Then
            
        'Verifica se a comissão existe na tabela de Comissoes Avulsas
        lErro = ComissoesAvulsas_Le(objComissoesAvulsas)
        If lErro <> SUCESSO And lErro <> 87632 Then gError 87883
            
        If lErro = 87362 Then gError 87884
            
        lNumIntDoc = objComissoesAvulsas.lNumIntDoc
        
        'Verifica se a comissão existe na tabela de Comissoes
        lErro = CF("Comissoes_Le", lNumIntDoc, colComissao, TIPO_COMISSAO_AVULSA)
        If lErro <> SUCESSO Then gError 87887
            
        'Verifica se é uma Comissao Baixada
        iStatus = colComissao.Item(1).iStatus
            
        'Preenche a tela com os dados recuperados no BD
        lErro = Preenche_Tela_ComissoesAvulsas(objComissoesAvulsas, iStatus)
        If lErro <> SUCESSO Then gError 87881
    
    End If

    Me.Show

    Exit Sub

Erro_objEventoComissoesAvulsas_evSelecao:

    Select Case gErr

        Case 87881, 87887

        Case 87884
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMISSAOAVULSA_INEXISTENTE", gErr, objComissoesAvulsas)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154348)

    End Select

    Exit Sub

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objVendedor As ClassVendedor

On Error GoTo Erro_objEventoVendedor_evSelecao

    Set objVendedor = obj1

    'Preenche vendedor no campo vendedor
    If Not (objVendedor Is Nothing) Then
        Vendedor.Text = CStr(objVendedor.iCodigo) & SEPARADOR & objVendedor.sNomeReduzido
    End If

    Me.Show

    Exit Sub

Erro_objEventoVendedor_evSelecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154349)

    End Select

    Exit Sub

End Sub

Private Sub ValorComissao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dValorTotal As Double, dBaseCalculo As Double, dAliquota As Double


On Error GoTo Erro_ValorComissao_Validate

    If Len(Trim(ValorComissao.Text)) = 0 Then Exit Sub

    'Verifica se o valor informado é válido
    lErro = Valor_Positivo_Critica(ValorComissao.Text)
    If lErro <> SUCESSO Then gError 87618
    
    'Verifica se a Alíquota foi preenchida
    If Len(Trim(Aliquota.Text)) > 0 Then
    
        'Recolhe os valores da tela
        dValorTotal = StrParaDbl(ValorComissao.Text)
        dBaseCalculo = StrParaDbl(BaseCalculo.Text)
        dAliquota = StrParaDbl(Aliquota.Text) / 100
    
        'Calcula a BaseCalculo com base na Alíquota e Valor total
        lErro = ComissaoAvulsa_CalculaBase(dBaseCalculo, dAliquota, dValorTotal)
        If lErro <> SUCESSO Then gError 87674
        
        'Se a variável dBaseCalculo não for nula ---> Preenche BaseCalculo
        If dBaseCalculo > 0 Then
            BaseCalculo.Text = Format(dBaseCalculo, "Standard")
        End If
    
    End If
    
    Exit Sub

Erro_ValorComissao_Validate:

    Cancel = True

    Select Case gErr

        Case 87618, 87674
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154350)

    End Select

    Exit Sub

End Sub

Private Sub Vendedor_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Vendedor, iAlterado)

End Sub

Private Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCria As Integer
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Vendedor_Validate

    'Verifica se o vendedor foi preenchido
    If Len(Trim(Vendedor.Text)) = 0 Then Exit Sub
    
    'Verifica existencia do vendedor
    lErro = TP_Vendedor_Le2(Vendedor, objVendedor, iCria)
    If lErro <> SUCESSO Then gError 87619
    
    Exit Sub

Erro_Vendedor_Validate:

    Cancel = True

    Select Case gErr

        Case 87619
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154351)

    End Select

    Exit Sub

End Sub


Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    Set objEventoVendedor = New AdmEvento
    Set objEventoComissoesAvulsas = New AdmEvento
             
    'Preenche a Combo Motivos no Carregamento da tela
    lErro = CarregaCombo_MotivoComissao()
    If lErro <> SUCESSO Then gError 87620
        
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
        
        Case 87620
            'Erro tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154352)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 87621

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case 87621
             Data.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154353)

    End Select

    Exit Sub


End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 87623
    
    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case 87623
            Data.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154354)

    End Select

    Exit Sub

End Sub

Private Sub Data_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Data_GotFocus()

    Call MaskEdBox_TrataGotFocus(Data, iAlterado)
    
End Sub

Private Sub Vendedor_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Referencia_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Motivo_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub BaseCalculo_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Aliquota_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ValorComissao_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub BaseCalculo_GotFocus()

    Call MaskEdBox_TrataGotFocus(BaseCalculo, iAlterado)
    
End Sub

Private Sub Aliquota_GotFocus()

    Call MaskEdBox_TrataGotFocus(Aliquota, iAlterado)
    
End Sub

Private Sub ValorComissao_GotFocus()

    Call MaskEdBox_TrataGotFocus(ValorComissao, iAlterado)
    
End Sub


'??Subir Função
Public Function ComissoesAvulsas_Le(objComissoesAvulsas As ClassComissoesAvulsas) As Long
'Lê tabela de ComissoesAvulsas

Dim lErro As Long
Dim lComando As Long
Dim tComissaoAvulsa As TypeComissaAvulsa

On Error GoTo Erro_ComissoesAvulsas_Le

    'Abre Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 87629
    
    With tComissaoAvulsa
    
        'Faz leitura nos campos tabela
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, CodigoMotivo, BaseCalculo, Aliquota, ValorComissao FROM ComissoesAvulsas WHERE Vendedor = ? AND Data = ? AND Referencia = ?" _
        , .lNumIntDoc, .iMotivo, .dBaseCalculo, .dAliquota, .dValorComissao, objComissoesAvulsas.iVendedor, objComissoesAvulsas.dtData, objComissoesAvulsas.sReferencia)
        If lErro <> AD_SQL_SUCESSO Then gError 87630
        
        'Busca o Primeiro registro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 87631
        
        'Se não existirem registros ---> erro
        If lErro = AD_SQL_SEM_DADOS Then gError 87632
        
        'Atribui os valores retornados da leitura para o objComissoesAvulsas
        objComissoesAvulsas.lNumIntDoc = .lNumIntDoc
        objComissoesAvulsas.iCodigoMotivo = .iMotivo
        objComissoesAvulsas.dBaseCalculo = .dBaseCalculo
        objComissoesAvulsas.dAliquota = .dAliquota
        objComissoesAvulsas.dValorComissao = .dValorComissao
    
    End With
    
    'Fecha Comando
    Call Comando_Fechar(lComando)
    
    ComissoesAvulsas_Le = SUCESSO
    
    Exit Function
    
Erro_ComissoesAvulsas_Le:

    ComissoesAvulsas_Le = gErr
    
    Select Case gErr
    
        Case 87629
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 87630, 87631
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMISSOESAVULSAS", gErr)
            
        Case 87632
            'Tratado na Rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154355)
            
    End Select
    
    'Fecha Comando ---> saída por erro
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

'?? Subir Função
Public Function ComissoesAvulsas_Exclui(objComissoesAvulsas As ClassComissoesAvulsas) As Long
'Excui ComissaoAvulsa
'ATENÇÃO: Se Comissao Avulsa estiver com data baixa <> DATA_NULA na tabela de Comissões, não será possível
'efetuar a exclusão

Dim lErro As Long
Dim lTransacao As Long
Dim lNumIntDoc As Long
Dim dtDataBaixa As Date
Dim alComando(1 To 4) As Long
Dim iIndice As Integer

On Error GoTo Erro_ComissoesAvulsas_Exclui

    'Abre Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 87633
    Next
        
    'Abre transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 87634
        
    'Lê tabela de Comissões Avulsas
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM ComissoesAvulsas WHERE Vendedor = ? AND Data = ? AND Referencia = ?" _
    , 0, lNumIntDoc, objComissoesAvulsas.iVendedor, objComissoesAvulsas.dtData, objComissoesAvulsas.sReferencia)
    If lErro <> AD_SQL_SUCESSO Then gError 87635

    'Seleciona a comissão
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 87636

    'Se não Encontrou erro
    If lErro = AD_SQL_SEM_DADOS Then gError 87637
    
    'Lock na tabela de ComissoesAvulsas
    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 87658

    'Verifica se a comissão está gravada na tabela de comissões
    lErro = Comando_ExecutarPos(alComando(2), "SELECT DataBaixa FROM Comissoes WHERE NumIntDoc = ? AND TipoTitulo = ?", 0, dtDataBaixa, lNumIntDoc, TIPO_COMISSAO_AVULSA)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 87659
    
    'Busca o registro
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 87660
        
    'Se não Encontrou erro
    If lErro = AD_SQL_SEM_DADOS Then gError 87661
    
    'Se a comissão já estiver baixada não pode excluir
    If dtDataBaixa <> DATA_NULA Then gError 87662
    
    'Exclui a comissão na tabela de comissões avulsas
    lErro = Comando_ExecutarPos(alComando(3), "DELETE FROM ComissoesAvulsas", alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 87638

    'Exclui a comissão na tabela de comissões
    lErro = Comando_ExecutarPos(alComando(4), "DELETE FROM Comissoes", alComando(2))
    If lErro <> AD_SQL_SUCESSO Then gError 87663

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 87664

    'Fecha Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    ComissoesAvulsas_Exclui = SUCESSO
    
    Exit Function

Erro_ComissoesAvulsas_Exclui:

    ComissoesAvulsas_Exclui = gErr
    
    Select Case gErr
    
        Case 87633
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 87634
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
                
        Case 87635, 87636
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMISSOESAVULSAS", gErr)

        Case 87637, 87661
            'Erro tratado na rotina chamadora
        
        Case 87638
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_COMISSOESAVULSAS", gErr)
                
        Case 87658
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_COMISSOESAVULSAS", gErr)
        
        Case 87659, 87660
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMISSOES", gErr)
        
        Case 87662
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMISSAO_BAIXADA_EXCLUI", gErr)
        
        Case 87663
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_COMISSOES", gErr)
        
        Case 87664
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154356)
            
    End Select
    
    'Desfaz Transação
    Call Transacao_Rollback
    
    'Fecha Comandos ---> saída por erro
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function Trata_Parametros(Optional objComissoesAvulsas As ClassComissoesAvulsas, Optional objVendedor As ClassVendedor) As Long

Dim lErro As Long
Dim iCria As Integer

On Error GoTo Erro_Trata_Parametros

    'Se existir uma Comissao passada como parametro, exibir seus dados
    If Not (objComissoesAvulsas Is Nothing) Then

        'Verifica se a comissão existe
        lErro = ComissoesAvulsas_Le(objComissoesAvulsas)
        If lErro <> SUCESSO And lErro <> 87632 Then gError 87885
            
        'Se não Existir ---> Erro
        If lErro = 87362 Then gError 87886
        
        'Preenche a tela com os dados retornados do BD
        lErro = Preenche_Tela_ComissoesAvulsas(objComissoesAvulsas)
        If lErro <> SUCESSO Then gError 87678
        
    End If
                
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 87678 'Tratado na rotina chamada

        Case 87886
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMISSAOAVULSA_INEXISTENTE", gErr, objComissoesAvulsas)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154357)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objComissoesAvulsas As New ClassComissoesAvulsas

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "ComissoesAvulsas"

    'Le os dados da Tela Comissões Avulsas
    lErro = Le_Dados_ComissoesAvulsas(objComissoesAvulsas)
    If lErro <> SUCESSO Then gError 87640

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Vendedor", objComissoesAvulsas.iVendedor, 0, "Vendedor"
    colCampoValor.Add "Data", objComissoesAvulsas.dtData, 0, "Data"
    colCampoValor.Add "Referencia", objComissoesAvulsas.sReferencia, STRING_COMISSAOAVULSA_REFERENCIA, "Referencia"
    colCampoValor.Add "CodigoMotivo", objComissoesAvulsas.iCodigoMotivo, 0, "CodigoMotivo"
    colCampoValor.Add "BaseCalculo", objComissoesAvulsas.dBaseCalculo, 0, "BaseCalculo"
    colCampoValor.Add "Aliquota", objComissoesAvulsas.dAliquota, 0, "Aliquota"
    colCampoValor.Add "ValorComissao", objComissoesAvulsas.dValorComissao, 0, "ValorComissao"
        
    iAlterado = 0
        
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 87641

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154358)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objComissoesAvulsas As New ClassComissoesAvulsas
Dim lNumIntDoc As Long
Dim colComissao As New colComissao
Dim iStatus As Integer

On Error GoTo Erro_Tela_Preenche

    'Atribui ao objComissoesAvulsas o valor existente nas coleções
    objComissoesAvulsas.iVendedor = colCampoValor.Item("Vendedor").vValor
    objComissoesAvulsas.dtData = colCampoValor.Item("Data").vValor
    objComissoesAvulsas.sReferencia = colCampoValor.Item("Referencia").vValor

    If objComissoesAvulsas.iVendedor <> 0 And objComissoesAvulsas.dtData <> 0 And Len(Trim(objComissoesAvulsas.sReferencia)) <> 0 Then

        'Verifica se a comissao existe na tabela de Comissões Avulsas
        lErro = ComissoesAvulsas_Le(objComissoesAvulsas)
        If lErro <> SUCESSO And lErro <> 87632 Then gError 87641

        'Se não ---> dispara erro
        If lErro = 87632 Then gError 87642
        
        lNumIntDoc = objComissoesAvulsas.lNumIntDoc
        
        'Verifica se a comissao existe na tabela de Comissões Avulsas
        lErro = CF("Comissoes_Le", lNumIntDoc, colComissao, TIPO_COMISSAO_AVULSA)
        If lErro <> SUCESSO Then gError 87887
        
        'Verifica se a Comissão esta Baixada
        iStatus = colComissao.Item(1).iStatus
        
        'Preenche a tela com os dados retornados do BD
        lErro = Preenche_Tela_ComissoesAvulsas(objComissoesAvulsas, iStatus)
        If lErro <> SUCESSO Then gError 87643
                
    End If

    iAlterado = 0

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 87641, 87643 'Tratados nas rotinas chamadas

        Case 87642
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMISSAOAVULSA_INEXISTENTE", gErr, objComissoesAvulsas)
                  
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154359)

    End Select

    Exit Sub
        
End Sub

Function Le_Dados_ComissoesAvulsas(objComissoesAvulsas As ClassComissoesAvulsas) As Long
'Le os dados da tela

Dim lErro As Long

On Error GoTo Erro_Le_Dados_ComissoesAvulsas

'---> Será recolhido da tela apenas os campos que estiverem preenchidos

    If Len(Trim(Vendedor.Text)) > 0 Then
        objComissoesAvulsas.iVendedor = Codigo_Extrai(Vendedor.Text)
    End If

    If Len(Trim(Data.ClipText)) > 0 Then
        objComissoesAvulsas.dtData = StrParaDate(Data.Text)
    End If

    If Len(Trim(Referencia.Text)) > 0 Then
        objComissoesAvulsas.sReferencia = Referencia.Text
    End If

    If Motivo.ListIndex = -1 Then
        objComissoesAvulsas.iCodigoMotivo = Codigo_Extrai(Motivo.Text)
    End If

    If Len(Trim(BaseCalculo.Text)) > 0 Then
        objComissoesAvulsas.dBaseCalculo = StrParaDbl(BaseCalculo.Text)
    End If

    If Len(Trim(Aliquota.Text)) > 0 Then
        objComissoesAvulsas.dAliquota = StrParaDbl(Aliquota.Text)
    End If

    If Len(Trim(ValorComissao.Text)) > 0 Then
        objComissoesAvulsas.dValorComissao = StrParaDbl(ValorComissao.Text)
    End If

    Le_Dados_ComissoesAvulsas = SUCESSO

    Exit Function
    
Erro_Le_Dados_ComissoesAvulsas:

    Le_Dados_ComissoesAvulsas = gErr
    
    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154360)
            
    End Select

    Exit Function

End Function


Function Preenche_Tela_ComissoesAvulsas(objComissoesAvulsas As ClassComissoesAvulsas, Optional iStatus As Integer) As Long
'Preenche a tela de Comissões Avulsas

Dim lErro As Long
Dim iIndice As Integer
Dim iCodigo As Integer

On Error GoTo Erro_Preenche_Tela_ComissoesAvulsas

    'Preenche o Vendedor e executa o validate
    Vendedor.Text = objComissoesAvulsas.iVendedor
    Call Vendedor_Validate(bSGECancelDummy)
    
    'Preenche a data
    Call DateParaMasked(Data, objComissoesAvulsas.dtData)
    
    'Preenche a referência
    Referencia.Text = objComissoesAvulsas.sReferencia
    
    'Preenche a Combo (Seleciona o ítem pois não é editavel)
    For iIndice = 0 To Motivo.ListCount - 1
        iCodigo = Codigo_Extrai(Motivo.List(iIndice))
        If iCodigo = objComissoesAvulsas.iCodigoMotivo Then
            Motivo.ListIndex = iIndice
            Exit For
        End If
    Next
    
    'Preenche a base de Calculo
    BaseCalculo.Text = objComissoesAvulsas.dBaseCalculo
    
    'Preenche a aliquota (Campo com formato %)
    Aliquota.Text = (objComissoesAvulsas.dAliquota * 100)
    ValorComissao.Text = Abs(objComissoesAvulsas.dValorComissao)
    
    'Preenche o Status ---> Baixado ou Aberto
    If iStatus = STATUS_BAIXADO Then
        Status.Caption = "BAIXADO"
    Else
        Status.Caption = "ABERTO"
    End If
    
    Preenche_Tela_ComissoesAvulsas = SUCESSO
    
    iAlterado = 0
    
    Exit Function
    
Erro_Preenche_Tela_ComissoesAvulsas:

    Preenche_Tela_ComissoesAvulsas = gErr
    
    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154361)
    
    End Select
    
    Exit Function

End Function

Public Sub LimpaTela_ComissoesAvulsas()

    Motivo.ListIndex = -1
    Status.Caption = ""
    
End Sub

Public Function ComissaoAvulsa_CalculaTotal(dBaseCalculo As Double, dAliquota As Double, dValorTotal As Double) As Long
'Calcula o Valor total da Comissão com base na Alíquota e Base de Cálculo

Dim lErro As Long

On Error GoTo Erro_ComissaoAvulsa_CalculaTotal

    If dBaseCalculo = 0 And dAliquota = 0 And dValorTotal = 0 Then Exit Function
    
    If dBaseCalculo <> 0 And dAliquota <> 0 Then
        dValorTotal = dBaseCalculo * dAliquota
    End If
    
    ComissaoAvulsa_CalculaTotal = SUCESSO
    
    Exit Function
    
Erro_ComissaoAvulsa_CalculaTotal:

    ComissaoAvulsa_CalculaTotal = gErr
    
    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154362)
            
    End Select
        
    Exit Function

End Function

Public Function ComissaoAvulsa_CalculaBase(dBaseCalculo As Double, dAliquota As Double, dValorTotal As Double) As Long
'Calcula a BaseCalculo da Comissão com base na Alíquota e Valor Total

Dim lErro As Long

On Error GoTo Erro_ComissaoAvulsa_CalculaBase

    If dValorTotal <> 0 And dAliquota <> 0 Then
        dBaseCalculo = dValorTotal / dAliquota
    End If

    ComissaoAvulsa_CalculaBase = SUCESSO
    
    Exit Function
    
Erro_ComissaoAvulsa_CalculaBase:

    ComissaoAvulsa_CalculaBase = gErr
    
    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154363)
            
    End Select
        
    Exit Function

End Function

'?? Subir Função
Public Function MotivoComissoes_Le_Todos(colMotivoComissoes As Collection) As Long
'Carrega todos os registros da tabela MotivoComissoes dentro de uma coleção

Dim objMotivoComissoes As ClassMotivoComissoes
Dim lComando As Long
Dim iIndice As Integer
Dim tMotivoComissoes As typeMotivoComissoes
Dim lErro As Long
Dim sDescricao As String

On Error GoTo Erro_MotivoComissoes_Le_Todos

    'Abre Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 87901
            
    With tMotivoComissoes
    
        'Inicia String
        .sDescricao = String(MOTIVOSCOMISSAO_DESCRICAO, 0)
    
        'Le a tabela de MotivoComissoes
        lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao, Deducao FROM MotivosComissao", .iCodigo, .sDescricao, .iDeducao)
        If lErro <> AD_SQL_SUCESSO Then gError 87902
        
        'Busca primeiro registro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 87903
                
        If lErro = AD_SQL_SEM_DADOS Then gError 87915
                
        '---> Enquanto existirem registros
        Do While lErro = AD_SQL_SUCESSO
        
            Set objMotivoComissoes = New ClassMotivoComissoes
        
            'Carrega no obj
            objMotivoComissoes.iCodigo = .iCodigo
            objMotivoComissoes.sDescricao = .sDescricao
            objMotivoComissoes.iDeducao = .iDeducao
            
            'Adiciona a coleção
            colMotivoComissoes.Add objMotivoComissoes
            
            'busca o próximo registro
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 87904
            
        Loop
            
    End With

    'Fecha Comando
    Call Comando_Fechar(lComando)
    
    MotivoComissoes_Le_Todos = SUCESSO
    
    Exit Function

Erro_MotivoComissoes_Le_Todos:

    MotivoComissoes_Le_Todos = gErr
    
    Select Case gErr
    
        Case 87901
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 87902 To 87904
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOTIVOCOMISSOES", gErr)
            
        Case 87915
            'Tratado na Rotina Chamadora
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154364)
            
    End Select
    
    'Fecha Comando ---> Saída por erro
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function CarregaCombo_MotivoComissao() As Long
'Carrega a combo de motivos

Dim lErro As Long
Dim colMotivoComissoes As New Collection
Dim objMotivoComissoes As ClassMotivoComissoes
Dim iIndice As Integer


On Error GoTo Erro_CarregaCombo_MotivoComissao

    'Le todos os motivo cadastrados na tabela
    lErro = MotivoComissoes_Le_Todos(colMotivoComissoes)
    If lErro <> SUCESSO And lErro <> 87915 Then gError 87905

    If lErro = 87915 Then gError 87916

    'Preencha a ComboMotivo
    For Each objMotivoComissoes In colMotivoComissoes
        
        Motivo.AddItem objMotivoComissoes.iCodigo & SEPARADOR & objMotivoComissoes.sDescricao
        Motivo.ItemData(Motivo.NewIndex) = objMotivoComissoes.iDeducao
            
    Next
            
    CarregaCombo_MotivoComissao = SUCESSO
    
    Exit Function
    
Erro_CarregaCombo_MotivoComissao:

    CarregaCombo_MotivoComissao = gErr
    
    Select Case gErr
    
        Case 87905
    
        Case 87916
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOTIVOCOMISSAO_SEM_DADOS", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154365)
            
    End Select
    
    Exit Function

End Function

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoComissoesAvulsas = Nothing
    Set objEventoVendedor = Nothing

    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub


'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Comissões Avulsas"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ComissaoAvulsa"
    
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

Private Sub LabelVendedor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelVendedor, Source, X, Y)
End Sub

Private Sub LabelVendedor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelVendedor, Button, Shift, X, Y)
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

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub



Function Verifica_ExistenciaDeducao(iCodigo As Integer, iDeducao As Integer) As Long
'Verifica se existe dedução na Comissao do Vendedor

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Verifica_ExistenciaDeducao

    'Loop em todos os itens da Combo Motivo
    For iIndice = 0 To Motivo.ListCount - 1
    
        'Se não houverem itens selecionados sai do for
        If iCodigo = 0 Then Exit For
        
        If iCodigo = Codigo_Extrai(Motivo.List(iIndice)) Then
            If Motivo.ItemData(iIndice) = 1 Then
                iDeducao = 1
                Exit For
            End If
        End If
    Next

    Verifica_ExistenciaDeducao = SUCESSO

    Exit Function
    
Erro_Verifica_ExistenciaDeducao:

    Verifica_ExistenciaDeducao = gErr
    
    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154366)
            
    End Select
    
    Exit Function

End Function

VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl GrafFaturamentoMensalOcx 
   ClientHeight    =   3672
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5148
   ScaleHeight     =   3672
   ScaleWidth      =   5148
   Begin VB.CheckBox Devolucoes 
      Caption         =   "Inclui Devoluções"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   252
      TabIndex        =   13
      Top             =   2772
      Width           =   4125
   End
   Begin VB.Frame Frame2 
      Caption         =   "Período"
      Height          =   2655
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   4575
      Begin VB.Frame Frame1 
         Caption         =   "Final"
         Height          =   840
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Width           =   3855
         Begin MSMask.MaskEdBox MesFinal 
            Height          =   315
            Left            =   840
            TabIndex        =   2
            Top             =   360
            Width           =   975
            _ExtentX        =   1715
            _ExtentY        =   550
            _Version        =   393216
            MaxLength       =   2
            Mask            =   "##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox AnoFinal 
            Height          =   315
            Left            =   2520
            TabIndex        =   3
            Top             =   360
            Width           =   975
            _ExtentX        =   1715
            _ExtentY        =   550
            _Version        =   393216
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            Caption         =   "Mês:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   360
            TabIndex        =   12
            Top             =   420
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Ano:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   2040
            TabIndex        =   11
            Top             =   420
            Width           =   450
         End
      End
      Begin VB.Frame FrameData 
         Caption         =   "Inicial"
         Height          =   840
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   3855
         Begin MSMask.MaskEdBox MesInicial 
            Height          =   315
            Left            =   840
            TabIndex        =   0
            Top             =   360
            Width           =   975
            _ExtentX        =   1715
            _ExtentY        =   550
            _Version        =   393216
            MaxLength       =   2
            Mask            =   "##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox AnoInicial 
            Height          =   315
            Left            =   2520
            TabIndex        =   1
            Top             =   360
            Width           =   975
            _ExtentX        =   1715
            _ExtentY        =   550
            _Version        =   393216
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin VB.Label PFim 
            Caption         =   "Ano:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   2040
            TabIndex        =   9
            Top             =   420
            Width           =   450
         End
         Begin VB.Label PIni 
            Caption         =   "Mês:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   360
            TabIndex        =   8
            Top             =   420
            Width           =   495
         End
      End
   End
   Begin VB.CommandButton BotaoFechar 
      Height          =   360
      Left            =   2640
      Picture         =   "GrafFaturamentoMensalOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Fechar"
      Top             =   3156
      Width           =   1305
   End
   Begin VB.CommandButton BotaoGrafico 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   720
      Picture         =   "GrafFaturamentoMensalOcx.ctx":017E
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Gerar Gráfico"
      Top             =   3156
      Width           =   1305
   End
End
Attribute VB_Name = "GrafFaturamentoMensalOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim iAlterado As Integer


'Property Variables:
Dim m_Caption As String
Event Unload()

Public Sub Form_Load()
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
End Sub

Public Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

Private Sub BotaoGrafico_Click()

Dim lErro As Long
Dim objGrafico As New ClassGrafico
Dim objItemGrafico As New ClassItemGrafico
Dim dTotal As Double
Dim iIndice As Integer
Dim iColunaSemValor As Integer
Dim sPeriodo As String

On Error GoTo Erro_BotaoGrafico_Click
    
    'Verifica se o Mes Inicial foi preenchido
    If Len(Trim(MesInicial.ClipText)) = 0 Then gError 90383
    
    'Verifica se o Ano Inicial foi preenchido
    If Len(Trim(AnoInicial.ClipText)) = 0 Then gError 90384
    
    'Verifica se o Mes Final preenchido
    If Len(Trim(MesFinal.ClipText)) = 0 Then gError 90385
    
    'Verifica se o Ano Inicial foi preenchido
    If Len(Trim(AnoFinal.ClipText)) = 0 Then gError 90386
    
    'Verifica se o Ano Inicial é maior que o Final
    If AnoInicial.Text > AnoFinal.Text Then gError 90387
    
    'Verifica se o Mes Inicial é maior que o Final, se os anos forem iguais
    If AnoInicial.Text = AnoFinal.Text Then If MesInicial > MesFinal Then gError 90388
    
    'Verifica se o Mes Inicial é igual ao Final, se os anos forem iguais
    If AnoInicial.Text = AnoFinal.Text Then If MesInicial = MesFinal Then gError 91934
        
    lErro = Grafico_Le_FaturamentoMensal(objGrafico, giFilialEmpresa, MesInicial.Text, AnoInicial.Text, MesFinal.Text, AnoFinal.Text)
    If lErro <> SUCESSO Then gError 90054
    
    If objGrafico.colcolItensGrafico.Count = 0 Then gError 90392
    
    iColunaSemValor = 0
    
    dTotal = 0
    For iIndice = 1 To objGrafico.colcolItensGrafico.Count
        For Each objItemGrafico In objGrafico.colcolItensGrafico(iIndice)
            dTotal = dTotal + objItemGrafico.dValorColuna
            If objItemGrafico.dValorColuna = 0 Then iColunaSemValor = iColunaSemValor + 1
        
        Next
    Next
    
    objGrafico.FootNote = "Valores em Real" & "TOTAL: " & Format(dTotal, "Standard")
    
    If (objGrafico.colcolItensGrafico.Count - iColunaSemValor) <> 0 Then
        dTotal = dTotal / (objGrafico.colcolItensGrafico.Count - iColunaSemValor)
    Else
        dTotal = 0
    End If
    
    For iIndice = 1 To objGrafico.colcolItensGrafico.Count
        Set objItemGrafico = New ClassItemGrafico
        objItemGrafico.dValorColuna = Format(dTotal, "STANDARD")
        objItemGrafico.LegendText = "Media"
        objItemGrafico.sNomeColuna = objGrafico.colcolItensGrafico(iIndice).Item(1).sNomeColuna
        objGrafico.colcolItensGrafico(iIndice).Add objItemGrafico
    Next
    
    objGrafico.FootNote = objGrafico.FootNote & "               MEDIA MENSAL: " & Format(dTotal, "standard")
    
    objGrafico.ChartType = 3 'Tipo do Gráfico
    objGrafico.TitleText = ""
    
'    Call Chama_Tela_Nova_Instancia("Grafico", objGrafico)
    
    sPeriodo = MesInicial.Text & "/" & AnoInicial.Text & _
            SEPARADOR & MesFinal.Text & "/" & AnoFinal.Text
        
    lErro = Gera_Grafico(objGrafico, sPeriodo)
    
    
    Unload Me
    
    Exit Sub
    
Erro_BotaoGrafico_Click:
    
    Select Case gErr

        Case 90383
            Call Rotina_Erro(vbOKOnly, "ERRO_MES_NAO_PREENCHIDO", gErr)
            MesInicial.SetFocus
        
        Case 90384
            Call Rotina_Erro(vbOKOnly, "ERRO_ANO_NAO_PREENCHIDO", gErr)
            AnoInicial.SetFocus
        
        Case 90385
            Call Rotina_Erro(vbOKOnly, "ERRO_MES_NAO_PREENCHIDO", gErr)
            MesFinal.SetFocus
        
        Case 90386
            Call Rotina_Erro(vbOKOnly, "ERRO_ANO_NAO_PREENCHIDO", gErr)
            AnoFinal.SetFocus
        
        Case 90387
            Call Rotina_Erro(vbOKOnly, "ERRO_ANOINIC_MAIOR_ANOFINAL", gErr)
        
        Case 90388
            Call Rotina_Erro(vbOKOnly, "ERRO_MESINIC_MAIOR_MESFINAL", gErr)
        
        Case 90392
            Call Rotina_Erro(vbOKOnly, "ERRO_SALDOPERIODO_NAO_ENCONTRADO", gErr)
                
        Case 91934
            Call Rotina_Erro(vbOKOnly, "ERRO_MESINIC_IGUAL_MESFINAL", gErr)
                
        Case 90054
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161697)
            
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    
    Unload Me

End Sub

Private Sub MesInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MesInicial_Validate

    If Len(MesInicial.ClipText) > 0 Then
        
        If StrParaInt(MesInicial.Text) > 12 Then gError 90389
        
    End If

    Exit Sub

Erro_MesInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 90389
            Call Rotina_Erro(vbOKOnly, "ERRO_MES_INVALIDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161698)

    End Select

    Exit Sub

End Sub

Private Sub AnoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AnoInicial_Validate

    If Len(AnoInicial.ClipText) > 0 Then
        
        If StrParaInt(AnoInicial.Text) < 4 Then gError 90393
        
    End If

    Exit Sub

Erro_AnoInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 90393
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_ANO_INVALIDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161699)

    End Select

    Exit Sub

End Sub

Private Sub MesFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MesFinal_Validate

    If Len(MesFinal.ClipText) > 0 Then
        
        If StrParaInt(MesFinal.Text) > 12 Then gError 90390
    
    End If

    Exit Sub

Erro_MesFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 90390
            Call Rotina_Erro(vbOKOnly, "ERRO_MES_INVALIDO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161700)

    End Select

    Exit Sub

End Sub

Private Sub AnoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AnoFinal_Validate

    If Len(AnoFinal.ClipText) > 0 Then
        
        If StrParaInt(AnoFinal.Text) < 4 Then gError 90394
    
    End If

    Exit Sub

Erro_AnoFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 90394
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_ANO_INVALIDO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161701)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Faturamento-Comparativo Mensal"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "GrafFaturamentoMensal"
    
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

Private Sub PIni_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(PIni, Source, X, Y)
End Sub

Private Sub PIni_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(PIni, Button, Shift, X, Y)
End Sub

Private Sub PFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(PFim, Source, X, Y)
End Sub

Private Sub PFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(PFim, Button, Shift, X, Y)
End Sub

Private Function Grafico_Le_FaturamentoMensal(objGrafico As ClassGrafico, iFilialEmpresa As Integer, iMesIni As Integer, iAnoIni As Integer, iMesFim As Integer, iAnoFim As Integer) As Long
'Lê o Faturamento Mensal da Filial Empresa passada (ou empresa toda) de acordo com o ano passado por parâmetro.

Dim lErro As Long
Dim lComando As Long
Dim objItemGrafico As ClassItemGrafico
Dim iCont As Integer
Dim dSoma(1 To 12) As Double
Dim sMes(1 To 12) As String
Dim colItensGrafico As Collection

Dim iMesInicial As Integer
Dim iMesFinal As Integer
Dim iAnoLido As Integer
Dim sFiltro As String

On Error GoTo Erro_Grafico_Le_FaturamentoMensal

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 90026
    
    sMes(1) = "Janeiro"
    sMes(2) = "Fevereiro"
    sMes(3) = "Março"
    sMes(4) = "Abril"
    sMes(5) = "Maio"
    sMes(6) = "Junho"
    sMes(7) = "Julho"
    sMes(8) = "Agosto"
    sMes(9) = "Setembro"
    sMes(10) = "Outubro"
    sMes(11) = "Novembro"
    sMes(12) = "Dezembro"
    
    If iFilialEmpresa = 0 Then
    
        lErro = CF("FilialEmpresa_Le_Filtro", sFiltro)
        If lErro <> SUCESSO Then gError 177597
    
        If Devolucoes.Value <> MARCADO Then
    
            'Lê os Dados de Saldo Mes Faturamento - Empresa Toda
            lErro = Comando_Executar(lComando, "SELECT Ano, SUM(ValorFaturado1), SUM(ValorFaturado2), SUM(ValorFaturado3), SUM(ValorFaturado4), SUM(ValorFaturado5), SUM(ValorFaturado6), SUM(ValorFaturado7), SUM(ValorFaturado8), SUM(ValorFaturado9), SUM(ValorFaturado10), SUM(ValorFaturado11), SUM(ValorFaturado12) FROM SldMesFat WHERE Ano >= ? AND Ano <= ? " & sFiltro & " GROUP BY Ano ORDER BY Ano", _
            iAnoLido, dSoma(1), dSoma(2), dSoma(3), dSoma(4), dSoma(5), dSoma(6), dSoma(7), dSoma(8), dSoma(9), dSoma(10), dSoma(11), dSoma(12), iAnoIni, iAnoFim)
            If lErro <> AD_SQL_SUCESSO Then gError 90391
        
        Else
        
            'Lê os Dados de Saldo Mes Faturamento - Empresa Toda
            lErro = Comando_Executar(lComando, "SELECT Ano, SUM(ValorFaturado1 - ValorDevolvido1), SUM(ValorFaturado2 - ValorDevolvido2), SUM(ValorFaturado3 - ValorDevolvido3), SUM(ValorFaturado4 - ValorDevolvido4), SUM(ValorFaturado5 - ValorDevolvido5), SUM(ValorFaturado6 - ValorDevolvido6), " & _
            "SUM(ValorFaturado7 - ValorDevolvido7), SUM(ValorFaturado8 - ValorDevolvido8), SUM(ValorFaturado9 - ValorDevolvido9), SUM(ValorFaturado10 - ValorDevolvido10), SUM(ValorFaturado11 - ValorDevolvido11), SUM(ValorFaturado12 - ValorDevolvido12) FROM SldMesFat WHERE Ano >= ? AND Ano <= ? " & sFiltro & " GROUP BY Ano ORDER BY Ano", _
            iAnoLido, dSoma(1), dSoma(2), dSoma(3), dSoma(4), dSoma(5), dSoma(6), dSoma(7), dSoma(8), dSoma(9), dSoma(10), dSoma(11), dSoma(12), iAnoIni, iAnoFim)
            If lErro <> AD_SQL_SUCESSO Then gError 90391
            
        End If
        
    Else
    
        If Devolucoes.Value <> MARCADO Then
    
            'Lê os Dados de Saldo Mes Faturamento
            lErro = Comando_Executar(lComando, "SELECT Ano, SUM(ValorFaturado1), SUM(ValorFaturado2), SUM(ValorFaturado3), SUM(ValorFaturado4), SUM(ValorFaturado5), SUM(ValorFaturado6), SUM(ValorFaturado7), SUM(ValorFaturado8), SUM(ValorFaturado9), SUM(ValorFaturado10), SUM(ValorFaturado11), SUM(ValorFaturado12) FROM SldMesFat WHERE FilialEmpresa=? AND Ano >= ? AND Ano <= ? GROUP BY Ano ORDER BY Ano", _
            iAnoLido, dSoma(1), dSoma(2), dSoma(3), dSoma(4), dSoma(5), dSoma(6), dSoma(7), dSoma(8), dSoma(9), dSoma(10), dSoma(11), dSoma(12), iFilialEmpresa, iAnoIni, iAnoFim)
            If lErro <> AD_SQL_SUCESSO Then gError 90391
            
        Else
        
            'Lê os Dados de Saldo Mes Faturamento
            lErro = Comando_Executar(lComando, "SELECT Ano, SUM(ValorFaturado1 - ValorDevolvido1), SUM(ValorFaturado2 - ValorDevolvido2), SUM(ValorFaturado3 - ValorDevolvido3), SUM(ValorFaturado4 - ValorDevolvido4), SUM(ValorFaturado5 - ValorDevolvido5), SUM(ValorFaturado6 - ValorDevolvido6), " & _
            "SUM(ValorFaturado7 - ValorDevolvido7), SUM(ValorFaturado8 - ValorDevolvido8), SUM(ValorFaturado9 - ValorDevolvido9), SUM(ValorFaturado10 - ValorDevolvido10), SUM(ValorFaturado11 - ValorDevolvido11), SUM(ValorFaturado12 - ValorDevolvido12) FROM SldMesFat WHERE FilialEmpresa=? AND Ano >= ? AND Ano <= ? GROUP BY Ano ORDER BY Ano", _
            iAnoLido, dSoma(1), dSoma(2), dSoma(3), dSoma(4), dSoma(5), dSoma(6), dSoma(7), dSoma(8), dSoma(9), dSoma(10), dSoma(11), dSoma(12), iFilialEmpresa, iAnoIni, iAnoFim)
            If lErro <> AD_SQL_SUCESSO Then gError 90391
            
        End If
        
    End If
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90333
    
    iMesFinal = 12
        
    If iAnoLido = iAnoIni Then
        iMesInicial = iMesIni
    Else
        iMesInicial = 1
    End If
    
    Do While lErro = AD_SQL_SUCESSO
        
        
        If iAnoLido <> iAnoIni Then iMesInicial = 1
        
        If iAnoLido = iAnoFim Then iMesFinal = iMesFim
                  
        For iCont = iMesInicial To iMesFinal
            'Carrega Dados em objItemGrafico
            Set colItensGrafico = New Collection
            Set objItemGrafico = New ClassItemGrafico
            
            objItemGrafico.sNomeColuna = sMes(iCont) & "/" & iAnoLido
            objItemGrafico.dValorColuna = dSoma(iCont)
            objItemGrafico.LegendText = "Saldo"
            colItensGrafico.Add objItemGrafico
            objGrafico.colcolItensGrafico.Add colItensGrafico
        Next
           
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90333
   
    Loop
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)

    Grafico_Le_FaturamentoMensal = SUCESSO

    Exit Function

Erro_Grafico_Le_FaturamentoMensal:

    Grafico_Le_FaturamentoMensal = gErr

    Select Case gErr

        Case 90026
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 90027, 90028, 90391
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESFAT", gErr)
        
        Case 177597
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161702)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Function Gera_Grafico(objGrafico As ClassGrafico, sPeriodo As String) As Long
'Obtém os dados necessários para gerar o gráfico
'Seta configurações do gráfico

Dim iColuna As Integer
Dim iLinha As Integer, iIndice As Integer
Dim lErro As Long
Dim sPercentual As String
Dim objPlanilha As New ClassPlanilhaExcel
Dim objColunasMeses As New ClassColunasExcel
Dim objColunasValores As New ClassColunasExcel
Dim objColunasValorMedio As New ClassColunasExcel
Dim objCelMeses As ClassCelulasExcel
Dim objCelValores As ClassCelulasExcel
Dim objCelValorMedio As ClassCelulasExcel
Dim objItemGrafico As ClassItemGrafico

On Error GoTo Erro_Gera_Grafico

    MousePointer = vbHourglass
    
    'Informa ao excel o nome do gráfico e o nome da planilha como ajustado
    objPlanilha.sNomeGrafico = "Gráfico-Fat-Comparativo Mensal"
    objPlanilha.sNomePlanilha = "Faturamento-Comparativo Mensal"
    
    ' *** COLUNA MESES ***
    'Configura essa coluna como integrante do eixo X
    objColunasMeses.iParticipaGrafico = EXCEL_PARTICIPA_GRAFICO_X
    
    'Configura o gráfico Linha 3D como gráfico a ser utilizado para essa coluna
    objColunasMeses.lTipoGraficoColuna = EXCEL_GRAFICO_3D_LINE
    
    'Configura o gráfico para não exibir os DataLabels referentes a essa coluna
    objColunasMeses.lDataLabels = EXCEL_NAO_EXIBE_LABELS
    '**********************
    
    ' *** COLUNA VALORES ***
    'Configura essa coluna como integrante do eixo Y
    objColunasValores.iParticipaGrafico = EXCEL_PARTICIPA_GRAFICO_Y
    
    'Configura o gráfico Linha 3D como gráfico a ser utilizado para essa coluna
    objColunasValores.lTipoGraficoColuna = EXCEL_GRAFICO_3D_LINE
    
    'Configura o gráfico para exibir valores como DataLabels referentes a essa coluna
    objColunasValores.lDataLabels = EXCEL_EXIBE_LABELS_VALOR
    
    'Configura o gráfico para exibir os DataLabels no sentido horizontal
    objColunasValores.lDataLabelsOrientacao = EXCEL_LABEL_ORIENTACAO_HORIZONTAL
    '***********************
    
    ' *** COLUNA VALOR MÉDIO ***
    'Informa ao excel em qual eixo essa coluna fará parte do gráfico
    objColunasValorMedio.iParticipaGrafico = EXCEL_PARTICIPA_GRAFICO_Y
    
    'Informa ao excel o tipo de gráfico que será usado para representar essa coluna
    objColunasValorMedio.lTipoGraficoColuna = EXCEL_GRAFICO_3D_LINE
    
    'Informa ao excel como serão exibidos os Datalabels para essa coluna
    objColunasValorMedio.lDataLabels = EXCEL_NAO_EXIBE_LABELS
    '***********************
    
    Set objCelMeses = New ClassCelulasExcel
    Set objCelValores = New ClassCelulasExcel
    Set objCelValorMedio = New ClassCelulasExcel

    objCelMeses.vValor = "Meses"
    objCelValores.vValor = "Faturamento"
    objCelValorMedio.vValor = "Média Anual"
        
    objColunasMeses.colCelulas.Add objCelMeses
    objColunasValores.colCelulas.Add objCelValores
    objColunasValorMedio.colCelulas.Add objCelValorMedio
    
    For iIndice = 1 To objGrafico.colcolItensGrafico.Count
    
        Set objItemGrafico = objGrafico.colcolItensGrafico(iIndice).Item(1)
        Set objCelMeses = New ClassCelulasExcel
        Set objCelValores = New ClassCelulasExcel
        Set objCelValorMedio = New ClassCelulasExcel
    
        objCelMeses.vValor = objItemGrafico.sNomeColuna
        objCelValores.vValor = Format(objItemGrafico.dValorColuna, "STANDARD")
        objCelValorMedio.vValor = objGrafico.colcolItensGrafico(iIndice).Item(2).dValorColuna
        
        objColunasMeses.colCelulas.Add objCelMeses
        objColunasValores.colCelulas.Add objCelValores
        objColunasValorMedio.colCelulas.Add objCelValorMedio
    
    Next
    
    'Adiciona as colunas à coleção de colunas
    objPlanilha.colColunas.Add objColunasMeses
    objPlanilha.colColunas.Add objColunasValores
    objPlanilha.colColunas.Add objColunasValorMedio
    
    'Instancia a coleção que guardará as seções de cabeçalho / rodapé
    Set objPlanilha.colCabecalhoRodape = New Collection

    'Monta o cabeçalho e o rodapé do Gráfico
    lErro = Grafico_Monta_Cabecalho_Rodape(objPlanilha.colCabecalhoRodape)
    If lErro <> SUCESSO Then gError 90538
    
    'Informa ao excel o título do gráfico
    objPlanilha.sTituloGrafico = "Período: " & sPeriodo & vbCrLf & "Média Mensal em Reais: R$ " & CStr(Format(objGrafico.colcolItensGrafico(1).Item(2).dValorColuna, "Standard"))
    
    'Informa ao excel a posição em que deverá ser exibida a legenda
    objPlanilha.lPosicaoLegenda = EXCEL_LEGENDA_DIREITA
    
    'Informa ao excel a posição dos labels do eixo X
    objPlanilha.lLabelsXPosicao = EXCEL_TICKLABEL_POSITION_LOW
    
    'Informa ao excel a orientação dos labels do eixo X
    objPlanilha.lLabelsXOrientacao = EXCEL_TICKLABEL_ORIENTATION_UPWARD
    
    'Informa ao excel que a plotagem do dados será por coluna
    objPlanilha.vPlotLinhaColuna = EXCEL_COLUMNS
    
    'Monta a planilha e o gráfico com os dados passados em objPlanilha
    lErro = CF("Excel_Cria_Grafico", objPlanilha)
    If lErro <> SUCESSO Then gError 79972

    MousePointer = vbDefault
    
    Gera_Grafico = SUCESSO
    
    Exit Function
    
Erro_Gera_Grafico:

    Gera_Grafico = gErr
    
    Select Case gErr
        
        Case 79972, 90538
        
        Case 79971
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORES_COLUNAS_NAO_TRATADOS_GRAFICO", gErr, iColuna)
            
        Case 79970
            Call Rotina_Erro(vbOKOnly, "ERRO_GRAFICO_VALORES_A_EXIBIR_NAO_DEFINIDOS2", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161703)
            
    End Select
    
    MousePointer = vbDefault
    
    Exit Function
    
End Function

Public Function Grafico_Monta_Cabecalho_Rodape(colLinhas As Collection) As Long
'Função criada em 29/05/01 por Luiz Gustavo de Freitas Nogueira
'Essa função preenche os objetos com os dados de cada linha que será exibida no cabeçalho

Dim objLinha As ClassLinhaCabecalhoExcel

On Error GoTo Erro_Grafico_Monta_Cabecalho_Rodape

        '*** PREENCHIMENTO DO CABEÇALHO ESQUERDO ***
        
        'Instancia um objeto para armazenar dados de uma nova linha
        Set objLinha = New ClassLinhaCabecalhoExcel
        
            ' *** LINHA 01 - CABEÇALHO ESQUERDO ***
            objLinha.iSecao = EXCEL_CABECALHO_ESQUERDO
            objLinha.sTexto = gsNomeEmpresa
            objLinha.sFonte = EXCEL_FONTE_COURIER_NEW
            objLinha.iTamanhoFonte = 9
            objLinha.iEspacoLinha = 1 'Indica quantas linhas devem existir entre essa linha e a próxima
            objLinha.iLinha = 1 'Indica a posição da linha no cabeçalho
            objLinha.sAlinhamento = EXCEL_CABECALHO_RODAPE_ALINHAMENTO_ESQUERDA
            
            'Adiciona a linha à coleção de linhas de cabeçalho / rodapé
            colLinhas.Add objLinha
        
        'Instancia um objeto para armazenar dados de uma nova linha
        Set objLinha = New ClassLinhaCabecalhoExcel
            
            ' *** LINHA 02 - CABEÇALHO ESQUERDO ***
            objLinha.iSecao = EXCEL_CABECALHO_ESQUERDO
            objLinha.sTexto = gsNomeFilialEmpresa
            objLinha.sFonte = EXCEL_FONTE_COURIER_NEW
            objLinha.iTamanhoFonte = 9
            objLinha.iEspacoLinha = 2 'Indica quantas linhas devem existir entre essa linha e a próxima
            objLinha.iLinha = 2 'Indica a posição da linha no cabeçalho
            objLinha.sAlinhamento = EXCEL_CABECALHO_RODAPE_ALINHAMENTO_ESQUERDA
            
            'Adiciona a linha à coleção de linhas de cabeçalho / rodapé
            colLinhas.Add objLinha
            
        ' *** FIM DO CABEÇALHO ESQUERDO ***
        
        ' *** PREENCHIMENTO DO CABEÇALHO CENTRAL ***
        
        'Instancia um objeto para armazenar dados de uma nova linha
        Set objLinha = New ClassLinhaCabecalhoExcel
            
            ' *** LINHA 01 - CABEÇALHO CENTRAL ***
            objLinha.iSecao = EXCEL_CABECALHO_CENTRAL
            objLinha.sTexto = "Faturamento - Comparativo Mensal"
            objLinha.sFonte = EXCEL_FONTE_BOOKMAN
            objLinha.iTamanhoFonte = 20
            objLinha.sNegrito = EXCEL_CABECALHO_RODAPE_NEGRITO
            objLinha.iEspacoLinha = 0 'Indica quantas linhas devem existir entre essa linha e a próxima
            objLinha.iLinha = 1 'Indica a posição da linha no cabeçalho
            objLinha.sAlinhamento = EXCEL_CABECALHO_RODAPE_ALINHAMENTO_CENTRAL
            
            'Adiciona a linha à coleção de linhas de cabeçalho / rodapé
            colLinhas.Add objLinha

        ' *** FIM DO CABEÇALHO CENTRAL ***
        
        ' *** PREENCHIMENTO DO CABEÇALHO DIREITO ***
        'Instancia um objeto para armazenar dados de uma nova linha
        Set objLinha = New ClassLinhaCabecalhoExcel
            
            ' *** LINHA 01 - CABEÇALHO DIREITO ***
            objLinha.iSecao = EXCEL_CABECALHO_DIREITO
            objLinha.sTexto = CStr(Date)
            objLinha.sFonte = EXCEL_FONTE_COURIER_NEW
            objLinha.iTamanhoFonte = 9
            objLinha.iEspacoLinha = EXCEL_CABECALHO_RODAPE_NAO_QUEBRA_LINHA 'Indica quantas linhas devem existir entre essa linha e a próxima
            objLinha.iLinha = 1 'Indica a posição da linha no cabeçalho
            objLinha.sAlinhamento = EXCEL_CABECALHO_RODAPE_ALINHAMENTO_DIREITA
            
            'Adiciona a linha à coleção de linhas de cabeçalho / rodapé
            colLinhas.Add objLinha

        ' *** FIM DO CABEÇALHO DIREITO ***
        
    Grafico_Monta_Cabecalho_Rodape = SUCESSO
    
    Exit Function
    
Erro_Grafico_Monta_Cabecalho_Rodape:

    Grafico_Monta_Cabecalho_Rodape = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161704)
            
    End Select
    
    Exit Function

End Function


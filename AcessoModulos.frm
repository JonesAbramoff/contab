VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form AcessoModulos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liberação de Acesso"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9480
   Icon            =   "AcessoModulos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   9480
   Begin VB.Frame Frame4 
      Caption         =   "PASSO 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   135
      TabIndex        =   12
      Top             =   3870
      Width           =   9165
      Begin VB.CommandButton BotaoFechar 
         Caption         =   "Fechar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   4838
         Picture         =   "AcessoModulos.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Fechar"
         Top             =   570
         Width           =   975
      End
      Begin VB.CommandButton BotaoOk 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   3668
         Picture         =   "AcessoModulos.frx":02C8
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   570
         Width           =   975
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Apertando botão OK, limites de uso e acessos aos módulos ficam liberados no Sistema."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1020
         TabIndex        =   15
         Top             =   225
         Width           =   7455
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "PASSO 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   135
      TabIndex        =   9
      Top             =   2940
      Width           =   9165
      Begin VB.TextBox Senha 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   6338
         MaxLength       =   5
         TabIndex        =   3
         Top             =   495
         Width           =   660
      End
      Begin VB.TextBox Senha 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   5558
         MaxLength       =   5
         TabIndex        =   4
         Top             =   495
         Width           =   660
      End
      Begin VB.TextBox Senha 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   4778
         MaxLength       =   5
         TabIndex        =   5
         Top             =   495
         Width           =   660
      End
      Begin VB.TextBox Senha 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   3998
         MaxLength       =   5
         TabIndex        =   6
         Top             =   495
         Width           =   660
      End
      Begin VB.TextBox Senha 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   3218
         MaxLength       =   5
         TabIndex        =   8
         Top             =   495
         Width           =   660
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Senha:"
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
         Left            =   2490
         TabIndex        =   16
         Top             =   555
         Width           =   645
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Preencha o campo de Senha com a senha fornecida pela FORPRINT."
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
         Left            =   1770
         TabIndex        =   17
         Top             =   195
         Width           =   5955
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "PASSO 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   135
      TabIndex        =   7
      Top             =   2325
      Width           =   9165
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Forneça a FORPRINT o  Nº Série, Razão Social, Nome Reduzido e CGC da Empresa escolhida."
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
         Left            =   675
         TabIndex        =   18
         Top             =   270
         Width           =   8145
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "PASSO 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   150
      TabIndex        =   0
      Top             =   30
      Width           =   9165
      Begin VB.CommandButton Command1 
         Caption         =   "Botão Temporário para gerar senhas para o sistema"
         Height          =   540
         Left            =   3990
         TabIndex        =   14
         ToolTipText     =   "Selecione uma empresa na Lista e um CGC na combo antes de apertar o botão. A senha aparecerá ao lado do nome da empresa."
         Top             =   1560
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.ComboBox CGC 
         Height          =   315
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1755
         Width           =   2520
      End
      Begin VB.ListBox Empresas 
         Height          =   1620
         Left            =   6345
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   435
         Width           =   2580
      End
      Begin MSMask.MaskEdBox Serie 
         Height          =   300
         Left            =   1170
         TabIndex        =   11
         Top             =   420
         Width           =   4830
         _ExtentX        =   8520
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nome Red:"
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
         TabIndex        =   25
         Top             =   1365
         Width           =   960
      End
      Begin VB.Label NomeReduzido 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1170
         TabIndex        =   24
         Top             =   1320
         Width           =   3180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Empresas"
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
         Left            =   6345
         TabIndex        =   19
         Top             =   210
         Width           =   855
      End
      Begin VB.Label Nome 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1170
         TabIndex        =   20
         Top             =   885
         Width           =   4830
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nº Série:"
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
         Left            =   330
         TabIndex        =   21
         Top             =   450
         Width           =   780
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
         Left            =   525
         TabIndex        =   22
         Top             =   930
         Width           =   585
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "CGC:"
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
         Left            =   660
         TabIndex        =   23
         Top             =   1815
         Width           =   450
      End
   End
End
Attribute VB_Name = "AcessoModulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim glEmpresaAtual As Long
Dim gColEmpresas As Collection

Private Sub BotaoFechar_Click()
    
    Unload Me

End Sub

Private Sub BotaoOk_Click()

Dim lErro As Long
Dim sCgc As String
Dim sNomeEmpresa As String, sSenha As String
Dim iIndice As Integer, iNumeroLogs As Integer
Dim iNumeroEmpresas As Integer, iNumeroFiliais As Integer
Dim colModulosLib As New Collection
Dim dtDataValidadeDe As Date
Dim dtDataValidadeAte As Date
Dim objDicConfig As New ClassDicConfig
Dim sSerie As String
Dim sTextoSenha As String, sCheckSum As String

On Error GoTo Erro_BotaoOk_Click

    'Verifica se os campos obrigatórios da tela estão preenchidos
    If Len(Trim(Serie.Text)) = 0 Then Error 62339
    If Len(Trim(Nome.Caption)) = 0 Then Error 62340
    If CGC.ListIndex = -1 Then Error 62341
    
    'Forma a senha
    For iIndice = Senha.LBound To Senha.UBound
        If Len(Trim(Senha(iIndice).Text)) = 0 Or Len(Trim(Senha(iIndice).Text)) < 5 Then Error 62342
        sSenha = sSenha & Senha(iIndice).Text
    Next
    
    'Verifica se a senha possui caracteres inválidos
    For iIndice = 1 To 25
        If Asc(Mid(sSenha, iIndice, 1)) < vbKey0 Or Asc(Mid(sSenha, iIndice, 1)) > vbKeyF Then Error 62343
    Next
        
    'Decifra a senha da empresa
    lErro = Senha_Empresa_Decifra(sSenha, sCgc, sNomeEmpresa, iNumeroLogs, iNumeroEmpresas, iNumeroFiliais, colModulosLib, dtDataValidadeAte, sTextoSenha)
    If lErro <> SUCESSO Then Error 62344
    
    'Calcula o Checksum
    Call Calcula_CheckSum(Mid(sTextoSenha, 1, 23), sCheckSum)
    
    'Se o CheckSum não coincidir c\om a senha --> Erro
    If sCheckSum <> Mid(sTextoSenha, 24, 2) Then Error 62338
    
    'Verifica se o trecho do nome da epresa extraído da senha
    'confere com o nome da empresa da tela5
    If UCase(Mid(Nome.Caption, 1, 2)) <> UCase(sNomeEmpresa) Then Error 62345
    
    'Verifica se o trecho do CGC da epresa extraído da senha
    'confere com o CGC informado na tela
    If Mid(CGC.Text, 1, 4) <> sCgc Then Error 62346
    
    'Lê as informações de configuração que estão no BD
    lErro = DicConfig_Le(objDicConfig)
    If lErro <> SUCESSO Then Error 62347
    
    sSerie = Trim(Serie.Text)
    
    'Se o número série estiver preenchido
    If Len(objDicConfig.sSerie) > 0 Then
        'Verifica se o número serie no BD coincide com o da tela
        If sSerie <> objDicConfig.sSerie Then Error 62348
    End If
    'Verifica se a Data da senha do BD é maior que a data de hoje.
    If objDicConfig.dtDataSenha <> DATA_NULA Then If objDicConfig.dtDataSenha > Date Then Error 62349
    
    If objDicConfig.dtValidadeDe <> DATA_NULA Then
        dtDataValidadeDe = objDicConfig.dtValidadeDe
    Else
        dtDataValidadeDe = Date
    End If
    
    Set objDicConfig = New ClassDicConfig
    
    objDicConfig.sSenha = sSenha
    objDicConfig.sSerie = sSerie
    objDicConfig.iLimiteEmpresas = iNumeroEmpresas
    objDicConfig.iLimiteFiliais = iNumeroFiliais
    objDicConfig.iLimiteLogs = iNumeroLogs
    objDicConfig.dtValidadeDe = dtDataValidadeDe
    objDicConfig.dtValidadeAte = dtDataValidadeAte
    objDicConfig.dtDataSenha = Date
    Set objDicConfig.colModulosLib = colModulosLib
    
    
    'Grava as Informações da configuração do sistema
    lErro = DicConfig_Grava(objDicConfig)
    If lErro <> SUCESSO Then Error 62350
    
    'Exibe os dados de informação na tela informativa AcessoDados
    Call AcessoDados.Trata_Parametros(objDicConfig)
    
    AcessoDados.Show
    
    Exit Sub
    
Erro_BotaoOk_Click:

    Select Case Err
    
        Case 62339
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_SERIE_NAO_PREENCHIDO", Err)
        
        Case 62340
            Call Rotina_Erro(vbOKOnly, "ERRO_EMPRESA_NAO_PREENCHIDA", Err)
        
        Case 62341
            Call Rotina_Erro(vbOKOnly, "ERRO_CGC_NAO_INFORMADO", Err)
        
        Case 62342
            Call Rotina_Erro(vbOKOnly, "ERRO_TRECHO_SENHA_INCOMPLETO", Err, iIndice)
        
        Case 62338, 62343, 62345, 62346
            Call Rotina_Erro(vbOKOnly, "ERRO_SENHA_INVALIDA", Err)
        
        Case 62344, 62347
        
        Case 62348
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_SERIE_DIFERENTE_BD", Err)
         
        Case 62349
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_SENHA_BD_MAIOR", Err)
        
        Case 62350
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142000)
             
    End Select
    
    Exit Sub
    
End Sub

Function DecParaHexa(vNumero As Variant, iTamanho As Integer) As String
'Faz a conversão de um Decimal para um HexaDecimal devolvendo o número
'formatado no tamanho passado

Dim sNumero As String

    sNumero = CStr(Trim(vNumero))
    sNumero = Hex(sNumero)
            
    DecParaHexa = FormataCpoNum(sNumero, iTamanho)
    
    Exit Function
    
End Function

Function HexaParaDec(vNumero As Variant, iTamanho As Integer) As String
'Faz a conversão de um HexaDecimal para um Decimal devolvendo o número
'formatado no tamanho passado

Dim sNumeroAux As String
Dim sAlgarismo As String
Dim iAlgarismo As String
Dim lNumero As Long
Dim iTamanhoIni As Integer
Dim iIndice As Integer

    sNumeroAux = CStr(Trim(vNumero))
    iTamanhoIni = Len(sNumeroAux)
    lNumero = 0
    
    For iIndice = 1 To iTamanhoIni
        
        sAlgarismo = Mid(sNumeroAux, iIndice, 1)
        
        Select Case Asc(sAlgarismo)
            
            Case Asc("A")
                iAlgarismo = 10
            
            Case Asc("B")
                iAlgarismo = 11
            
            Case Asc("C")
                iAlgarismo = 12
            
            Case Asc("D")
                iAlgarismo = 13
            
            Case Asc("E")
                iAlgarismo = 14
            
            Case Asc("F")
                iAlgarismo = 15
            
            Case Else
                iAlgarismo = CInt(sAlgarismo)
        
        End Select
        
        lNumero = lNumero + (iAlgarismo * (16 ^ (iTamanhoIni - iIndice)))
    
    Next
        
    HexaParaDec = FormataCpoNum(lNumero, iTamanho)
    
    Exit Function
    
End Function

Function BinParaDec(vNumero As Variant, iTamanho As Integer) As String
'Faz a conversão de um Binário para um Decimal devolvento o número
'formatado no tamanho passado

Dim sNumeroAux As String
Dim sAlgarismo As String
Dim iAlgarismo As String
Dim lNumero As Long
Dim iIndice As Integer
Dim iTamanhoIni As Integer

    sNumeroAux = CStr(Trim(vNumero))
    iTamanhoIni = Len(sNumeroAux)
    lNumero = 0
    
    For iIndice = 1 To iTamanhoIni
        
        sAlgarismo = Mid(sNumeroAux, iIndice, 1)
        
        iAlgarismo = CInt(sAlgarismo)
        
        lNumero = lNumero + (iAlgarismo * (2 ^ (iTamanhoIni - iIndice)))
    
    Next
        
    BinParaDec = FormataCpoNum(lNumero, iTamanho)
    
    Exit Function
    
End Function

Function DecParaBin(vNumero As Variant, iTamanho As Integer) As String
'Faz a conversão de um Decimal para um Binário devolvento o número
'formatado no tamanho passado

Dim iNumeroAux As String
Dim iIndice As Integer
Dim iResto As Integer
Dim iQuociente As Integer
Dim colRestos As New Collection
Dim sNumero As String

    iNumeroAux = vNumero
    
    iQuociente = iNumeroAux \ 2
    iResto = iNumeroAux Mod 2
    
    colRestos.Add iResto
    
    Do While iQuociente <> 0
        
        iResto = iQuociente Mod 2
        iQuociente = iQuociente \ 2

        colRestos.Add iResto

    Loop
    
    sNumero = ""
    
    For iIndice = colRestos.Count To 1 Step -1
        sNumero = sNumero & colRestos(iIndice)
    Next
        
    If iTamanho > 0 Then sNumero = FormataCpoNum(sNumero, iTamanho)
        
    DecParaBin = sNumero
    
    Exit Function
    
End Function


'Já existe (Criada por Raphael na CLassGeracaoArqIcms)
Private Function FormataCpoNum(vData As Variant, iTam As Integer) As String
'formata campo numerico alinhado-o à direita sem ponto e decimais, com zeros a esquerda

Dim iData As Integer
Dim sData As String

    If Len(vData) = iTam Then

        FormataCpoNum = vData
        Exit Function

    End If

    iData = iTam - Len(vData)
    
    If iData > 0 Then sData = String(iData, "0")

    FormataCpoNum = sData & vData

    Exit Function

End Function

Function Intercala_Texto(sTexto1 As String, sTexto2 As String) As String
'Intercala o sTexto1 e o sTexto2
'Ex. "1111", "0000", Res. = "10101010"
'    "11111", "000", REs. = "10101011"
'    "111", "00000", Res. = "10101000"

Dim sTextoFinal As String
Dim iTamanhoMenor As Integer
Dim iIndice As Integer
    
    'Veriifca o tamanho da menor string
    iTamanhoMenor = IIf(Len(sTexto1) > Len(sTexto2), Len(sTexto2), Len(sTexto1))

    'Intercala até a menor string acabar
    For iIndice = 1 To iTamanhoMenor
        sTextoFinal = sTextoFinal & Mid(sTexto1, iIndice, 1)
        sTextoFinal = sTextoFinal & Mid(sTexto2, iIndice, 1)
    Next
    
    'Intercala o que sobrou
    sTextoFinal = sTextoFinal & Mid(sTexto1, iTamanhoMenor + 1)
    sTextoFinal = sTextoFinal & Mid(sTexto2, iTamanhoMenor + 1)
    
    Intercala_Texto = sTextoFinal
    
    Exit Function

End Function

Sub Desintercala_Texto(iTamanhoTexto2 As Integer, sTextoTotal, sTexto1 As String, sTexto2 As String)
'Desintercala o textototal em sTexto1 e o sTexto2 de acordo com o tam do 2 texto que é passado
'Ex.  "10101010", tam 4 => "1111", "0000"
'     "10101011", tam 3 => "11111", "000"
'     "10101000", tam 5 => "111", "00000"

Dim iTamanhoTotal As Integer
Dim iIndice As Integer
Dim iTamanhoTexto1 As Integer
Dim iTamanhoMenor As Integer
    
    'Deduz o tamanho de texto 1
    iTamanhoTexto1 = Len(sTextoTotal) - iTamanhoTexto2
    
    sTexto1 = ""
    sTexto2 = ""
    
    'Verifica qual dos 2 textos é o menor
    iTamanhoMenor = IIf(iTamanhoTexto1 > iTamanhoTexto2, iTamanhoTexto2, iTamanhoTexto1)
    
    'Desintercala até o tamanho do menor texto
    For iIndice = 1 To iTamanhoMenor
        
        sTexto1 = sTexto1 & Mid(sTextoTotal, 2 * iIndice - 1, 1)
        sTexto2 = sTexto2 & Mid(sTextoTotal, 2 * iIndice, 1)
           
    Next
    
    'Atribui o restante do texto a strign maior
    If iTamanhoMenor = iTamanhoTexto1 Then
        sTexto2 = sTexto2 & Mid(sTextoTotal, (2 * iTamanhoMenor) + 1)
    Else
        sTexto1 = sTexto1 & Mid(sTextoTotal, (2 * iTamanhoMenor) + 1)
    End If

    Exit Sub

End Sub

Private Sub Calcula_CheckSum(sTexto As String, sCheckSum As String)
'Calcula o conteúdo do campo CheckSum

Dim iTamanhoTexto As Integer
Dim iIndice As Integer
Dim iDigito As Integer
Dim iSoma As Integer
Dim iModulo As Integer
Dim iResto As Integer
    
    iTamanhoTexto = Len(sTexto)
    
    iSoma = 0
        
    For iIndice = 1 To iTamanhoTexto
          
       iDigito = HexaParaDec(Mid(sTexto, iIndice, 1), 2)
       
       iModulo = IIf((iIndice Mod 9) = 0, 9, iIndice Mod 9)
       
       iSoma = iSoma + (iDigito * iModulo)
       
    Next
    
    iResto = iSoma Mod 256
    
    sCheckSum = DecParaHexa(iResto, 2)
        
    Exit Sub
    
End Sub

Function Senha_Empresa_Gera(sCgc As String, sNomeEmpresa As String, iNumeroLogs As Integer, iNumeroEmpresas As Integer, iNumeroFiliais As Integer, colSiglasModulosLib As Collection, dtDataValidade As Date, sSenha As String) As Long
'Gera a senha para o sistema através do CGC, NomeEmpresa, Limites de Logs , empresas e filiais e a Data de validadae.

Dim lErro As Long
Dim asPartesSenha(1 To 7) As String
Dim iIndice As Integer
Dim colModulos As New Collection
Dim objModulos As AdmModulo
Dim bModuloLiberado As Boolean
Dim sTextoAux As String
Dim sCheckSum As String
Dim sTextoSenha As String

On Error GoTo Erro_Senha_Empresa_Gera
    
    For iIndice = LBound(asPartesSenha) To UBound(asPartesSenha)
        asPartesSenha(iIndice) = ""
    Next
    
    'Extrai os 4 primeiros dígitos do CGC
    asPartesSenha(1) = DecParaHexa(Left(sCgc, 4), 4)
       
    'Lê todos os módulos
    lErro = CF("Modulos_Le_Todos", colModulos)
    If lErro <> SUCESSO Then Error 62333
    
    'Para cada módulo
    For Each objModulos In colModulos
        
        bModuloLiberado = False
        'MOnta o trecho da senha de liberação dos módulos
        For iIndice = 1 To colSiglasModulosLib.Count
            If objModulos.sSigla = colSiglasModulosLib(iIndice) Then
                bModuloLiberado = True
            End If
        Next
        
        If bModuloLiberado Then
            sTextoAux = sTextoAux & MODULO_LIBERADO
        Else
            sTextoAux = sTextoAux & MODULO_NAO_LIBERADO
        End If
        
    Next
    
    sTextoAux = sTextoAux & FormataCpoNum("", 16 - Len(sTextoAux))
    
    sTextoAux = BinParaDec(sTextoAux, 5)
    sTextoAux = DecParaHexa(sTextoAux, 4)
    asPartesSenha(2) = sTextoAux
    
    sTextoAux = DecParaHexa(Asc(Mid(sNomeEmpresa, 1, 1)), 2) & DecParaHexa(Asc(Mid(sNomeEmpresa, 2, 1)), 2)
    asPartesSenha(3) = sTextoAux
    
    sTextoAux = DecParaHexa(Month(dtDataValidade), 1) & DecParaHexa(Format(dtDataValidade, "YY"), 2)
    asPartesSenha(4) = sTextoAux
    
    asPartesSenha(5) = DecParaHexa(iNumeroFiliais, 2)
    asPartesSenha(6) = DecParaHexa(iNumeroLogs, 4)
    asPartesSenha(7) = DecParaHexa(iNumeroEmpresas, 2)
    
    sTextoAux = ""
    For iIndice = LBound(asPartesSenha) To UBound(asPartesSenha)
        sTextoAux = sTextoAux & asPartesSenha(iIndice)
    Next
    
    Call Calcula_CheckSum(sTextoAux, sCheckSum)
    
    sTextoAux = sTextoAux & sCheckSum
    
    sTextoSenha = Intercala_Texto(Mid(sTextoAux, 1, 8), Mid(sTextoAux, 9, 7))
    sTextoSenha = Intercala_Texto(sTextoSenha, Mid(sTextoAux, 16, 10))
    
    sSenha = sTextoSenha
    
    Senha_Empresa_Gera = SUCESSO
    
    Exit Function

Erro_Senha_Empresa_Gera:

    Senha_Empresa_Gera = Err
    
    Select Case Err
    
        Case 62333
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142001)
            
    End Select
    
    Exit Function
    
End Function

Function Senha_Empresa_Decifra(sSenha As String, sCgc As String, sNomeEmpresa, iNumeroLogs As Integer, iNumeroEmpresas As Integer, iNumeroFiliais As Integer, colModulosLib As Collection, dtDataValidade As Date, sTextoSenha As String) As Long

Dim iIndice As Integer
Dim colModulos As New Collection
Dim objModulos As AdmModulo
Dim bModuloLiberado As Boolean
Dim lErro As Long
Dim sCheckSum As String
Dim sTextoAux As String
Dim sTextoAux2 As String
Dim sTextoAux3 As String

On Error GoTo Erro_Senha_Empresa_Decifra
    
    'Desdobra a senha
    Call Desintercala_Texto(10, sSenha, sTextoSenha, sTextoAux)
    Call Desintercala_Texto(7, sTextoSenha, sTextoAux2, sTextoAux3)
    sTextoSenha = sTextoAux2 & sTextoAux3 & sTextoAux
        
    sCgc = HexaParaDec(Mid(sTextoSenha, 1, 4), 4)
    'Lê todos os módulos
    lErro = CF("Modulos_Le_Todos", colModulos)
    If lErro <> SUCESSO Then Error 62333

    sTextoAux = DecParaBin(HexaParaDec(Mid(sTextoSenha, 5, 4), 16), 16)
   
    Set colModulosLib = New Collection
    'Para cada mósulo liberado adiciona o modulo na coleção de módulos
    For Each objModulos In colModulos
        iIndice = iIndice + 1
        
        '###################################
        'Alterado por Wagner - 15/04/2008 (estourou o limite de 16 - Libera tudo)
        'If Mid(sTextoAux, iIndice, 1) = MODULO_LIBERADO Then colModulosLib.Add objModulos
        colModulosLib.Add objModulos
        '###################################
    Next

    'Recolhe os dados retirados da senha
    sNomeEmpresa = Chr(HexaParaDec(Mid(sTextoSenha, 9, 2), 3)) & Chr(HexaParaDec(Mid(sTextoSenha, 11, 2), 3))
    sTextoAux = "01/" & HexaParaDec(Mid(sTextoSenha, 13, 1), 2) & "/" & "20" & HexaParaDec(Mid(sTextoSenha, 14, 2), 2)
    dtDataValidade = DateAdd("m", 1, CDate(sTextoAux)) - 1
    iNumeroFiliais = HexaParaDec(Mid(sTextoSenha, 16, 2), 3)
    iNumeroLogs = HexaParaDec(Mid(sTextoSenha, 18, 4), 5)
    iNumeroEmpresas = HexaParaDec(Mid(sTextoSenha, 22, 2), 3)
    sCheckSum = HexaParaDec(Mid(sTextoSenha, 24, 2), 3)
    
    Senha_Empresa_Decifra = SUCESSO
    
    Exit Function

Erro_Senha_Empresa_Decifra:

    Senha_Empresa_Decifra = Err
    
    Select Case Err
    
        Case 62333
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142002)
            
    End Select
    
    Exit Function
    
End Function

Private Sub Command1_Click()

    If Empresas.ListIndex = -1 Then
        MsgBox "Selecione uma Empresa!", vbCritical, "SGE"
        Exit Sub
    End If
    If CGC.ListIndex = -1 Then
        MsgBox "Selecione um CGC!", vbCritical, "SGE"
        Exit Sub
    End If

    Call Teste

End Sub

Private Sub Empresas_Click()

Dim lErro As Long
Dim colFiliais As New Collection
Dim objFilialEmpresa As AdmFiliais
Dim bJaEsta As Boolean

On Error GoTo Erro_Empresas_Click

    'Se a empresa não foi alterada , Sai.
    If glEmpresaAtual = Empresas.ListIndex Then Exit Sub
    
    'Limpa a combo de CGC
    CGC.Clear
    'Se não tiver empresa selecionada , Sai
    If Empresas.ListIndex = -1 Then Exit Sub
    
    'Coloca o nome da empresa selecionada na Tela
    Nome.Caption = gColEmpresas(Empresas.ListIndex + 1).sNome
    NomeReduzido.Caption = gColEmpresas(Empresas.ListIndex + 1).sNomeReduzido
    
    'Para poder fazer transacao no bd da empresa
    lErro = Sistema_DefEmpresa(gColEmpresas(Empresas.ListIndex + 1).sNome, Empresas.ItemData(Empresas.ListIndex), EMPRESA_TODA_NOME, EMPRESA_TODA)
    If lErro <> AD_BOOL_TRUE Then Error 62350
    
    'Lê do SGEDados os CGCs da Filiais da empresa selecionada
    lErro = FiliaisEmpresa_Le_Dados_Empresa(Empresas.ItemData(Empresas.ListIndex), colFiliais)
    If lErro <> SUCESSO Then Error 62351
    'para cada filial da empresa
    For Each objFilialEmpresa In colFiliais
        bJaEsta = False
        'Verifica se o CGC já está na combo
        Call Verifica_CGC(objFilialEmpresa.sCgc, bJaEsta)
        'Se não estiver
        If Not bJaEsta Then
            'Adiciona o CGC na combo
            CGC.AddItem objFilialEmpresa.sCgc
            CGC.ItemData(CGC.NewIndex) = objFilialEmpresa.iCodFilial
        End If
    Next
    'Guarda a empresa que está selecionada
    glEmpresaAtual = Empresas.ListIndex

    Exit Sub

Erro_Empresas_Click:

    Select Case Err
    
        Case 62350, 62351
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142003)
            
    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colEmpresas As New Collection
Dim objEmpresa As ClassDicEmpresa
Dim objDicConfig As New ClassDicConfig

On Error GoTo Erro_Form_Load
       
    glEmpresaAtual = -1
    
    'Lê as configurações do Sistema no BD
    lErro = DicConfig_Le(objDicConfig)
    If lErro <> SUCESSO Then Error 62448
    
    'Se a série estiver preenchida, exibe na tela.
    If Len(objDicConfig.sSerie) > 0 Then Serie.Text = objDicConfig.sSerie
    
    'Lê todas as empresas
    lErro = Empresas_Le_Todas(colEmpresas)
    If lErro <> SUCESSO And lErro <> 6179 Then Error 62449
    If lErro <> SUCESSO Then Error 62450
    
    'Para cada empresa lida
    For Each objEmpresa In colEmpresas
        'Adiciona a empresa na List de Empresas
        Empresas.AddItem objEmpresa.sNomeReduzido
        Empresas.ItemData(Empresas.NewIndex) = objEmpresa.lCodigo
    Next
    
    Set gColEmpresas = colEmpresas
    
    'Seleciona a primeira emrpesa
    Empresas.ListIndex = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
    
        Case 62448, 62449
        
        Case 62450
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_EMPRESAS", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142004)
    
    End Select
       
    Exit Sub
            
End Sub

Private Sub Senha_Change(Index As Integer)
    'Facilita o preenchimento da senha passando o cursor para o
    'próximo campo de senha quando esse está totalmente preenchido.
    If Index < Senha.UBound Then
        If Len(Senha(Index).Text) = 5 Then Senha(Index + 1).SetFocus
    End If
End Sub

Private Sub Senha_KeyPress(Index As Integer, KeyAscii As Integer)
    'Passa os caracteres digitados para maiúsculos
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

Private Sub Verifica_CGC(sCgc As String, bJaEsta As Boolean)

Dim iIndice As Integer
    
    bJaEsta = False

    For iIndice = 0 To CGC.ListCount - 1
        If sCgc = CGC.List(iIndice) Then
            bJaEsta = True
            Exit For
        End If
    Next

End Sub




'========= apagar ======
Public Sub Teste()

Dim coll As New Collection
Dim lErro As Long
Dim sSenha As String
Dim s1$, s2$, s3$, s4 As String
Dim i1 As Integer, i2 As Integer, i3 As Integer
Dim dt As Date

coll.Add "ADM"
coll.Add "COM"
coll.Add "CP"
coll.Add "CR"
coll.Add "CTB"
coll.Add "EST"
coll.Add "FAT"
coll.Add "TES"
    
    dt = DateAdd("yyyy", 3, Date)

    lErro = Senha_Empresa_Gera(CGC.Text, Nome.Caption, 20, 10, 50, coll, dt, sSenha)

    Nome = Nome & " - " & sSenha

End Sub

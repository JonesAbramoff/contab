VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   8460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   525
      Left            =   1470
      TabIndex        =   1
      Top             =   1665
      Width           =   1155
   End
   Begin VB.TextBox Text1 
      Height          =   570
      Left            =   1665
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   510
      Width           =   5445
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   570
      Left            =   1695
      TabIndex        =   2
      Top             =   2685
      Width           =   6060
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim sNumerosCodBarra As String
    Call Calcula_NumerosCodBarras(Text1.Text, sNumerosCodBarra)
    Label1.Caption = sNumerosCodBarra
End Sub

Function Calcula_CodigoBarras(ByVal sBanco As String, ByVal sMoeda As String, ByVal dValor As Double, ByVal dtVencimento As Date, ByVal sLivre As String, sCodigoBarra As String) As Long
'Calcula o Código de Barras para geração de Boleto Bancário - Alterado por Jorge Specian - 08/03/2005

Dim lErro As Long
Dim sCodigoSequencia As String
Dim dtDataBase As Date
Dim iFator As Integer
Dim iDac As Integer

On Error GoTo Erro_Calcula_CodigoBarras
    
    'Database para calculo do fator
    dtDataBase = CDate("7/10/1997")
    iFator = DateDiff("d", dtDataBase, Format(dtVencimento, "dd/mm/yyyy"))
    dValor = Arredonda_Moeda(dValor * 100, 0)
    sBanco = Format(sBanco, "000")
    sLivre = Format(sLivre, "0000000000000000000000000")
    
    'Sequencia sem o DV
    sCodigoSequencia = sBanco & sMoeda & iFator & Format(dValor, "0000000000") & sLivre
    
    'Calculo do DV
    lErro = Calcula_DV_CodBarras(sCodigoSequencia, iDac)
    If lErro <> 0 Then MsgBox ("erro")
    
    'Monta a sequencia para o codigo de barras com o DV
    sCodigoBarra = Left(sCodigoSequencia, 4) & iDac & Right(sCodigoSequencia, 39)

    Calcula_CodigoBarras = 0

    Exit Function
    
Erro_Calcula_CodigoBarras:

    MsgBox ("erro")
    
    Exit Function
    
End Function

Function Calcula_DV11(ByVal sSequencia As String, ByVal iBase As Integer, sDigito As String) As Long
'Calcula o Dígito Verificador do Nosso Numero - Alterado por Jorge Specian - 15/03/2005
'Cálculo através do módulo 11

Dim lErro As Long
Dim iContador As Integer
Dim iNumero As Integer
Dim iTotalNumero As Integer
Dim iMultiplicador As Integer
Dim iResto As Integer
Dim iResultado As Integer
Dim sCaracter As String

On Error GoTo Erro_Calcula_DV11

    iMultiplicador = 2
    
    For iContador = 1 To Len(sSequencia)
        sCaracter = Mid(Right(sSequencia, iContador), 1, 1)
        If iMultiplicador > iBase Then
            iMultiplicador = 2
        End If
        iNumero = sCaracter * iMultiplicador
        iTotalNumero = iTotalNumero + iNumero
        iMultiplicador = iMultiplicador + 1
    Next
    
    iResto = iTotalNumero Mod 11
    
    iResultado = 11 - iResto
    
    If iResultado = 10 Then
        sDigito = "P"
    ElseIf iResultado = 11 Then
        sDigito = "0"   'zero
    Else
        sDigito = CStr(iResultado)
    End If

    Calcula_DV11 = 0

    Exit Function

Erro_Calcula_DV11:

    MsgBox ("erro")
    
    Exit Function

End Function

Private Function Calcula_DV_CodBarras(ByVal sSequencia As String, iDac As Integer) As Long
'Calcula o Dígito Verificador do Código de Barras do Boleto Bancário - Alterado por Jorge Specian - 08/03/2005
'Cálculo através do módulo 11, com base de cálculo igual a 9

Dim lErro As Long
Dim iContador As Integer
Dim iNumero As Integer
Dim iTotalNumero As Integer
Dim iMultiplicador As Integer
Dim iResto As Integer
Dim iResultado As Integer
Dim sCaracter As String

On Error GoTo Erro_Calcula_DV_CodBarras

    iMultiplicador = 2
    
    For iContador = 1 To 43
       sCaracter = Mid(Right(sSequencia, iContador), 1, 1)
       If iMultiplicador > 9 Then
              iMultiplicador = 2
             iNumero = 0
       End If
       iNumero = sCaracter * iMultiplicador
       iTotalNumero = iTotalNumero + iNumero
       iMultiplicador = iMultiplicador + 1
    Next
    
    iResto = iTotalNumero Mod 11
    
    iResultado = 11 - iResto
    
    If iResultado = 10 Or iResultado = 11 Then
        iDac = 1
    Else
        iDac = iResultado
    End If

    Calcula_DV_CodBarras = 0

    Exit Function

Erro_Calcula_DV_CodBarras:

    MsgBox ("erro")

End Function

Function Calcula_NumerosCodBarras(ByVal sSequencia As String, sNumerosCodBarra As String) As Long
'Calcula os Números da Linha Digitável de um Boleto Bancário - Alterado por Jorge Specian - 08/03/2005

Dim lErro As Long
Dim sSeq1 As String
Dim sSeq2 As String
Dim sSeq3 As String
Dim sDVCodBarras As String
Dim sFatorVenc As String
Dim sValor As String
Dim iDv1 As Integer
Dim iDv2 As Integer
Dim iDv3 As Integer

On Error GoTo Erro_Calcula_NumerosCodBarras
        
    'Separa as sequencias
    sSeq1 = Left(sSequencia, 4) & Mid(sSequencia, 20, 5)
    
    sDVCodBarras = Mid(sSequencia, 5, 1)
    
    Call Calcula_DV_CodBarras(
    
    sFatorVenc = Mid(sSequencia, 6, 4)
    sValor = Mid(sSequencia, 10, 10)
    sSeq2 = Mid(sSequencia, 25, 10)
    sSeq3 = Right(sSequencia, 10)
    
    'Calcula os DVs
    lErro = Calcula_DV10(sSeq1, iDv1)
    
    lErro = Calcula_DV10(sSeq2, iDv2)
    
    lErro = Calcula_DV10(sSeq3, iDv3)
    
    'Formata as sequencias
    sSeq1 = Left(sSeq1 & iDv1, 5) & "." & Right(sSeq1 & iDv1, 5)
    sSeq2 = Left(sSeq2 & iDv2, 5) & "." & Right(sSeq2 & iDv2, 6)
    sSeq3 = Left(sSeq3 & iDv3, 5) & "." & Right(sSeq3 & iDv3, 6)
    
    'Concatena as sequencias
    sNumerosCodBarra = sSeq1 & " " & sSeq2 & " " & sSeq3 & " " & sDVCodBarras & " " & sFatorVenc & sValor

    Calcula_NumerosCodBarras = 0

    Exit Function
    
Erro_Calcula_NumerosCodBarras:

    MsgBox ("erro")
    
End Function

Function Calcula_DV10(ByVal sSequencia As String, iDigito As Integer) As Long
'Calcula o Digito Verificador no Módulo 10 para Linha Digitável de um Boleto Bancário
'Alterado por Jorge Specian - 09/03/2005

Dim lErro As Long
Dim iContador As Integer
Dim iNumero As Integer
Dim iTotalNumero As Integer
Dim iMultiplicador As Integer
Dim DezenaSuperior As Integer

On Error GoTo Erro_Calcula_DV10

    'Se nao for um valor numerico -> erro
    If Not IsNumeric(sSequencia) Then Error 134236
        
    'Inicia o multiplicador
    iMultiplicador = 2
    
    'Pega cada caracter do numero a partir da direita
    For iContador = Len(sSequencia) To 1 Step -1
        
        'Extrai o caracter e multiplica pelo multiplicador
        iNumero = Val(Mid(sSequencia, iContador, 1)) * iMultiplicador
        
        'Se o resultado for maior que nove soma os algarismos do resultado
        If iNumero > 9 Then
            
            iNumero = Val(Left(iNumero, 1)) + Val(Right(iNumero, 1))
        
        End If
        
        'Soma o resultado para totalização
        iTotalNumero = iTotalNumero + iNumero
        
        'Se o multiplicador for igual a 2 atribuir valor 1 se for 1 atribui 2
        iMultiplicador = IIf(iMultiplicador = 2, 1, 2)
        
    Next

    If iTotalNumero < 10 Then
        DezenaSuperior = 10
    Else
        DezenaSuperior = 10 * (Val(Left(CStr(iTotalNumero), 1)) + 1)
    End If
    
    iDigito = DezenaSuperior - iTotalNumero

    'verifica as exceções ( 10 -> DV=0 )
    If iDigito = 10 Then iDigito = 0
    
    Calcula_DV10 = 0

    Exit Function

Erro_Calcula_DV10:

    MsgBox ("erro")
    
    Exit Function

End Function


Public Function Arredonda_Moeda(dValor As Double, Optional ByVal iNumDigitos As Integer = 2) As Double

    If dValor >= 0 Then
        Arredonda_Moeda = Round(dValor + 0.0000000001, iNumDigitos)
    Else
        Arredonda_Moeda = Round(dValor - 0.0000000001, iNumDigitos)
    End If

End Function



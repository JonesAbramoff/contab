VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFiscalFaturaVGMgz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public gobjTelaPadrao As Object 'ct com codigo padrao
Public gobjTelaUsu As Object 'ct com customizacao

Private WithEvents objEventoNatureza As AdmEvento
Attribute objEventoNatureza.VB_VarHelpID = -1

Private gobjCT As Object
Private gobjCTUsu As Object

Public Sub objEventoNatureza_New()
    Set objEventoNatureza = New AdmEvento
End Sub

Public Sub objEventoNatureza_Nothing()
    Set objEventoNatureza = Nothing
End Sub

Sub LabelNatureza_Click(ByVal objCT As Object, ByVal objCTUsu As Object)

    Dim objNatMovCta As New ClassNatMovCta
    Dim colSelecao As New Collection

    Set gobjCT = objCT
    Set gobjCTUsu = objCTUsu
    
    objNatMovCta.sCodigo = gobjCTUsu.Natureza(objCT).ClipText
    
    colSelecao.Add NATUREZA_TIPO_RECEBIMENTO
    
    Call Chama_Tela("NatMovCtaLista", colSelecao, objNatMovCta, objEventoNatureza, "Tipo = ?")

End Sub

Private Sub objEventoNatureza_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNatMovCta As ClassNatMovCta
Dim sNaturezaEnxuta As String

On Error GoTo Erro_objEventoNatureza_evSelecao

    Set objNatMovCta = obj1

    If objNatMovCta.sCodigo = "" Then
        
        gobjCTUsu.Natureza(gobjCT).PromptInclude = False
        gobjCTUsu.Natureza(gobjCT).Text = ""
        gobjCTUsu.Natureza(gobjCT).PromptInclude = True
    
    Else

        sNaturezaEnxuta = String(STRING_NATMOVCTA_CODIGO, 0)
    
        lErro = Mascara_RetornaItemEnxuto(SEGMENTO_NATMOVCTA, objNatMovCta.sCodigo, sNaturezaEnxuta)
        If lErro <> SUCESSO Then gError 122833

        gobjCTUsu.Natureza(gobjCT).PromptInclude = False
        gobjCTUsu.Natureza(gobjCT).Text = sNaturezaEnxuta
        gobjCTUsu.Natureza(gobjCT).PromptInclude = True
    
    End If

    Call gobjCTUsu.Natureza_Validate(gobjCT, bSGECancelDummy)
    
    gobjCT.Show

    Exit Sub

Erro_objEventoNatureza_evSelecao:

    Select Case gErr

        Case 122833

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRVProdTafCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarcolTarifas As New Collection
Private mvarsProduto As String

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Set colTarifas(ByVal vData As Collection)
    Set mvarcolTarifas = vData
End Property

Public Property Get colTarifas() As Collection
    Set colTarifas = mvarcolTarifas
End Property

Public Function Calcula_Tarifa(ByVal iAno As Integer, ByVal iVigencia As Integer, dValor As Double) As Long

Dim objTaf As ClassTRVProdTarifa
Dim dTarifa As Double

On Error GoTo Erro_Calcula_Tarifa

    For Each objTaf In mvarcolTarifas
        
        If objTaf.iDiasDe <= iVigencia And (objTaf.iDiasAte >= iVigencia Or objTaf.iDiario = MARCADO) And objTaf.iAno = iAno Then
        
            If objTaf.iDiasAte < iVigencia Then
                dTarifa = objTaf.dValor + (objTaf.dValorAdicional * (objTaf.iDiasAte - iVigencia))
            Else
                dTarifa = objTaf.dValor
            End If
        
            Exit For
        End If
    
    Next
    
    dValor = dTarifa

    Calcula_Tarifa = SUCESSO

    Exit Function

Erro_Calcula_Tarifa:

    Calcula_Tarifa = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192205)

    End Select

    Exit Function
    
End Function


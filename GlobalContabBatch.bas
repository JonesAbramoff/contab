Attribute VB_Name = "GlobalContabBatch"
Option Explicit

Public Const INSERE_LANCAMENTO_AGLUTINADO = 1 'indica se deve inserir um novo lançamento aglutinado
Public Const TEM_LANCAMENTO_AGLUTINADO = 1 'indica se tem lançamento aglutinado acumulado para ser gravado

Type typeFaixaPeriodo
    dtDataInicio As Date
    dtDataFim As Date
End Type

Type typeLote_batch
    iFilialEmpresa As Integer
    sOrigem As String
    iExercicio As Integer
    iPeriodo As Integer
    iLote As Integer
End Type

Type typeProcessa_Lancamento
    tLancamento As typeLancamento
    iFilialEmpresa As Integer
    iExercicio As Integer
    iPeriodo As Integer
    lID_Arq_Temp As Long
    lID_Arq_Sort As Long
    lID_Arq_Sort1 As Long
    lID_Arq_Sort2 As Long
    dDebito As Double
    dCredito As Double
    lComando2 As Long
    lComando3 As Long
    lComando4 As Long
    lComando5 As Long
    lComando6 As Long
    lComando7 As Long
    lComando8 As Long
    lComando9 As Long
    lComando10 As Long
    lComando11 As Long
    lComando12 As Long
    lComando13 As Long
    lComando14 As Long
    lComando15 As Long
    lComando16 As Long
    lComando17 As Long
    lComando18 As Long
    lComando19 As Long
    lComando20 As Long
    lComando21 As Long
    iFim_de_Arquivo As Integer
    sPeriodo As String
    iOperacao As Integer
    iUsoCcl As Integer
    iOperacao1 As Integer
    lDocAglutinado As Long
    iSeqAglutinado As Integer
    dValorAglutinado As Double
    iAglutinaLancamPorDia As Integer 'indica se o modulo que gerou o lote quer gerar lançamentos aglutinados por dia ou não
    iUltSeqAglutinado As Integer 'ultimo sequencial de documento de aglutinacao de uma data utilizado
    dtDataAglutinado As Date
    iInsereLancamAglutinado As Integer
    sCclAglutinado As String
    iSeqAglutinadoContaCcl As Integer
    iTemLancamAglutinado As Integer
    alComando(1 To 15) As Long
    alComando1(4 To 15) As Long 'esta numeração é para compatiblizar com a rotina CustoMedio_Le que também é usada por LanPendente_Trata_Produto
    colExercicio As Collection 'usado no reprocessamento - guarda os exercicios que foram verificados quanto ao fechamento
End Type





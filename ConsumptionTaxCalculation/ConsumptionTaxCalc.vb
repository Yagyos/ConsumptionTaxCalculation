Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

Public Class ConsumptionTaxCalc

    Public OriginalAmount As Decimal '算出元の金額
    Public TaxRate As Decimal '税率 8%→[8] 10%→[10]
    Public TaxRound As RoundCalc.RoundType '丸め種類
    Public RoundDigits As Integer '丸め桁数 0:小数点以下丸め(12.3→12) 1:(12→10) -1:(1.23→1.2)
    Public OutsideTaxAmount As Decimal '外税税額
    Public InsideTaxAmount As Decimal '内税税額
    Public InsideNetAmount As Decimal '内税正味金額

    ''' <summary>
    ''' 税計算実施
    ''' </summary>
    ''' <remarks></remarks>
    ''' 
    Public Sub Calculation()

        '外税税額算出
        Call CalculationOutside()

        '内税税額算出
        Call CalculationInside()

    End Sub

    ''' <summary>
    ''' 外税税額算出 丸め処理も実施
    ''' </summary>
    ''' <returns>外税税額</returns>
    ''' <remarks></remarks>
    Public Function CalculationOutside() As Decimal

        Dim TaxAmount As Decimal

        '外税税額算出
        TaxAmount = CalculationOutsideTax(OriginalAmount, TaxRate)
        OutsideTaxAmount = RoundCalc.Calculation(TaxAmount, TaxRound, RoundDigits)

        Return OutsideTaxAmount
    End Function

    ''' <summary>
    ''' 外税税額算出 丸め処理前
    ''' </summary>
    ''' <param name="OutsideNetAmount"></param>
    ''' <param name="TaxRate"></param>
    ''' <returns>外税税額</returns>
    ''' <remarks></remarks>
    Public Function CalculationOutsideTax(ByVal OutsideNetAmount As Decimal, ByVal TaxRate As Decimal) As Decimal

        '[8%]を[0.08]に
        '[10%]を[0.1]に
        TaxRate = TaxRate / 100

        '丸め前の税額を返す
        Return OutsideNetAmount * TaxRate
    End Function

    ''' <summary>
    ''' 内税税額算出 丸め処理も実施 税別金額はInsideNetAmount
    ''' </summary>
    ''' <returns>内税税額 InsideTaxAmount:内税税額 InsideNetAmount:税別金額</returns>
    ''' <remarks></remarks>
    Public Function CalculationInside() As Decimal

        Dim TaxAmount As Decimal

        '内税税額算出
        TaxAmount = CalculationInsideTax(OriginalAmount, TaxRate)
        InsideTaxAmount = RoundCalc.Calculation(TaxAmount, TaxRound, RoundDigits)

        '内税正味金額算出
        InsideNetAmount = OriginalAmount - InsideTaxAmount

        Return InsideTaxAmount
    End Function

    ''' <summary>
    ''' 内税税額算出 丸め処理前
    ''' </summary>
    ''' <param name="InsideNetAmount"></param>
    ''' <param name="TaxRate"></param>
    ''' <returns>内税税額</returns>
    ''' <remarks></remarks>
    Public Function CalculationInsideTax(ByVal InsideNetAmount As Decimal, ByVal TaxRate As Decimal) As Decimal

        '[8%]を[1.08]に
        '[10%]を[1.1]に
        TaxRate = (100 + TaxRate) / 100

        '内税正味金額算出
        Dim OutsideNetAmount As Decimal
        OutsideNetAmount = InsideNetAmount / TaxRate

        '丸め前の税額を返す
        Return InsideNetAmount - OutsideNetAmount
    End Function

End Class

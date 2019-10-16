Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

Public Module ConsumptionTaxCalc

    Structure DataSet
        Dim OriginalAmount As Decimal '算出元の金額
        Dim TaxRate As Decimal '税率 8%→[8] 10%→[10]
        Dim TaxRound As RoundCalc.RoundType '丸め種類
        Dim RoundDigits As Integer '丸め桁数 0:小数点以下丸め(12.3→12) 1:(12→10) -1:(1.23→1.2)
        Dim OutsideTaxAmount As Decimal '外税税額
        Dim InsideTaxAmount As Decimal '内税税額
        Dim InsideNetAmount As Decimal '内税正味金額
    End Structure

    ''' <summary>
    ''' 税計算実施
    ''' </summary>
    ''' <param name="ConTaxCalData"></param>
    ''' <remarks></remarks>
    Public Sub Calculation(ByRef ConTaxCalData As DataSet)

        '外税税額算出
        Call CalculationOutside(ConTaxCalData)

        '内税税額算出
        Call CalculationInside(ConTaxCalData)

    End Sub

    ''' <summary>
    ''' 外税税額算出 丸め処理も実施
    ''' </summary>
    ''' <param name="ConTaxCalData"></param>
    ''' <returns>外税税額</returns>
    ''' <remarks></remarks>
    Public Function CalculationOutside(ByRef ConTaxCalData As DataSet) As Decimal

        Dim TaxAmount As Decimal

        '外税税額算出
        TaxAmount = CalculationOutsideTax(ConTaxCalData.OriginalAmount, ConTaxCalData.TaxRate)
        ConTaxCalData.OutsideTaxAmount = RoundCalc.Calculation(TaxAmount, ConTaxCalData.TaxRound, ConTaxCalData.RoundDigits)

        Return ConTaxCalData.OutsideTaxAmount
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
    ''' <param name="ConTaxCalData"></param>
    ''' <returns>内税税額 InsideTaxAmount:内税税額 InsideNetAmount:税別金額</returns>
    ''' <remarks></remarks>
    Public Function CalculationInside(ByRef ConTaxCalData As DataSet) As Decimal

        Dim TaxAmount As Decimal

        '内税税額算出
        TaxAmount = CalculationInsideTax(ConTaxCalData.OriginalAmount, ConTaxCalData.TaxRate)
        ConTaxCalData.InsideTaxAmount = RoundCalc.Calculation(TaxAmount, ConTaxCalData.TaxRound, ConTaxCalData.RoundDigits)

        '内税正味金額算出
        ConTaxCalData.InsideNetAmount = ConTaxCalData.OriginalAmount - ConTaxCalData.InsideTaxAmount

        Return ConTaxCalData.InsideTaxAmount
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

End Module

Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim ConTaxCalData As New ConsumptionTaxCalculation_Str

        ConTaxCalData.OriginalAmount = 100
        ConTaxCalData.TaxRound = CTC_RoundType.Round_off
        ConTaxCalData.RoundDigits = 0
        ConTaxCalData.TaxRate = 10

        Call ConsumptionTaxCalculation(ConTaxCalData)

        Button1.Text = ConTaxCalData.OriginalAmount & "/" _
                     & ConTaxCalData.OutsideTaxAmount & vbCrLf _
                     & ConTaxCalData.InsideNetAmount & "/" _
                     & ConTaxCalData.InsideTaxAmount

    End Sub

    Enum CTC_RoundType
        Round_off  '四捨五入
        Round_down '切り捨て
        Round_up   '切り上げ
    End Enum

    Structure ConsumptionTaxCalculation_Str
        Dim OriginalAmount As Decimal '算出元の金額
        Dim TaxRate As Decimal '税率 8%→[8] 10%→[10]
        Dim TaxRound As CTC_RoundType '丸め種類
        Dim RoundDigits As Integer '丸め桁数 0:小数点以下丸め(12.3→12) 1:(12→10) -1:(1.23→1.2)
        Dim OutsideTaxAmount As Decimal '外税税額
        Dim InsideTaxAmount As Decimal '内税税額
        Dim InsideNetAmount As Decimal '内税正味金額
    End Structure

    Public Sub ConsumptionTaxCalculation(ByRef ConTaxCalData As ConsumptionTaxCalculation_Str)

        Dim TaxAmount As Decimal

        '外税税額算出
        TaxAmount = CalcOutsideTax(ConTaxCalData.OriginalAmount, ConTaxCalData.TaxRate)
        ConTaxCalData.OutsideTaxAmount = RoundCalc(TaxAmount, ConTaxCalData.TaxRound, ConTaxCalData.RoundDigits)

        '内税税額算出
        TaxAmount = CalcInsideTax(ConTaxCalData.OriginalAmount, ConTaxCalData.TaxRate)
        ConTaxCalData.InsideTaxAmount = RoundCalc(TaxAmount, ConTaxCalData.TaxRound, ConTaxCalData.RoundDigits)

        '内税正味金額算出
        ConTaxCalData.InsideNetAmount = ConTaxCalData.OriginalAmount - ConTaxCalData.InsideTaxAmount

    End Sub

    '外税税額算出
    Public Function CalcOutsideTax(ByVal OutsideNetAmount As Decimal, ByVal TaxRate As Decimal) As Decimal

        '[8%]を[0.08]に
        '[10%]を[0.1]に
        TaxRate = TaxRate / 100

        '丸め前の税額を返す
        Return OutsideNetAmount * TaxRate
    End Function

    '内税税額算出
    Public Function CalcInsideTax(ByVal InsideNetAmount As Decimal, ByVal TaxRate As Decimal) As Decimal

        '[8%]を[1.08]に
        '[10%]を[1.1]に
        TaxRate = (100 + TaxRate) / 100

        '内税正味金額算出
        Dim OutsideNetAmount As Decimal
        OutsideNetAmount = InsideNetAmount / TaxRate

        '丸め前の税額を返す
        Return InsideNetAmount - OutsideNetAmount
    End Function

    '丸め実施
    Public Function RoundCalc(ByVal TaxAmount As Decimal, ByVal TaxRound As CTC_RoundType, ByVal RoundDigits As Integer) As Decimal

        '負や正の無限大への丸めや
        '銀行丸めがあるので注意

        '無限大方向丸め対策として
        'マイナスはプラスとして丸め処理する
        Dim Minus As Decimal
        If TaxAmount < 0 Then
            Minus = -1
            TaxAmount = TaxAmount * -1
        Else
            Minus = 1
        End If

        '丸め計算用に桁数を合わせる
        TaxAmount = TaxAmount / CDec(10 ^ RoundDigits)

        Select Case TaxRound
            Case CTC_RoundType.Round_down
                '切り捨て
                TaxAmount = Math.Truncate(TaxAmount)
            Case CTC_RoundType.Round_up
                '切り上げ
                TaxAmount = Math.Ceiling(TaxAmount)
            Case Else
                '四捨五入

                '銀行丸めに対処するため
                '0.5を足して切り捨てを実施
                TaxAmount = TaxAmount + 0.5D
                TaxAmount = Math.Truncate(TaxAmount)

        End Select

        '合わせた桁数を戻す
        TaxAmount = TaxAmount * CDec(10 ^ RoundDigits)

        Return TaxAmount * Minus
    End Function

End Class

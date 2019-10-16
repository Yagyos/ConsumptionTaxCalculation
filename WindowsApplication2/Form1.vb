Public Class Form1

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        Dim ConTaxCalData As New ConsumptionTaxCalculation_Str

        ConTaxCalData.OriginalAmount = 100
        ConTaxCalData.Round = CTC_RoundType.Round_off
        ConTaxCalData.TaxRate = 10

        Call ConsumptionTaxCalculation(ConTaxCalData)

        Button1.Text = 0

    End Sub

    Enum CTC_RoundType
        Round_off  '四捨五入
        Round_down '切り捨て
        Round_up   '切り上げ
    End Enum

    Structure ConsumptionTaxCalculation_Str
        Dim OriginalAmount As Decimal '算出元の金額
        Dim TaxRate As Decimal '税率 8%→[8] 10%→[10]
        Dim Round As CTC_RoundType
        Dim OutsideTaxAmount As Decimal '外税税額
        Dim InsideTaxAmount As Decimal '内税税額
        Dim InsideNetAmount As Decimal '内税正味金額
    End Structure

    Public Sub ConsumptionTaxCalculation(ByRef ConTaxCalData As ConsumptionTaxCalculation_Str)

        'これから実装

        'テスト修正


    End Sub

End Class

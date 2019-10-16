Public Class Form1

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

End Class

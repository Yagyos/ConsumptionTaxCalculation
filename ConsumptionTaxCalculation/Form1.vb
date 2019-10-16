Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        'Call RoundCalc.SelfCheck()

        Dim ConTaxCalData As New ConsumptionTaxCalc.DataSet

        ConTaxCalData.OriginalAmount = 100
        ConTaxCalData.TaxRound = RoundCalc.RoundType.Round_off
        ConTaxCalData.RoundDigits = 0
        ConTaxCalData.TaxRate = 10

        Call ConsumptionTaxCalc.Calculation(ConTaxCalData)

        Button1.Text = ConTaxCalData.OriginalAmount & "/" _
                     & ConTaxCalData.OutsideTaxAmount & vbCrLf _
                     & ConTaxCalData.InsideNetAmount & "/" _
                     & ConTaxCalData.InsideTaxAmount

    End Sub

End Class

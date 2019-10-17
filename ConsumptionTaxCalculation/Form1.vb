Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        'Call RoundCalc.SelfCheck()

        Dim ConTaxCalData As New ConsumptionTaxCalc

        With ConTaxCalData

            .OriginalAmount = 100
            .TaxRound = RoundCalc.RoundType.Round_off
            .RoundDigits = 0
            .TaxRate = 10

            Call .Calculation()

            Button1.Text = .OriginalAmount & "/" _
                         & .OutsideTaxAmount & vbCrLf _
                         & .InsideNetAmount & "/" _
                         & .InsideTaxAmount

        End With

    End Sub

End Class

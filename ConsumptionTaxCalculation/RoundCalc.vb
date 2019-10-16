Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

Public Module RoundCalc

    ''' <summary>
    ''' 丸め種類 off:四捨五入 down:切り捨て up:切り上げ
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum RoundType
        Round_off  '四捨五入
        Round_down '切り捨て
        Round_up   '切り上げ
    End Enum

    ''' <summary>
    ''' 丸め実施(小数点以下を丸め)
    ''' </summary>
    ''' <param name="TaxAmount">丸め元の金額</param>
    ''' <param name="TaxRound"></param>
    ''' <returns>丸められた金額</returns>
    ''' <remarks></remarks>
    Public Function Calculation(ByVal TaxAmount As Decimal, ByVal TaxRound As RoundType) As Decimal
        Return Calculation(TaxAmount, TaxRound, 0)
    End Function

    ''' <summary>
    ''' 丸め実施
    ''' </summary>
    ''' <param name="TaxAmount">丸め元の金額</param>
    ''' <param name="TaxRound"></param>
    ''' <param name="RoundDigits">丸め桁数 0:小数点以下丸め(12.3→12) 1:(12→10) -1:(1.23→1.2)</param>
    ''' <returns>丸められた金額</returns>
    ''' <remarks></remarks>
    Public Function Calculation(ByVal TaxAmount As Decimal, ByVal TaxRound As RoundType, ByVal RoundDigits As Integer) As Decimal

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
            Case RoundType.Round_down
                '切り捨て
                TaxAmount = Math.Truncate(TaxAmount)
            Case RoundType.Round_up
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

    'セルフチェック
    Public Sub SelfCheck()

        If Calculation(123.456D, RoundType.Round_off) <> 123 Then
            Debug.Assert(False)
        End If
        If Calculation(123.56D, RoundType.Round_off) <> 124 Then
            Debug.Assert(False)
        End If
        If Calculation(123.456D, RoundType.Round_up) <> 124 Then
            Debug.Assert(False)
        End If
        If Calculation(123.56D, RoundType.Round_down) <> 123 Then
            Debug.Assert(False)
        End If
        If Calculation(555D, RoundType.Round_off, 1) <> 560 Then
            Debug.Assert(False)
        End If
        If Calculation(555.555D, RoundType.Round_off, -1) <> 555.6 Then
            Debug.Assert(False)
        End If

    End Sub

End Module

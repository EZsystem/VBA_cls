VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "acc_clsDataCleaner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================
' Class Module: acc_clsDataCleaner
' 説明　：データ型変換とクレンジングを一括で提供する汎用クラスにゃ！
' 作成日：2025/04/26
' 作成者：そうじろう
'=====================================
Option Compare Database
Option Explicit

'=====================================
' 関数名 : CleanText
' 説明　 : NULL補完＋Trimして安全な文字列を返すにゃ
'=====================================
Public Function CleanText(val As Variant) As String
    If IsNull(val) Then
        CleanText = ""
    Else
        CleanText = Trim(CStr(val))
    End If
End Function

'=====================================
' 関数名 : CleanNumber
' 説明　 : NULL補完＋数値型に安全変換するにゃ
'=====================================
Public Function CleanNumber(val As Variant, Optional defaultValue As Double = 0) As Double
    If IsNull(val) Or Not IsNumeric(val) Then
        CleanNumber = defaultValue
    Else
        CleanNumber = CDbl(val)
    End If
End Function

'=====================================
' 関数名 : CleanDate
' 説明　 : NULL補完＋日付型に安全変換するにゃ
'=====================================
Public Function CleanDate(val As Variant, Optional defaultDate As Date = #1/1/1900#) As Date
    If IsNull(val) Or Not IsDate(val) Then
        CleanDate = defaultDate
    Else
        CleanDate = CDate(val)
    End If
End Function

'=====================================
' 関数名 : TextToInteger
' 説明　 : テキスト型をInteger型に安全変換するにゃ
'=====================================
Public Function TextToInteger(val As Variant, Optional defaultValue As Integer = 0) As Integer
    If IsNull(val) Or Trim(val) = "" Then
        TextToInteger = defaultValue
    ElseIf IsNumeric(val) Then
        TextToInteger = CInt(val)
    Else
        TextToInteger = defaultValue
    End If
End Function

'=====================================
' 関数名 : TextToLong
' 説明　 : テキスト型をLong型に安全変換するにゃ
'=====================================
Public Function TextToLong(val As Variant, Optional defaultValue As Long = 0) As Long
    If IsNull(val) Or Trim(val) = "" Then
        TextToLong = defaultValue
    ElseIf IsNumeric(val) Then
        TextToLong = CLng(val)
    Else
        TextToLong = defaultValue
    End If
End Function

'=====================================
' 関数名 : TextToSingle
' 説明　 : テキスト型をSingle型に安全変換するにゃ
'=====================================
Public Function TextToSingle(val As Variant, Optional defaultValue As Single = 0) As Single
    If IsNull(val) Or Trim(val) = "" Then
        TextToSingle = defaultValue
    ElseIf IsNumeric(val) Then
        TextToSingle = CSng(val)
    Else
        TextToSingle = defaultValue
    End If
End Function

'=====================================
' 関数名 : TextToDouble
' 説明　 : テキスト型をDouble型に安全変換するにゃ
'=====================================
Public Function TextToDouble(val As Variant, Optional defaultValue As Double = 0) As Double
    If IsNull(val) Or Trim(val) = "" Then
        TextToDouble = defaultValue
    ElseIf IsNumeric(val) Then
        TextToDouble = CDbl(val)
    Else
        TextToDouble = defaultValue
    End If
End Function

'=====================================
' 関数名 : TextToCurrency
' 説明　 : テキスト型をCurrency型に安全変換するにゃ
'=====================================
Public Function TextToCurrency(val As Variant, Optional defaultValue As Currency = 0) As Currency
    If IsNull(val) Or Trim(val) = "" Then
        TextToCurrency = defaultValue
    ElseIf IsNumeric(val) Then
        TextToCurrency = CCur(val)
    Else
        TextToCurrency = defaultValue
    End If
End Function

'=====================================
' 関数名 : TextToDate
' 説明　 : テキスト型を日付型に安全変換するにゃ
'=====================================
Public Function TextToDate(val As Variant, Optional defaultDate As Date = #1/1/1900#) As Date
    If IsNull(val) Or Trim(val) = "" Then
        TextToDate = defaultDate
    ElseIf IsDate(val) Then
        TextToDate = CDate(val)
    Else
        TextToDate = defaultDate
    End If
End Function

'=====================================
' 関数名 : TextToBoolean
' 説明　 : テキスト型をブール型に変換するにゃ
'=====================================
Public Function TextToBoolean(val As Variant, Optional defaultValue As Boolean = False) As Boolean
    Dim txt As String
    If IsNull(val) Then
        TextToBoolean = defaultValue
        Exit Function
    End If
    txt = LCase(Trim(CStr(val)))
    Select Case txt
        Case "true", "yes", "1"
            TextToBoolean = True
        Case "false", "no", "0"
            TextToBoolean = False
        Case Else
            TextToBoolean = defaultValue
    End Select
End Function

'=====================================
' 関数名 : IsNumericStrict
' 説明　 : 厳密な数値判定をするにゃ
'=====================================
Public Function IsNumericStrict(val As Variant) As Boolean
    IsNumericStrict = False
    If Not IsNull(val) Then
        If IsNumeric(val) Then IsNumericStrict = True
    End If
End Function

'=====================================
' 関数名 : IsDateStrict
' 説明　 : 厳密な日付判定をするにゃ
'=====================================
Public Function IsDateStrict(val As Variant) As Boolean
    IsDateStrict = False
    If Not IsNull(val) Then
        If IsDate(val) Then IsDateStrict = True
    End If
End Function

'=================================================
' 関数名 : GetSupportedVBAType
' 説明   : 日本語のAccess型名をVBA型名に変換する
'         対応していない型名の場合はエラーを返す
' 引数   : jpType（String） 日本語型名（例："長整数型"）
' 戻り値 : String（"Long", "Date" など）
'=================================================
Public Function GetSupportedVBAType(ByVal jptype As String) As String
    Dim typeName As String
    typeName = Trim(jptype)

    Select Case typeName
        Case "テキスト型"
            GetSupportedVBAType = "String"

        Case "長整数型"
            GetSupportedVBAType = "Long"

        Case "整数型"
            GetSupportedVBAType = "Integer"

        Case "単精度浮動小数点型"
            GetSupportedVBAType = "Single"

        Case "倍精度浮動小数点型"
            GetSupportedVBAType = "Double"

        Case "通貨型"
            GetSupportedVBAType = "Currency"

        Case "日付/時刻型"
            GetSupportedVBAType = "Date"

        Case "Yes/No型"
            GetSupportedVBAType = "Boolean"

        Case Else
            Debug.Print "【GetSupportedVBAType エラー】未対応の型：" & jptype
            MsgBox "未対応の型です：" & jptype, vbExclamation, "型変換エラー"
            Err.Raise vbObjectError + 500, "GetSupportedVBAType", "未対応の型：" & jptype
    End Select
End Function

'=====================================
' 関数名 : TextToString
' 説明　 : テキスト型に安全に変換するにゃ（Null→""）
'=====================================
Public Function TextToString(val As Variant, Optional defaultValue As String = "") As String
    If IsNull(val) Then
        TextToString = defaultValue
    Else
        TextToString = CStr(val)
    End If
End Function




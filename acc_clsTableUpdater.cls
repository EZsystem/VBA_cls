VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "acc_clsTableUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================
' Class Module: acc_clsTableUpdater
' 説明　：テーブルのレコードセットを操作しフィールド更新を行うクラス
' 作成日：2025/04/25
' 更新日：再生成 by そうじろう
'=====================================
Option Compare Database
Option Explicit

' --- 内部保持変数 ---
Private pDB    As DAO.Database
Private pTable As String

'=================================================
' サブルーチン名 : Init
' 引数 : tableName - 操作対象のテーブル名
' 機能 : CurrentDb を初期化し、対象テーブル名を内部に保持するにゃ
'=================================================
Public Sub Init(tableName As String)
    Set pDB = CurrentDb
    pTable = tableName
End Sub

'=================================================
' 関数名 : OpenRS
' 戻り値 : DAO.Recordset - 対象テーブルのレコードセット（Dynaset）
' 機能 : 指定テーブルを開いて編集可能なレコードセットを返すにゃ
'=================================================
Private Function OpenRS() As DAO.Recordset
    Set OpenRS = pDB.OpenRecordset(pTable, dbOpenDynaset)
End Function

'=================================================
' サブルーチン名 : UpdateFieldFromField
' 引数 : src - コピー元のフィールド名
'        tgt - 更新対象のフィールド名
' 機能 : src の値を tgt にコピーするフィールド更新処理にゃ（NULL補完付き）
'=================================================
Public Sub UpdateFieldFromField(src As String, tgt As String)
    Dim rs As DAO.Recordset
    Set rs = OpenRS()
    Do Until rs.EOF
        rs.Edit
        rs.Fields(tgt).value = Nz(rs.Fields(src).value, Null)
        rs.Update
        rs.MoveNext
    Loop
    rs.Close
End Sub

'=================================================
' サブルーチン名 : UpdateFieldFromDiff
' 引数 : a - 被減数フィールド名
'        b - 減数フィールド名
'        tgt - 結果を格納するフィールド名
' 機能 : a - b の結果を tgt に格納する差分計算処理にゃ
'=================================================
Public Sub UpdateFieldFromDiff(a As String, b As String, tgt As String)
    Dim rs As DAO.Recordset
    Set rs = OpenRS()
    Do Until rs.EOF
        rs.Edit
        rs.Fields(tgt).value = Nz(rs.Fields(a).value, 0) - Nz(rs.Fields(b).value, 0)
        rs.Update
        rs.MoveNext
    Loop
    rs.Close
End Sub

'=================================================
' サブルーチン名 : UpdateFieldByConcat
' 引数 : a, b - 結合対象のフィールド名
'        tgt - 結果格納先フィールド名
' 機能 : a と b の値を "a_b" の形式で連結し、tgt に格納するにゃ
'=================================================
Public Sub UpdateFieldByConcat(a As String, b As String, tgt As String)
    Dim rs As DAO.Recordset
    Set rs = OpenRS()
    Do Until rs.EOF
        rs.Edit
        rs.Fields(tgt).value = Nz(rs.Fields(a).value, "") & "_" & Nz(rs.Fields(b).value, "")
        rs.Update
        rs.MoveNext
    Loop
    rs.Close
End Sub

'=================================================
' サブルーチン名 : UpdateFieldWithCondition
' 引数 : condField - 判定用フィールド名
'        condVal - 比較値
'        tgt - 更新対象フィールド名
'        valField - 条件成立時の値取得元フィールド
'        altFunc - 条件不成立時に呼び出す関数名（Application.Run）
' 機能 : 条件に応じて valField か altFunc の結果を tgt に設定するにゃ
'=================================================
Public Sub UpdateFieldWithCondition(condField As String, condVal As String, tgt As String, valField As String, altFunc As String)
    Dim rs As DAO.Recordset
    Set rs = OpenRS()
    Do Until rs.EOF
        Dim v As Variant
        If Nz(rs.Fields(condField).value, "") = condVal Then
            v = Nz(rs.Fields(valField).value, 0)
        Else
            v = Application.Run(altFunc, rs)
        End If
        rs.Edit
        rs.Fields(tgt).value = v
        rs.Update
        rs.MoveNext
    Loop
    rs.Close
End Sub


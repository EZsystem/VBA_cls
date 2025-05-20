VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "acc_clsFieldReplacer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================
' Class Module: acc_clsFieldReplacer
' 説明　：マスタ定義に基づき、フィールド値またはタイトル行を置換する汎用クラス
' 作成日：2025/04/30
' 更新日：-
'=====================================
Option Compare Database
Option Explicit

' --- 内部保持変数 ---
Private pDB As DAO.Database
Private pTableName As String
Private pTargetFields As Variant
Private pDictMapping As Object

'=================================================
' 関数名 : Init
' 説明   : DAOデータベースの初期化
' 引数   : なし
'=================================================
Public Sub Init()
    Set pDB = CurrentDb
    Set pDictMapping = CreateObject("Scripting.Dictionary")
End Sub

'=================================================
' 関数名 : SetTargetTable
' 説明   : 対象となるテーブル名を設定する
' 引数   : tableName（String）
'=================================================
Public Sub SetTargetTable(tableName As String)
    pTableName = tableName
End Sub

'=================================================
' 関数名 : SetTargetFields
' 説明   : 処理対象のフィールド名を限定的に指定する
' 引数   : fieldNames（Variant：配列）
'=================================================
Public Sub SetTargetFields(fieldNames As Variant)
    pTargetFields = fieldNames
End Sub

'=================================================
' 関数名 : LoadMapping
' 説明   : マスタテーブルから置換マッピングを読み込む
' 引数   : mappingTable（String）
'        : keyField（String）
'        : valueField（String）
'=================================================
Public Sub LoadMapping(mappingTable As String, keyField As String, valueField As String)
    Dim rs As DAO.Recordset
    Set rs = pDB.OpenRecordset(mappingTable, dbOpenSnapshot)

    Dim k As String
    Do Until rs.EOF
        k = Nz(rs.Fields(keyField).value, "")
        If Len(k) > 0 Then
            If pDictMapping.Exists(k) Then
                Debug.Print "【重複キー検出】：" & k & " → " & pDictMapping(k)
                Err.Raise vbObjectError + 1001, , "マスタに重複キーがあります：" & k
            End If
            pDictMapping.Add k, Nz(rs.Fields(valueField).value, "")
        End If
        rs.MoveNext
    Loop
    rs.Close
End Sub

'=================================================
' 関数名 : ReplaceFieldNames
' 説明   : フィールド名（F1など）をキーに、1レコードの値を置換する
' 引数   : なし
'=================================================
Public Sub ReplaceFieldNames()
    Dim rs As DAO.Recordset
    Set rs = pDB.OpenRecordset(pTableName, dbOpenDynaset)

    If rs.EOF Then Exit Sub
    rs.MoveFirst
    rs.Edit

    Dim fld As DAO.Field
    For Each fld In rs.Fields
        If pDictMapping.Exists(fld.Name) Then
            rs.Fields(fld.Name).value = pDictMapping(fld.Name)
        End If
    Next fld

    rs.Update
    rs.Close
End Sub

'=================================================
' 関数名 : ReplaceFieldValues
' 説明   : 各レコードの中身を置換マスタに基づいて更新する
' 引数   : Optional whereCondition（String）← レコードの絞り込み条件
'=================================================
Public Sub ReplaceFieldValues(Optional whereCondition As String = "")
    Dim sql As String
    sql = "SELECT * FROM [" & pTableName & "]"
    If Len(whereCondition) > 0 Then
        sql = sql & " WHERE " & whereCondition
    End If

    Dim rs As DAO.Recordset
    Set rs = pDB.OpenRecordset(sql, dbOpenDynaset)

    If rs.EOF Then Exit Sub

    Do Until rs.EOF
        rs.Edit
        Dim fld As DAO.Field
        For Each fld In rs.Fields
            If Not IsMissing(pTargetFields) Then
                If IsInArray(fld.Name, pTargetFields) Then
                    ReplaceIfMatch rs, fld
                End If
            Else
                ReplaceIfMatch rs, fld
            End If
        Next fld
        rs.Update
        rs.MoveNext
    Loop

    rs.Close
End Sub

'=================================================
' 関数名 : ReplaceIfMatch
' 説明   : 指定されたフィールドに対して、辞書に一致すれば置換する
' 引数   : rs（DAO.Recordset）, fld（DAO.Field）
'=================================================
Private Sub ReplaceIfMatch(rs As DAO.Recordset, fld As DAO.Field)
    Dim val As String
    val = Nz(fld.value, "")
    If pDictMapping.Exists(val) Then
        rs.Fields(fld.Name).value = pDictMapping(val)
    End If
End Sub

'=================================================
' 関数名 : IsInArray
' 説明   : 指定した文字列が配列内に存在するか確認する
' 引数   : val（String）, arr（Variant）
'=================================================
Private Function IsInArray(val As String, arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = val Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function



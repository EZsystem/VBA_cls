VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "acc_clsTableCleaner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================
' Class Module: acc_clsTableCleaner
' 説明　：転写元のキーに基づいて、転写先から対象行を削除する専用クラスにゃ
' 作成日：2025/04/25
' 更新日：2025/05/08
'=====================================
Option Compare Database
Option Explicit

' --- 内部保持変数 ---
Private pDB        As DAO.Database
Private pSettings  As Collection
Private pTables As Collection

'=================================================
' サブルーチン名 : Init
' 引数 : DB（省略可）- 使用する DAO.Database（指定なし時は CurrentDb）
' 機能 : データベースを初期化し、削除設定リストを作成するにゃ
'=================================================
Public Sub Init(Optional db As DAO.Database)
    If db Is Nothing Then
        Set pDB = CurrentDb
    Else
        Set pDB = db
    End If
    Set pSettings = New Collection
    Set pTables = New Collection
End Sub

'=================================================
' サブルーチン名 : AddSetting
' 引数 : srcTable（String）- 転写元テーブル名
'        tgtTable（String）- 転写先テーブル名
'        keyField（String）- 削除対象キー列名
' 機能 : 1件の削除設定（元・先・キー）を登録するにゃ
'=================================================
Public Sub AddSetting(srcTable As String, tgtTable As String, keyField As String)
    Dim setting As Collection
    Set setting = New Collection
    setting.Add srcTable, "Source"
    setting.Add tgtTable, "Target"
    setting.Add keyField, "Key"
    pSettings.Add setting
End Sub

'=================================================
' サブルーチン名 : CleanTarget
' 機能 : 登録されたすべての設定に対して、転写先テーブルの対象行を削除するにゃ
'=================================================
Public Sub CleanTarget()
    Dim setting As Collection
    For Each setting In pSettings
        Call DeleteTargetByKey( _
            CStr(setting("Source")), _
            CStr(setting("Target")), _
            CStr(setting("Key")))
    Next
End Sub

'=================================================
' サブルーチン名 : DeleteTargetByKey
' 引数 : src（String）- 転写元テーブル名
'        tgt（String）- 転写先テーブル名
'        key（String）- 比較用キー列名
' 機能 : src に存在するキーと一致する tgt の行を削除するにゃ
'=================================================
Private Sub DeleteTargetByKey(src As String, tgt As String, key As String)
    If pDB Is Nothing Then
        Err.Raise vbObjectError + 1101, "acc_clsTableCleaner", _
                  "DBが初期化されていないにゃ。Init を先に呼んでにゃ。"
    End If

    Dim sql As String
    sql = "DELETE FROM [" & tgt & "] " & _
          "WHERE [" & key & "] IN (" & _
          "SELECT [" & key & "] FROM [" & src & "] WHERE [" & key & "] IS NOT NULL);"

    pDB.Execute sql, dbFailOnError
End Sub

'=================================================
' サブルーチン名 : AddTable
' 引数 : tableName - 操作対象として管理するテーブル名
' 機能 : クレンジングや削除対象として保持するテーブル名をリストに追加するにゃ
'=================================================
Public Sub AddTable(TableName As String)
    pTables.Add TableName
End Sub

'=================================================
' サブルーチン名 : ClearAll
' 機能 : 登録されている全テーブルに対して DELETE 文を発行し、中身を全削除するにゃ
'        登録済みテーブルは `AddTable` 経由で追加されたものにゃ
'=================================================
Public Sub ClearAll()
    Dim t As Variant
    For Each t In pTables
        pDB.Execute "DELETE FROM [" & t & "];", dbFailOnError
    Next
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "acc_clsTableTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================
' Class Module: acc_clsTableTransfer
' 説明　：指定テーブル間のデータ削除および挿入による転送機能を提供するクラス
' 作成日：2025/04/25
' 更新日：再生成 by そうじろう
'=====================================
Option Compare Database
Option Explicit

' --- 内部保持変数 ---
Private pDB        As DAO.Database
Private pSettings  As Collection

'=================================================
' サブルーチン名 : Init
' 引数 : DB（省略可）- 使用する DAO.Database（指定なし時は CurrentDb）
' 機能 : データベースを初期化し、転送設定（pSettings）用の Collection を作成するにゃ
'=================================================
Public Sub Init(Optional DB As DAO.Database)
    If DB Is Nothing Then
        Set pDB = CurrentDb
    Else
        Set pDB = DB
    End If
    Set pSettings = New Collection
End Sub

'=================================================
' サブルーチン名 : AddSetting
' 引数 : srcTable - 転送元テーブル名
'        tgtTable - 転送先テーブル名
'        keyField - 削除対象を特定するためのキー列名
' 機能 : テーブル転送の設定（転送元・転送先・キー）を1件分コレクションに追加するにゃ
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
' サブルーチン名 : ExecuteAll
' 機能 : 登録されたすべての転送設定を元に、テーブル転送（TransferTable）を実行するにゃ
'=================================================
Public Sub ExecuteAll()
    Dim setting As Collection
    For Each setting In pSettings
        TransferTable CStr(setting("Source")), CStr(setting("Target")), CStr(setting("Key"))
    Next
End Sub

'=================================================
' サブルーチン名 : TransferTable
' 引数 : src - 転送元テーブル名
'        tgt - 転送先テーブル名
'        key - 削除対象キー列名
' 機能 : 転送先の既存データをキーで削除後、転送元の全件を挿入するにゃ
'        DELETE + INSERT INTO を連続実行する構造にゃ
'=================================================
Private Sub TransferTable(src As String, tgt As String, key As String)
    Dim sqlDel As String, sqlIns As String
    sqlDel = "DELETE FROM [" & tgt & "] WHERE [" & key & "] IN (SELECT [" & key & "] FROM [" & src & "]);"
    pDB.Execute sqlDel, dbFailOnError
    sqlIns = "INSERT INTO [" & tgt & "] SELECT * FROM [" & src & "];"
    pDB.Execute sqlIns, dbFailOnError
End Sub


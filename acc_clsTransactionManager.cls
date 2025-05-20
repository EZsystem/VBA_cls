VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "acc_clsTransactionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================
' Class Module: acc_clsTransactionManager
' 説明　：DAOトランザクションの開始・コミット・ロールバックを管理するクラス
' 作成日：2025/04/25
' 更新日：再生成 by そうじろう
'=====================================
Option Compare Database
Option Explicit

' --- 内部保持変数 ---
Private pDB            As DAO.Database
Private pInTransaction As Boolean

'=================================================
' サブルーチン名 : Init
' 引数 : DB - 操作対象の DAO.Database オブジェクト
' 機能 : データベースを初期化し、トランザクション未開始状態に設定するにゃ
'=================================================
Public Sub Init(DB As DAO.Database)
    Set pDB = DB
    pInTransaction = False
End Sub

'=================================================
' サブルーチン名 : BeginTransaction
' 機能 : トランザクションを開始するにゃ
'        すでに開始済みの場合は何もしないように制御してるにゃ
'=================================================
Public Sub BeginTransaction()
    If Not pInTransaction Then
        pDB.BeginTrans
        pInTransaction = True
    End If
End Sub

'=================================================
' サブルーチン名 : Commit
' 機能 : トランザクションをコミットして確定するにゃ
'        実行中のトランザクションがある場合のみ動作するにゃ
'=================================================
Public Sub Commit()
    If pInTransaction Then
        pDB.CommitTrans
        pInTransaction = False
    End If
End Sub

'=================================================
' サブルーチン名 : Rollback
' 機能 : トランザクションをロールバックして元に戻すにゃ
'        実行中のトランザクションがある場合のみ動作するにゃ
'=================================================
Public Sub Rollback()
    If pInTransaction Then
        pDB.Rollback
        pInTransaction = False
    End If
End Sub


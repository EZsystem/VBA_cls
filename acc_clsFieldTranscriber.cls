VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "acc_clsFieldTranscriber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================
' Class Module: acc_clsFieldTranscriber
' 説明　：テーブルの任意フィールド間で値を転写する汎用ユーティリティ
' 作成日：2025/05/12
' 更新日：2025/05/12
' 修正履歴：
' ・Rollback方式をDAO.Workspaceに修正
' ・フィールド存在チェック（FieldExists）を追加
' ・マッピング登録時にTrimと空チェックを追加
'=====================================
Option Compare Database
Option Explicit

' --- 内部保持変数 ---
Private pDB             As DAO.Database
Private pTableName      As String
Private pWhereCondition As String
Private pSkipNulls      As Boolean
Private pMappings       As Object  ' Scripting.Dictionary

'=================================================
' サブルーチン名 : Init
' 説明   : DAO.Database を取得し、テーブル名・設定を初期化する
'=================================================
Public Sub Init(Optional TableName As String = "Icube_", _
                Optional WhereCondition As String = "", _
                Optional SkipNulls As Boolean = False)
    Set pDB = CurrentDb
    pTableName = TableName
    pWhereCondition = WhereCondition
    pSkipNulls = SkipNulls
    Set pMappings = CreateObject("Scripting.Dictionary")
End Sub

'=================================================
' サブルーチン名 : AddMapping
' 説明   : フィールド間転写のマッピングを登録する
'=================================================
Public Sub AddMapping(srcField As String, tgtField As String)
    srcField = Trim(srcField)
    tgtField = Trim(tgtField)
    If srcField = "" Or tgtField = "" Then
        MsgBox "AddMappingエラー：空のフィールド名は登録できませんにゃ", vbExclamation
        Exit Sub
    End If
    pMappings(srcField) = tgtField
End Sub

'=================================================
' サブルーチン名 : ClearMappings
' 説明   : 登録済みの全てのマッピングをクリアする
'=================================================
Public Sub ClearMappings()
    pMappings.RemoveAll
End Sub

'=================================================
' サブルーチン名 : TranscribeAll
' 説明   : 登録されたマッピングに従い、レコード毎にフィールド値を転写する
'=================================================
Public Sub TranscribeAll()
    Dim rs    As DAO.Recordset
    Dim ws    As DAO.Workspace
    Dim src   As Variant
    Dim tgt   As String
    Dim val   As Variant
    Dim safeSrc As String, safeTgt As String

    On Error GoTo Err_Handler
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans

    Set rs = OpenRS()
    With rs
        Do While Not .EOF
            .Edit
            For Each src In pMappings.Keys
                safeSrc = Trim(CStr(src))
                safeTgt = Trim(CStr(pMappings(src)))
                If safeSrc = "" Or safeTgt = "" Then GoTo NextField
                If Not FieldExists(.Fields, safeSrc) Then
                    Debug.Print "srcフィールドなし：" & "[" & safeSrc & "]"
                    GoTo NextField
                End If
                If Not FieldExists(.Fields, safeTgt) Then
                    Debug.Print "tgtフィールドなし：" & "[" & safeTgt & "]"
                    GoTo NextField
                End If
                val = .Fields(safeSrc).value
                If Not (pSkipNulls And IsNull(val)) Then
                    .Fields(safeTgt).value = val
                End If
NextField:
            Next
            .Update
            .MoveNext
        Loop
        .Close
    End With

    ws.CommitTrans
    Exit Sub

Err_Handler:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not ws Is Nothing Then
        If Err.Number > 0 Then ws.Rollback
    End If
    MsgBox "【FieldTranscriber】エラー：" & Err.Number & " - " & Err.description, vbCritical
    Debug.Print "【FieldTranscriber】Err:" & Err.Number & " - " & Err.description

End Sub

'=================================================
' 関数名 : OpenRS
' 説明   : 対象テーブルのレコードセットを返す
'=================================================
Private Function OpenRS() As DAO.Recordset
    Dim sql As String
    sql = "SELECT * FROM [" & pTableName & "]"
    If pWhereCondition <> "" Then sql = sql & " WHERE " & pWhereCondition
    Set OpenRS = pDB.OpenRecordset(sql, dbOpenDynaset)
End Function

'=================================================
' 関数名 : FieldExists
' 説明   : 指定フィールド名が存在するか判定する
'=================================================
Private Function FieldExists(flds As Fields, fldName As String) As Boolean
    On Error GoTo Err_Handler
    Dim f As Field
    Set f = flds(fldName)
    FieldExists = True
    Exit Function
Err_Handler:
    FieldExists = False
End Function

'=================================================
' プロパティ Let/Get : TableName, WhereCondition, SkipNulls
'=================================================
Public Property Let TableName(val As String): pTableName = val: End Property
Public Property Get TableName() As String: TableName = pTableName: End Property

Public Property Let WhereCondition(val As String): pWhereCondition = val: End Property
Public Property Get WhereCondition() As String: WhereCondition = pWhereCondition: End Property

Public Property Let SkipNulls(val As Boolean): pSkipNulls = val: End Property
Public Property Get SkipNulls() As Boolean: SkipNulls = pSkipNulls: End Property

'=================================================
' クリーンアップ：Class_Terminate
'=================================================
Private Sub Class_Terminate()
    On Error Resume Next
    Set pMappings = Nothing
    Set pDB = Nothing
End Sub



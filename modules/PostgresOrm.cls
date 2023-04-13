VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PostgresOrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' **************************************************
' 参照設定
'   Microsoft Activex Data Objects 6.1 Library
' **************************************************

' ==================================================
' メンバ変数
' ==================================================
Private context_ As PostgresOrmContext

Private is_initialzed_ As Boolean

Private connection_ As New ADODB.Connection
Private schema_ As String
Private table_ As String

' ==================================================
' メソッド
' ==================================================
' --------------------------------------------------
' クラス生成後の初期化(クラスとしての機能)
' インスタンス化した後、必ず実行する
' --------------------------------------------------
Public Sub Init( _
            ByVal ArgServer As String, _
            ByVal ArgDatabase As String, _
            ByVal ArgPassword As String, _
            Optional ByVal ArgUser As String = "postgres", _
            Optional ByVal ArgPort As Long = 5432, _
            Optional ByVal ArgConnectionTimeout As Long = 15, _
            Optional ByVal ArgCommandTimeout As Long = 30, _
            Optional ByVal ArgRetry As Long = 0 _
            )
    
    If is_initialzed_ Then PostgresOrmUtil.RaiseError Me, PostgresOrmErrorCode.poTwiceInitialized

    context_.poRetry = ArgRetry

    schema_ = "public"

    connection_.ConnectionString = _
        "Provider=MSDASQL;" & _
        "Driver=PostgreSQL Unicode;" & _
        "Server=" & ArgServer & ";" & _
        "Port=" & ArgPort & ";" & _
        "Database=" & ArgDatabase & ";" & _
        "UID=" & ArgUser & ";" & _
        "PWD=" & ArgPassword & ";" & _
        "Connect Timeout=" & ArgConnectionTimeout & ";" & _
        "Command Timeout=" & ArgCommandTimeout & ";"

    is_initialzed_ = True

End Sub

' --------------------------------------------------
' スキーマのsetter
' --------------------------------------------------
Public Function SetSchema(ByVal ArgSchema As String) As PostgresOrm

    CheckInit

    schema_ = ArgSchema
    Set SetSchema = Me

End Function

' --------------------------------------------------
' スキーマのgetter
' --------------------------------------------------
Public Function GetSchema() As String

    CheckInit

    GetSchema = schema_

End Function

' --------------------------------------------------
' テーブルのsetter
' --------------------------------------------------
Public Function SetTable(ByVal ArgTable As String) As PostgresOrm

    CheckInit

    table_ = ArgTable
    Set SetTable = Me

End Function

' --------------------------------------------------
' テーブルのgetter
' --------------------------------------------------

Public Function GetTable() As String

    CheckInit

    GetTable = table_

End Function

' --------------------------------------------------
' SQLを直接発行
' --------------------------------------------------
Public Function ExecuteSql( _
                        ByVal ArgSql As String, _
                        Optional ByVal ArgCursorType As CursorTypeEnum = adOpenUnspecified, _
                        Optional ByVal ArgLockType As LockTypeEnum = adLockUnspecified, _
                        Optional ByVal ArgShouldOutputSql As Boolean = False _
                        ) As ADODB.Recordset

On Error GoTo 0:

    CheckInit

    Dim expiration As Long: expiration = 1

    Dim rs As New ADODB.Recordset

    If ArgShouldOutputSql Then Debug.Print ArgSql

ConnectionOpen:
On Error GoTo ConnectionOpenErr:

    connection_.Open

RecoredsetOpen:
On Error GoTo RecordsetOpenErr:

    rs.Open ArgSql, connection_, ArgCursorType, ArgLockType

On Error GoTo 0:

    Set ExecuteSql = rs

    Exit Function

ConnectionOpenErr:

    If context_.poRetry >= expiration Then
        expiration = expiration + 1
        Resume ConnectionOpen
    Else
        With Err
            .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
        End With
    End If

    Exit Function

RecordsetOpenErr:

    If context_.poRetry >= expiration Then
        expiration = expiration + 1
        Resume RecoredsetOpen
    Else
        With Err
            .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
        End With
    End If

    Exit Function

End Function

' --------------------------------------------------
' データの作成
' --------------------------------------------------
Public Function Create() As PostgresOrmCreate

    CheckInit

    Dim create_class As New PostgresOrmCreate

    create_class.Init context_, connection_, schema_, table_

    Set Create = create_class

End Function

' --------------------------------------------------
' データの読み込み
' --------------------------------------------------
Public Function Read() As PostgresOrmRead

    CheckInit

    Dim read_class As New PostgresOrmRead

    read_class.Init context_, connection_, schema_, table_

    Set Read = read_class

End Function

' --------------------------------------------------
' データの更新
' --------------------------------------------------
Public Function Update() As PostgresOrmUpdate

    CheckInit

    Dim update_class As New PostgresOrmUpdate

    update_class.Init context_, connection_, schema_, table_

    Set Update = update_class

End Function

' --------------------------------------------------
' データの削除
' --------------------------------------------------
Public Function Delete() As PostgresOrmDelete

    CheckInit

    Dim delete_class As New PostgresOrmDelete

    delete_class.Init context_, connection_, schema_, table_

    Set Delete = delete_class

End Function

' --------------------------------------------------
' メタ情報
' --------------------------------------------------
Public Function Meta() As PostgresOrmMeta

    CheckInit

    Dim meta_class As New PostgresOrmMeta

    meta_class.Init context_, connection_, schema_, table_

    Set Meta = meta_class

End Function

' ==================================================
' 非公開メソッド
' ==================================================
' --------------------------------------------------
' Initメソッドを実行したか確認
' 未初期化の場合、エラーを発生
' --------------------------------------------------
Private Sub CheckInit()

    If Not is_initialzed_ Then PostgresOrmUtil.RaiseError Me, PostgresOrmErrorCode.poUnInitialized

End Sub

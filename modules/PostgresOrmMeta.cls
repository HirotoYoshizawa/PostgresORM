VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PostgresOrmMeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==================================================
' メンバ変数
' ==================================================
Private context_ As PostgresOrmContext

Private is_initialzed_ As Boolean

Private connection_ As New ADODB.Connection
Private schema_ As String
Private table_ As String

' ==================================================
' 初期化
' ==================================================
' --------------------------------------------------
' クラス生成後の初期化(クラスとしての機能)
' インスタンス化した後、必ず実行する
' --------------------------------------------------
Public Sub Init( _
            ByRef ArgContext As PostgresOrmContext, _
            ByRef ArgConnection As ADODB.Connection, _
            ByVal ArgSchema As String, _
            ByVal ArgTable As String _
            )

    If is_initialzed_ Then PostgresOrmUtil.RaiseError Me, PostgresOrmErrorCode.poTwiceInitialized

    context_ = ArgContext
    Set connection_ = ArgConnection
    schema_ = ArgSchema
    table_ = ArgTable

    is_initialzed_ = True

End Sub

' ==================================================
' 終了処理
' ==================================================
' --------------------------------------------------
' クラス破棄時の終了処理
' --------------------------------------------------
Private Sub Class_Terminate()

    If connection_.State = adStateOpen Then connection_.Close

End Sub


' ==================================================
' メソッド
' ==================================================
' --------------------------------------------------
' スキーマ名のリスト
' 配列、またはコレクションで返す
' --------------------------------------------------
Public Function GetSchemas( _
                        Optional ByVal ArgReturnType As PostgresOrmReturnType = PostgresOrmReturnType.poArray _
                        ) As Variant

    CheckInit

    Dim rs As New ADODB.Recordset
    Dim sql As String

    Dim i As Long
    Dim schemas_array() As String
    Dim schemas_collection As New Collection

    sql = _
        "select nspname " & _
        "from pg_namespace " & _
        "where nspname not like 'pg_%' " & _
        "and nspname != 'information_schema' " & _
        "order by nspname"
    connection_.Open
    rs.Open sql, connection_, adOpenStatic, adLockReadOnly

    For i = 0 To rs.RecordCount - 1
        ReDim Preserve schemas_array(i)
        schemas_array(i) = rs![nspname]

        schemas_collection.Add rs![nspname]

        rs.MoveNext
    Next

    Select Case ArgReturnType
        Case PostgresOrmReturnType.poArray
            GetSchemas = schemas_array

        Case PostgresOrmReturnType.poCollection
            Set GetSchemas = schemas_collection
    End Select

End Function

' --------------------------------------------------
' テーブル名のリスト
' 配列、またはコレクションで返す
' --------------------------------------------------
Public Function GetTables( _
                        Optional ByVal ArgSchema As String = "", _
                        Optional ByVal ArgReturnType As PostgresOrmReturnType = PostgresOrmReturnType.poArray _
                        ) As Variant

    CheckInit

    Dim rs As New ADODB.Recordset
    Dim sql As String

    Dim i As Long
    Dim tables_array() As String
    Dim tables_collection As New Collection

    If ArgSchema = "" Then ArgSchema = schema_

    sql = _
        "select tablename " & _
        "from pg_tables " & _
        "where schemaname = '" & ArgSchema & "' " & _
        "order by tablename"
    connection_.Open
    rs.Open sql, connection_, , adLockReadOnly

    i = 0
    Do Until rs.EOF
        ReDim Preserve tables_array(i)
        tables_array(i) = rs![tableName]

        tables_collection.Add rs![tableName]

        i = i + 1
        rs.MoveNext
    Loop

    Select Case ArgReturnType
        Case PostgresOrmReturnType.poArray
            GetTables = tables_array

        Case PostgresOrmReturnType.poCollection
            Set GetTables = tables_collection
    End Select

End Function

' --------------------------------------------------
' カラム名のリスト
' 配列、またはコレクションで返す
' --------------------------------------------------
Public Function GetColumns( _
                        Optional ByVal ArgSchema As String = "", _
                        Optional ByVal ArgTable As String = "", _
                        Optional ByVal ArgReturnType As PostgresOrmReturnType = PostgresOrmReturnType.poArray _
                        ) As Variant

    CheckInit

    Dim rs As New ADODB.Recordset
    Dim sql As String
    
    Dim i As Long
    Dim columns_array() As String
    Dim columns_collection As New Collection

    If ArgSchema = "" Then ArgSchema = schema_
    If ArgTable = "" Then ArgTable = table_

    sql = _
        "select * " & _
        "from " & ArgSchema & "." & ArgTable & " " & _
        "limit 0"
    connection_.Open
    rs.Open sql, connection_, adOpenStatic, adLockReadOnly

    For i = 0 To rs.Fields.Count - 1
        ReDim Preserve columns_array(i)
        columns_array(i) = rs.Fields(i).Name

        columns_collection.Add rs.Fields(i).Name
    Next

    Select Case ArgReturnType
        Case PostgresOrmReturnType.poArray
            GetColumns = columns_array

        Case PostgresOrmReturnType.poCollection
            Set GetColumns = columns_collection
    End Select

End Function

' --------------------------------------------------
' テーブルの存在チェック
' --------------------------------------------------
Public Function ExistsTable( _
                        Optional ByVal ArgSchema As String = "", _
                        Optional ByVal ArgTable As String = "" _
                        ) As Boolean

    CheckInit

    If ArgSchema = "" Then ArgSchema = schema_
    If ArgTable = "" Then ArgTable = table_

    Dim rs As New ADODB.Recordset
    Dim sql As String

    sql = _
        "select schemaname, tablename, tableowner " & _
        "from pg_tables " & _
        "where pg_tables.schemaname='" & ArgSchema & "' " & _
        "and pg_tables.tablename='" & ArgTable & "'"
    connection_.Open
    rs.Open sql, connection_, , adLockReadOnly

    If Not rs.EOF Then
        ExistsTable = True
    Else
        ExistsTable = False
    End If

End Function

' --------------------------------------------------
' フィールドの存在チェック
' --------------------------------------------------
Public Function ExistsField( _
                        ByVal ArgColumn As String, _
                        ByVal ArgValue As String, _
                        Optional ByVal ArgSchema As String = "", _
                        Optional ByVal ArgTable As String = "" _
                        ) As Boolean

    CheckInit

    If ArgSchema = "" Then ArgSchema = schema_
    If ArgTable = "" Then ArgTable = table_

    Dim rs As New ADODB.Recordset
    Dim sql As String

    sql = _
        "select * " & _
        "from " & ArgSchema & "." & ArgTable & " " & _
        "where " & ArgColumn & " = '" & ArgValue & "'"
    connection_.Open
    rs.Open sql, connection_, , adLockReadOnly

    If Not rs.EOF Then
        ExistsField = True
    Else
        ExistsField = False
    End If

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

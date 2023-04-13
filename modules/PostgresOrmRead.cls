VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PostgresOrmRead"
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

Private distinct_string_ As String
Private columns_string_ As String
Private where_string_ As String
Private group_by_string_ As String
Private having_string_ As String
Private order_by_string_ As String
Private limit_string_ As String
Private offset_string_ As String

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
    columns_string_ = "* "

    is_initialzed_ = True

End Sub

' ==================================================
' メソッド
' ==================================================
' --------------------------------------------------
' DISTINCT句のsetter
' --------------------------------------------------
Public Function SetDistinct() As PostgresOrmRead

    CheckInit

    distinct_string_ = "distinct "

    Set SetDistinct = Me

End Function

' --------------------------------------------------
' カラムのsetter
' --------------------------------------------------
Public Function SetAddColumn( _
                        ByVal ArgColumn As String, _
                        Optional ByVal ArgSqlFucntion As PostgresOrmSqlFunction = PostgresOrmSqlFunction.poNoFunction, _
                        Optional ByVal ArgColumnName As String = "" _
                        ) As PostgresOrmRead

    CheckInit

    Dim tmp_string As String

    tmp_string = ToSqlFunction(ArgColumn, ArgSqlFucntion)

    If ArgColumnName <> "" Then
        tmp_string = tmp_string & " as " & ArgColumnName
    End If

    If columns_string_ = "* " Then
        columns_string_ = tmp_string & " "
    Else
        columns_string_ = Left(columns_string_, Len(columns_string_) - 1)
        columns_string_ = columns_string_ & ", " & tmp_string & " "
    End If

    Set SetAddColumn = Me

End Function

' --------------------------------------------------
' カラム(複数)のsetter
' --------------------------------------------------
Public Function SetAddColumns(ByVal ArgColumns As Variant) As PostgresOrmRead

    CheckInit

    If columns_string_ = "* " Then
        columns_string_ = PostgresOrmUtil.ToColumnsString(ArgColumns) & " "
    Else
        columns_string_ = Left(columns_string_, Len(columns_string_) - 1)
        columns_string_ = columns_string_ & ", " & PostgresOrmUtil.ToColumnsString(ArgColumns) & " "
    End If

    Set SetAddColumns = Me

End Function

' --------------------------------------------------
' WHERE句のsetter
' --------------------------------------------------
Public Function SetAddWhere( _
                    ByVal ArgColumn As String, _
                    ByVal ArgWhereType As PostgresOrmWhereType, _
                    ByVal ArgValue As Variant, _
                    Optional ByVal ArgConnectType As PostgresOrmConnectType = PostgresOrmConnectType.poAnd _
                    ) As PostgresOrmRead

    CheckInit

    Dim connect_string As String

    If where_string_ = "" Then
        where_string_ = "where " & ToWhereString(ArgColumn, ArgWhereType, ArgValue)
    Else
        Select Case ArgConnectType
            Case PostgresOrmConnectType.poAnd
                connect_string = " and "

            Case PostgresOrmConnectType.poOr
                connect_string = " or "
        End Select

        where_string_ = where_string_ & connect_string & ToWhereString(ArgColumn, ArgWhereType, ArgValue)
    End If

    Set SetAddWhere = Me

End Function

' --------------------------------------------------
' GROUP BY句のsetter
' --------------------------------------------------
Public Function SetAddGroupBy(ByVal ArgColumn As String) As PostgresOrmRead

    CheckInit

    If group_by_string_ = "" Then
        group_by_string_ = "group by " & ArgColumn
    Else
        group_by_string_ = group_by_string_ & ", " & ArgColumn
    End If

    Set SetAddGroupBy = Me

End Function

' --------------------------------------------------
' HAVING句のsetter
' --------------------------------------------------
Public Function SetAddHaving( _
                    ByVal ArgColumn As String, _
                    ByVal ArgWhereType As PostgresOrmWhereType, _
                    ByVal ArgValue As Variant, _
                    Optional ByVal ArgSqlFucntion As PostgresOrmSqlFunction = PostgresOrmSqlFunction.poNoFunction, _
                    Optional ByVal ArgConnectType As PostgresOrmConnectType = PostgresOrmConnectType.poAnd _
                    ) As PostgresOrmRead

    CheckInit

    Dim connect_string As String

    If having_string_ = "" Then
        having_string_ = "having " & ToWhereString(ToSqlFunction(ArgColumn, ArgSqlFucntion), ArgWhereType, ArgValue)
    Else
        Select Case ArgConnectType
            Case PostgresOrmConnectType.poAnd
                connect_string = " and "

            Case PostgresOrmConnectType.poOr
                connect_string = " or "
        End Select

        having_string_ = having_string_ & connect_string & ToWhereString(ToSqlFunction(ArgColumn, ArgSqlFucntion), ArgWhereType, ArgValue)
    End If

    Set SetAddHaving = Me

End Function

' --------------------------------------------------
' ORDER BY句のsetter
' --------------------------------------------------
Public Function SetAddOrderBy( _
                    ByVal ArgColumn As String, _
                    Optional ByVal sortType As PostgresOrmSortType = PostgresOrmSortType.poAsc _
                    ) As PostgresOrmRead

    CheckInit

    Dim sort_string As String

    Select Case sortType
        Case PostgresOrmSortType.poAsc
            sort_string = " asc"

        Case PostgresOrmSortType.poDesc
            sort_string = " desc"
    End Select

    If order_by_string_ = "" Then
        order_by_string_ = "order by " & ArgColumn & sort_string
    Else
        order_by_string_ = order_by_string_ & ", " & ArgColumn & sort_string
    End If

    Set SetAddOrderBy = Me

End Function

' --------------------------------------------------
' LIMIT句のsetter
' --------------------------------------------------
Public Function SetLimit(ByVal ArgValue As Long) As PostgresOrmRead

    CheckInit

    limit_string_ = "limit " & ArgValue

    Set SetLimit = Me

End Function

' --------------------------------------------------
' OFFSET句のsetter
' --------------------------------------------------
Public Function SetOffset(ByVal ArgValue As Long) As PostgresOrmRead

    CheckInit

    offset_string_ = "offset " & ArgValue

    Set SetOffset = Me

End Function

' --------------------------------------------------
' DISTINCT句の削除
' --------------------------------------------------
Public Sub ClearDistinct()

    CheckInit

    distinct_string_ = ""

End Sub

' --------------------------------------------------
' カラムの削除
' --------------------------------------------------
Public Sub ClearColumns()

    CheckInit

    columns_string_ = "* "

End Sub

' --------------------------------------------------
' WHERE句の削除
' --------------------------------------------------
Public Sub ClearWhere()

    CheckInit

    where_string_ = ""

End Sub

' --------------------------------------------------
' GROUP BY句の削除
' --------------------------------------------------
Public Sub ClearGroupBy()

    CheckInit

    group_by_string_ = ""

End Sub

' --------------------------------------------------
' HAVING句の削除
' --------------------------------------------------
Public Sub ClearHaving()

    CheckInit

    having_string_ = ""

End Sub

' --------------------------------------------------
' ORDER BY句の削除
' --------------------------------------------------
Public Sub ClearOrderBy()

    CheckInit

    order_by_string_ = ""

End Sub

' --------------------------------------------------
' LIMIT句の削除
' --------------------------------------------------
Public Sub ClearLimit()

    CheckInit

    limit_string_ = ""

End Sub

' --------------------------------------------------
' OFFSETの削除
' --------------------------------------------------
Public Sub ClearOffset()

    CheckInit

    offset_string_ = ""

End Sub

' --------------------------------------------------
' SQLのgetter
' --------------------------------------------------
Public Function GetSql() As String

    CheckInit

    GetSql = GetGeneratedSql

End Function

' --------------------------------------------------
' SQLを発行
' --------------------------------------------------
Public Function Commit( _
                    Optional ByVal ArgCursorType As CursorTypeEnum = adOpenUnspecified, _
                    Optional ByVal ArgLockType As LockTypeEnum = adLockReadOnly, _
                    Optional ByVal ArgShouldOutputSql As Boolean = False _
                    ) As ADODB.Recordset

On Error GoTo 0:

    CheckInit

    Dim expiration As Long: expiration = 1

    Dim rs As New ADODB.Recordset

    If ArgShouldOutputSql Then Debug.Print GetGeneratedSql

ConnectionOpen:
On Error GoTo ConnectionOpenErr:

    If connection_.State = adStateClosed Then connection_.Open

RecoredsetOpen:
On Error GoTo RecordsetOpenErr:

    rs.Open GetGeneratedSql, connection_, ArgCursorType, ArgLockType

On Error GoTo 0:

    Set Commit = rs

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

' --------------------------------------------------
' カラムをSQL関数にして返す
' --------------------------------------------------
Private Function ToSqlFunction( _
                            ByVal ArgColumn As String, _
                            ByVal ArgSqlFunction As PostgresOrmSqlFunction _
                            ) As String

    Dim tmp_string As String

    tmp_string = ArgColumn

    Select Case ArgSqlFunction
        Case PostgresOrmSqlFunction.poCount
            tmp_string = "count(" & ArgColumn & ")"

        Case PostgresOrmSqlFunction.poSum
            tmp_string = "sum(" & ArgColumn & ")"

        Case PostgresOrmSqlFunction.poAvg
            tmp_string = "avg(" & ArgColumn & ")"

        Case PostgresOrmSqlFunction.poMax
            tmp_string = "max(" & ArgColumn & ")"

        Case PostgresOrmSqlFunction.poMin
            tmp_string = "min(" & ArgColumn & ")"

        Case PostgresOrmSqlFunction.poAbs
            tmp_string = "abs(" & ArgColumn & ")"

        Case PostgresOrmSqlFunction.poRound
            tmp_string = "round(" & ArgColumn & ")"
    End Select

    ToSqlFunction = tmp_string

End Function

' --------------------------------------------------
' SQL文に変換して返す
' --------------------------------------------------
Private Function GetGeneratedSql() As String

    Dim sql As String

    sql = _
        "select " & distinct_string_ & columns_string_ & _
        "from " & schema_ & "." & table_

    If where_string_ <> "" Then sql = sql & " " & where_string_
    If group_by_string_ <> "" Then sql = sql & " " & group_by_string_
    If having_string_ <> "" Then sql = sql & " " & having_string_
    If order_by_string_ <> "" Then sql = sql & " " & order_by_string_
    If limit_string_ <> "" Then sql = sql & " " & limit_string_
    If offset_string_ <> "" Then sql = sql & " " & offset_string_

    GetGeneratedSql = sql

End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PostgresOrmUpdate"
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

Private columns_and_values_string_ As String
Private where_string_ As String

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
' メソッド
' ==================================================
' --------------------------------------------------
' カラムとバリューのsetter
' --------------------------------------------------
Public Function SetAddColumnAndValue( _
                                ByVal ArgColumn As String, _
                                ByVal ArgValue As String _
                                ) As PostgresOrmUpdate

    CheckInit

    Dim tmp_string As String

    If ArgValue = "" Then
        tmp_string = ArgColumn & " = null"
    Else
        tmp_string = ArgColumn & " = '" & ArgValue & "'"
    End If

    If columns_and_values_string_ = "" Then
        columns_and_values_string_ = tmp_string
    Else
        columns_and_values_string_ = columns_and_values_string_ & ", " & tmp_string
    End If

    Set SetAddColumnAndValue = Me

End Function

' --------------------------------------------------
' WHERE句のsetter
' --------------------------------------------------
Public Function SetAddWhere( _
                    ByVal ArgColumn As String, _
                    ByVal ArgWhereType As PostgresOrmWhereType, _
                    ByVal ArgValue As Variant, _
                    Optional ByVal ArgConnectType As PostgresOrmConnectType = PostgresOrmConnectType.poAnd _
                    ) As PostgresOrmUpdate

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
' カラムとバリューの削除
' --------------------------------------------------
Public Sub ClearColumnsAndValues()

    CheckInit

    columns_and_values_string_ = ""

End Sub

' --------------------------------------------------
' WHERE句の削除
' --------------------------------------------------
Public Sub ClearWhere()

    CheckInit

    where_string_ = ""

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
                Optional ByVal ArgLockType As LockTypeEnum = adLockPessimistic, _
                Optional ByVal ArgShouldOutputSql As Boolean = False _
                )

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
' SQL文に変換して返す
' --------------------------------------------------
Private Function GetGeneratedSql() As String

    Dim sql As String

    sql = _
        "update " & schema_ & "." & table_ & " " & _
        "set " & columns_and_values_string_

    If where_string_ <> "" Then sql = sql & " " & where_string_
 
    GetGeneratedSql = sql

End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PostgresOrmCreate"
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

Private columns_string_ As String
Private values_string_ As String

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
' カラムのsetter
' --------------------------------------------------
Public Function SetAddColumn(ByVal ArgColumn As String) As PostgresOrmCreate

    CheckInit

    If columns_string_ = "" Then
        columns_string_ = "(" & ArgColumn & ") "
    Else
        columns_string_ = Left(columns_string_, Len(columns_string_) - 1)
        columns_string_ = Replace(columns_string_, "(", "")
        columns_string_ = Replace(columns_string_, ")", "")
        columns_string_ = "(" & columns_string_ & ", " & ArgColumn & ") "
    End If

    Set SetAddColumn = Me

End Function

' --------------------------------------------------
' カラム(複数)のsetter
' --------------------------------------------------
Public Function SetAddColumns(ByVal ArgColumns As Variant) As PostgresOrmCreate

    CheckInit
    CheckArrayOrCollection ArgColumns

    If columns_string_ = "" Then
        columns_string_ = "(" & PostgresOrmUtil.ToColumnsString(ArgColumns) & ") "
    Else
        columns_string_ = Left(columns_string_, Len(columns_string_) - 1)
        columns_string_ = Replace(columns_string_, "(", "")
        columns_string_ = Replace(columns_string_, ")", "")
        columns_string_ = "(" & columns_string_ & ", " & PostgresOrmUtil.ToColumnsString(ArgColumns) & ") "
    End If

    Set SetAddColumns = Me

End Function

' --------------------------------------------------
' バリュー(複数)のsetter
' --------------------------------------------------
Public Function SetAddValues(ByVal ArgValues As Variant) As PostgresOrmCreate

    CheckInit
    CheckArrayOrCollection ArgValues
    
    If values_string_ = "" Then
        values_string_ = "(" & PostgresOrmUtil.ToValuesString(ArgValues) & ")"
    Else
        values_string_ = values_string_ & ", (" & PostgresOrmUtil.ToValuesString(ArgValues) & ")"
    End If

    Set SetAddValues = Me

End Function

' --------------------------------------------------
' カラムの削除
' --------------------------------------------------
Public Sub ClearColumns()

    CheckInit

    columns_string_ = ""

End Sub

' --------------------------------------------------
' バリューの削除
' --------------------------------------------------
Public Sub ClearValues()

    CheckInit

    values_string_ = ""

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
' 配列かcollection型か確認
' それ以外の場合、エラーを発生
' --------------------------------------------------
Private Sub CheckArrayOrCollection(ByVal ArgObject As Variant)

    Dim is_valid As Boolean

    is_valid = PostgresOrmUtil.IsArrayOrCollection(ArgObject)

    If Not is_valid Then PostgresOrmUtil.RaiseError Me, PostgresOrmErrorCode.poNotArrayOrCollection

End Sub

' --------------------------------------------------
' SQL文の変換して返す
' --------------------------------------------------
Private Function GetGeneratedSql() As String

    Dim sql As String

    sql = _
        "insert into " & schema_ & "." & table_ & columns_string_ & _
        "values" & values_string_

    GetGeneratedSql = sql

End Function

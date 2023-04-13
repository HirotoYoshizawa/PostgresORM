Attribute VB_Name = "PostgresOrmUtil"
Option Explicit

' ==================================================
' 共通変数
' ==================================================
Public Enum PostgresOrmErrorCode
    poTwiceInitialized = 8010
    poUnInitialized = 8011
    poNotArrayOrCollection = 8012

End Enum

Public Enum PostgresOrmSqlFunction
    poNoFunction = 0
    poCount = 1
    poSum = 2
    poAvg = 3
    poMax = 4
    poMin = 5
    poAbs = 6
    poRound = 7
End Enum

Public Enum PostgresOrmWhereType
    poEqual = 0
    poNotEqual = 1
    poLess = 2
    poLessOrEqual = 3
    poGreater = 4
    poGreaterOrEqual = 5
    poLike = 6
    poIs = 7
    poIsNot = 8
    poIn = 9
    poBetween = 10
End Enum

Public Enum PostgresOrmConnectType
    poAnd = 0
    poOr = 1
End Enum

Public Enum PostgresOrmSortType
    poAsc = 0
    poDesc = 1
End Enum

Public Enum PostgresJoinType
    poInnerJoin = 0
    poLeftOuterJoin = 1
    poRightOuterJoin = 2
End Enum

Public Enum PostgresOrmReturnType
    poArray = 0
    poCollection = 1
End Enum

Public Type PostgresOrmContext
    poRetry As Long
End Type

' ==================================================
' 関数
' ==================================================
' --------------------------------------------------
' エラーを発生させる
' 列挙型「ErrorCode」を参照
' --------------------------------------------------
Public Sub RaiseError(ByRef ArgClassModule As Object, ByVal ArgErrorCode As Long)

    ' エラーメッセージ内容
    Dim message As String

    Select Case ArgErrorCode
        Case PostgresOrmErrorCode.poTwiceInitialized
            message = "既に初期化されています"

        Case PostgresOrmErrorCode.poUnInitialized
            message = "Initメソッドでクラスを初期化して下さい"

        Case PostgresOrmErrorCode.poNotArrayOrCollection
            message = "配列かCollection型を指定して下さい"

        Case Else
            message = "PostgresORMでの不明なエラー"
    End Select

    ' エラーを発生
    Err.Raise _
        Number:=ArgErrorCode, _
        Description:=TypeName(ArgClassModule) & ": " & message

End Sub

' --------------------------------------------------
' 配列かcollection型かを判定して真偽値を返す
' --------------------------------------------------
Public Function IsArrayOrCollection(ByRef ArgObject As Variant) As Boolean

    Dim is_valid As Boolean

    Select Case True
        Case IsArray(ArgObject)
            is_valid = True

        Case TypeName(ArgObject) = "Collection"
            is_valid = True
    End Select

    IsArrayOrCollection = is_valid

End Function

' --------------------------------------------------
' collection型から配列に変換して返す
' --------------------------------------------------
Public Function ToArrayFromCollecton(ByVal ArgCollection As Variant) As Variant

    Dim i As Long
    Dim tmp_array() As Variant

    For i = 0 To ArgCollection.Count - 1
        ReDim Preserve tmp_array(i)
        tmp_array(i) = ArgCollection(i + 1)
    Next

    ToArrayFromCollecton = tmp_array

End Function

' --------------------------------------------------
' 配列かcollection型からcolumns文字列に変換して返す
' --------------------------------------------------
Public Function ToColumnsString(ByVal ArgColumns As Variant) As String

    Dim tmp_array As Variant
    
    If TypeName(ArgColumns) = "Collection" Then
        tmp_array = PostgresOrmUtil.ToArrayFromCollecton(ArgColumns)
    Else
        tmp_array = ArgColumns
    End If

    ToColumnsString = Join(tmp_array, ", ")

End Function

' --------------------------------------------------
' 配列かcollection型からvalue文字列に変換して返す
' --------------------------------------------------
Public Function ToValuesString(ByVal ArgValues As Variant) As String

    Dim tmp_array As Variant

    If TypeName(ArgValues) = "Collection" Then
        tmp_array = PostgresOrmUtil.ToArrayFromCollecton(ArgValues)
    Else
        tmp_array = ArgValues
    End If

    Dim i As Long
    Dim min As Long, max As Long

    min = LBound(tmp_array)
    max = UBound(tmp_array)

    For i = min To max
        If tmp_array(i) = "" Then
            tmp_array(i) = "null"
        Else
            tmp_array(i) = "'" & tmp_array(i) & "'"
        End If
    Next

    ToValuesString = Join(tmp_array, ", ")

End Function

' --------------------------------------------------
' 配列かcollection型からbetween文字列に変換して返す
' --------------------------------------------------
Public Function ToBetweenString(ByVal ArgValues As Variant) As String

    Dim tmp_array As Variant

    If TypeName(ArgValues) = "Collection" Then
        tmp_array = PostgresOrmUtil.ToArrayFromCollecton(ArgValues)
    Else
        tmp_array = ArgValues
    End If
    
    Dim i As Long
    Dim min As Long, max As Long

    min = LBound(tmp_array)
    max = UBound(tmp_array)

    For i = min To max
        If tmp_array(i) = "" Then
            tmp_array(i) = "null"
        Else
            tmp_array(i) = "'" & tmp_array(i) & "'"
        End If
    Next

    ToBetweenString = Join(tmp_array, " and ")

End Function

' --------------------------------------------------
' columnとvalueからWHERE文字列に変換して返す
' --------------------------------------------------
Public Function ToWhereString( _
                            ByVal ArgColumn As String, _
                            ByVal ArgWhereType As PostgresOrmWhereType, _
                            ByVal ArgValue As Variant _
                            ) As String

    Dim tmp_string As String

    Select Case ArgWhereType
        Case PostgresOrmWhereType.poEqual
            tmp_string = ArgColumn & " = '" & ArgValue & "'"

        Case PostgresOrmWhereType.poNotEqual
            tmp_string = ArgColumn & " <> '" & ArgValue & "'"

        Case PostgresOrmWhereType.poLess
            tmp_string = ArgColumn & " < '" & ArgValue & "'"

        Case PostgresOrmWhereType.poLessOrEqual
            tmp_string = ArgColumn & " <= '" & ArgValue & "'"

        Case PostgresOrmWhereType.poGreater
            tmp_string = ArgColumn & " > '" & ArgValue & "'"

        Case PostgresOrmWhereType.poGreaterOrEqual
            tmp_string = ArgColumn & " >= '" & ArgValue & "'"

        Case PostgresOrmWhereType.poLike
            tmp_string = ArgColumn & " like '" & ArgValue & "'"

        Case PostgresOrmWhereType.poIs
            tmp_string = ArgColumn & " is " & ArgValue

        Case PostgresOrmWhereType.poIsNot
            tmp_string = ArgColumn & " is not " & ArgValue

        Case PostgresOrmWhereType.poIn
            tmp_string = ArgColumn & " in(" & ToValuesString(ArgValue) & ")"

        Case PostgresOrmWhereType.poBetween
            tmp_string = ArgColumn & " between " & ToBetweenString(ArgValue)

    End Select

    ToWhereString = tmp_string

End Function

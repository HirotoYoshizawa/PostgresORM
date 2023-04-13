# VBA PostgreSQL ORM

## 概要

このモジュールはVBAに、PostgreSQLへのORM(Object Relational Mapping)を提供します。SQLを記述する事無く、PostgreSQLに対してCRUD処理が可能です。尚、Postgresに対するSQLについて、ある程度理解している者を対象に制作しています。

## 準備

modulesフォルダ内にあるファイルを全てインポートします。

```tex
PostgresOrm
│  documentation.md
│  postgres.xlsm
│
└─modules
        PostgresOrm.cls
        PostgresOrmCreate.cls
        PostgresOrmDelete.cls
        PostgresOrmMeta.cls
        PostgresOrmRead.cls
        PostgresOrmUpdate.cls
        PostgresOrmUtil.bas
```

以下のライブラリを参照します。

- Microsoft Activex Data Objects 6.1 Library

## 参照

### PosrtgresOrm

Postgres用のORMとして基になるクラスです。このクラスをインスタンス化する事でデータベースに対して様々な操作が可能になります。



#### Init

インスタンス化したクラスを初期化します。インスタンス化した後に必ず実行して下さい。クラスにある各メソッドはInitメソッド後に実行出来ます。またInitメソッドを2回以上実行する事は出来ません。

- パラメータ

| 引数                 | 型     | 初期値     |
| -------------------- | ------ | ---------- |
| ArgServer            | String |            |
| ArgDatabase          | String |            |
| ArgPassword          | String |            |
| ArgUser              | String | "postgres" |
| ArgPort              | Long   | 5432       |
| ArgConnectionTimeout | Long   | 15         |
| ArgCommandTimeout    | Long   | 30         |
| ArgRetry             | Long   | 0          |



#### SetSchema

使用するスキーマ名を設定します。チェーンメソッドに対応します。

- パラメータ

| 引数      | 型     | 初期値 |
| --------- | ------ | ------ |
| ArgSchema | String |        |



#### GetSchema

設定したスキーマ名を返します。

- 戻り値

String

#### SetTable

使用するテーブル名を設定します。チェーンメソッドに対応します。

- パラメータ

| 引数     | 型     | 初期値 |
| -------- | ------ | ------ |
| ArgTable | String |        |



#### GetTable

設定したテーブル名を返します。

- 戻り値

String



#### ExecuteSql

引数に代入したSQL文からレコードセットを生成して返します。SQL文を直接実行する為、SetSchemaやSetTableで設定した値は考慮されずに実行されます。

- パラメータ

| 引数   | 型     | 初期値 |
| ------ | ------ | ------ |
| ArgSql | String |        |

- 戻り値

ADODO.Recordset



#### Create

データを作成するサブクラスを呼び出します。

PostgresOrmCreate項目を参照



#### Read

データを読み込むサブクラスを呼び出します。

PostgresOrmRead項目を参照



#### Update

データを更新するサブクラスを呼び出します。

PostgresOrmUpdate項目を参照



#### Delete

データを削除するサブクラスを呼び出します。

PostgresOrmDelete項目を参照



#### Meta

メタ情報を取得するサブクラスを呼び出します。

PostgresOrmMeta項目を参照



### PostgresOrmCreate

PostgresOrmクラスが内包しているサブクラスです。データベースに対してレコードを作成する各種メソッドを提供します。



#### Init

インスタンス化したクラスを初期化しますが、Createサブクラスを処理する際に自動で実行されます。Initメソッドを2回以上実行する事は出来ません。

- パラメータ

| 引数          | 型                 | 初期値 |
| ------------- | ------------------ | ------ |
| ArgContext    | PostgresOrmContext |        |
| ArgConnection | ADODB.Connection   |        |
| ArgSchema     | String             |        |
| ArgTable      | String             |        |



#### SetAddColumn

データを作成するカラム名を追加します。複数回実行すると、カラム名が順次追加されます。チェーンメソッドに対応します。

- パラメータ

| 引数      | 型     | 初期値 |
| --------- | ------ | ------ |
| ArgColumn | String |        |



#### SetAddColumns

データを作成するカラム名を複数追加します。複数回実行すると、カラム名が順次追加されます。チェーンメソッドに対応します。

- パラメータ

| 引数       | 型                       | 初期値 |
| ---------- | ------------------------ | ------ |
| ArgColumns | String()またはCollection |        |



#### SetAddValues

作成するデータを追加します。SetAddColumnおよびSetAddColumnsで設定したカラム名に対応したデータを用意する必要があります。複数回実行すると、作成するデータが順次追加されます。チェーンメソッドに対応します。

- パラメータ

| 引数      | 型                       | 初期値 |
| --------- | ------------------------ | ------ |
| ArgValues | String()またはCollection |        |



#### ClearColumns

SetAddColumnおよびSetAddColumnsで設定したカラム情報を削除します。



#### ClearValues

SetAddValuesで設定したデータ情報を削除します。



#### GetSql

各種メソッドで設定した情報から実際に実行するSQL文を返します。

- 戻り値

String



#### Commit

各種メソッドで設定した情報からPostgresへデータを作成します。

- パラメータ

| 引数               | 型             | 初期値            |
| ------------------ | -------------- | ----------------- |
| ArgCursorType      | CursorTypeEnum | adOpenUnspecified |
| ArgLockType        | LockTypeEnum   | adLockUnspecified |
| ArgShouldOutputSql | Boolean        | False             |

- 戻り値

ADODB.Recordset



### PostgresOrmRead

PostgresOrmクラスが内包しているサブクラスです。データベースに対してレコードを読み込む各種メソッドを提供します。尚、現段階ではテーブル結合には対応していません。



#### Init

インスタンス化したクラスを初期化しますが、Readサブクラスを処理する際に自動で実行されます。Initメソッドを2回以上実行する事は出来ません。

- パラメータ

| 引数          | 型               | 初期値 |
| ------------- | ---------------- | ------ |
| ArgContext    | PostgresOrmContext |        
| ArgConnection | ADODB.Connection |        |
| ArgSchema     | String           |        |
| ArgTable      | String           |        |

#### SetDistinct

読み込み後の重複レコードを纏めます。



#### SetAddColumn

読み込むカラム名を追加します。複数回実行すると、カラム名が順次追加されます。オプション引数でSQL関数の使用及び、カラム名のを変更が出来ます。チェーンメソッドに対応します。

- パラメータ

| 引数           | 型                                       | 初期値       |
| -------------- | ---------------------------------------- | ------------ |
| ArgColumn      | String                                   |              |
| ArgSqlFunction | PostgresOrmSqlFunction(ユーザー定義列挙) | poNoFunction |
| ArgColumnName  | String                                   | ""           |



#### SetAddColumns

読み込むカラム名を複数追加します。SetAddColumnとは異なり、SQL関数の使用及び、カラム名の変更は出来ません。チェーンメソッドに対応します。

- パラメータ

| 引数       | 型                     | 初期値 |
| ---------- | ---------------------- | ------ |
| ArgColumns | String()及びCollection |        |



#### SetAddWhere

読み込みの条件を追加します。複数回実行すると、条件が順次追加されます。チェーンメソッドに対応します。

- パラメータ

| 引数           | 型                                       | 初期値 |
| -------------- | ---------------------------------------- | ------ |
| ArgColumn      | String                                   |        |
| ArgWhereType   | PostgresOrmWhereType(ユーザー定義列挙)   |        |
| ArgValue       | Variant                                  |        |
| ArgConnectType | PostgresOrmConnectType(ユーザー定義列挙) | poAnd  |

>  whereTypeがpoInの場合は、要素数が複数の配列またはCollectionを指定します。またpoBetweenの場合は、要素数が2つの配列またはCollectionを指定します。それ以外はStringを指定します



#### SetAddGroupBy

読み込み前にグループ化する条件を追加します。複数回実行すると、条件が順次追加されます。チェーンメソッドに対応します。

- パラメータ

| 引数      | 型     | 初期値 |
| --------- | ------ | ------ |
| ArgColumn | String |        |

​	

#### SetAddHaving

読み込み後にグループ化する条件を追加します。複数回実行すると、条件が順次追加されます。オプション引数でSQL関数の使用及び、カラム名のを変更が出来ます。チェーンメソッドに対応します。

- パラメータ

| 引数           | 型                                       | 初期値 |
| -------------- | ---------------------------------------- | ------ |
| ArgColumn      | String                                   |        |
| ArgWhereType   | PostgresOrmWhrerType(ユーザー定義列挙)   |        |
| ArgValue       | string                                   |        |
| ArgSqlFunction | PostgresOrmSqlFunction(ユーザー定義列挙) |        |
| ArgConnectType | PostgresOrmConnectType(ユーザー定義列挙) | poAnd  |



#### SetAddOrderBy

並び替えの条件を追加します。複数回実行すると、条件が順次追加されます。

- パラメータ

| 引数        | 型                                    | 初期値 |
| ----------- | ------------------------------------- | ------ |
| ArgColumn   | String                                |        |
| ArgSortType | PostgresOrmSortType(ユーザー定義列挙) | poAsc  |



#### SetLimit

読み込むデータ行数の上限を設定します。 

- パラメータ

| 引数     | 型   | 初期値 |
| -------- | ---- | ------ |
| ArgValue | Long |        |



#### SetOffset

読み込むデータの開始行数の設定します。

- パラメータ

| 引数     | 型   | 初期値 |
| -------- | ---- | ------ |
| ArgValue | Long |        |



#### ClearDistinct

SetDistinctの設定を削除します。



#### ClearColumns

SetAddColumnおよびSetAddColumnsで設定したカラム情報を削除します。



#### ClearWhere

SetAddWhereで設定した条件を削除します。



#### ClearGroupBy

SetGroupByで設定した条件を削除します。



#### ClearHaving

SetHavingで設定した条件を削除します。



#### ClearOrderBy

SetOrderByで設定した条件を削除します。



#### ClearLimit

SetLimitで設定した条件を削除します。



#### ClearOffset

SetOffsetで設定した条件を削除します。



#### GetSql

各種メソッドで設定した情報から実際に実行するSQL文を返します。

- 戻り値

String



#### Commit

各種メソッドで設定した情報からPostgresのデータを更新します。

- パラメータ

| 引数               | 型             | 初期値            |
| ------------------ | -------------- | ----------------- |
| ArgCursorType      | CursorTypeEnum | adOpenUnspecified |
| ArgLockType        | LockTypeEnum   | adLockUnspecified |
| ArgShouldOutputSql | Boolean        | False             |

- 戻り値

ADODB.Recordset



### PostgresOrmUpdate

PostgresOrmクラスが内包しているサブクラスです。データベースに対してレコードを更新する各種メソッドを提供します。



#### Init

インスタンス化したクラスを初期化しますが、Updateサブクラスを処理する際に自動で実行されます。Initメソッドを2回以上実行する事は出来ません。

- パラメータ

| 引数          | 型               | 初期値 |
| ------------- | ---------------- | ------ |
| ArgContext    | PostgresOrmContext |        
| ArgConnection | ADODB.Connection |        |
| ArgSchema     | String           |        |
| ArgTable      | String           |        |



#### SetColumnAndValue

更新するカラム名と、そのデータを追加します。複数回実行すると、カラム名とデータが順次追加されます。チェーンメソッドに対応します。

- パラメータ

| 引数      | 型     | 初期値 |
| --------- | ------ | ------ |
| ArgColumn | String |        |
| ArgValue  | String |        |



#### SetAddWhere

削除の条件を追加します。複数回実行すると、条件が順次追加されます。チェーンメソッドに対応します。

- パラメータ

| 引数           | 型                                       | 初期値 |
| -------------- | ---------------------------------------- | ------ |
| ArgColumn      | String                                   |        |
| ArgWhereType   | PostgresOrmWhereType(ユーザー定義列挙)   |        |
| ArgValue       | String                                   |        |
| ArgConnectType | PostgresOrmConnectType(ユーザー定義列挙) | poAnd  |

> whereTypeがpoInの場合は、要素数が複数の配列またはCollectionを指定します。またpoBetweenの場合は、要素数が2つの配列またはCollectionを指定します。それ以外はStringを指定します



#### ClearColumnsAndValues

SetColumnAndValueで設定したカラムとデータ情報を削除します。



#### ClearWhere

SetAddWhereで設定した条件を削除します。



#### GetSql

各種メソッドで設定した情報から実際に実行するSQL文を返します。

- 戻り値

String



#### Commit

各種メソッドで設定した情報からPostgresのデータを更新します。

- パラメータ

| 引数               | 型             | 初期値            |
| ------------------ | -------------- | ----------------- |
| ArgCursorType      | CursorTypeEnum | adOpenUnspecified |
| ArgLockType        | LockTypeEnum   | adLockUnspecified |
| ArgShouldOutputSql | Boolean        | False             |

- 戻り値

ADODB.Recordset



### PostgresOrmDelete

PostgresOrmクラスが内包しているサブクラスです。データベースに対してレコードを削除する各種メソッドを提供します。



#### Init

インスタンス化したクラスを初期化しますが、Deleteサブクラスを処理する際に自動で実行されます。Initメソッドを2回以上実行する事は出来ません。

- パラメータ

| 引数          | 型               | 初期値 |
| ------------- | ---------------- | ------ |
| ArgContext    | PostgresOrmContext |        
| ArgConnection | ADODB.Connection |        |
| ArgSchema     | String           |        |
| ArgTable      | String           |        |



#### SetAddWhere

削除の条件を追加します。複数回実行すると、条件が順次追加されます。チェーンメソッドに対応します。

- パラメータ

| 引数           | 型                                       | 初期値 |
| -------------- | ---------------------------------------- | ------ |
| ArgColumn      | String                                   |        |
| ArgWhereType   | PostgresOrmWhereType(ユーザー定義列挙)   |        |
| ArgValue       | String                                   |        |
| ArgConnectType | PostgresOrmConnectType(ユーザー定義列挙) | poAnd  |

> whereTypeがpoInの場合は、要素数が複数の配列またはCollectionを指定します。またpoBetweenの場合は、要素数が2つの配列またはCollectionを指定します。それ以外はStringを指定します



#### ClearWhere

SetAddWhereで設定した条件を削除します。



#### GetSql

各種メソッドで設定した情報から実際に実行するSQL文を返します。

- 戻り値

String



#### Commit

各種メソッドで設定した情報からPostgresのデータを削除します。

- パラメータ

| 引数               | 型             | 初期値            |
| ------------------ | -------------- | ----------------- |
| ArgCursorType      | CursorTypeEnum | adOpenUnspecified |
| ArgLockType        | LockTypeEnum   | adLockUnspecified |
| ArgShouldOutputSql | Boolean        | False             |

- 戻り値

ADODB.Recordset



### PostgresOrmMeta

PostgresOrmクラスが内包しているサブクラスです。データベースのメタ情報を取得する、各種メソッドを提供します。



#### Init

インスタンス化したクラスを初期化しますが、Metaサブクラスを処理する際に自動で実行されます。Initメソッドを2回以上実行する事は出来ません。

- パラメータ

| 引数          | 型               | 初期値 |
| ------------- | ---------------- | ------ |
| ArgContext    | PostgresOrmContext |        
| ArgConnection | ADODB.Connection |        |
| ArgSchema     | String           |        |
| ArgTable      | String           |        |



#### GetSchemas

PostgresOrmクラスで接続したデータベースのスキーマ一覧を返します。

- パラメータ

| 引数          | 型                                      | 初期値  |
| ------------- | --------------------------------------- | ------- |
| ArgReturnType | PostgresOrmReturnType(ユーザー定義列挙) | poArray |

- 戻り値

String()またはCollection (returnTypeで指定した型)



#### GetTables

PostgresOrm.SetSchemaで設定したスキーマ内のテーブル一覧を返します。オプション引数で他スキーマを指定出来ます。

- パラメータ

| 引数          | 型                                      | 初期値                |
| ------------- | --------------------------------------- | --------------------- |
| ArgSchema     | String                                  | PostgresOrm.GetSchema |
| ArgReturnType | PostgresOrmReturnType(ユーザー定義列挙) | poArray               |

- 戻り値

String()またはCollection (returnTypeで指定した型)



#### GetColumns

PostgresOrm.SetSchema.SetTableで設定したテーブルのカラム一覧を返します。オプション引数で他スキーマ、他テーブルを指定出来ます。

- パラメータ

| 引数          | 型                                      | 初期値                |
| ------------- | --------------------------------------- | --------------------- |
| ArgSchema     | String                                  | PostgresOrm.GetSchema |
| ArgTable      | String                                  | PostgresOrm.GetTable  |
| ArgReturnType | PostgresOrmReturnType(ユーザー定義列挙) | poArray               |

- 戻り値

String()またはCollection (returnTypeで指定した型)



#### ExistsTable

PostgresOrm.SetSchema.SetTableで設定したテーブルの存在を真偽値で返します。オプション引数で他スキーマ、他テーブルを指定出来ます。

- パラメータ

| 引数      | 型     | 初期値                |
| --------- | ------ | --------------------- |
| ArgSchema | String | PostgresOrm.GetSchema |
| ArgTable  | String | PostgresOrm.GetTable  |

- 戻り値

Boolean



#### ExtstsField

PostgresOrm.SetSchema.SetTableで設定したテーブルに対し、データの存在を真偽値で返します。オプション引数で他スキーマ、他テーブルを指定出来ます。

- パラメータ

| 引数      | 型     | 初期値                |
| --------- | ------ | --------------------- |
| ArgColumn | String |                       |
| ArgValue  | String |                       |
| ArgSchema | String | PostgresOrm.GetSchema |
| ArgTable  | String | PostgresOrm.GetTable  |

- 戻り値

Boolean



### ユーザー定義列挙型

主に引数で使用するユーザー定義列挙型です。



#### PostgresOrmSqlFunction

SQL関数を使用する際に指定する定義です。

| メンバ       | 備考                                       |
| ------------ | ------------------------------------------ |
| poNoFunction | SQL関数の使用しません                      |
| poCount      | グループ化した行数を返します               |
| poSum        | グループ化したフィールドの合計を返します   |
| poAvg        | グループ化したフィールドの平均を返します   |
| poMax        | グループ化したフィールドの最大値を返します |
| poMin        | グループ化しらフィールドの最小値を返します |
| poAbs        | 絶対値を返します                           |
| poRound      | 小数点の位置で四捨五入して返します         |



#### PostgresOrmWhereType

条件式を指定する際に使用する、条件式記号の定義です。

| メンバ           | 条件式記号          | 備考         |
| ---------------- | ------------------- | ------------ |
| poEuqal          | =                   | 等式         |
| poNotEqual       | <>                  | 非等式       |
| poLess           | <                   | より小さい   |
| poLessOrEqual    | <=                  | 以下         |
| poGraater        | >                   | より大きい   |
| poGreaterOrEqual | >=                  | 以上         |
| poLike           | Like                | 含む         |
| poIs             | Is                  | Nullと一致   |
| poNotIs          | Not Is              | Nullと不一致 |
| poIn             | In('X', 'Y', 'Z' )  | 複数の条件   |
| poBetween        | Between 'X' And 'Y' | 条件の間     |



#### PostgresOrmConnectType

複数の条件式を指定する際に使用する、論理式の定義です。

| メンバ | 条件式記号 | 備考   |
| ------ | ---------- | ------ |
| poAnd  | And        | および |
| poOr   | Or         | または |



#### PostgresOrmSortType

並び替えを指定する際に使用する定義です。

| メンバ | 備考 |
| ------ | ---- |
| poAsc  | 昇順 |
| poDesc | 降順 |



#### PostgresOrmReturnType

連続した戻り値の際に、型を指定する定義です。

​	

| メンバ       | 備考       |
| ------------ | ---------- |
| poArray      | 配列       |
| poCollection | Collection |



PostgresOrmContext

各処理に必要な各種情報の定義です。

| メンバ  | 備考                                |
| ------- | ----------------------------------- |
| poRetry | コネクションSQLコマンドの再実行回数 |



## 例

### データベース構成

```tex
example
├─public1
│  ├─test1
│  ├─test2
│  └─test3
├─public2
└─public3
```



### table1

```sql
CREATE TABLE public1.table1 (
	id text NULL,
	name text NULL,
	"number" int4 NULL
);
```

| id     | name  | number |
| ------ | ----- | ------ |
| 000001 | user1 | 10     |
| 000002 | user2 | 20     |
| 000003 | user3 | 10     |



### SQLを直接実行する

```vbscript
Sub ExampleExecuteSql()

    ' 変数の定義
    Dim postgres As PostgresOrm
    Dim rs As New ADODB.Recordset

    ' PostgresOrmをインスタンス化
    Set postgres = New PostgresOrm

    ' 接続するデータベースの情報を定義
    ' ExampleではローカルのPostgresに対しログイン
    ' データベースはexample、パスワードは*****とする
    postgres.Init "localhost", "example", "*****"

    ' SQLを直接発行
    ' 実行結果はADODB.Recordsetとして返る
    Set rs = postgres.ExecuteSql("select * from public1.table1")

    ' 取得したレコードセットをShees1にコピーする
    Sheet1.Cells(1, 1).CurrentRegion.ClearContents
    Sheet1.Cells(1, 1).CopyFromRecordset rs

End Sub
```

イミディエイト

```tex
select * from public1.table1
```



### データベース情報を取得する

```vbscript
Sub ExampleMeta()

    ' 変数の定義
    Dim postgres As New PostgresOrm
    Dim ary As Variant

    ' 接続するデータベース情報を定義
    ' スキーマとテーブルを定義
    With postgres
        .Init "localhost", "example", "*****"
        .SetSchema "public1"
        .SetTable "table1"

        ' 情報取得のMetaサブクラスをインスタンス化
        With .Meta
        
            ' 接続したデータベースのスキーマ一覧を取得
            ary = .GetSchemas
            Debug.Print Join(ary, ", ")

            ' SetSchemaで設定したpublic1内のテーブルを一覧を取得
            ary = .GetTables
            Debug.Print Join(ary, ", ")

            ' SetTableで設定したtable1のカラム名を一覧で取得
            ary = .GetColumns
            Debug.Print Join(ary, ", ")
        End With
    End With

    ' 以下の様な記述も可能
    ary = postgres.SetSchema("public1").Meta.GetTables
    Debug.Print Join(ary, ", ")
    
    ary = postgres.SetSchema("public1").SetTable("table1").Meta.GetColumns
    Debug.Print Join(ary)

End Sub
```

イミディエイト

```tex
public1, public2, public3
table1, table2, table3
id, name, number
table1, table2, table3
id name number
```

### データを追加する

```vbscript
Sub ExampleCreate()

    ' 変数の定義
    Dim postgres As New PostgresOrm

    ' 接続するデータベース情報を定義
    ' スキーマとテーブルを定義
    With postgres
        .Init "localhost", "example", "*****"
        .SetSchema "public1"
        .SetTable "table1"

        ' データ追加のCreateサブクラスをインスタンス化
        With .Create

            ' データを挿入するカラム名を配列で定義
            .SetAddColumns Array("id", "name", "number")
            
            ' データを挿入する値を配列で定義
            ' カラム名の順序と合わせる必要有り
            .SetAddValues Array("000004", "user4", 30)

            ' データベースへ書き込み
            .Commit
        End With

        ' 複数行のデータを一度に追加する事も可能
        With .Create
            .SetAddColumns Array("id", "name", "number")
            .SetAddValues Array("000005", "user5", 20)
            .SetAddValues Array("000006", "user6", 10)
            .Commit
        End With
    End With

End Sub
```

イミディエイト

```tex
insert into public1.table1(id, name, number) values('000004', 'user4', '30')
insert into public1.table1(id, name, number) values('000005', 'user5', '20'), ('000006', 'user6', '10')
```

### データを取得する

```vbscript
Sub ExampleRead()

    ' 変数の定義
    Dim postgres As New PostgresOrm
    Dim rs As New Recordset

    ' 接続するデータベース情報を定義
    ' スキーマとテーブルを定義
    With postgres
        .Init "localhost", "example", "*****"
        .SetSchema "public1"
        .SetTable "table1"

        ' データ取得のReadサブクラスをインスタンス化
        ' 全てのデータを取得
        Set rs = .Read.Commit()

        ' カラムと結果の上限を指定して取得
        With .Read
            .SetAddColumns Array("id", "name")
            .SetLimit 2
            Set rs = .Commit()
        End With

        ' カラムを指定してグループ化し、その行数をカウント
        With .Read
            .SetAddColumn "number", , "番号"
            .SetAddColumn "*", poCount, "カウント"
            .SetAddGroupBy "number"
            Set rs = .Commit()
        End With

        ' 条件を指定して取得
        With .Read
            .SetAddWhere "number", poEqual, 10
            Set rs = .Commit()
        End With

        ' 複数の条件を指定して取得
        With .Read
            .SetAddWhere "id", poLike, "%3"
            .SetAddWhere "number", poGreaterOrEqual, 20, poOr
            Set rs = .Commit()
        End With
    End With

End Sub
```

イミディエイト

```tex
select * from public1.table1
select id, name from public1.table1 limit 2
select number as 番号, count(*) as カウント from public1.table1 group by number
select * from public1.table1 where number = '10'
select * from public1.table1 where id like '%3' or number >= '20'
```



### データを更新する

```vbscript
Sub ExampleUpdate()

    ' 変数の定義
    Dim postgres As New PostgresOrm
    Dim rs As New Recordset

    ' 接続するデータベース情報を定義
    ' スキーマとテーブルを定義
    With postgres
        .Init "localhost", "example", "******"
        .SetSchema "public1"
        .SetTable "table1"

        ' データ取得のUpdateサブクラスをインスタンス化
        ' 条件を指定して更新
        With .Update
            .SetAddColumnAndValue "number", 100
            .SetAddWhere "id", poEqual, "000001"
            .Commit
        End With
    End With

End Sub
```

イミディエイト

```tex
update public1.table1 set number = '100' where id = '000001'
```



### データを削除する

```vbscript
Sub ExampleDelete()

    ' 変数の定義
    Dim postgres As New PostgresOrm
    Dim rs As New Recordset

    ' 接続するデータベース情報を定義
    ' スキーマとテーブルを定義
    With postgres
        .Init "localhost", "example", "*****"
        .SetSchema "public1"
        .SetTable "table1"

        ' データ削除のDeleteサブクラスをインスタンス化
        ' 全てを削除
        .Delete.Commit

        ' 条件を指定して削除
        With .Delete
            .SetAddWhere "number", poNotEqual, 10
            .Commit
        End With
    End With

End Sub
```

イミディエイト

```tex
delete from public1.table1
delete from public1.table1 where number <> '10'
```



## 著者

- 作成者：やまと興業株式会社 生産支援部 吉澤洋人
- e-mail：h.yoshizawa@yamato-industrial.co.jp



## 免責

このプログラムについては動作、性能の保証はしていません。また、使用によって生じた、いかなる損害に於いては一切の責任を負いかねます。使用の前にバックアップ等のデータの保全対策や、運用のテストを行って下さい。なお改変、再配布については自由にして構いませんが、その場合も一切の責任を負いかねます。

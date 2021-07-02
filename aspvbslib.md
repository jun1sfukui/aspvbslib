# ASPVBSLIB
レガシーとなったClassic ASP(VBScript)でのコーディングを、可能な限りVB.NET的に行えるようにするライブラリです。

VBScriptにおいてストレスになるものの一つが非常に使いにくい配列ですが、リスト処理周りを補助するArrayListクラスを中心に、.NET Frameworkの Linq風メソッドや「Entity Framework」的なO/Rマッピング処理を行えるようにしました。

#### 例) ArrayListの使用1
```VBScript
Dim list
Set list = New ArrayList()
list.Add "Tokyo"
list.Add "Osaka"
list.Add "Kyoto"
list.Add "Fukui"

Dim item
For Each item In list.Items
    Response.Write "<li>" & item & "</li>"
Next

'先頭要素を出力
Response.Write list.FirstOrDefault("") ' "Tokyo"を出力

'リストが空でも大丈夫
Set list = New ArrayList()
Response.Write list.FirstOrDefault("") ' ""を出力
```

#### 例) ArrayListの使用2
```VBScript
Class Person
    Public fullName
    Public age

    'コンストラクタ風デフォルトプロパティ
    Public Default Property Get Constructor(fullName_, age_)
        fullName = fullName_
        age = age_
        Set Constructor = Me
    End Property
End Class

Dim list
Set list = New ArrayList()
list.Add (New Person)("田中", 24)
list.Add (New Person)("斎藤", 31)
list.Add (New Person)("佐藤", 19)
list.Add (New Person)("鈴木", 48)

'Whereメソッドで age < 30 の要素だけに絞り込む
Dim item
For Each item In list.Where("item.age < p", 30).Items
    Response.Write "<li>" & item.fullName & "</li>"  '田中, 佐藤 が出力される
Next
```

#### 例) RequestとDBの読み込み、JSON出力
```VBScript
'リクエストの内容をユーザー定義オブジェクトに読み込む（同名のプロパティに自動設定される）
Dim req
Set req = New RequestModel
LoadRequest(req)

'レコードセットをユーザー定義のEmpクラスのArrayListとして取得する
Dim employees
Set employees = SqlQuery(DbConn, "Emp", "SELECT * FROM emp")

'リクエストにempnoが設定されていたら、employeesを絞り込み、結果をユーザー定義のViewModelに設定する。
Dim vm
Set vm = New ViewModel
If req.empno <> 0 Then
    Set vm.employees = employees.Where("item.empno = p", req.empno)
Else
    Set vm.employees = employees
End If

'ViewModelをJSON形式でブラウザに出力する。
Response.Write JSONValue(vm)
```

今から新規のClassic ASPシステムを作ることはないと思いますが、既存のClassic ASPシステムをどうしても延命しなくてはならなくなった場合などに、新規開発部分でこのライブラリを活用し、開発コストを下げてください。

皆様のClassic ASP開発現場が少しでも明るくなることを祈っております。

# Precautions

VBScriptにはラムダ式や関数オブジェクトがない為、文字列としてラムダ式（に相当する評価式）を渡し、`Execute`関数や`Eval`関数を使って実行しています。

また同様に、動的なプロパティへのアクセスも`Execute`や`Eval`で実現しています。
 
`Execute`関数や`Eval`関数は実行コストが高い為、CPU負荷がシビアな環境には適していません。 

しかし、CPUパワーに余裕のある環境であれば、バグの少ない簡潔な記述でClassic ASPのシステムを構築することができるでしょう。 

使用の際には、本格的な導入の前に局所的に導入し、十分なパフォーマンス計測を行ってください。

#### このライブラリが適している処理
- アクセス頻度のそれほど高くない、社内用システム
- 1件分の情報を表示するような画面および処理
- 1件分の情報を入力して登録するような画面および処理
- 最大でも数百件程度のレコードを読み込んで表示するような画面および処理

#### このライブラリが適さない処理
- 日常的に大量のアクセスを処理する必要のある公開システム => 1件1件の処理はそれほど重く感じなくても、確実にCPUパワーを消費している為、大量のアクセスを捌く際には十分なパフォーマンス計測が必要です。
- 1度に数千～数十万件のデータを読み込む処理 => `SqlQuery`を使う際には十分なパフォーマンス計測が必要です。恐らく現実的な速度は出ないと思われます。あまりにも処理が遅くなる箇所については、既存の`ADODB.Recordset`を直接使用した一般的な記述に置き換えてください。
- 大量のレコード処理の速度が遅いと感じた場合`<ClassName>Props`クラスの定義によってパフォーマンスが大きく改善することがあります。詳しくは`SqlQuery`の説明を参照してください。

# Usage
ライブラリの組み込みに必要なファイルは以下のファイルのみです。それ以外はユニットテストやサンプル用のファイルです。

- adovbs.inc
- aspvbslib-core.inc
- aspvbslib-data.inc
- aspvbslib-web.inc
- aspvbslib-test.inc

以下のように、必要な機能のみをincludeして使用します。

## aspvbslib-coreのみ使用
```Classic ASP
<!--#include file="adovbs.inc"-->
<!--#include file="aspvbslib-core.inc"-->
```

## aspvbslib-dataを使用
```Classic ASP
<!--#include file="adovbs.inc"-->
<!--#include file="aspvbslib-core.inc"-->
<!--#include file="aspvbslib-data.inc"-->
```

## aspvbslib-webを使用
```Classic ASP
<!--#include file="adovbs.inc"-->
<!--#include file="aspvbslib-core.inc"-->
<!--#include file="aspvbslib-web.inc"-->
```

## aspvbslib-testを使用
```Classic ASP
<!--#include file="adovbs.inc"-->
<!--#include file="aspvbslib-core.inc"-->
<!--#include file="aspvbslib-test.inc"-->
```

## 全て使用
```Classic ASP
<!--#include file="adovbs.inc"-->
<!--#include file="aspvbslib-core.inc"-->
<!--#include file="aspvbslib-data.inc"-->
<!--#include file="aspvbslib-web.inc"-->
<!--#include file="aspvbslib-test.inc"-->
```


# Features

## aspvbslib-core.inc
便利な型付き`ArrayList`クラスや、VBにはあるのにVBScriptになくて困る`IIf`関数、配列を便利に扱える`InArray`関数や`ArrayIndexOf`関数、文字列編集に便利な`VBDateFormat`関数、`VBStringFormat`関数など、VBScriptをもっとシンプルに、宣言的に記述できるようになるものを揃えました。

---
### ArrayList
動的配列のように使えるリストクラスです。この例ではプリミティブな値を使っていますが、もちろん、ユーザー定義クラスのインスタンスオブジェクトも格納可能です。内部では`Dictionary`を使ってリストを管理しています。

Itemsプロパティで、リストの内容をイテレーション可能な配列として取得します。

```VBScript
Dim list
Set list = New ArrayList()
list.Add "Tokyo"
list.Add "Osaka"
list.Add "Kyoto"
list.Add "Fukui"

Dim item
For Each item In list.Items
    Response.Write "<li>" & item & "</li>"
Next
```

### *ArrayList*.ItemType
リストの要素の型定数を取得・設定します。型定数は、`vbInteger`, `vbLong`, `vbDecimal`, `vbSingle`, `vbDouble`, `vbDate`, `vbVariant` 等です。

ユーザー定義型の場合は`Nothing`の場合には、`vbObject` になります。
`VarType`関数で取得できる値です。

このプロパティは`Add`時に自動的に設定されます。空のリストの状態で型を指定したい場合にのみ、ユーザーが明示的に指定することができます。

### *ArrayList*.ItemClassName
`ItemType`が `vbObject` の場合に、リストの要素のクラス名を取得・設定します。
`TypeName`関数で取得できる値です。

`ItemType`が `vbObject` 以外の時は、空白が返されます。

このプロパティは`Add`時に自動的に設定されます。空のリストの状態で型を指定したい場合にのみ、ユーザーが明示的に指定することができます。

### *ArrayList*.Add(item)
リストに要素を追加します。

`ItemType`が空の場合、最初に`Add`を呼び出した際の要素の型が自動的に`ItemType`、`ItemClassName`に設定されます。

`ItemType`が設定された後、`ItemType`が異なるitemを`Add`しようとすると、例外が発生します。

`ArrayList`の生成後、リストが空の状態で`ItemType`や`ItemClassName`を設定する事もできます。

### *ArrayList*.Count
リストの要素数を返します。空要素の場合は0が返されます。

```VBScript
If list.Count > 0 Then
    :
End If
```

### *ArrayList*.Items
リストの内容を配列で取得します。実際には内部で保持している`Dictionary`型の`Items`プロパティがそのまま渡されます。

このプロパティで`For Each`ループを使うことができます。

```VBScript
Dim item
For Each item In list.Items
    Response.Write "<li>" & item & "</li>"
Next
```

### *ArrayList*.Item(index)
indexの位置の要素を取得します。indexは0から始まる整数です。最大値はCount - 1です。

```VBScript
Dim list
Set list = New ArrayList()
list.Add "Tokyo"    'index = 0
list.Add "Osaka"    'index = 1
list.Add "Kyoto"    'index = 2
list.Add "Fukui"    'index = 3

Dim item
item = list.Item(2) '"Kyoto"を取得します
```

### *ArrayList*.Remove(item)
リストから指定した要素を削除します。

```VBScript
Dim list
Set list = New ArrayList()
list.Add "Tokyo"    'index = 0
list.Add "Osaka"    'index = 1
list.Add "Kyoto"    'index = 2
list.Add "Fukui"    'index = 3

Dim count
count = list.Count ' 4を取得します

list.Remove "Kyoto"
count = list.Count ' 3を取得します

Dim item
item = list.Item(2) '"Fukui"を取得します
```

### *ArrayList*.RemoveAt(index)
リストから指定したindexの要素を削除します。

```VBScript
Dim list
Set list = New ArrayList()
list.Add "Tokyo"    'index = 0
list.Add "Osaka"    'index = 1
list.Add "Kyoto"    'index = 2
list.Add "Fukui"    'index = 3

Dim count
count = list.Count ' 4を取得します

list.RemoveAt 0
count = list.Count ' 3を取得します

Dim item
item = list.Item(0) '"Osaka"を取得します
```

### *ArrayList*.Contains(item)
指定した値に一致する要素があれば`True`を、なければ`False`を返します。

ユーザー定義型を格納している場合、同一インスタンスの場合にのみ一致とみなされます（つまり、`A Is B`で判定する）。

プリミティブ型の場合、値が同じかどうかで判定します。

```VBScript
If list.Contains("abc") Then
    :
End If
```

### *ArrayList*.IndexOf(item)
指定した値に一致する要素があればその位置（0～）を、なければ-1を返します。

ユーザー定義型を格納している場合、同一インスタンスの場合にのみ一致とみなされます（つまり、`A Is B`で判定する）。

プリミティブ型の場合、値が同じかどうかで判定します。

```VBScript
Dim list
Set list = New ArrayList()
list.Add "Tokyo"    'index = 0
list.Add "Osaka"    'index = 1
list.Add "Kyoto"    'index = 2
list.Add "Fukui"    'index = 3

Dim index
index = list.IndexOf("Kyoto") '2を取得します。
index = list.IndexOf("France")  '-1を取得します。
```

### *ArrayList*.Where(expr, p)
リストを条件に合致するよう絞り込んだ結果を返します。

- 引数のラムダ式文字列では、`"item"`というキーワードで各要素にアクセスできます。
- また、第二引数を`"p"`というキーワードで参照できます。
- 自分自身への変更は行いません。

```VBScript
'年齢が30際未満の社員名一覧を出力
Set listU30 = list.Where("item.Age < p", 30)

Dim item
For Each item In listU30.Items
    Response.Write "<li>" & item.Name & "</li>"
Next
```

### *ArrayList*.Map(expr, p)
リストの写像を作ります。Linqでいう`Select`に該当します。Selectという文字列はVBScriptでは予約語の為、関数型言語で一般的に使われている`Map`というメソッド名に変更しています。

- 引数のラムダ式文字列では、`"item"`というキーワードで各要素にアクセスできます。
- また、第二引数を`"p"`というキーワードで参照できます。
- 自分自身への変更は行いません。

通常は、要素の特定のプロパティのみのリストを作りたい時などに使用します。

```VBScript
'"名前(年齢)"形式で出力。
Dim item
For Each item In list.Map("item.Name & ""("" & item.Age & "")""", Nothing)
    Response.Write "<li>" & item & "</li>"
Next
```

### *ArrayList*.Select(expr, p)
`Select`はVBScriptの予約語ですが、いちおう、`Map`メソッドの別名として`Select`も用意しました。こちらがお好みの場合はこちらをご利用下さい。使い方はMapと同様です。

### *ArrayList*.OrderByAsc(expr)
リストの要素を`expr`で評価した値で昇順ソートした結果を返します。

自分自身への変更は行いません。

```VBScript
Dim sortedList
Set sortedList = list.OrderByAsc("list.empno")
```

### *ArrayList*.OrderByDesc(expr)
リストの要素を`expr`で評価した値で降順ソートした結果を返します。

自分自身への変更は行いません。

```VBScript
Dim sortedList
Set sortedList = list.OrderByDesc("list.empno")
```

### *ArrayList*.FirstOrDefault(defaultValue)
リストの要素の先頭を取得します。存在しない場合、指定した`defaultValue`を返します。

主に、`Where`で特定の要素を検索した後、`Map`で文字列や数値等のプリミティブ値に変換し、その要素を取得する際に用います。

オブジェクトのまま取得したい場合は、`FindFirst`メソッドを使った方がシンプルに記述できることがあるでしょう。

```VBScript
Dim empname
empname = list.Where("item.empno = p", 3).Map("item.empname", Nothing).FirstOrDefault("")
```

### *ArrayList*.LastOrDefault(defaultValue)
`FirstOrDefault`とは違い、末尾の項目を取得します。存在しない場合の扱いは`FirstOrDefault`と同じです。

```VBScript
If selectedEmpno = list.Map("item.empno", Nothing).LastOrDefault(0) Then
    Response.Write "最終項目です"
End If
```

### *ArrayList*.FindFirst(expr, p)
`exprt`が`True`になる最初の要素のみを含む`ArrayList`を返します。存在しない場合、（`Nothing`ではなく）空の`ArrayList`を返します。

- 引数のラムダ式文字列では、`"item"`というキーワードで各要素にアクセスできます。
- また、第二引数を`"p"`というキーワードで参照できます。

```VBScript
Dim first
Set first = list.FindFirst("item.empno = p", 3)
If first.Count > 0 Then
    Response.Write first(0).empname
End If
```

### *ArrayList*.Reverse()
リストの順序を末尾から先頭へと反転させます。
このメソッドは、自身の内部の値を反転させる、副作用のあるメソッドです。

```VBScript
list.Reverse
```

---

### IIf(expr, TruePart, FalsePart)
3項演算子で余計な変数を増やさず宣言的な記述を可能にします。

```VBScript
Dim message
message = IIf(list.Count > 0, list.Count & "件のメッセージ", "メッセージがありません")
```

### IsNullOrEmpty(value) 
値が`Null`又は`Empty`又は空文字（`""`）かどうかを示します。

```VBScript
If IsNullOrEmpty(v) Then
    :
End If
```

### NullValue(value, nullval) 
値が`Null`又は`Empty`又は空文字（`""`）の際に、指定した規定値を返します。

```VBScript
Dim inputID
inputID = NullValue(Request.QueryString("ID"), "999")
```

### InArray(searchKey, arr)
配列内に指定した`searchKey`があれば`True`、なければ`False`を返します。

a = 1 Or a = 2 Or a = 5 Or... といった冗長な記述をシンプルにします。SQLでいう IN句 のような感じです。

```VBScript
If InArray(item.deptno, InArray(1, 2, 5)) Then
    :
End If
```

### ArrayIndexOf(searchKey, arr)
VB.NETの`Array.IndexOf`のように使います。

指定した値が配列中にあればその位置を(0～)、なければ-1を返します。

```VBScript
Dim index
index = ArrayIndexOf("Tokyo", arr)
If index <> -1 Then
    :
End If
```

### FindKeyValue(searchKey, arrKeys, arrValues)
Keyの配列とその位置に対応するValueの配列がある場合に（古いVBScript資産によくある形です）、それらをKey-Valueの対として、Keyで検索して同じ位置にあるValueを取得します。見つからなければ空文字を返します。

既存のVBScript資産をどうしても使わなければならない場合に利用します。

新規開発の場合、この関数に頼らず`ArrayList`や`Dictionary`を使ってください。

```VBScript
Dim deptname
deptname= FindKeyValue(item.deptno, arrDeptNo, arrDeptName)
```

### VBDateFormat(dt, format)
VB.NETの日付書式指定文字列で整形できる関数です。複雑な指定はできない簡易版ですが、VBScriptに不足している機能なので重宝します。

```VBScript
Dim dtString
dtString = VBDateFormat(Now, "yyyy年MM月dd日 HH時mm分ss秒")
```
#### 対応パラメータ


| パラメータ | 説明 |
| ---- | ---- |
| yyyy | 西暦4桁 |
| yy | 西暦下2桁 |
| y | 西暦 |
| MM | 月2桁 |
| M | 月 |
| dd | 日2桁 |
| d | 日 |
| HH | 時(24H)2桁 |
| H | 時(24H) |
| mm | 分2桁 |
| m | 分 |
| ss | 秒2桁 |
| s | 秒 |


### VBStringFormat(format, valueOrArray)
VB.NETのようなプレースホルダー付き書式設定（`{0}`や`{1}`等）を行います。プレースホルダーが複数の場合の値は配列で渡します。

```VBScript
Dim message
message = VBStringFormat("Cd={0}は{1}です", Array( "1234", "編集中" ) ) ' => "Cd=1234は編集中です"
```

### VBPadLeftZero(value, length)
左ゼロ埋めを行います。

```VBScript
Dim amount
amout = "12"
amout = VBPadLeftZero(amount, 5) '=> "00012"
```

### RegexIsMatch(input, pattern)
正規表現補助関数（指定したパターンにマッチするかを返す）です。

簡単な正規表現でのパターンマッチングであれば、1行で書くことができます。

```VBScript
If RegexIsMatch(input, "[^a-zA-Z0-9]") Then
 :
End If
```

### GetPropNames(obj)
指定したユーザー定義オブジェクトのプロパティ名一覧を配列で取得します。

クラス名+PropNamesという名前の、プロパティ名を配列で返す関数を事前に定義しておく必要があります。

プロパティ名一覧取得用関数が定義されていない場合、例外が発生します。

基本的にはライブラリ内部で使用する関数ですが、必要に応じてリフレクション機能的に利用することができるでしょう。

```VBScript
Dim propNames
propNames = GetPropNames(obj)
```

### CopyProps(source, target)
指定したオブジェクトの各プロパティを、対象となるオブジェクトにコピーします。

- コピー元とコピー先のオブジェクトは同じプロパティがあれば良く、必ずしも同じクラスである必要はありません。
- コピー先に存在しないプロパティは無視されます。
- プロパティがオブジェクト型の場合、参照がコピーされます（浅いコピー）。
- コピー元のクラス名+PropNamesという名前のプロパティ名取得用関数の定義が必要です。
- プロパティ名取得用関数の定義がない場合、例外が発生します。

```VBScript
Dim emp1
Set emp1 = New Emp

emp1.empno = 5
emp1.empno = "田中"

Dim emp2
Set emp2 = New Emp
CopyProps emp1, emp2

' emp2.empno => 5
' emp2.empname => "田中"
```

### CloneObject(obj)
指定したオブジェクトを元に新しいオブジェクトを複製して返します。

- 指定したオブジェクトの各プロパティを、複製したオブジェクトにコピーします。
- プロパティがオブジェクト型の場合、参照がコピーされます（浅いコピー）。
- 複製元のクラス名+PropNamesという名前のプロパティ名取得用関数の定義が必要です。
- プロパティ名取得用関数の定義がない場合、例外が発生します。

```VBScript
Dim emp1
Set emp1 = New Emp

emp1.empno = 5
emp1.empno = "田中"

Dim emp2
Set emp2 = CloneObject(emp1)

' emp2.empno => 5
' emp2.empname => "田中"

```
---

## aspvbslib-data.inc
データベース等へのアクセス用のライブラリです。
aspvbslib-core.inc が必要です。

### SqlQuery(DbConn, className, SQLorCmd)
Entity Framework風のSQL実行（O/Rマッピング）を行います。

DBへの問い合わせ結果を、指定したユーザー定義クラスのオブジェクトの`ArrayList`へとマッピングします。

指定したユーザー定義クラスにフィールド名と同名のプロパティがあれば、自動的に値が設定されます。

問い合わせ結果が存在しない場合、（`Nothing`ではなく）空の`ArrayList`が返されます。

```VBScript
Class Emp
    Public empno
    Public empname
    Public deptno
End Class

Dim className
className = "Emp"
Dim list
Set list = SqlQuery(DbConn, className, "SELECT * FROM EMP")

Dim item
For Each item In list.Items
    Response.Write "<td>" & item.empno & "</td>"
    Response.Write "<td>" & item.empname & "</td>"
    Response.Write "<td>" & item.deptno & "</td>"
Next
```

`ADODB.Command`オブジェクトや`ADODB.Parameter`オブジェクトを利用したパラメータ付きクエリにも対応します。
```VBScript

Dim deptno = CInt(Request.Form("deptno"))

Dim className
className = "Emp"

Dim cmd
Set cmd = Server.CreateObject("ADODB.Command")
cmd.CommandText = "SELECT * FROM emp WHERE deptno=?"
cmd.parameters.Append cmd.CreateParameter("@deptno", adInteger, adParamInput, , deptno)
Set employees = SqlQuery(DbConn, className, cmd)
```

### *\<ClassName\>*Propsクラス
`SqlQuery`関数は`ADODB.Recordset`の`Fields`コレクションからフィールド名を取得し、与えられたクラス型のオブジェクトに同名のプロパティがある場合にフィールド値をコピーします。

この時、内部では`Execute`関数が使用されますが、実行時に文字列をスクリプトとして実行する為にCPUに通常よりも大きな負荷がかかります。

CPU負荷を減らす為に、プロパティ名を示す文字列と実際のプロパティを対応付ける`<ClassName>Props`クラスを定義することで、`Execute`関数を使わずにO/Rマッピングを行うことが可能です。

#### *\<ClassName\>*Propsクラスの例(Empクラス)
自作のクラス用の`<ClassName>Props`クラスを作成する場合には、`EmpProps`の定義をそのままコピーし、クラス名および、コメントで指定されているユーザー定義部分のプロパティ名を適切に変更してください。

```VBScript
''' <summary>
''' Empクラス。
''' </summary>
Class Emp
    Public empno
    Public empname
End Class

''' <summary>
''' Empクラス用のPropsクラス。
''' Propプロパティを持ち、GetとLetにて、プロパティ名と実際のプロパティを関連付ける。
''' 定義しておくとSqlQuery関数の速度が向上する。
''' </summary>
Class EmpProps
    '-------------------------------------------------------
    ' 通常、Propsクラスはクラス名以外はそのまま流用し、
    ' Propプロパティのユーザー定義部のみ対象クラス用に実装する。
    '-------------------------------------------------------
    '※obj, Target は変更不可
	private obj

	Public Property Get Target
		Set Target = obj
	End Property

	Public Property Set Target(newObj)
		Set obj = newObj
	End Property

	Public Property Get Prop(propName)
		Dim value

		Select Case propName
            '-- ここからユーザー定義部分 ------
			Case "empno"
				value = obj.empno
			Case "empname"
				value = obj.empname
            '-- ここまで ---------------------
			Case Else
				value = Eval("obj." & propName)
		End Select
		Prop = value
	End Property

	Public Property Let Prop( propName, value )
		Select Case propName
            '-- ここからユーザー定義部分 ------
			Case "empno"
				obj.empno = value
			Case "empname"
				obj.empname = value
            '-- ここまで ---------------------
			Case Else
				Execute("obj." & propName & " = value")
		End Select
	End Property

End Class
```



### DBSqlValue(val)
型に応じたSQL用のリテラル表現を返します。

> **注意**: valはサニタイズ（無害化）処理されません。SQLインジェクション攻撃に備える為にも、外部からの入力値を`DBSqlValue`に与えてはいけません。外部からの入力値を使ってSQLクエリを組み立てる場合には、必ずパラメータ付きクエリを使ってください。



```VBScript
Dim searchName, age, isTarget, startDate, nval
searchName = "Mike"
age = 45
isTarget = True
startDate = #6/21/2021#
nval = Nothing

Dim sqlSearchName, sqlAge, sqlIsTarget, sqlStartDate, sqlNval
sqlSearchName = DBSqlValue(searchName) ' => "'Mike'"
sqlAge = DBSqlValue(age) ' => "45"
sqlIsTarget = DBSqlValue(isTarget) ' => "True"
sqlStartDate = DBSqlValue(startDate) ' => "'2021/6/21 00:00:00'" ※DBSqlValue_DATE_FORMAT 変数で出力形式を変更できる。
sqlNval = DBSqlValue(nval) ' => "Null"
```

---

## aspvbslib-web.inc
ウェブ用のライブラリです。リクエストのO/Rマッピングや、テンプレート記述補助、オブジェクトのJSON出力を用意しました。

aspvbslib-core.inc が必要です。

### LoadRequest(req)
リクエストオブジェクトのO/Rマッピングを行います。

`Request`のパラメータに、ユーザー定義オブジェクトのプロパティと同じものがあれば、自動的にプロパティを設定します。プロパティの型を決めていない場合、値は文字列として設定されます。
```VBScript
'リクエスト取得用モデルクラス
Class RequestModel
    Public empno
    Public deptno
End Class

'RequestModelのプロパティ名一覧を返す関数
Function RequestModelPropNames
    RequestModelPropNames = Array("empno", "deptno")
End Function

Dim req
LoadRequest(req)   ' ?empno=1&deptno=123 のリクエストの際に、req.empno = "1" 及び req.deptno = "123" を実行。

Response.Write req.empno    'empnoは文字列型
Response.Write req.deptno   'deptnoは文字列型
```

型指定によるマッピングも可能です。`Class_Initialize`で各プロパティを指定した型で初期化しておくと、`Request`から読み込んだ値をその型に変換して設定します。変換できなかった場合、自動的にその型の規定値が設定されます。

```VBScript
'リクエスト取得用モデルクラス（型指定あり）
Class RequestModel
    Public empno
    Public deptno

    Sub Class_Initialize
        empno = CLng(0)     'Long型として初期化
        deptno = CLng(0)    'Long型として初期化
    End Sub
End Class

'RequestModelのプロパティ名一覧を返す関数（マッピング対象のクラスと一緒にこの関数を定義すると、マッピング可能になります）
Function RequestModelPropNames
    RequestModelPropNames = Array("empno", "deptno")
End Function

Dim req
Set req = New RequestModel
LoadRequest(req)   ' ?empno=1&deptno=123 のリクエストの際に、req.empno = 1 及び req.deptno = 123 を実行。

Response.Write req.empno    'empnoはLong型
Response.Write req.deptno   'deptnoはLong型
```

また、以下のようにプロパティを初期化することで、同名のリクエストパラメータが複数あるような場合に`ArrayList`として取得する事が可能です。値が1つも指定されなかった場合、空の`ArrayList`が設定されたままになります。

```VBScript
'リクエスト取得用モデルクラス（型指定あり）
Class RequestModel
    Public empno
    Public departments

    Sub Class_Initialize
        empno = CLng(0)     'Long型として初期化
        Set departments = New ArrayList()
        departments.ItemType = vbLong  'Long型のArrayListとして初期化
    End Sub
End Class

'RequestModelのプロパティ名一覧を返す関数（マッピング対象のクラスと一緒にこの関数を定義すると、マッピング可能になります）
Function RequestModelPropNames
    RequestModelPropsName = Array("empno", "departments")
End Function

' ?empno=1&departments=123&departments=456&departments=789 のリクエストの際に、
' req.empno = 1 及び req.departments.Add 123, req.deprtments.Add 456, req.departments.Add 789 を実行。
Dim req
Set req = New RequestModel
LoadRequest(req)   

Response.Write req.empno    'empnoはLong型
Dim deptno
For Each deptno In req.departments.Items
    Response.Write deptno   'deptnoはLong型
Next
```

#### その他の注意事項

- 型変換に失敗すると、その型の規定値（`Long`型ならば0）が設定されます。
- リクエストパラメータに存在しない場合、`Class_Initialize`で初期化されたまま、何も設定されません。これを利用して、パラメータなしの場合の値を定義することができます。

### LoadQueryString, LoadForm 
`LoadRequest`の、`Request.QueryString`, `Request.Form` 版です。

それぞれGET, POSTに対応します。GETかPOSTかが確定している場合は、セキュリティの為にもこちらを使うべきです。

### write_if(output_str, expr)
HTMLテンプレート記述補助関数です。

`expr`に指定した条件が`True`の場合に`output_str`を返します。

<%= %>と組み合わせて、HTMLテンプレート内で使用します。

```Classic ASP
<!-- dept.deptnoとreq.selectedDeptnoが一致する場合に "selected" を出力する -->
<select name="selectedDeptno">
    <% Dim dept: For Each dept In departments.Items %>
    <option <%= write_if("selected", dept.deptno = req.selectedDeptno) %>>
    <% Next %>
</select>
```

### css_display(expr)
HTMLテンプレート記述補助関数です。

`expr`に指定した条件が`False`の場合に`"display:none;"`を返します。

<%= %>と組み合わせて、HTMLテンプレート内のstyle属性の中で使用します。

```Classic ASP
<div style="css_display(emp.empno > 0)">
    <span><%=emp.Name%>
</div>
```

### css_visibility(expr)
HTMLテンプレート記述補助関数です。

`expr`に指定した条件が`True`の時に`"visibility:visible;"`、`False`の時に `"visibility:hidden;"`を返します。

<%= %>と組み合わせて、HTMLテンプレート内のstyle属性の中で使用します。

```Classic ASP
<div style="css_visibility(emp.empno > 0)">
    <span><%=emp.Name%>
</div>
```

### html_options(arrKeys, arrCaptions, selectedKey)
HTMLテンプレート記述補助関数です。

HTMLの`<option>`要素を出力する専用のヘルパー関数です。`value`属性のリスト、見出し文字列のリスト、選択状態としたい`value`値、をそれぞれ渡します。

```Classic ASP
<select name="selectedCountry">
    <%=html_options(Array("US","JP","FR"), Array("米国", "日本", "フランス"), req.selectedCountry) %>
</select>
```

出力（`req.selectedCountry = "JP"`の時）
```HTML
<select name="selectedCountry">
    <option value="US">米国</option>
    <option value="JP" selected>日本</option>
    <option value="FR">フランス</option>
</select>
```

### JSONValue(obj)
VBScriptのオブジェクトをJSON形式の文字列で返します。配列、`ArrayList`、ユーザー定義クラスに対応しています。

変換対象がユーザー定義クラスの場合、プロパティ名一覧を返す為の`<ClassName>PropNames`関数の定義が必要です。例えば`Emp`クラスをJSON化したい場合、`Emp`クラスのプロパティ名一覧を文字列の配列で返す`EmpPropNames`関数の定義が別途必要です。

```VBScript
'社員クラス
Class Emp
    Public empno
    Public empname

    Public Default Property Get Constructor(empno, empname)
        Me.empno = empno
        Me.empname = empname
        Set Constructor = Me
    End Property
End Class

'社員クラスのプロパティ名取得関数
Function EmpPropNames
    EmpPropsNames = Array("empno", "empname")
End Function

'ViewModelクラス
Class ViewModel
    Public selectedEmpno
    Public employees

    Sub Class_Initialize
        selectedDeptno = CLng(0)
        Set employees = New ArrayList
    End Sub

End Class
'ViewModel用プロパティ名取得関数
Function ViewModelPropNames
    ViewModelPropNames = Array("selectedEmpno", "employees")
End Function

Dim vm
Set vm = New ViewModel
vm.departments.Add (New Emp)(1,"emp1")
vm.departments.Add (New Emp)(2,"emp2")
vm.departments.Add (New Emp)(3,"emp3")
vm.selectedEmpno = 2

Response.Write JSONValue(vm)
```

出力
```JSON
{"selectedEmpno":2, "employees":[{"empno":1, "empname":"emp1"}, {"empno":2, "empname":"emp2"}, {"empno":3, "empname":"emp3"}]}
```

#### その他の注意事項

- ダブルバイト文字や特殊文字は `\uXXXX`形式でエンコードされます。
- `Date`型は、規定では`"2021-12-31 23:58:59"`形式の文字列で出力されます。`Date`型の出力形式を変更したい場合には、`JSONValue_DATE_FORMAT` 変数に、`VBDateFormat`関数の第二引数に与える日付整形文字列を設定して下さい。

---

## aspvbslib-test
ユニットテスト用のライブラリですが、基本的にこのライブラリ自身をテストする為に作られた簡易的なもので、`AssertEquals`しか用意されていません。使い方は各テスト用のaspファイルをご覧ください。

```VBScript
Dim test
Set test = New UnitTest

test.Add("MyFunction1Test")
test.Add("MyFunction2Test1")
test.Add("MyFunction2Test2")

test.RunTest
test.ResultHtml

Sub MyFunction1Test
    AssertEquals MyFunction1("hoge"), "fuga"
    AssertEquals MyFunction1(""), ""
End Sub

Sub MyFunction2Test1
    AssertEquals MyFunction2("hoge"), "fuga"
    AssertEquals MyFunction2(""), ""
End Sub

Sub MyFunction2Test2
    Dim obj
    Set obj = New MyClass
    AssertEquals MyFunction1(Nothing), Nothing
    AssertEquals MyFunction1(obj), "message123"
End Sub
```


# Samples

## sample.asp
`ArrayList`と`SqlQuery`を使ったサンプルです。
aspvbslibを使ってどのようにClassic ASPシステムを構築するかの良いサンプルになるでしょう。

同梱のaspvbslib_sample.mdb をADODBのMicrosoft Access Driver (*.mdb)経由で読み込み、ユーザー定義クラスの値を持つ`ArrayList`へと変換しています。

変換した値はHTMLテンプレート上でドロップダウンリストボックスとして表示し、選択された値は`LoadRequest`を使って`RequestModel`にマッピングした上で再度`ViewModel`に反映しています。

処理全体を`Main`という名前のプロシージャでラッピングしていたり、HTML出力処理を`RenderHTML`という名前のプロシージャにしていたり、表示内容を`ViewModel`という名前のクラスにしていたり、リクエスト取得用のモデルを`RequestModel`という名前のクラスにしていたりしますが、これはただの作者の好みで、その辺の命名規約や処理の構造に特に制限はありません。

## sample_checkboxes.asp
チェックボックス配列を`LoadRequest`で`ArrayList`として読み込むサンプルです。

`LoadRequest`が具体的にどのようにブラウザからのリクエスト情報をユーザー定義クラスにマッピングするかの確認にご利用ください。

## sample_vue.asp
Vueを使ってsample.aspを作り直しています。
Classic ASPでも、使い方次第で綺麗にVueとの融合が可能であることを示しています。

## sample_vue_ajax.asp
Vue+Ajaxでsample.aspを作り直しています。

Vueと一緒に使われていることの多いaxiosをAjaxライブラリとして用いています。

最初に一度画面全体が読み込まれた後は、画面のリロードは行われず、データをAjaxで随時読み込んでいるSPA(Single Page Application)です。

Classic ASPでも、使い方次第でSPAを作ることが可能である事を示しています。

## sample_emp_json.asp
sample_vue_ajax.aspから参照しているサービスです。指定した部署に所属する社員リストをJSON形式で返します。

sample_vue_ajax.aspでは、受け取った社員リストのJSONを、画面上の社員ドロップダウンリストボックスとしてバインドしています。


# Unit Tests
aspvbslib-testライブラリを用いて、自分自身のユニットテストを行っています。

もしこのライブラリを修正した場合には、以下のテストを実行して後方互換性を確認して下さい。

- test_aspvbs-core.asp
- test_aspvbs-data.asp
- test_aspvbs-web.asp
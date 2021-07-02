<%
'--------------------------------------------------------------------
' ASPVBSLib Data
'--------------------------------------------------------------------
'==========================================================
' SqlQuery クラス
'==========================================================

''' <summary>
''' SQLクエリの結果を型付きオブジェクトのArrayListにマッピングする。
''' 　- 与えるSQLのSELECT句のフィールド名が変換後のクラスのフィールド名と一致している場合に自動的にO/Rマッピングを行う
'''	  - フィールドがNullの場合、Nullがマッピングされる。（IsNull関数でチェックできる）
''' 　- SELECT句を * にしていると、全フィールドのマッピングを試みて遅くなる為、明示的にSELECT句にフィールド名を指定した方が良い。
''' 　- 内部的にExecute関数を使っている為CPUパワーがないマシンでは遅い。速度が気になる場合には、Propプロパティを実装したPropsクラスを別途定義すると高速化する。
''' </summary>
''' <param name="DbConn">DbConnectionオブジェクト</param>
''' <param name="className">ORマッピング先のクラス名。SELECT句のフィールド名と一致するメンバを定義しておくと自動的に値が読み込まれる。</param>
''' <param name="cmdOrSql">SQLクエリ文字列 又は ADODB.Command オブジェクト</param>
''' <returns>SQLの結果がマッピングされた、classNameで指定したクラスのオブジェクトのArrayList。結果が0件の場合は空のArrayListを返す。</returns>	
Public Function SqlQuery(DbConn, className, cmdOrSql )
	Dim RS
	Set RS = Server.CreateObject("ADODB.Recordset")

    If VarType(cmdOrSql) = vbString Then
    	RS.Open cmdOrSql, DbConn, adOpenKeyset, adLockReadOnly    '静的スクロール、読取り専用
    Else
        cmdOrSql.ActiveConnection = DbConn
        Set RS = cmdOrSql.Execute
    End If

	Dim list
	Set list = New ArrayList

    'Propsクラスの生成（存在しない場合はNothing）
    Dim props
    Set props = Nothing
    Dim propsClassName
    propsClassName = className & "Props"
    On Error Resume Next
	Set props = Eval("New " & propsClassName)
    'Execute("Set props = New " & propsClassName)
    On Error Goto 0

	Do Until RS.EOF
		Dim item
		Set item = Eval("New " & className )

        'Propsクラスがある場合、オブジェクトと関連づける
        If Not props Is Nothing Then
            Set props.Target = item
        End If

		Dim field
		Dim fieldValue

		For Each field In RS.Fields

			If IsNull(RS(field.Name)) Then
				'DBの値がNullの時はNullのまま取得
				On Error Resume Next
				If Not props Is Nothing Then
					props.Prop(field.Name) = RS(field.Name)
				Else
					'CallByName item, field.Name, vbLet, RS(field.Name)
					Execute("item." & field.Name & " = RS(field.Name)") 'Nullを設定。IsNull(item.[FieldName])でNullチェックできる。
				End If
				On Error Goto 0
			Else
				'値が設定されている時
				On Error Resume Next
				Select Case field.Type
					Case adChar
						'RTrimして取得
						fieldValue = RTrim(RS(field.Name).Value)
						If Not props Is Nothing Then
							props.Prop(field.Name) = fieldValue
						Else
							'CallByName item, field.Name, vbLet, fieldValue '(RTrim(RS(field.name).Value))
							Execute("item." & field.Name & " = fieldValue")
						End If

					Case adDBTimeStamp
						'CDateして取得
						fieldValue = CDate(RS(field.Name).Value)
						If Not props Is Nothing Then
							props.Prop(field.Name) = fieldValue
						Else
							'CallByName item, field.Name, vbLet, CDate(RS(field.Name).Value)
							Execute("item." & field.Name & " = fieldValue")
						End If

					Case Else
						'そのまま取得
						fieldValue = RS(field.Name).Value
						If Not props Is Nothing Then
							props.Prop(field.Name) = fieldValue
						Else
							'CallByName item, field.Name, vbLet, RS(field.Name).Value
							Execute("item." & field.Name & " = fieldValue")
						End If

				End Select
				On Error Goto 0
			End If
		Next

		list.Add item
		RS.MoveNext
	Loop

	RS.Close
	Set RS = Nothing

	Set SqlQuery = list

End Function

'==========================================================
' ArrayList や SqlQuery と組み合わせて使用できる代表的なエンティティクラス
' これらの例を元に、ユーザー定義のエンティティクラスを定義する事ができる。
'==========================================================

''' <summary>
''' キーバリュークラス
''' キーと値の組を持つ共通クラス。
''' </summary>
Class KeyAndValue
	''' <summary>キー</summary>
	Public Key

	'値
	Public Value

    ''' <summary>
    ''' コンストラクタ風 値設定用デフォルトプロパティ。
    ''' 例) Set keyValue = (New KeyAndValue)(key, value)
    ''' </summary>
	Public Default Property Get Constructor(Key, Value)
        Me.Key = Key
		If IsObject(Value) Then
			Set Me.Value = Value
		Else
			Me.Value = Value
		End If
		Set Constructor = Me
	End Property

End Class

''' <summary>
''' キーバリュークラス用のプロパティ定義関数
''' 定義しておくとJSONValue関数やLoadRequest関数等が利用できる。
''' </summary>
Function KeyAndValuePropNames
	KeyAndValuePropNames = Array("Key", "Value")
End Function

''' <summary>
''' キーバリュークラス用のPropsクラス。
''' Propプロパティを持ち、GetとLetにて、プロパティ名と実際のプロパティを関連付ける。
''' 定義しておくとSqlQuery関数の速度が向上する。
''' </summary>
Class KeyAndValueProps
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
			Case "Key"
				value = obj.Key
			Case "Value"
				value = obj.Value
            '-- ここまで ---------------------
			Case Else
				value = Eval("obj." & propName)
		End Select
		Prop = value
	End Property

	Public Property Let Prop( propName, value )
		Select Case propName
            '-- ここからユーザー定義部分 ------
			Case "Key"
				obj.Key = value
			Case "Value"
				obj.Value = value
            '-- ここまで ---------------------
			Case Else
				Execute("obj." & propName & " = value")
		End Select
	End Property

End Class

'==========================================================
' DBヘルパー
'==========================================================

''' <summary>
''' JSONValue関数がDate型の値を整形する際の整形文字列。
''' </summary>
Dim DBSqlValue_DATE_FORMAT: DBSqlValue_DATE_FORMAT = "'yyyy-MM-dd HH:mm:ss'"

''' <summary>
''' 値の型に応じたSQL上のリテラル表現を返す
''' </summary>
''' <param name="val">値</param>
''' <returns>SQLリテラル文字列</returns>
Function DBSqlValue(value)

	If IsObject(value) Then
		Err.Raise 13	'型が一致しません
	End If

	dim v
	v = value

	Select Case VarType(value)
		Case vbString
			v = "'" & v & "'"
		
		Case vbEmpty, vbNull, vbError
			v = "Null"

		Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbByte, vbVariant
			v = CStr(v)

		Case vbDate
			v = VBDateFormat(v, DBSqlValue_DATE_FORMAT)

		Case vbBoolean
			If v Then
				v = "True"
			Else
				v = "False"
			End If
		
		Case Else
			Err.Raise 13	'型が一致しません

	End Select

	DBSqlValue = v
End Function

%>
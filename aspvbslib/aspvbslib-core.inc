<%
'--------------------------------------------------------------------
' ASPVBSLib Core
'--------------------------------------------------------------------

'==========================================================
' コーディングを補助する関数
'==========================================================

''' <summary>
''' 式の評価に基づいて、2つの部分のいずれかを返す。
''' VBのIIf関数 代替。VBのIIfと同様に、TruePartとFalsePartはどちらも必ず評価される事に注意。
''' </summary>
''' <param name="expr">評価する式</param>
''' <param name="TruePart">exprがTrueの時に返される値</param>
''' <param name="FalsePart">exprがFalseの時に返される値</param>
''' <returns></returns>
Function IIf(expr, TruePart, FalsePart)

	If expr Then
		IIf = TruePart
	Else
		IIf = FalsePart
	End If

End Function

'==========================================================
' 配列を扱う関数
'==========================================================

''' <summary>
''' 指定した値が配列の中にあればTrue、なければFalseを返す。
''' </summary>
''' <param name="value">検索したい値</param>
''' <param name="arr">配列</param>
''' <returns></returns>
Function InArray(value, arr)

	If IsNull(arr) Then
		IsArray = False
		Exit function
	End If

	Dim i
	InArray = False
	For i = LBound(arr) To UBound(arr)
		If arr(i) = value Then
			InArray = True
			Exit For
		End If
	Next

End Function

''' <summary>
''' 指定した値が配列の中にあれば、その要素のindexを返す。見つからなければ-1を返す。
''' </summary>
''' <param name="searchKey">検索したい値</param>
''' <param name="arr">配列</param>
''' <returns>配列中のindex</returns>
Function ArrayIndexOf(searchKey, arr)

	Dim i
	ArrayIndexOf = -1
	For i = LBound(arr) To UBound(arr)
		If arr(i) = searchKey Then
			ArrayIndexOf = i
			Exit For
		End If
	Next

End Function

''' <summary>
''' 指定した値が配列arrKeysの中にあれば、arrValuesの同じ位置の要素を返す。見つからなければ空文字を返す。
''' </summary>
''' <param name="searchKey">検索したい値</param>
''' <param name="arrKeys">キーの配列</param>
''' <param name="arrValues">値の配列</param>
''' <returns>見つかった値</returns>
Function FindKeyValue(searchKey, arrKeys, arrValues)

	Dim i
	FindKeyValue = ""
	For i = LBound(arrKeys) To UBound(arrKeys)
		If arrKeys(i) = searchKey Then
			FindKeyValue = arrValues(i)
			Exit For
		End If
	Next

End Function

'==========================================================
' 値のチェックや整形を行う関数
'==========================================================

''' <summary>
''' 値がNull又はEmpty又は空文字かどうかを示す。
''' </summary>
''' <param name="value">値として渡す文字列</param>
''' <returns></returns>
Function IsNullOrEmpty(value)
	Dim result : result = False

	If IsObject(value) Then
		If value Is Nothing Then
			result = True
		Else
			result = False
		End If
	ElseIf IsArray(value) Then
		result = False
	ElseIf IsEmpty(value) Then
		result = True
	ElseIf IsNull(value) Then
		result = True
	Else
		If value = "" Then
			result = True
		End If
	End If

	IsNullOrEmpty = result
End Function

''' <summary>
''' 値がNull又はEmpty又は空文字の時に規定値となる値を返す。それ以外の時は値そのものを返す。
''' </summary>
''' <param name="value">値</param>
''' <param name="nullval">値がNull又はEmpty又は空文字の時に返す値</param>
''' <returns></returns>
Function NullValue(value, nullval)
	If IsNullOrEmpty(value) Then
		If IsObject(nullval) Then
			Set NullValue = nullval
		Else
			NullValue = nullval
		End If
	Else
		If IsObject(value) Then
			Set NullValue = value
		Else
			NullValue = value
		End If
	End If
End Function

''' <summary>
''' 日付型の値をVB.NETのDateTime.ToStringのような書式指定で整形する。
'''  例) VBDateFormat( Now, "yyyy年MM月dd日 HH時mm分ss秒")
''' </summary>
''' <param name="dt">日付型の値</param>
''' <param name="formatString">書式指定文字列。例) "yyyy年MM月dd日 HH時mm分ss秒", "yyyy.M.d H:m:s"</param>
''' <returns></returns>
Function VBDateFormat(dt, formatString)

	Dim iYear, iMonth, iDay, iHour, iMinute, iSecond
	If IsDate(dt) Then
		iYear = Year(dt)
		iMonth = Month(dt)
		iDay = Day(dt)
		iHour = Hour(dt)
		iMinute = Minute(dt)
		iSecond = Second(dt)
	Else
		iYear = 0
		iMonth = 0
		iDay = 0
		iHour = 0
		iMinute = 0
		iSecond = 0
	End If

	Dim result
	result = formatString

	result = Replace(result, "yyyy", Right("0000" & iYear, 4))
	result = Replace(result, "yy", Right("00" & iYear, 2))
	result = Replace(result, "y", iYear)
	result = Replace(result, "MM", Right("00" & iMonth, 2))
	result = Replace(result, "M", iMonth)
	result = Replace(result, "dd", Right("00" & iDay, 2))
	result = Replace(result, "d", iDay)
	result = Replace(result, "HH", Right("00" & iHour, 2))
	result = Replace(result, "H", iHour)
	result = Replace(result, "mm", Right("00" & iMinute, 2))
	result = Replace(result, "m", iMinute)
	result = Replace(result, "ss", Right("00" & iSecond, 2))
	result = Replace(result, "s", iSecond)

	VBDateFormat = result

End Function

''' <summary>
''' プレースホルダー付き文字列整形。{0}、{1}、…のように指定すると、引数の同じ位置の値をその場所に埋め込む。
''' 例) VBStringFormat( "Cd={0}は{1}です", Array( "1234", "編集中" ) )
''' 例) VBStringFormat( "{0}が存在しません", "〇〇情報" )
''' </summary>
''' <param name="formatString">プレースホルダー付き文字列</param>
''' <param name="args">引数が1つの場合は単一で、複数の場合は配列で渡す。</param>
''' <returns></returns>
Function VBStringFormat(formatString, args)
	Dim result, i
	result = formatString

	If IsArray(args) Then
		For i = 0 To UBound(args)
			result = Replace(result, "{" & i & "}", args(i))
		Next
	Else
		i = 0
		result = Replace(result, "{0}", args)
	End If

	VBStringFormat = result

End Function

''' <summary>
''' 文字列を指定した長さまで左側にゼロ埋めする。例）VBPadLeftZero("12", 5) => "00012"
''' </summary>
''' <param name="value">文字列</param>
''' <param name="length">ゼロ埋めする長さ。</param>
''' <returns>lengthよりvalueの方が長い場合、valueがそのまま返される。</returns>
Function VBPadLeftZero(value, length)
	Dim valueLen
	valueLen = Len(Trim(value))
	
	If valueLen < length Then
		VBPadLeftZero = String(length - valueLen, "0") & Trim(value)
	Else
		VBPadLeftZero = Trim(value)
	End If
End Function

' ''★★★！！！　64bit環境ではScriptControlがない為実行できない ！！！★★★
' ''' <summary>
' ''' プレースホルダー付き文字列整形。{0}、{1}、…のように指定すると、引数の同じ位置の値をその場所に埋め込む。
' ''' 例) VBDateFormat( "Cd={0}は{1}です", "1234", "編集中" )
' ''' 例) VBDateFormat( "{0}が存在しません", "〇〇情報" )
' ''' </summary>
' ''' <param name="formatString">プレースホルダー付き文字列</param>
' ''' <param name="args">引数。複数指定可能。</param>
' ''' <returns></returns>
' Dim objScriptControl
' Set objScriptControl = Server.CreateObject("MSScriptControl.ScriptControl")
' Function CreateJSFunction(funcName, funcString)
' 	objScriptControl.Language = "JavaScript"
' 	objScriptControl.AddCode funcName & " = " & funcString & ";"
' 	Set CreateJSFunction = objScriptControl.Eval( funcName )
' End Function

' Dim VBStringFormat2
' Set VBStringFormat2 = CreateJSFunction( "VBStringFormat2", _
' 	"function() {" & _
'  	"	var result = arguments[0];" & _
' 	"	for ( var i = 1; i < arguments.length; i++ ) {" & _
' 	"		result = result.replace( ('{' + (i - 1) + '}'), arguments[i] );" & _
' 	"	};" & _
' 	"	return result;" & _
' 	"}" ) 

''' <summary>
''' 型を指定しての型変換
''' </summary>
''' <param name="v">入力値</param>
''' <param name="vType">変換する型を表す定数</param>
Function ConvertTo(v, vType)
	On Error Resume Next

    Select Case vType
		Case vbBoolean
			ConvertTo = CBool(v)

        Case vbString
            ConvertTo = CStr(v)
		
		Case vbBoolean
			ConvertTo = CByte(v)

		Case vbCurrency
			ConvertTo = CCur(v)
		
		Case vbDate
			ConvertTo = CDate(v)
		
		Case vbDouble
			ConvertTo = CDbl(v)
		
		Case vbInteger
			ConvertTo = CInt(v)
		
		Case vbLong
			ConvertTo = CLng(v)
		
		Case vbSingle
			ConvertTo = CSng(v)

        Case vbObject
			ConvertTo = CStr(v)

        Case vbArray + vbInteger, vbArray + vbLong, vbArray + vbSingle, vbArray + vbDouble, vbArray + vbCurrency, vbArray + vbString, vbArray + vbObject, vbArray + vbVariant
			Dim list: Set list = New ArrayList
			list.ItemType = vType - vbArray
			Dim item
			For Each item In Split(v, ",")
				list.Add ConvertTo(Trim(item), vType - vbArray)
			Next
			ConvertTo = list.Items

        Case Else
            ConvertTo = CStr(v)

    End Select

	If Err.Number <> 0 Then
		ConvertTo = DefaultValue(vType)
	End If

	On Error Goto 0

End Function

''' <summary>
''' 型の規定値を取得する
''' </summary>
''' <param name="vType">型を表す定数</param>
''' <returns>型の規定値</returns>
Function DefaultValue(vType)

    Select Case vType
		Case vbBoolean
			DefaultValue = False

        Case vbString
			DefaultValue = ""
		
		Case vbBoolean
			DefaultValue = CByte(0)

		Case vbCurrency
			DefaultValue = CCur(0)
		
		Case vbDate
			DefaultValue = CDate(0)
		
		Case vbDouble
			DefaultValue = CDbl(0)
		
		Case vbInteger
			DefaultValue = CInt(0)
		
		Case vbLong
			DefaultValue = CLng(0)
		
		Case vbSingle
			DefaultValue = CSng(0)

        Case vbObject
			DefaultValue = ""

        Case vbArray + vbInteger, vbArray + vbLong, vbArray + vbSingle, vbArray + vbDouble, vbArray + vbCurrency, vbArray + vbString, vbArray + vbObject, vbArray + vbVariant
			DefaultValue = Array()

        Case Else
            DefaultValue = ""

    End Select

End Function


'==========================================================
' 正規表現を扱う関数
'==========================================================

''' <summary>
''' 文字列の中に、パターンにマッチする部分があるかどうかを返す。
''' </summary>
''' <param name="input">テストしたい文字列</param>
''' <param name="pattern">正規表現パターン文字列</param>
''' <returns></returns>
Function RegexIsMatch(input, pattern)
	Dim retVal
	retVal = True

	Dim re, FoundMatch
	Set re = New RegExp
	re.Pattern = pattern
	FoundMatch = re.Test(input)

	RegexIsMatch = FoundMatch

End Function

''' <summary>
''' 文字列に数字以外が含まれていないかどうかを返す。
''' </summary>
''' <param name="test">テストしたい文字列</param>
''' <returns>True:含まれていない False:含まれている</returns>
Function ChkNum(test)
	ChkNum = Not RegexIsMatch(test, "[^0-9]")
End Function

''' <summary>
''' 文字列が数字か英文字以外が含まれていないかどうかを返す。
''' </summary>
''' <param name="test">テストしたい文字列</param>
''' <returns>True:含まれていない False:含まれている</returns>
Function ChkNumAlphabet(test)
	ChkNumAlphabet = Not RegexIsMatch(test, "[^a-zA-Z0-9]")
End Function


'==========================================================
' ArrayList クラス
'==========================================================
''' <summary>
''' VB.NETのArrayListに似た使い方ができるクラス(内部でDictionaryを使用）
''' Linq的に使える Where, Map(Selectに相当), FirstOrDefault, OrderBy, Reverse を実装。
''' </summary>
Class ArrayList
	Private pItems
	Private pKeyMax
    Private pItemType
    Private pItemClassName

	Private Sub Class_Initialize()
        Set pItems = Server.CreateObject("Scripting.Dictionary")
        pKeyMax = 0
	End Sub

	Private Sub Class_Terminate()
        Set pItems = Nothing
    End Sub

    ''' <summary>
    ''' リストの件数。
    ''' </summary>
    ''' <returns></returns>
	Public Property Get Count
        Count = pItems.Count
    End Property

    ''' <summary>
    ''' Indexで要素を取得
    ''' </summary>
    ''' <returns></returns>
	Public Default Property Get Item(index)
		Dim arr
		arr = pItems.Items
        If IsObject(arr(index)) Then
			Set Item = arr(index)
		Else
			Item = arr(index)
		End If
	End Property

    ''' <summary>
    ''' 要素のリストを配列で取得
    ''' </summary>
    ''' <returns></returns>
	Public Property Get Items
        Items = pItems.Items
    End Property

    ''' <summary>
    ''' 要素の型を設定
    ''' </summary>
    ''' <returns></returns>
    Public Property Let ItemType(aType)
        pItemType = aType
    End Property

    ''' <summary>
    ''' 要素の型を取得
    ''' </summary>
    ''' <returns></returns>
    Public Property Get ItemType
        ItemType = pItemType
    End Property

    ''' <summary>
    ''' 要素がオブジェクト型の場合、クラス名を設定
    ''' </summary>
    ''' <returns></returns>
    Public Property Let ItemClassName(aClassName)
        pItemClassName = aClassName
		If Len(aClassName) > 0 Then
	        pItemType = vbObject
		End If
    End Property

    ''' <summary>
    ''' 要素がオブジェクト型の場合、クラス名を取得
    ''' </summary>
    ''' <returns></returns>
    Public Property Get ItemClassName
        ItemClassName = pItemClassName
    End Property

    ''' <summary>
    ''' 引数の型が要素の型と一致しているか
    ''' </summary>
    ''' <returns></returns>
    Private Function CheckType(item)
        Dim aType
        aType = VarType(item)
        If Len(pItemType) = 0 Then
            pItemType = aType
        ElseIf pItemType <> aType Then
            CheckType = False
            Exit Function
        End If

        If aType = vbObject Then
            Dim aClassName
            aClassName = TypeName(item)
            If Len(pItemClassName) = 0 Then
                pItemClassName = aClassName
            ElseIf pItemClassName <> aClassName Then
                CheckType = False
                Exit Function
            End If
        End If

        CheckType = True
    End Function

    ''' <summary>
    ''' リストの最初の要素を取得（Items(0)に相当するが、リストが空の場合には指定した規定値を返す。）
    ''' </summary>
	''' <param name="defaultValue">リストが空の場合の規定値</param>
    ''' <returns></returns>
	Public Function FirstOrDefault(defaultValue)
		If pItems.Count = 0 Then
			If IsObject(defaultValue) Then
				Set FirstOrDefault = defaultValue
			Else
				FirstOrDefault = defaultValue
			End If
		Else
			If IsObject(pItems(0)) Then
				Set FirstOrDefault = pItems(0)
			Else
				FirstOrDefault = pItems(0)
			End If
		End If
	End Function

    ''' <summary>
    ''' リストの最初の要素を取得（リストが空の場合には指定した規定値を返す。）
    ''' </summary>
	''' <param name="defaultValue">リストが空の場合の規定値</param>
    ''' <returns></returns>
	Public Function LastOrDefault(defaultValue)
		If pItems.Count = 0 Then
			If IsObject(defaultValue) Then
				Set LastOrDefault = defaultValue
			Else
				LastOrDefault = defaultValue
			End If
		Else
			If IsObject(pItems(pItems.Count - 1)) Then
				Set LastOrDefault = pItems(pItems.Count - 1)
			Else
				LastOrDefault = pItems(pItems.Count - 1)
			End If
		End If
	End Function	

    ''' <summary>
    ''' 要素を末尾に追加
    ''' </summary>
	Public Sub Add(item)
        If Not CheckType(item) Then
            Err.Raise vbObjectError + 1,,"type of item mismatch. adding item type = " & VarType(item) & ", list item type = " & pItemType
        End If
		pItems.Add pKeyMax, item
		pKeyMax = pKeyMax + 1
	End Sub

    ''' <summary>
    ''' リストを末尾に追加
    ''' </summary>
	Public Sub AddList(list)
		Dim item
		For Each item In list.Items
			Call Me.Add(item)
		Next
	End Sub

    ''' <summary>
    ''' リストをクリアする
    ''' </summary>
	Public Sub Clear()
		pItems.RemoveAll
		pKeyMax = 0
	End Sub

    ''' <summary>
    ''' 指定した要素をリストから削除する
    ''' </summary>
	''' <param name="item">削除する要素</param>
	Public Sub Remove(item)
		Dim key
		key = FindItem(item)
		If key <> -1 Then
			pItems.Remove(key)
		Else
			Err.Raise 9 'インデックスが有効範囲にありません
		End If
	End Sub

    ''' <summary>
    ''' 指定した要素をリストから削除する
    ''' </summary>
	''' <param name="index">削除する要素の位置(0～)</param>
	Public Sub RemoveAt(index)
		If index < 0 Or index >= Me.Count Then
			Err.Raise 9 'インデックスが有効範囲にありません
		Else
			Dim arr
			arr = pItems.Items
			Call Me.Remove(arr(index))
		End If
	End Sub

    ''' <summary>
    ''' 指定した要素がリストに含まれているか
    ''' </summary>
	''' <param name="item">含まれているか確認したい要素</param>
	''' <returns>True: 含まれている, False: 含まれていない</returns>
	Public Function Contains(item)
		Contains = (IndexOf(item) <> -1)
	End Function

    ''' <summary>
    ''' 指定した要素をリストから検索し、0から始まる位置を返す。存在しない場合は-1を返す。
    ''' </summary>
	''' <param name="searchItem">要素</param>
	''' <returns>要素の位置</returns>
	Public Function IndexOf(searchItem)

		Dim index
		Dim arr
		arr = pItems.Items
		For index = 0 To UBound(arr)
			If IsObject(arr(index)) Then
				If arr(index) Is searchItem Then
					IndexOf = index
					Exit Function
				End If
			Else
				If arr(index) = searchItem Then
					IndexOf = index
					Exit Function
				End If
			End If
		Next

		IndexOf = -1

	End Function

    ''' <summary>
    ''' 指定した要素をリストから検索し、内部コレクションのkeyを返す
    ''' </summary>
	''' <param name="searchItem">含まれているか確認したい要素</param>
	''' <returns>内部コレクションのkey</returns>
	Private Function FindItem(searchItem)
		Dim key
		For Each key In pItems.Keys
			If IsObject(pItems(key)) Then
				If pItems(key) Is searchItem Then
					FindItem = key
					Exit Function
				End If
			Else
				If pItems(key) = searchItem Then
					FindItem = key
					Exit Function
				End If
			End If
		Next

		FindItem = -1
	End Function


    ''' <summary>
	''' リストをDictionaryに変換
	''' 例) ToDictionary( "item.ID", "item.Name" )
    ''' </summary>
    ''' <returns></returns>
	Public Function ToDictionary(keyExpr, valueExpr)
		Dim dic
		Set dic = Server.CreateObject("Scripting.Dictionary")

        Dim item
		For Each item In pItems.Items
			Dim key, value
			key = Eval(keyExpr)
			value = Eval(valueExpr)
			dic.Add key, value
		Next

		Set ToDictionary = dic

	End Function

    ''' <summary>
	''' リストの絞り込み。VB.NETのEnumerable(Of T).Whereに相当
	''' 例) list.Where( "item.Age < p", 20 )
	''' 例) list.Where( "item.SectionCd = p.SectionCd", targetSection )
    ''' </summary>
    ''' <param name="expr">各要素が選択対象となるかを判断する為の、真偽値を返す条件式文字列。"item"は各要素、"p"は第二引数を表す。</param>
    ''' <param name="p">パラメータとして与える引数。expr中で"p"として参照できる。</param>
    ''' <returns>条件に一致するArrayListを返す。条件に一致する結果がなかった場合、空のArrayListを返す。</returns>
	Public Function Where(expr, p)
		Dim list
        Set list = New ArrayList
		list.ItemType = pItemType
		list.ItemClassName = pItemClassName
        Dim item
		For Each item In pItems.Items
			If Eval(expr) Then
				list.Add item
			End If
		Next
        Set Where = list
    End Function

    ''' <summary>
	''' リストの絞り込み。
	''' 条件に一致する最初の要素を抜き出す。戻り値は空のリストか1件のみのリスト。
	''' 例) list.FindFirst( "item.Age < p", 20 )
	''' 例) list.FindFirst( "item.SectionCd = p.SectionCd", targetSection )
    ''' </summary>
    ''' <param name="expr">各要素が選択対象となるかを判断する為の、真偽値を返す条件式文字列。"item"は各要素、"p"は第二引数を表す。</param>
    ''' <param name="p">パラメータとして与える引数。expr中で"p"として参照できる。</param>
    ''' <returns>条件に一致する最初の値を1つ含むArrayListを返す。条件に一致する結果がなかった場合、空のArrayListを返す。</returns>
	Public Function FindFirst(expr, p)
		Dim list
        Set list = New ArrayList
		list.ItemType = Me.ItemType
		list.ItemClassName = Me.ItemClassName
        Dim item
		For Each item In pItems.Items
			If Eval(expr) Then
				list.Add item
                Exit For
			End If
		Next
        
        Set FindFirst = list

	End function

    ''' <summary>
	''' リスト中の要素を別の形式の要素に置き換える（射影）。VB.NETのEnumerable(Of T).Selectに相当。
	''' 例) Map( "item.ID & "":"" & item.Name" )
    ''' </summary>
    ''' <param name="expr">新しい要素となる値を返す式。"item"は各要素を表す。</param>
    ''' <returns></returns>
	Public Function Map(expr, p)
		Dim list
        Set list = New ArrayList
        Dim item
		For Each item In pItems.Items
			list.Add Eval(expr)
		Next
        Set Map = list
    End Function

    ''' <summary>
	''' Mapの別名。
	''' 例) Select( "item.ID & "":"" & item.Name" )
    ''' </summary>
    ''' <param name="expr">新しい要素となる値を返す式。"item"は各要素を表す。</param>
    ''' <returns></returns>
	Public Function [Select](expr, p)
		Set [Select] = Map(expr, p)
	End Function

    ''' <summary>
	''' ソート
	''' 例) OrderBy("item.Name", True) →Nameプロパティでソート
	''' 例) OrderBy("item", True)		→要素自体でソート
    ''' </summary>
    ''' <param name="expr">比較対象とする式。"item"は各要素を表す。</param>
    ''' <param name="ascending">True: 昇順、False:降順</param>
    ''' <returns></returns>
	Public Function OrderBy(expr, ascending)
		Dim itms, i, tmpDic
		itms = pItems.Items
		For i = 0 To UBound(itms) - 1
			Dim j
			For j = i + 1 To UBound(itms)
				Dim itemj, itemi
				itemj = ItemToScalar(itms(j), expr)
				itemi = ItemToScalar(itms(i), expr)
				If CompareTo(itemj, itemi) < 0 Then
					Call swap(itms(i), itms(j))
				End If
			Next
		Next

		Dim list
		Set list = New ArrayList
		list.ItemType = Me.ItemType
		list.ItemClassName = Me.ItemClassName
		For i = 0 To UBound(itms)
			list.Add itms(i)
		Next

        If Not ascending Then
            list.Reverse()
        End If

		Set OrderBy = list
	End Function

    ''' <summary>
	''' ソート（昇順）
	''' 例) OrderByAsc("item.Name") →Nameプロパティでソート
	''' 例) OrderByAsc("item")		→要素自体でソート
    ''' </summary>
    ''' <param name="expr">比較対象とする式。"item"は各要素を表す。</param>
    ''' <returns></returns>
	Public Function OrderByAsc(expr)
        Set OrderByAsc = Me.OrderBy(expr, True)
    End Function

    ''' <summary>
	''' ソート（降順）
	''' 例) OrderByAsc("item.Name") →Nameプロパティでソート
	''' 例) OrderByAsc("item")		→要素自体でソート
    ''' </summary>
    ''' <param name="expr">比較対象とする式。"item"は各要素を表す。</param>
    ''' <returns></returns>
	Public Function OrderByDesc(expr)
        Set OrderByDesc = Me.OrderBy(expr, False)
    End Function

    ''' <summary>
	''' リストの順序を反転
    ''' </summary>
    ''' <returns></returns>
	Public Sub Reverse()
		Dim list
		Set list = New ArrayList
		list.ItemType = Me.ItemType
		list.ItemClassName = Me.ItemClassName

		If Me.Count > 0 Then
			Dim i
			For i = UBound(pItems.Items) To 0 Step -1
				list.Add pItems.Item(i)
			Next
		End If

		Me.Clear
		Me.AddList(list)

	End Sub

	Public Property Get [_NewEnum]
		Set [_NewEnum] = pItem.[_NewEnum]
	End Property

    ''' <summary>
	''' オブジェクトを指定した式（文字列）でスカラー値に変換する
	''' 例) ItemToScalar( item, "item.ID & "":"" & item.Name" )
    ''' </summary>
    ''' <param name="item">オブジェクト値。</param>
    ''' <param name="expr">比較対象とする式。"item"は各要素を表す。</param>
    ''' <returns></returns>
	Private Function ItemToScalar(item, expr)
		ItemToScalar = Eval(expr)
	End Function

    ''' <summary>
	''' 値の比較
    ''' </summary>
    ''' <param name="x">比較する値x</param>
    ''' <param name="y">比較する値y</param>
    ''' <returns>x = y の時 0、x > y の時 1、x < y の時 -1 を返す。</returns>	
    Private Function CompareTo(ByRef x, ByRef y)
		Dim xType, yType, xVal, yVal
		xType = TypeName(x)
		yType = TypeName(y)
		If xType = yType Then
			If xType = "Boolean" Then
				If x = y Then
					CompareTo = 0
				ElseIf x = False And y = True Then
					CompareTo = -1
				Else
					CompareTo = 1
				End If

				Exit Function
			End If

            xVal = x
            yVal = y

			If xVal = yVal Then
				CompareTo = 0
			ElseIf xVal < yVal Then
				CompareTo = -1
			Else
				CompareTo = 1
			End If
		Else
			' 型の違うオブジェクトは型名でソート
			If xType = yType Then
				CompareTo = 0
			ElseIf xType < yType Then
				CompareTo = -1
			Else
				CompareTo = 1
			End If
		End If
	End Function

    ''' <summary>
	''' 要素の入れ替え
    ''' </summary>
    ''' <param name="x">入れ替えるする値x</param>
    ''' <param name="y">入れ替える値y</param>
	Private Sub swap(ByRef x, ByRef y)
		Dim d
		SetVariable d, x
		SetVariable x, y
		SetVariable y, d
	End Sub

    ''' <summary>
	''' 値の設定
    ''' </summary>
    ''' <param name="TargetVal">設定先の変数</param>
    ''' <param name="SrcVal">設定する値</param>
	Private Sub SetVariable(ByRef TargetVal, ByRef SrcVal)
		If IsObject(SrcVal) Then
			Set TargetVal = SrcVal
		Else
			TargetVal = SrcVal
		End If
	End Sub

End Class

'==========================================================
' リフレクション・動的プログラミング補助機能
'==========================================================

''' <summary>
''' 指定したオブジェクトのプロパティ名一覧を配列で取得する
''' ※クラス名+PropNamesという名前のプロパティ名取得用関数の定義が必要。
''' ※プロパティ名取得用関数の定義がない場合、例外が発生する。
''' </summary>
''' <param name="obj">ユーザー定義オブジェクト</param>
''' <returns>プロパティ名の配列</returns>
Function GetPropNames(obj)
	Dim className: className = TypeName(obj)
	Dim propNamesFuncName: propNamesFuncName = className & "PropNames"

	Dim propNames
	On Error Resume Next
	propNames = Eval(propNamesFuncName)
	On Error Goto 0

    If IsEmpty(propNames) Then
        Err.Raise vbObjectError + 513, className & "PropNames関数の定義がないか、定義が完全ではありません"
    End If

	GetPropNames = propNames

End Function

''' <summary>
''' 指定したオブジェクトの各プロパティを、対象となるオブジェクトにコピーする。
''' コピー元とコピー先のオブジェクトは同じプロパティがあれば良く、必ずしも同じクラスである必要はない。
''' また、コピー先に存在しないプロパティは無視される。
''' プロパティがオブジェクト型の場合、参照がコピーされる（浅いコピー）。
''' ※コピー元のクラス名+PropNamesという名前のプロパティ名取得用関数の定義が必要。
''' ※プロパティ名取得用関数の定義がない場合、例外が発生する。
''' </summary>
''' <param name="source">コピー元のオブジェクト</param>
''' <param name="target">コピー先のオブジェクト</param>
Sub CopyProps(ByRef source, ByRef target)

	Dim propNames: propNames = GetPropNames(source)

	Dim propName
	On Error Resume Next
	For Each propName In propNames
		If Eval("VarType(source." & propName & ")") = vbObject Then
			Execute("Set target." & propName & " = source." & propName)
		Else
			Execute("target." & propName & " = source." & propName)
		End If
	Next
	On Error Goto 0

End Sub

''' <summary>
''' 指定したオブジェクトを元に新しいオブジェクトを複製して返す。
''' 指定したオブジェクトの各プロパティを、複製したオブジェクトにコピーする。
''' プロパティがオブジェクト型の場合、参照がコピーされる（浅いコピー）。
''' ※複製元のクラス名+PropNamesという名前のプロパティ名取得用関数の定義が必要。
''' ※プロパティ名取得用関数の定義がない場合、例外が発生する。
''' </summary>
''' <param name="obj">コピー元のオブジェクト</param>
''' <returns>複製したオブジェクト</returns>
Function CloneObject(obj)
	Dim className: className = TypeName(obj)
	Dim target: Set target = Eval("New " & className)

	CopyProps obj, target
	Set CloneObject = target

End Function
%>
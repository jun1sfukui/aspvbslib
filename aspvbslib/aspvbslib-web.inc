<%

'--------------------------------------------------------------------
' ASPVBSLib Web
'--------------------------------------------------------------------

'==========================================================
' 入出力補助
'==========================================================

''' <summary>
''' オブジェクトの各プロパティ名と一致するIDの値をRequest.Form(つまりPOST）から取得し、プロパティに設定する。
''' ※xxxxPropNames関数の設定が必要。
''' </summary>
''' <param name="req">リクエストを受け取るオブジェクト</param>
Sub LoadForm(ByRef req)
    Dim propNames: propNames = GetPropNames(req)

    Dim propName
    Dim propClassName
    For Each propName In propNames
        If Request.Form(propName).Count > 0 Then
            Dim propType
            propClassName = Eval("TypeName(req." & propName & ")")
            If propClassName = "ArrayList" Then
                propType = Eval("req." & propName & ".ItemType")
                Dim value
                Dim list: Set list = New ArrayList
                For Each value In Request.Form(propName)
                    list.Add ConvertTo(value, propType)
                Next
                Execute("Set req." & propName & " = list")
            Else
                propType = Eval("VarType(req." & propName & ")")
                Execute("req." & propName & " = ConvertTo(Request.Form(propName), propType)")
            End If
        End If
    Next
End Sub

''' <summary>
''' オブジェクトの各プロパティ名と一致するIDの値をRequest.QueryString（つまりGET）から取得し、プロパティに設定する。
''' ※xxxxPropNames関数の設定が必要。
''' </summary>
''' <param name="req">リクエストを受け取るオブジェクト</param>
Sub LoadQueryString(ByRef req)
    Dim propNames: propNames = GetPropNames(req)

    Dim propName
    Dim propClassName
    For Each propName In propNames
        If Request.QueryString(propName).Count > 0 Then
            Dim propType
            propClassName = Eval("TypeName(req." & propName & ")")
            If propClassName = "ArrayList" Then
                propType = Eval("req." & propName & ".ItemType")
                Dim value
                Dim list: Set list = New ArrayList
                For Each value In Request.QueryString(propName)
                    list.Add ConvertTo(value, propType)
                Next
                Execute("Set req." & propName & " = list")
            Else
                propType = Eval("VarType(req." & propName & ")")
                Execute("req." & propName & " = ConvertTo(Request.QueryString(propName), propType)")
            End If
        End If
    Next
End Sub

''' <summary>
''' オブジェクトの各プロパティ名と一致するIDの値をRequestから取得し、プロパティに設定する。
''' ※xxxxPropNames関数の設定が必要。
''' </summary>
''' <param name="req">リクエストを受け取るオブジェクト</param>
Sub LoadRequest(ByRef req)
    Dim propNames: propNames = GetPropNames(req)

    Dim propName
    Dim propClassName
    For Each propName In propNames
        If Request(propName).Count > 0 Then
            Dim propType
            propClassName = Eval("TypeName(req." & propName & ")")
            If propClassName = "ArrayList" Then
                propType = Eval("req." & propName & ".ItemType")
                Dim value
                Dim list: Set list = New ArrayList
                For Each value In Request(propName)
                    list.Add ConvertTo(value, propType)
                Next
                Execute("Set req." & propName & " = list")
            Else
                propType = Eval("VarType(req." & propName & ")")
                Execute("req." & propName & " = ConvertTo(Request(propName), propType)")
            End If
        End If
    Next
End Sub

''' <summary>
''' JSONValue関数がDate型の値を整形する際の整形文字列。
''' </summary>
Dim JSONValue_DATE_FORMAT: JSONValue_DATE_FORMAT = """yyyy-MM-dd HH:mm:ss"""

''' <summary>
''' Object→JSON変換
'''   - Class定義したオブジェクトをJSON文字列に変換する。
'''   - 渡す値はオブジェクト値、配列、ArrayListが可能。
'''   - 渡すオブジェクトのクラス型名+"PropNames"という名前の関数を事前に定義する必要がある。関数の戻り値は、そのクラスのプロパティ名の配列。
'''       例) Function MyClassPropNames()
'''              Return Array("prop01", "prop02", "prop03")
'''           End Function
'''   - 内部的にExecute関数を使っている為CPUパワーがないマシンでは遅い。速度が気になる場合には、Propプロパティを実装したPropsクラスを別途定義すると高速化する。
''' </summary>
''' <param name="data">オブジェクト型の値、又はオブジェクト型の値のArrayList</param>
''' <param name="fieldNames">dataの持つプロパティ名の配列。</param>
''' <returns>JSON文字列</returns>
Function JSONValue(v)

    JSONValue = "null"

    Select Case VarType(v)
        Case vbString
            JSONValue = """" & jsEncode(v) & """"

        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbByte, vbVariant
            JSONValue = CStr(v)

        Case vbDate
            JSONValue = VBDateFormat(v, JSONValue_DATE_FORMAT)

        Case vbBoolean
            If v Then
                JSONValue = "true"
            Else
                JSONValue = "false"
            End If

        Case vbObject
            If v Is Nothing Then
                JSONValue = "null"
                Exit Function
            End If

            Dim className: className = TypeName(v)

            If className = "ArrayList" Then
                JSONValue = JSONValue(v.Items)
            Else
                Dim propNames: propNames = GetPropNames(v)

                'Propsクラスの生成（存在しない場合はNothing）
                Dim props
                Set props = Nothing
                Dim propsClassName
                propsClassName = className & "Props"
                On Error Resume Next
                Execute("Set props = New " & propsClassName)
                On Error Goto 0

                'Propsクラスがある場合、オブジェクトと関連づける
                If Not props Is Nothing Then
                    Call props(v)
                End If

                Dim jsonPairs: Set jsonPairs = New ArrayList
                Dim propName
                For Each propName In propNames
                    If Not props Is Nothing Then
                        jsonPairs.Add """" & propName & """:" & JSONValue(props.Prop(propName))
                    Else
                        jsonPairs.Add """" & propName & """:" & JSONValue(Eval("v." & propName))
                    End If
                Next
                JSONValue = "{" & Join(jsonPairs.Items, ", ") & "}"
            End If

        Case vbArray + vbInteger, vbArray + vbLong, vbArray + vbSingle, vbArray + vbDouble, vbArray + vbCurrency, vbArray + vbString, vbArray + vbObject, vbArray + vbVariant
            Dim item
            Dim arrList: Set arrList = New ArrayList
            For Each item In v
                arrList.Add JSONValue(item)
            Next
            JSONValue = "[" & Join(arrList.Items, ", ") & "]"

        Case vbEmpty, vbNull, vbError, vbDataObject
            JSONValue = "null"

        Case Else
            JSONValue = "null"

    End Select

End Function

''' <summary>
''' 文字列をJSON用にエンコードする。
''' This function is from "aspjson" by tugrul in GitHub
''' </summary>
''' <param name="str">文字列</param>
''' <returns>JSONエンコードされた文字列</returns>
Function jsEncode(str)
	Dim charmap(127), haystack()
	charmap(8) = "\b"
	charmap(9) = "\t"
	charmap(10) = "\n"
	charmap(12) = "\f"
	charmap(13) = "\r"
	charmap(34) = "\"""
	charmap(47) = "\/"
	charmap(92) = "\\"

	Dim strlen : strlen = Len(str) - 1
	ReDim haystack(strlen)

	Dim i, charcode
	For i = 0 To strlen
		haystack(i) = Mid(str, i + 1, 1)

		charcode = AscW(haystack(i)) And 65535
		If charcode < 127 Then
			If Not IsEmpty(charmap(charcode)) Then
				haystack(i) = charmap(charcode)
			ElseIf charcode < 32 Then
				haystack(i) = "\u" & Right("000" & Hex(charcode), 4)
			End If
		Else
			haystack(i) = "\u" & Right("000" & Hex(charcode), 4)
		End If
	Next

	jsEncode = Join(haystack, "")
End Function

'==========================================================
' HTMLテンプレート記述補助
'==========================================================

''' <summary>
''' display:noneを出力するかどうか。falseの場合のみ、display:noneを出力する。IIf関数よりシンプルに書ける。
''' </summary>
''' <param name="expr">noneを出力するかどうか決める条件式</param>
''' <returns>Trueの時、空文字。Falseの時、display:none;</returns>
Function css_display(expr)

	If expr Then
		css_display = ""
	Else
		css_display = "display:none;"
	End If

End Function

''' <summary>
''' visibilityの切替。IIf関数よりシンプルに書ける。
''' </summary>
''' <param name="expr">visibilityをvisibleにするかどうかを決める条件式</param>
''' <returns>exprがTrueの時、visibility:visible:、Falseの時、visibility:hidden;</returns>
Function css_visibility(expr)

	If expr Then
		css_visibility = "visibility:visible;"
	Else
		css_visibility = "visibility:hidden;"
	End If

End Function

''' <summary>
''' 指定文字列をを出力するかどうか。IIf関数よりシンプルに書ける。
''' </summary>
''' <param name="str">出力する文字列</param>
''' <param name="expr">strを出力する条件式</param>
''' <returns>exprがTrueの時、strを返す。Falseの時、空文字を返す。</returns>
Function write_if(str, expr)

	If expr Then
		write_if = str
	Else
		write_if = ""
	End If

End Function

''' <summary>
''' optionタグを以下の形式でリスト出力した文字列を返す。
''' <option value="arrKeys(i)">arrCaptions(i)</option>
''' arrKeys(i)がselectedKeyと一致した場合、selected属性を付与する。
''' </summary>
''' <param name="arrKeys">optionタグのvalueの配列</param>
''' <param name="arrCaptions">optionタグのラベル</param>
''' <param name="selectKey">selectedにするkey</param>
''' <returns></returns>
Function html_options(arrKeys, arrCaptions, selectedKey)
	Dim html
	html = ""

	Dim i
	For i = LBound(arrKeys) To UBound(arrKeys)
		If IsNullOrEmpty(arrKeys(i)) Then
			Exit For    '配列のキーが空文字ならばそこから先は出力しない（空の配列キーを配列の終端として扱っている為）
		End If
		html = html & "<option value=""" & arrKeys(i) & """" & write_if(" selected", arrKeys(i) = selectedKey) & ">"
		html = html & arrCaptions(i) & "</option>"
	Next

	html_options = html

End Function

%>
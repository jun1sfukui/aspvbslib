<%

'--------------------------------------------------------------------
' ASPVBSLib Web
'--------------------------------------------------------------------

'==========================================================
' ���o�͕⏕
'==========================================================

''' <summary>
''' �I�u�W�F�N�g�̊e�v���p�e�B���ƈ�v����ID�̒l��Request.Form(�܂�POST�j����擾���A�v���p�e�B�ɐݒ肷��B
''' ��xxxxPropNames�֐��̐ݒ肪�K�v�B
''' </summary>
''' <param name="req">���N�G�X�g���󂯎��I�u�W�F�N�g</param>
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
''' �I�u�W�F�N�g�̊e�v���p�e�B���ƈ�v����ID�̒l��Request.QueryString�i�܂�GET�j����擾���A�v���p�e�B�ɐݒ肷��B
''' ��xxxxPropNames�֐��̐ݒ肪�K�v�B
''' </summary>
''' <param name="req">���N�G�X�g���󂯎��I�u�W�F�N�g</param>
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
''' �I�u�W�F�N�g�̊e�v���p�e�B���ƈ�v����ID�̒l��Request����擾���A�v���p�e�B�ɐݒ肷��B
''' ��xxxxPropNames�֐��̐ݒ肪�K�v�B
''' </summary>
''' <param name="req">���N�G�X�g���󂯎��I�u�W�F�N�g</param>
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
''' JSONValue�֐���Date�^�̒l�𐮌`����ۂ̐��`������B
''' </summary>
Dim JSONValue_DATE_FORMAT: JSONValue_DATE_FORMAT = """yyyy-MM-dd HH:mm:ss"""

''' <summary>
''' Object��JSON�ϊ�
'''   - Class��`�����I�u�W�F�N�g��JSON������ɕϊ�����B
'''   - �n���l�̓I�u�W�F�N�g�l�A�z��AArrayList���\�B
'''   - �n���I�u�W�F�N�g�̃N���X�^��+"PropNames"�Ƃ������O�̊֐������O�ɒ�`����K�v������B�֐��̖߂�l�́A���̃N���X�̃v���p�e�B���̔z��B
'''       ��) Function MyClassPropNames()
'''              Return Array("prop01", "prop02", "prop03")
'''           End Function
'''   - �����I��Execute�֐����g���Ă����CPU�p���[���Ȃ��}�V���ł͒x���B���x���C�ɂȂ�ꍇ�ɂ́AProp�v���p�e�B����������Props�N���X��ʓr��`����ƍ���������B
''' </summary>
''' <param name="data">�I�u�W�F�N�g�^�̒l�A���̓I�u�W�F�N�g�^�̒l��ArrayList</param>
''' <param name="fieldNames">data�̎��v���p�e�B���̔z��B</param>
''' <returns>JSON������</returns>
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

                'Props�N���X�̐����i���݂��Ȃ��ꍇ��Nothing�j
                Dim props
                Set props = Nothing
                Dim propsClassName
                propsClassName = className & "Props"
                On Error Resume Next
                Execute("Set props = New " & propsClassName)
                On Error Goto 0

                'Props�N���X������ꍇ�A�I�u�W�F�N�g�Ɗ֘A�Â���
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
''' �������JSON�p�ɃG���R�[�h����B
''' This function is from "aspjson" by tugrul in GitHub
''' </summary>
''' <param name="str">������</param>
''' <returns>JSON�G���R�[�h���ꂽ������</returns>
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
' HTML�e���v���[�g�L�q�⏕
'==========================================================

''' <summary>
''' display:none���o�͂��邩�ǂ����Bfalse�̏ꍇ�̂݁Adisplay:none���o�͂���BIIf�֐����V���v���ɏ�����B
''' </summary>
''' <param name="expr">none���o�͂��邩�ǂ������߂������</param>
''' <returns>True�̎��A�󕶎��BFalse�̎��Adisplay:none;</returns>
Function css_display(expr)

	If expr Then
		css_display = ""
	Else
		css_display = "display:none;"
	End If

End Function

''' <summary>
''' visibility�̐ؑցBIIf�֐����V���v���ɏ�����B
''' </summary>
''' <param name="expr">visibility��visible�ɂ��邩�ǂ��������߂������</param>
''' <returns>expr��True�̎��Avisibility:visible:�AFalse�̎��Avisibility:hidden;</returns>
Function css_visibility(expr)

	If expr Then
		css_visibility = "visibility:visible;"
	Else
		css_visibility = "visibility:hidden;"
	End If

End Function

''' <summary>
''' �w�蕶��������o�͂��邩�ǂ����BIIf�֐����V���v���ɏ�����B
''' </summary>
''' <param name="str">�o�͂��镶����</param>
''' <param name="expr">str���o�͂��������</param>
''' <returns>expr��True�̎��Astr��Ԃ��BFalse�̎��A�󕶎���Ԃ��B</returns>
Function write_if(str, expr)

	If expr Then
		write_if = str
	Else
		write_if = ""
	End If

End Function

''' <summary>
''' option�^�O���ȉ��̌`���Ń��X�g�o�͂����������Ԃ��B
''' <option value="arrKeys(i)">arrCaptions(i)</option>
''' arrKeys(i)��selectedKey�ƈ�v�����ꍇ�Aselected������t�^����B
''' </summary>
''' <param name="arrKeys">option�^�O��value�̔z��</param>
''' <param name="arrCaptions">option�^�O�̃��x��</param>
''' <param name="selectKey">selected�ɂ���key</param>
''' <returns></returns>
Function html_options(arrKeys, arrCaptions, selectedKey)
	Dim html
	html = ""

	Dim i
	For i = LBound(arrKeys) To UBound(arrKeys)
		If IsNullOrEmpty(arrKeys(i)) Then
			Exit For    '�z��̃L�[���󕶎��Ȃ�΂��������͏o�͂��Ȃ��i��̔z��L�[��z��̏I�[�Ƃ��Ĉ����Ă���ׁj
		End If
		html = html & "<option value=""" & arrKeys(i) & """" & write_if(" selected", arrKeys(i) = selectedKey) & ">"
		html = html & arrCaptions(i) & "</option>"
	Next

	html_options = html

End Function

%>
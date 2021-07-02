<%
'--------------------------------------------------------------------
' ASPVBSLib Core
'--------------------------------------------------------------------

'==========================================================
' �R�[�f�B���O��⏕����֐�
'==========================================================

''' <summary>
''' ���̕]���Ɋ�Â��āA2�̕����̂����ꂩ��Ԃ��B
''' VB��IIf�֐� ��ցBVB��IIf�Ɠ��l�ɁATruePart��FalsePart�͂ǂ�����K���]������鎖�ɒ��ӁB
''' </summary>
''' <param name="expr">�]�����鎮</param>
''' <param name="TruePart">expr��True�̎��ɕԂ����l</param>
''' <param name="FalsePart">expr��False�̎��ɕԂ����l</param>
''' <returns></returns>
Function IIf(expr, TruePart, FalsePart)

	If expr Then
		IIf = TruePart
	Else
		IIf = FalsePart
	End If

End Function

'==========================================================
' �z��������֐�
'==========================================================

''' <summary>
''' �w�肵���l���z��̒��ɂ����True�A�Ȃ����False��Ԃ��B
''' </summary>
''' <param name="value">�����������l</param>
''' <param name="arr">�z��</param>
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
''' �w�肵���l���z��̒��ɂ���΁A���̗v�f��index��Ԃ��B������Ȃ����-1��Ԃ��B
''' </summary>
''' <param name="searchKey">�����������l</param>
''' <param name="arr">�z��</param>
''' <returns>�z�񒆂�index</returns>
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
''' �w�肵���l���z��arrKeys�̒��ɂ���΁AarrValues�̓����ʒu�̗v�f��Ԃ��B������Ȃ���΋󕶎���Ԃ��B
''' </summary>
''' <param name="searchKey">�����������l</param>
''' <param name="arrKeys">�L�[�̔z��</param>
''' <param name="arrValues">�l�̔z��</param>
''' <returns>���������l</returns>
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
' �l�̃`�F�b�N�␮�`���s���֐�
'==========================================================

''' <summary>
''' �l��Null����Empty���͋󕶎����ǂ����������B
''' </summary>
''' <param name="value">�l�Ƃ��ēn��������</param>
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
''' �l��Null����Empty���͋󕶎��̎��ɋK��l�ƂȂ�l��Ԃ��B����ȊO�̎��͒l���̂��̂�Ԃ��B
''' </summary>
''' <param name="value">�l</param>
''' <param name="nullval">�l��Null����Empty���͋󕶎��̎��ɕԂ��l</param>
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
''' ���t�^�̒l��VB.NET��DateTime.ToString�̂悤�ȏ����w��Ő��`����B
'''  ��) VBDateFormat( Now, "yyyy�NMM��dd�� HH��mm��ss�b")
''' </summary>
''' <param name="dt">���t�^�̒l</param>
''' <param name="formatString">�����w�蕶����B��) "yyyy�NMM��dd�� HH��mm��ss�b", "yyyy.M.d H:m:s"</param>
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
''' �v���[�X�z���_�[�t�������񐮌`�B{0}�A{1}�A�c�̂悤�Ɏw�肷��ƁA�����̓����ʒu�̒l�����̏ꏊ�ɖ��ߍ��ށB
''' ��) VBStringFormat( "Cd={0}��{1}�ł�", Array( "1234", "�ҏW��" ) )
''' ��) VBStringFormat( "{0}�����݂��܂���", "�Z�Z���" )
''' </summary>
''' <param name="formatString">�v���[�X�z���_�[�t��������</param>
''' <param name="args">������1�̏ꍇ�͒P��ŁA�����̏ꍇ�͔z��œn���B</param>
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
''' ��������w�肵�������܂ō����Ƀ[�����߂���B��jVBPadLeftZero("12", 5) => "00012"
''' </summary>
''' <param name="value">������</param>
''' <param name="length">�[�����߂��钷���B</param>
''' <returns>length���value�̕��������ꍇ�Avalue�����̂܂ܕԂ����B</returns>
Function VBPadLeftZero(value, length)
	Dim valueLen
	valueLen = Len(Trim(value))
	
	If valueLen < length Then
		VBPadLeftZero = String(length - valueLen, "0") & Trim(value)
	Else
		VBPadLeftZero = Trim(value)
	End If
End Function

' ''�������I�I�I�@64bit���ł�ScriptControl���Ȃ��׎��s�ł��Ȃ� �I�I�I������
' ''' <summary>
' ''' �v���[�X�z���_�[�t�������񐮌`�B{0}�A{1}�A�c�̂悤�Ɏw�肷��ƁA�����̓����ʒu�̒l�����̏ꏊ�ɖ��ߍ��ށB
' ''' ��) VBDateFormat( "Cd={0}��{1}�ł�", "1234", "�ҏW��" )
' ''' ��) VBDateFormat( "{0}�����݂��܂���", "�Z�Z���" )
' ''' </summary>
' ''' <param name="formatString">�v���[�X�z���_�[�t��������</param>
' ''' <param name="args">�����B�����w��\�B</param>
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
''' �^���w�肵�Ă̌^�ϊ�
''' </summary>
''' <param name="v">���͒l</param>
''' <param name="vType">�ϊ�����^��\���萔</param>
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
''' �^�̋K��l���擾����
''' </summary>
''' <param name="vType">�^��\���萔</param>
''' <returns>�^�̋K��l</returns>
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
' ���K�\���������֐�
'==========================================================

''' <summary>
''' ������̒��ɁA�p�^�[���Ƀ}�b�`���镔�������邩�ǂ�����Ԃ��B
''' </summary>
''' <param name="input">�e�X�g������������</param>
''' <param name="pattern">���K�\���p�^�[��������</param>
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
''' ������ɐ����ȊO���܂܂�Ă��Ȃ����ǂ�����Ԃ��B
''' </summary>
''' <param name="test">�e�X�g������������</param>
''' <returns>True:�܂܂�Ă��Ȃ� False:�܂܂�Ă���</returns>
Function ChkNum(test)
	ChkNum = Not RegexIsMatch(test, "[^0-9]")
End Function

''' <summary>
''' �����񂪐������p�����ȊO���܂܂�Ă��Ȃ����ǂ�����Ԃ��B
''' </summary>
''' <param name="test">�e�X�g������������</param>
''' <returns>True:�܂܂�Ă��Ȃ� False:�܂܂�Ă���</returns>
Function ChkNumAlphabet(test)
	ChkNumAlphabet = Not RegexIsMatch(test, "[^a-zA-Z0-9]")
End Function


'==========================================================
' ArrayList �N���X
'==========================================================
''' <summary>
''' VB.NET��ArrayList�Ɏ����g�������ł���N���X(������Dictionary���g�p�j
''' Linq�I�Ɏg���� Where, Map(Select�ɑ���), FirstOrDefault, OrderBy, Reverse �������B
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
    ''' ���X�g�̌����B
    ''' </summary>
    ''' <returns></returns>
	Public Property Get Count
        Count = pItems.Count
    End Property

    ''' <summary>
    ''' Index�ŗv�f���擾
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
    ''' �v�f�̃��X�g��z��Ŏ擾
    ''' </summary>
    ''' <returns></returns>
	Public Property Get Items
        Items = pItems.Items
    End Property

    ''' <summary>
    ''' �v�f�̌^��ݒ�
    ''' </summary>
    ''' <returns></returns>
    Public Property Let ItemType(aType)
        pItemType = aType
    End Property

    ''' <summary>
    ''' �v�f�̌^���擾
    ''' </summary>
    ''' <returns></returns>
    Public Property Get ItemType
        ItemType = pItemType
    End Property

    ''' <summary>
    ''' �v�f���I�u�W�F�N�g�^�̏ꍇ�A�N���X����ݒ�
    ''' </summary>
    ''' <returns></returns>
    Public Property Let ItemClassName(aClassName)
        pItemClassName = aClassName
		If Len(aClassName) > 0 Then
	        pItemType = vbObject
		End If
    End Property

    ''' <summary>
    ''' �v�f���I�u�W�F�N�g�^�̏ꍇ�A�N���X�����擾
    ''' </summary>
    ''' <returns></returns>
    Public Property Get ItemClassName
        ItemClassName = pItemClassName
    End Property

    ''' <summary>
    ''' �����̌^���v�f�̌^�ƈ�v���Ă��邩
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
    ''' ���X�g�̍ŏ��̗v�f���擾�iItems(0)�ɑ������邪�A���X�g����̏ꍇ�ɂ͎w�肵���K��l��Ԃ��B�j
    ''' </summary>
	''' <param name="defaultValue">���X�g����̏ꍇ�̋K��l</param>
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
    ''' ���X�g�̍ŏ��̗v�f���擾�i���X�g����̏ꍇ�ɂ͎w�肵���K��l��Ԃ��B�j
    ''' </summary>
	''' <param name="defaultValue">���X�g����̏ꍇ�̋K��l</param>
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
    ''' �v�f�𖖔��ɒǉ�
    ''' </summary>
	Public Sub Add(item)
        If Not CheckType(item) Then
            Err.Raise vbObjectError + 1,,"type of item mismatch. adding item type = " & VarType(item) & ", list item type = " & pItemType
        End If
		pItems.Add pKeyMax, item
		pKeyMax = pKeyMax + 1
	End Sub

    ''' <summary>
    ''' ���X�g�𖖔��ɒǉ�
    ''' </summary>
	Public Sub AddList(list)
		Dim item
		For Each item In list.Items
			Call Me.Add(item)
		Next
	End Sub

    ''' <summary>
    ''' ���X�g���N���A����
    ''' </summary>
	Public Sub Clear()
		pItems.RemoveAll
		pKeyMax = 0
	End Sub

    ''' <summary>
    ''' �w�肵���v�f�����X�g����폜����
    ''' </summary>
	''' <param name="item">�폜����v�f</param>
	Public Sub Remove(item)
		Dim key
		key = FindItem(item)
		If key <> -1 Then
			pItems.Remove(key)
		Else
			Err.Raise 9 '�C���f�b�N�X���L���͈͂ɂ���܂���
		End If
	End Sub

    ''' <summary>
    ''' �w�肵���v�f�����X�g����폜����
    ''' </summary>
	''' <param name="index">�폜����v�f�̈ʒu(0�`)</param>
	Public Sub RemoveAt(index)
		If index < 0 Or index >= Me.Count Then
			Err.Raise 9 '�C���f�b�N�X���L���͈͂ɂ���܂���
		Else
			Dim arr
			arr = pItems.Items
			Call Me.Remove(arr(index))
		End If
	End Sub

    ''' <summary>
    ''' �w�肵���v�f�����X�g�Ɋ܂܂�Ă��邩
    ''' </summary>
	''' <param name="item">�܂܂�Ă��邩�m�F�������v�f</param>
	''' <returns>True: �܂܂�Ă���, False: �܂܂�Ă��Ȃ�</returns>
	Public Function Contains(item)
		Contains = (IndexOf(item) <> -1)
	End Function

    ''' <summary>
    ''' �w�肵���v�f�����X�g���猟�����A0����n�܂�ʒu��Ԃ��B���݂��Ȃ��ꍇ��-1��Ԃ��B
    ''' </summary>
	''' <param name="searchItem">�v�f</param>
	''' <returns>�v�f�̈ʒu</returns>
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
    ''' �w�肵���v�f�����X�g���猟�����A�����R���N�V������key��Ԃ�
    ''' </summary>
	''' <param name="searchItem">�܂܂�Ă��邩�m�F�������v�f</param>
	''' <returns>�����R���N�V������key</returns>
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
	''' ���X�g��Dictionary�ɕϊ�
	''' ��) ToDictionary( "item.ID", "item.Name" )
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
	''' ���X�g�̍i�荞�݁BVB.NET��Enumerable(Of T).Where�ɑ���
	''' ��) list.Where( "item.Age < p", 20 )
	''' ��) list.Where( "item.SectionCd = p.SectionCd", targetSection )
    ''' </summary>
    ''' <param name="expr">�e�v�f���I��ΏۂƂȂ邩�𔻒f����ׂ́A�^�U�l��Ԃ�������������B"item"�͊e�v�f�A"p"�͑�������\���B</param>
    ''' <param name="p">�p�����[�^�Ƃ��ė^��������Bexpr����"p"�Ƃ��ĎQ�Ƃł���B</param>
    ''' <returns>�����Ɉ�v����ArrayList��Ԃ��B�����Ɉ�v���錋�ʂ��Ȃ������ꍇ�A���ArrayList��Ԃ��B</returns>
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
	''' ���X�g�̍i�荞�݁B
	''' �����Ɉ�v����ŏ��̗v�f�𔲂��o���B�߂�l�͋�̃��X�g��1���݂̂̃��X�g�B
	''' ��) list.FindFirst( "item.Age < p", 20 )
	''' ��) list.FindFirst( "item.SectionCd = p.SectionCd", targetSection )
    ''' </summary>
    ''' <param name="expr">�e�v�f���I��ΏۂƂȂ邩�𔻒f����ׂ́A�^�U�l��Ԃ�������������B"item"�͊e�v�f�A"p"�͑�������\���B</param>
    ''' <param name="p">�p�����[�^�Ƃ��ė^��������Bexpr����"p"�Ƃ��ĎQ�Ƃł���B</param>
    ''' <returns>�����Ɉ�v����ŏ��̒l��1�܂�ArrayList��Ԃ��B�����Ɉ�v���錋�ʂ��Ȃ������ꍇ�A���ArrayList��Ԃ��B</returns>
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
	''' ���X�g���̗v�f��ʂ̌`���̗v�f�ɒu��������i�ˉe�j�BVB.NET��Enumerable(Of T).Select�ɑ����B
	''' ��) Map( "item.ID & "":"" & item.Name" )
    ''' </summary>
    ''' <param name="expr">�V�����v�f�ƂȂ�l��Ԃ����B"item"�͊e�v�f��\���B</param>
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
	''' Map�̕ʖ��B
	''' ��) Select( "item.ID & "":"" & item.Name" )
    ''' </summary>
    ''' <param name="expr">�V�����v�f�ƂȂ�l��Ԃ����B"item"�͊e�v�f��\���B</param>
    ''' <returns></returns>
	Public Function [Select](expr, p)
		Set [Select] = Map(expr, p)
	End Function

    ''' <summary>
	''' �\�[�g
	''' ��) OrderBy("item.Name", True) ��Name�v���p�e�B�Ń\�[�g
	''' ��) OrderBy("item", True)		���v�f���̂Ń\�[�g
    ''' </summary>
    ''' <param name="expr">��r�ΏۂƂ��鎮�B"item"�͊e�v�f��\���B</param>
    ''' <param name="ascending">True: �����AFalse:�~��</param>
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
	''' �\�[�g�i�����j
	''' ��) OrderByAsc("item.Name") ��Name�v���p�e�B�Ń\�[�g
	''' ��) OrderByAsc("item")		���v�f���̂Ń\�[�g
    ''' </summary>
    ''' <param name="expr">��r�ΏۂƂ��鎮�B"item"�͊e�v�f��\���B</param>
    ''' <returns></returns>
	Public Function OrderByAsc(expr)
        Set OrderByAsc = Me.OrderBy(expr, True)
    End Function

    ''' <summary>
	''' �\�[�g�i�~���j
	''' ��) OrderByAsc("item.Name") ��Name�v���p�e�B�Ń\�[�g
	''' ��) OrderByAsc("item")		���v�f���̂Ń\�[�g
    ''' </summary>
    ''' <param name="expr">��r�ΏۂƂ��鎮�B"item"�͊e�v�f��\���B</param>
    ''' <returns></returns>
	Public Function OrderByDesc(expr)
        Set OrderByDesc = Me.OrderBy(expr, False)
    End Function

    ''' <summary>
	''' ���X�g�̏����𔽓]
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
	''' �I�u�W�F�N�g���w�肵�����i������j�ŃX�J���[�l�ɕϊ�����
	''' ��) ItemToScalar( item, "item.ID & "":"" & item.Name" )
    ''' </summary>
    ''' <param name="item">�I�u�W�F�N�g�l�B</param>
    ''' <param name="expr">��r�ΏۂƂ��鎮�B"item"�͊e�v�f��\���B</param>
    ''' <returns></returns>
	Private Function ItemToScalar(item, expr)
		ItemToScalar = Eval(expr)
	End Function

    ''' <summary>
	''' �l�̔�r
    ''' </summary>
    ''' <param name="x">��r����lx</param>
    ''' <param name="y">��r����ly</param>
    ''' <returns>x = y �̎� 0�Ax > y �̎� 1�Ax < y �̎� -1 ��Ԃ��B</returns>	
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
			' �^�̈Ⴄ�I�u�W�F�N�g�͌^���Ń\�[�g
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
	''' �v�f�̓���ւ�
    ''' </summary>
    ''' <param name="x">����ւ��邷��lx</param>
    ''' <param name="y">����ւ���ly</param>
	Private Sub swap(ByRef x, ByRef y)
		Dim d
		SetVariable d, x
		SetVariable x, y
		SetVariable y, d
	End Sub

    ''' <summary>
	''' �l�̐ݒ�
    ''' </summary>
    ''' <param name="TargetVal">�ݒ��̕ϐ�</param>
    ''' <param name="SrcVal">�ݒ肷��l</param>
	Private Sub SetVariable(ByRef TargetVal, ByRef SrcVal)
		If IsObject(SrcVal) Then
			Set TargetVal = SrcVal
		Else
			TargetVal = SrcVal
		End If
	End Sub

End Class

'==========================================================
' ���t���N�V�����E���I�v���O���~���O�⏕�@�\
'==========================================================

''' <summary>
''' �w�肵���I�u�W�F�N�g�̃v���p�e�B���ꗗ��z��Ŏ擾����
''' ���N���X��+PropNames�Ƃ������O�̃v���p�e�B���擾�p�֐��̒�`���K�v�B
''' ���v���p�e�B���擾�p�֐��̒�`���Ȃ��ꍇ�A��O����������B
''' </summary>
''' <param name="obj">���[�U�[��`�I�u�W�F�N�g</param>
''' <returns>�v���p�e�B���̔z��</returns>
Function GetPropNames(obj)
	Dim className: className = TypeName(obj)
	Dim propNamesFuncName: propNamesFuncName = className & "PropNames"

	Dim propNames
	On Error Resume Next
	propNames = Eval(propNamesFuncName)
	On Error Goto 0

    If IsEmpty(propNames) Then
        Err.Raise vbObjectError + 513, className & "PropNames�֐��̒�`���Ȃ����A��`�����S�ł͂���܂���"
    End If

	GetPropNames = propNames

End Function

''' <summary>
''' �w�肵���I�u�W�F�N�g�̊e�v���p�e�B���A�ΏۂƂȂ�I�u�W�F�N�g�ɃR�s�[����B
''' �R�s�[���ƃR�s�[��̃I�u�W�F�N�g�͓����v���p�e�B������Ηǂ��A�K�����������N���X�ł���K�v�͂Ȃ��B
''' �܂��A�R�s�[��ɑ��݂��Ȃ��v���p�e�B�͖��������B
''' �v���p�e�B���I�u�W�F�N�g�^�̏ꍇ�A�Q�Ƃ��R�s�[�����i�󂢃R�s�[�j�B
''' ���R�s�[���̃N���X��+PropNames�Ƃ������O�̃v���p�e�B���擾�p�֐��̒�`���K�v�B
''' ���v���p�e�B���擾�p�֐��̒�`���Ȃ��ꍇ�A��O����������B
''' </summary>
''' <param name="source">�R�s�[���̃I�u�W�F�N�g</param>
''' <param name="target">�R�s�[��̃I�u�W�F�N�g</param>
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
''' �w�肵���I�u�W�F�N�g�����ɐV�����I�u�W�F�N�g�𕡐����ĕԂ��B
''' �w�肵���I�u�W�F�N�g�̊e�v���p�e�B���A���������I�u�W�F�N�g�ɃR�s�[����B
''' �v���p�e�B���I�u�W�F�N�g�^�̏ꍇ�A�Q�Ƃ��R�s�[�����i�󂢃R�s�[�j�B
''' ���������̃N���X��+PropNames�Ƃ������O�̃v���p�e�B���擾�p�֐��̒�`���K�v�B
''' ���v���p�e�B���擾�p�֐��̒�`���Ȃ��ꍇ�A��O����������B
''' </summary>
''' <param name="obj">�R�s�[���̃I�u�W�F�N�g</param>
''' <returns>���������I�u�W�F�N�g</returns>
Function CloneObject(obj)
	Dim className: className = TypeName(obj)
	Dim target: Set target = Eval("New " & className)

	CopyProps obj, target
	Set CloneObject = target

End Function
%>
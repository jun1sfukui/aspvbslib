<%
'--------------------------------------------------------------------
' ASPVBSLib Data
'--------------------------------------------------------------------
'==========================================================
' SqlQuery �N���X
'==========================================================

''' <summary>
''' SQL�N�G���̌��ʂ��^�t���I�u�W�F�N�g��ArrayList�Ƀ}�b�s���O����B
''' �@- �^����SQL��SELECT��̃t�B�[���h�����ϊ���̃N���X�̃t�B�[���h���ƈ�v���Ă���ꍇ�Ɏ����I��O/R�}�b�s���O���s��
'''	  - �t�B�[���h��Null�̏ꍇ�ANull���}�b�s���O�����B�iIsNull�֐��Ń`�F�b�N�ł���j
''' �@- SELECT��� * �ɂ��Ă���ƁA�S�t�B�[���h�̃}�b�s���O�����݂Ēx���Ȃ�ׁA�����I��SELECT��Ƀt�B�[���h�����w�肵�������ǂ��B
''' �@- �����I��Execute�֐����g���Ă����CPU�p���[���Ȃ��}�V���ł͒x���B���x���C�ɂȂ�ꍇ�ɂ́AProp�v���p�e�B����������Props�N���X��ʓr��`����ƍ���������B
''' </summary>
''' <param name="DbConn">DbConnection�I�u�W�F�N�g</param>
''' <param name="className">OR�}�b�s���O��̃N���X���BSELECT��̃t�B�[���h���ƈ�v���郁���o���`���Ă����Ǝ����I�ɒl���ǂݍ��܂��B</param>
''' <param name="cmdOrSql">SQL�N�G�������� ���� ADODB.Command �I�u�W�F�N�g</param>
''' <returns>SQL�̌��ʂ��}�b�s���O���ꂽ�AclassName�Ŏw�肵���N���X�̃I�u�W�F�N�g��ArrayList�B���ʂ�0���̏ꍇ�͋��ArrayList��Ԃ��B</returns>	
Public Function SqlQuery(DbConn, className, cmdOrSql )
	Dim RS
	Set RS = Server.CreateObject("ADODB.Recordset")

    If VarType(cmdOrSql) = vbString Then
    	RS.Open cmdOrSql, DbConn, adOpenKeyset, adLockReadOnly    '�ÓI�X�N���[���A�ǎ���p
    Else
        cmdOrSql.ActiveConnection = DbConn
        Set RS = cmdOrSql.Execute
    End If

	Dim list
	Set list = New ArrayList

    'Props�N���X�̐����i���݂��Ȃ��ꍇ��Nothing�j
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

        'Props�N���X������ꍇ�A�I�u�W�F�N�g�Ɗ֘A�Â���
        If Not props Is Nothing Then
            Set props.Target = item
        End If

		Dim field
		Dim fieldValue

		For Each field In RS.Fields

			If IsNull(RS(field.Name)) Then
				'DB�̒l��Null�̎���Null�̂܂܎擾
				On Error Resume Next
				If Not props Is Nothing Then
					props.Prop(field.Name) = RS(field.Name)
				Else
					'CallByName item, field.Name, vbLet, RS(field.Name)
					Execute("item." & field.Name & " = RS(field.Name)") 'Null��ݒ�BIsNull(item.[FieldName])��Null�`�F�b�N�ł���B
				End If
				On Error Goto 0
			Else
				'�l���ݒ肳��Ă��鎞
				On Error Resume Next
				Select Case field.Type
					Case adChar
						'RTrim���Ď擾
						fieldValue = RTrim(RS(field.Name).Value)
						If Not props Is Nothing Then
							props.Prop(field.Name) = fieldValue
						Else
							'CallByName item, field.Name, vbLet, fieldValue '(RTrim(RS(field.name).Value))
							Execute("item." & field.Name & " = fieldValue")
						End If

					Case adDBTimeStamp
						'CDate���Ď擾
						fieldValue = CDate(RS(field.Name).Value)
						If Not props Is Nothing Then
							props.Prop(field.Name) = fieldValue
						Else
							'CallByName item, field.Name, vbLet, CDate(RS(field.Name).Value)
							Execute("item." & field.Name & " = fieldValue")
						End If

					Case Else
						'���̂܂܎擾
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
' ArrayList �� SqlQuery �Ƒg�ݍ��킹�Ďg�p�ł����\�I�ȃG���e�B�e�B�N���X
' �����̗�����ɁA���[�U�[��`�̃G���e�B�e�B�N���X���`���鎖���ł���B
'==========================================================

''' <summary>
''' �L�[�o�����[�N���X
''' �L�[�ƒl�̑g�������ʃN���X�B
''' </summary>
Class KeyAndValue
	''' <summary>�L�[</summary>
	Public Key

	'�l
	Public Value

    ''' <summary>
    ''' �R���X�g���N�^�� �l�ݒ�p�f�t�H���g�v���p�e�B�B
    ''' ��) Set keyValue = (New KeyAndValue)(key, value)
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
''' �L�[�o�����[�N���X�p�̃v���p�e�B��`�֐�
''' ��`���Ă�����JSONValue�֐���LoadRequest�֐��������p�ł���B
''' </summary>
Function KeyAndValuePropNames
	KeyAndValuePropNames = Array("Key", "Value")
End Function

''' <summary>
''' �L�[�o�����[�N���X�p��Props�N���X�B
''' Prop�v���p�e�B�������AGet��Let�ɂāA�v���p�e�B���Ǝ��ۂ̃v���p�e�B���֘A�t����B
''' ��`���Ă�����SqlQuery�֐��̑��x�����シ��B
''' </summary>
Class KeyAndValueProps
    '-------------------------------------------------------
    ' �ʏ�AProps�N���X�̓N���X���ȊO�͂��̂܂ܗ��p���A
    ' Prop�v���p�e�B�̃��[�U�[��`���̂ݑΏۃN���X�p�Ɏ�������B
    '-------------------------------------------------------
    '��obj, Target �͕ύX�s��
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
            '-- �������烆�[�U�[��`���� ------
			Case "Key"
				value = obj.Key
			Case "Value"
				value = obj.Value
            '-- �����܂� ---------------------
			Case Else
				value = Eval("obj." & propName)
		End Select
		Prop = value
	End Property

	Public Property Let Prop( propName, value )
		Select Case propName
            '-- �������烆�[�U�[��`���� ------
			Case "Key"
				obj.Key = value
			Case "Value"
				obj.Value = value
            '-- �����܂� ---------------------
			Case Else
				Execute("obj." & propName & " = value")
		End Select
	End Property

End Class

'==========================================================
' DB�w���p�[
'==========================================================

''' <summary>
''' JSONValue�֐���Date�^�̒l�𐮌`����ۂ̐��`������B
''' </summary>
Dim DBSqlValue_DATE_FORMAT: DBSqlValue_DATE_FORMAT = "'yyyy-MM-dd HH:mm:ss'"

''' <summary>
''' �l�̌^�ɉ�����SQL��̃��e�����\����Ԃ�
''' </summary>
''' <param name="val">�l</param>
''' <returns>SQL���e����������</returns>
Function DBSqlValue(value)

	If IsObject(value) Then
		Err.Raise 13	'�^����v���܂���
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
			Err.Raise 13	'�^����v���܂���

	End Select

	DBSqlValue = v
End Function

%>
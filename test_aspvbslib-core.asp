<% Option Explicit %>
<!--#include file="./aspvbslib/adovbs.inc"-->
<!--#include file="./aspvbslib/aspvbslib-core.inc"-->
<!--#include file="./aspvbslib/aspvbslib-test.inc"-->
<%
'--------------------------------------------------------------------
' ASPVBSLib Core のユニットテスト
'--------------------------------------------------------------------

Sub TestMain()
	Dim test
	Set test = New UnitTest
	
	test.Add("IIfTest")
	test.Add("InArrayTest")
	test.Add("ArrayIndexOfTest")
	test.Add("FindKeyValueTest")
	test.Add("IsNullOrEmptyTest")
	test.Add("NullValueTest")
	test.Add("VBDateFormatTest")
	test.Add("VBStringFormatTest")
	test.Add("VBPadLeftZeroTest")
	test.Add("ConvertToTest")
	test.Add("DefaultValueTest")
	test.Add("RegexIsMatchTest")
	test.Add("ChkNumTest")
	test.Add("ChkNumAlphabetTest")
	test.Add("ArrayListTest")
	test.Add("ArrayListTest2")
	test.Add("CopyPropsTest")
	test.Add("CloneObjectTest")

	test.RunTest
	test.ResultHtml

End Sub

Call TestMain()


Sub IIfTest()
	AssertEquals IIf(True, "OK", "OK"), "OK"
	AssertEquals IIf(False, "OK", "NG"), "NG"
End Sub

Sub InArrayTest()
	Dim arr
	arr = Array("1","2","c", "test")

	AssertEquals InArray("1", arr), True
	AssertEquals InArray("2", arr), True
	AssertEquals InArray("c", arr), True
	AssertEquals InArray("test", arr), True
	AssertEquals InArray("d", arr), False
	AssertEquals InArray(1, arr), False
	AssertEquals InArray(0, arr), False
	AssertEquals InArray("", arr), False
End Sub

Sub ArrayIndexOfTest()
	Dim arr
	arr = Array("1","2","c", "test")

	AssertEquals ArrayIndexOf("1", arr), 0
	AssertEquals ArrayIndexOf("2", arr), 1
	AssertEquals ArrayIndexOf("c", arr), 2
	AssertEquals ArrayIndexOf("test", arr), 3
	AssertEquals ArrayIndexOf("d", arr), -1
	AssertEquals ArrayIndexOf(1, arr), -1
	AssertEquals ArrayIndexOf(0, arr), -1
	AssertEquals ArrayIndexOf("", arr), -1
End Sub

Sub FindKeyValueTest()
	Dim arrKeys, arrValues
	arrKeys = Array("a", "b", "c")
	arrValues = Array("name_a", "name_b", "name_c")

	AssertEquals FindKeyValue("a", arrKeys, arrValues), "name_a"
	AssertEquals FindKeyValue("b", arrKeys, arrValues), "name_b"
	AssertEquals FindKeyValue("c", arrKeys, arrValues), "name_c"
	AssertEquals FindKeyValue("d", arrKeys, arrValues), ""
	AssertEquals FindKeyValue("", arrKeys, arrValues), ""

	Dim value
	AssertEquals FindKeyValue(value, arrKeys, arrValues), ""

End Sub

Sub IsNullOrEmptyTest()
	Dim value
	AssertEquals IsNullOrEmpty(value), True

	Set value = Nothing
	AssertEquals IsNullOrEmpty(value), True

	value = ""
	AssertEquals IsNullOrEmpty(value), True

	value = "a"
	AssertEquals IsNullOrEmpty(value), False

	value = "0"
	AssertEquals IsNullOrEmpty(value), False

	value = 0
	AssertEquals IsNullOrEmpty(value), False

	value = array()
	AssertEquals IsNullOrEmpty(value), False

	Set value = New ArrayList
	AssertEquals IsNullOrEmpty(value), False

End Sub

Sub NullValueTest()
	Dim value, nullval

	AssertEquals NullValue(value, "N"), "N"

	Set nullval = New ArrayList
	AssertEquals NullValue(value, nullval), nullval

	Set value = Nothing
	AssertEquals NullValue(value, "N"), "N"

	value = ""
	AssertEquals NullValue(value, "N"), "N"

	value = "Y"
	AssertEquals NullValue(value, "N"), "Y"

	Set value = Nothing
	Set nullval = New ArrayList
	AssertEquals NullValue(value, nullval), nullval

	Set value = New ArrayList
	Set nullval = New ArrayList
	AssertEquals NullValue(value, nullval), value

End Sub


Sub VBDateFormatTest()
	Dim dt1, dt2

	dt1 = #2021/1/2 3:4:5#
	dt2 = #12/31/1999 23:59:58#

	AssertEquals VBDateFormat(dt1, "y年M月d日"), "2021年1月2日"
	AssertEquals VBDateFormat(dt1, "yy年MM月dd日"), "21年01月02日"
	AssertEquals VBDateFormat(dt1, "yyyy年MM月dd日 H時m分s秒"), "2021年01月02日 3時4分5秒"
	AssertEquals VBDateFormat(dt1, "yyyy年MM月dd日 HH時mm分ss秒"), "2021年01月02日 03時04分05秒"
	AssertEquals VBDateFormat(dt1, "HH時mm分ss秒"), "03時04分05秒"

	AssertEquals VBDateFormat(dt2, "y年M月d日"), "1999年12月31日"
	AssertEquals VBDateFormat(dt2, "yy年MM月dd日"), "99年12月31日"
	AssertEquals VBDateFormat(dt2, "yyyy年MM月dd日 H時m分s秒"), "1999年12月31日 23時59分58秒"
	AssertEquals VBDateFormat(dt2, "yyyy年MM月dd日 HH時mm分ss秒"), "1999年12月31日 23時59分58秒"
	AssertEquals VBDateFormat(dt2, "HH時mm分ss秒"), "23時59分58秒"

End Sub

Sub VBStringFormatTest()

	AssertEquals VBStringFormat("ID:{0}", 123), "ID:123"

	Dim params
	params = array(123, "メッセージ123")

	AssertEquals VBStringFormat("ID:{0}, Message:{1}", params), "ID:123, Message:メッセージ123"
	AssertEquals VBStringFormat("ID:{0}", params), "ID:123"
	AssertEquals VBStringFormat("Message:{1}", params), "Message:メッセージ123"

End Sub

Sub VBPadLeftZeroTest()

	AssertEquals VBPadLeftZero(123, 5), "00123"
	AssertEquals VBPadLeftZero("abc", 5), "00abc"
	AssertEquals VBPadLeftZero("", 5), "00000"
	AssertEquals VBPadLeftZero("", 0), ""
	AssertEquals VBPadLeftZero("12345", 0), "12345"
	AssertEquals VBPadLeftZero("12345", 2), "12345"
	AssertEquals VBPadLeftZero("12345", 5), "12345"
	AssertEquals VBPadLeftZero("12345", 6), "012345"

End Sub

Sub ConvertToTest()

	AssertEquals ConvertTo("True", vbBoolean), True
	AssertEquals ConvertTo("False", vbBoolean), False
	AssertEquals ConvertTo("123", vbInteger), 123
	AssertEquals ConvertTo("123", vbCurrency), 123
	AssertEquals ConvertTo("123", vbLong), 123
	AssertEquals ConvertTo("2021/01/02 3:4:5", vbDate), #2021/01/02 3:4:5#
	AssertEquals ConvertTo("123.456", vbSingle), 123.456
	AssertEquals ConvertTo("123.456", vbDouble), 123.456

	AssertEquals ConvertTo("", vbBoolean), False
	AssertEquals ConvertTo("", vbInteger), 0
	AssertEquals ConvertTo("", vbCurrency), 0
	AssertEquals ConvertTo("", vbLong), 0
	AssertEquals ConvertTo("", vbDate), CDate(0)
	AssertEquals ConvertTo("", vbSingle), 0
	AssertEquals ConvertTo("", vbDouble), 0

End Sub

Sub DefaultValueTest()

	AssertEquals DefaultValue(vbBoolean), False
	AssertEquals DefaultValue(vbInteger), 0
	AssertEquals DefaultValue(vbLong), 0
	AssertEquals DefaultValue(vbDate), CDate(0)
	AssertEquals DefaultValue(vbSingle), 0
	AssertEquals DefaultValue(vbDouble), 0

End Sub

Sub RegexIsMatchTest()

	AssertEquals RegexIsMatch("abc123edf", "[0-9]+"), True
	AssertEquals RegexIsMatch("abcedf", "[0-9]+"), False
	AssertEquals RegexIsMatch("050-1234-5678", "0\d{1,4}-\d{1,4}-\d{4}"), True
	AssertEquals RegexIsMatch("a050-1234-5678", "0\d{1,4}-\d{1,4}-\d{4}"), True
	AssertEquals RegexIsMatch("a050-1234-5678", "^0\d{1,4}-\d{1,4}-\d{4}$"), False

End Sub

Sub ChkNumTest()

	AssertEquals ChkNum("123456789"), True
	AssertEquals ChkNum("12345a6789"), False
	AssertEquals ChkNum("-123456789"), False
	AssertEquals ChkNum("+123456789"), False
	AssertEquals ChkNum(""), True
	AssertEquals ChkNum("１２３４５"), False

End Sub

Sub ChkNumAlphabetTest()

	AssertEquals ChkNumAlphabet("abcDEF"), True
	AssertEquals ChkNumAlphabet("12345abcDEF"), True
	AssertEquals ChkNumAlphabet("+12345abcDEF"), False
	AssertEquals ChkNumAlphabet("+"), False
	AssertEquals ChkNumAlphabet(""), True

End Sub

Sub ArrayListTest()

	'----------------------------
	' プリミティブ型の要素
	'----------------------------

	'0件時の処理
	Dim list
	Set list = New ArrayList

	AssertEquals list.Count, 0
	AssertEquals list.FirstOrDefault("abc"), "abc"
	AssertEquals list.LastOrDefault("end"), "end"
	AssertEquals list.Where("item = p", "abc").Count, 0
	AssertEquals list.Map("item & p", "123").FirstOrDefault("abc"), "abc"
	AssertEquals list.FindFirst("item = p", "").FirstOrDefault("abc"), "abc"
	AssertEquals list.OrderByAsc("item").Count, 0
	AssertEquals list.OrderByDesc("item").Count, 0

	'1件時の処理	
	Dim item1
	item1 = "onlyone"
	list.Add item1

	AssertEquals list.Count, 1
	AssertEquals list.Item(0), item1
	AssertEquals list.ItemType, vbString
	AssertEquals list.ItemClassName, ""
	AssertEquals list.FirstOrDefault(""), item1
	AssertEquals list.LastOrDefault(""), item1

	Dim item

	For Each item In list.Items
		AssertEquals VarType(item), vbString
	Next

	For Each item In list.Where("item = p", item1).Items
		AssertEquals VarType(item), vbString
	Next

	AssertEquals list.Where("item = p", item1).Item(0), item1
	AssertEquals list.Where("item = p", "b").Count, 0
	AssertEquals list.Map("item & p", "123").FirstOrDefault("abc"), "onlyone123"
	AssertEquals list.FindFirst("item = p", item1).FirstOrDefault(""), "onlyone"
	AssertEquals list.FindFirst("item = p", "b").FirstOrDefault(""), ""
	AssertEquals list.OrderByAsc("item").Item(0), item1
	AssertEquals list.OrderByDesc("item").Item(0), item1

	'Add, Remove, RemoveAt, Clear, Contains
	Set list = New ArrayList
	list.Add "aaa"
	list.Add "bbb"
	list.Add "ccc"
	list.Add "ddd"
	list.Add "eee"

	AssertEquals list.Count, 5
	
	Call list.Remove("aaa")
	AssertEquals list.Count, 4
	AssertEquals list.Item(0), "bbb"
	AssertEquals list.Contains("aaa"), False
	AssertEquals list.Contains("eee"), True
	AssertEquals list.IndexOf("aaa"), -1
	AssertEquals list.IndexOf("eee"), 3

	AssertEquals list.Contains("ccc"), True
	list.RemoveAt(1)	'remove "ccc"
	AssertEquals list.Count, 3
	AssertEquals list.Contains("ccc"), False

	list.RemoveAt(0)
	list.RemoveAt(0)
	list.RemoveAt(0)
	AssertEquals list.Count, 0
	AssertEquals list.Contains("aaa"), False
	AssertEquals list.IndexOf("aaa"), -1

	list.Add "aaa"
	list.Add "bbb"
	list.Add "ccc"
	list.Add "ddd"
	list.Add "eee"
	AssertEquals list.Count, 5

	Dim sorted
	Set sorted = list.OrderByDesc("item")
	AssertEquals sorted.Item(0), "eee"
	AssertEquals sorted.Item(1), "ddd"
	AssertEquals sorted.Item(2), "ccc"
	AssertEquals sorted.Item(3), "bbb"
	AssertEquals sorted.Item(4), "aaa"

	Set sorted = sorted.OrderByAsc("item")
	AssertEquals sorted.Item(0), "aaa"
	AssertEquals sorted.Item(1), "bbb"
	AssertEquals sorted.Item(2), "ccc"
	AssertEquals sorted.Item(3), "ddd"
	AssertEquals sorted.Item(4), "eee"

	list.Clear
	AssertEquals list.Count, 0

End Sub


Class Emp
	Public empno
	Public empname

	Public Default Property Get Constructor(empno, empname)
		Me.empno = empno
		Me.empname = empname
		Set Constructor = Me
	End Property

End Class

Sub ArrayListTest2()

	'----------------------------
	' ユーザー定義型の要素
	'----------------------------

	Dim list
	Set list = New ArrayList

	'1件時の処理	
	Dim item1
	Set item1 = (New Emp)("123", "empname123")
	list.Add item1

	AssertEquals list.Count, 1
	AssertEquals list.Item(0), item1
	AssertEquals list.ItemType, vbObject
	AssertEquals list.ItemClassName, "Emp"
	AssertEquals list.FirstOrDefault(""), item1
	AssertEquals list.LastOrDefault(""), item1

	Dim item
	For Each item In list.Items
		AssertEquals VarType(item), vbObject
		AssertEquals TypeName(item), "Emp"
	Next

	For Each item In list.Where("item.empno = p", "123").Items
		AssertEquals TypeName(item), "Emp"
	Next


	AssertEquals list.Where("item.empno = p", "123").Item(0), item1
	AssertEquals list.Where("item.empno = p", "b").Count, 0
	AssertEquals list.Map("p & item.empname", "Name: ").FirstOrDefault(""), "Name: empname123"
	AssertEquals list.FindFirst("item.empno = p", "b").FirstOrDefault(Nothing), Nothing
	AssertEquals list.OrderByAsc("item.empno").Item(0), item1
	AssertEquals list.OrderByDesc("item.empno").Item(0), item1

	'Add, Remove, RemoveAt, Clear, Contains
	Set list = New ArrayList
	Call list.Add((New Emp)("ID001", "empname1"))
	Call list.Add((New Emp)("ID002", "empname2"))
	Call list.Add((New Emp)("ID003", "empname3"))
	Call list.Add((New Emp)("ID004", "empname4"))
	Call list.Add((New Emp)("ID005", "empname5"))

	AssertEquals list.Count, 5
	
	Set item = list.Item(0)
	Call list.Remove(item)
	AssertEquals list.Count, 4
	AssertEquals list.Item(0).empno, "ID002"
	AssertEquals list.Contains(list.Item(1)), True
	AssertEquals list.Contains(item), False
	AssertEquals list.IndexOf(list.Item(1)), 1
	AssertEquals list.IndexOf(item), -1

	list.RemoveAt(0)
	list.RemoveAt(0)
	list.RemoveAt(0)
	AssertEquals list.Count, 1

	list.RemoveAt(0)
	AssertEquals list.Count, 0
	AssertEquals list.Contains(item), False
	AssertEquals list.IndexOf(item), -1

	Call list.Add((New Emp)("ID001", "empname1"))
	Call list.Add((New Emp)("ID002", "empname2"))
	Call list.Add((New Emp)("ID003", "empname3"))
	Call list.Add((New Emp)("ID004", "empname4"))
	Call list.Add((New Emp)("ID005", "empname5"))
	AssertEquals list.Count, 5

	Dim sorted
	Set sorted = list.OrderByDesc("item.empno")
	AssertEquals sorted.Item(0).empname, "empname5"
	AssertEquals sorted.Item(1).empname, "empname4"
	AssertEquals sorted.Item(2).empname, "empname3"
	AssertEquals sorted.Item(3).empname, "empname2"
	AssertEquals sorted.Item(4).empname, "empname1"

	Set sorted = sorted.OrderByAsc("item.empname")
	AssertEquals sorted.Item(0).empname, "empname1"
	AssertEquals sorted.Item(1).empname, "empname2"
	AssertEquals sorted.Item(2).empname, "empname3"
	AssertEquals sorted.Item(3).empname, "empname4"
	AssertEquals sorted.Item(4).empname, "empname5"

	list.Clear
	AssertEquals list.Count, 0

End Sub


Function EmpPropNames
	EmpPropNames = Array("empno", "empname")
End Function

Class SpecialEmp
	Public empno
	Public empname
	Public boss
End Class
Function SpecialEmpPropNames
	SpecialEmpPropNames = Array("empno", "empname", "boss")
End Function

Class EmpName
	Public empname
End Class

Sub CopyPropsTest
	Dim emp1: Set emp1 = New Emp
	emp1.empno = 5
	emp1.empname = "田中"

	'コピー
	Dim emp2: Set emp2 = New Emp
	CopyProps emp1, emp2
	AssertEquals emp1.empno, emp2.empno
	AssertEquals emp1.empname, emp2.empname

	'コピー先にプロパティがない
	Dim empname1: Set empname1 = New EmpName
	CopyProps emp1, empname1
	AssertEquals emp1.empname, empname1.empname

	'コピー先に追加のプロパティがある
	Dim spemp1: Set spemp1 = New SpecialEmp
	CopyProps emp1, spemp1
	AssertEquals emp1.empno, spemp1.empno
	AssertEquals emp1.empname, spemp1.empname
	AssertEquals IsEmpty(spemp1.boss), True

	'オブジェクト型のプロパティ
	Dim spemp2: Set spemp2 = New SpecialEmp
	Set spemp1.boss = emp2
	CopyProps spemp1, spemp2
	AssertEquals spemp1.empno, spemp2.empno
	AssertEquals spemp1.empname, spemp2.empname
	AssertEquals spemp1.boss, spemp2.boss
	AssertEquals spemp1.boss.empno, spemp2.boss.empno
	AssertEquals spemp1.boss.empname, spemp2.boss.empname

End Sub

Sub CloneObjectTest
	Dim spemp1, spemp2
	Set spemp1 = New SpecialEmp
	spemp1.empno = 1234
	spemp1.empname = "斎藤"
	Set spemp1.boss = (New Emp)(5, "田中")

	'複製
	Set spemp2 = CloneObject(spemp1)

	AssertEquals spemp1.empno, spemp2.empno
	AssertEquals spemp1.empname, spemp2.empname
	AssertEquals spemp1.boss, spemp2.boss
	AssertEquals spemp1.boss.empno, spemp2.boss.empno
	AssertEquals spemp1.boss.empname, spemp2.boss.empname

End Sub

%>
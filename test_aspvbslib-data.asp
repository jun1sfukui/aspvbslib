<!--#include file="./aspvbslib/adovbs.inc"-->
<!--#include file="./aspvbslib/aspvbslib-core.inc"-->
<!--#include file="./aspvbslib/aspvbslib-data.inc"-->
<!--#include file="./aspvbslib/aspvbslib-test.inc"-->
<%
'--------------------------------------------------------------------
' ASPVBSLib Data のユニットテスト
'--------------------------------------------------------------------

'データベース接続
Dim DbConn
Set DbConn = Server.CreateObject("ADODB.Connection")
Dim dbfilepath :dbfilepath = Server.MapPath("aspvbslib_sample.mdb")
dim dbconnString :dbconnString = "Driver={Microsoft Access Driver (*.mdb)}; DBQ=" & dbfilepath & ";"
DbConn.Open dbconnString

'社員クラス
Class Emp 
    '社員番号
    Public empno
    '社員名
    Public empname
    '部署番号
    Public deptno
    '説明
    Public description
End Class

'部署クラス
Class Dept
    '部署番号
    Public deptno
    '部署名
    Public deptname
End Class

Class DeptProps
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
			Case "deptno"
				value = obj.deptno
			Case "deptname"
				value = obj.deptname
            '-- ここまで ---------------------
			Case Else
				value = Eval("obj." & propName)
		End Select
		Prop = value
	End Property

	Public Property Let Prop( propName, value )
		Select Case propName
            '-- ここからユーザー定義部分 ------
			Case "deptno"
				obj.deptno = value
			Case "deptname"
				obj.deptname = value
            '-- ここまで ---------------------
			Case Else
				Execute("obj." & propName & " = value")
		End Select
	End Property

End Class

Sub TestMain()
	Dim test
	Set test = New UnitTest
	
	test.Add("SqlQueryTest")
	test.Add("DBSqlValueTest")

	test.RunTest
	test.ResultHtml

End Sub

Call TestMain()


Sub SqlQueryTest
	Dim list
	Set list = SqlQuery(DbConn, "Emp", "SELECT * FROM emp")

	'全ての列
	AssertEquals list.Count, 12
	AssertEquals list.ItemClassName, "Emp"
	AssertEquals list.Item(0).empno, 1
	AssertEquals list.Item(0).empname, "社長花子"
	AssertEquals list.Item(0).deptno, 1
	AssertEquals TypeName(list.Item(0)), "Emp"

	'足りないプロパティ有り
	Set list = SqlQuery(DbConn, "Emp", "SELECT empno, deptno, description FROM emp")
	AssertEquals list.Count, 12
	AssertEquals list.ItemClassName, "Emp"
	AssertEquals list.Item(0).empno, 1
	AssertEquals list.Item(0).empname, ""
	AssertEquals list.Item(0).deptno, 1
	AssertEquals TypeName(list.Item(0)), "Emp"

	'マッピング先に存在しないプロパティ有り
	Set list = SqlQuery(DbConn, "Emp", "SELECT empno, deptno AS departmentno, description FROM emp")
	AssertEquals list.Count, 12
	AssertEquals list.ItemClassName, "Emp"
	AssertEquals list.Item(0).empno, 1
	AssertEquals list.Item(0).empname, ""
	AssertEquals list.Item(0).deptno, ""
	AssertEquals TypeName(list.Item(0)), "Emp"

	'値がNothingの列
	Set list = SqlQuery(DbConn, "Emp", "SELECT empno, NULL AS empname, NULL AS deptno, description FROM emp")
	AssertEquals list.Count, 12
	AssertEquals list.ItemClassName, "Emp"
	AssertEquals list.Item(0).empno, 1
	AssertEquals list.Item(0).empname, Null
	AssertEquals list.Item(0).deptno, Null
	AssertEquals TypeName(list.Item(0)), "Emp"

	'Propsクラスを定義Dept
	Set list = SqlQuery(DbConn, "Dept", "SELECT * FROM dept")
	AssertEquals list.Count, 6
	AssertEquals list.ItemClassName, "Dept"
	AssertEquals list.Item(0).deptno, 1
	AssertEquals list.Item(0).deptname, "経営企画部"
	AssertEquals TypeName(list.Item(0)), "Dept"

	'結果が0件のクエリ
	Set list = SqlQuery(DbConn, "Emp", "SELECT * FROM emp WHERE empno = 99999")
	AssertEquals list.Count,  0

End Sub

Sub DBSqlValueTest

	'数値型
	AssertEquals DBSqlValue(CLng(123)), "123"
	AssertEquals DBSqlValue(CInt(123)), "123"
	AssertEquals DBSqlValue(CCur(123)), "123"
	AssertEquals DBSqlValue(CSng(123.456)), "123.456"
	AssertEquals DBSqlValue(CDbl(123.456)), "123.456"
	AssertEquals DBSqlValue(CDbl(123.456)), "123.456"

	'文字列型
	AssertEquals DBSqlValue(""), "''"
	AssertEquals DBSqlValue("abc"), "'abc'"

	'日付型
	Dim dt
	dt = #12/31/2021 23:58:59#
	AssertEquals DBSqlValue(dt), "'2021-12-31 23:58:59'"
	dt = #12/31/2021#
	AssertEquals DBSqlValue(dt), "'2021-12-31 00:00:00'"

	'論理型
	AssertEquals DBSqlValue(True), "True"
	AssertEquals DBSqlValue(False), "False"

	'Null Values
	Dim nullVal
	AssertEquals DBSqlValue(nullVal), "Null"
	nullVal = Null
	AssertEquals DBSqlValue(nullVal), "Null"

End Sub

%>
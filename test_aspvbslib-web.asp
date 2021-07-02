<!--#include file="./aspvbslib/adovbs.inc"-->
<!--#include file="./aspvbslib/aspvbslib-core.inc"-->
<!--#include file="./aspvbslib/aspvbslib-web.inc"-->
<!--#include file="./aspvbslib/aspvbslib-test.inc"-->
<%
'--------------------------------------------------------------------
' ASPVBSLib Web のユニットテスト
'--------------------------------------------------------------------

Sub DispReloadPage()
%>
<form method="post" name="frm1" action="?reloaded=true&selectedEmpno=5&searchKeyword=apple123&employees=2&employees=4&employees=6&param1=getmethod">
	<button onclick="this.form.submit()">リロード</button>
	<input type="hidden" name="reloaded" value="true" />
	<input type="hidden" name="selectedEmpno" value="8" />
	<input type="hidden" name="searchKeyword" value="This is an apple." />
	<input type="hidden" name="employees" value="1" />
	<input type="hidden" name="employees" value="3" />
	<input type="hidden" name="employees" value="5" />
	<input type="hidden" name="employees" value="7" />
	<input type="hidden" name="param2" value="postmethod" />
</form>
<%
End Sub

Sub TestMain()
	Response.Write "<!DOCTYPE html>"
	Response.Write "<html><body>"

	DispReloadPage
	If Request("reloaded").Count = 0 Then 
		Exit Sub
	End If

	Dim test
	Set test = New UnitTest
	
	test.Add("LoadQueryStringTest")
	test.Add("LoadFormTest")
	test.Add("LoadRequestTest")
	test.Add("JSONValueTest")
	test.Add("css_displayTest")
	test.Add("css_visibilityTest")
	test.Add("write_ifTest")
	test.Add("html_optionsTest")

	test.RunTest
	test.ResultHtml

	Response.Write "</body></html>"

End Sub

Call TestMain()

'リクエストモデル
Class RequestModel
	'選択中の社員No
	Public selectedEmpno

	'社員リスト
	Public employees

	'検索キーワード
	Public searchKeyword

	'パラメータ1
	Public param1

	'パラメータ2
	Public param2

	Sub Class_Initialize
		'Long
		selectedEmpno = CLng(0)

		'ArrayList(Of Long)
		Set employees = New ArrayList
		employees.ItemType = vbLong

		'String
		searchKeyword = ""

		param1 = ""
		param2 = ""
	End Sub

End Class
Function RequestModelPropNames
	RequestModelPropNames = Array("selectedEmpno", "employees", "searchKeyword", "param1", "param2")
End Function


Sub LoadQueryStringTest
	'QueryStringの取得(GETメソッド)
	Dim req
	Set req = New RequestModel
	LoadQueryString(req)

	AssertEquals req.selectedEmpno, 5
	AssertEquals req.searchKeyword, "apple123"
	AssertEquals req.employees.Count, 3
	AssertEquals req.employees.Item(0), 2
	AssertEquals req.employees.Item(1), 4
	AssertEquals req.employees.Item(2), 6
	AssertEquals req.param1, "getmethod"
	AssertEquals req.param2, ""

End Sub

Sub LoadFormTest
	'Formの取得(POSTメソッド)
	Dim req
	Set req = New RequestModel
	LoadForm(req)

	AssertEquals req.selectedEmpno, 8
	AssertEquals req.searchKeyword, "This is an apple."
	AssertEquals req.employees.Count, 4
	AssertEquals req.employees.Item(0), 1
	AssertEquals req.employees.Item(1), 3
	AssertEquals req.employees.Item(2), 5
	AssertEquals req.employees.Item(3), 7
	AssertEquals req.param1, ""
	AssertEquals req.param2, "postmethod"

End Sub

Sub LoadRequestTest
	'Requestの取得(QueryStringを優先)
	Dim req
	Set req = New RequestModel
	LoadRequest(req)

	AssertEquals req.selectedEmpno, 5
	AssertEquals req.searchKeyword, "apple123"
	AssertEquals req.employees.Count, 3
	AssertEquals req.employees.Item(0), 2
	AssertEquals req.employees.Item(1), 4
	AssertEquals req.employees.Item(2), 6
	AssertEquals req.param1, "getmethod"
	AssertEquals req.param2, "postmethod"

End Sub

'ツリーノードクラス
Class TreeNode
	'ノードNo
	Public nodeno

	'ノード名
	Public nodename

	'子ノードリスト
	Public childNodes

	'親ノード
	Public nodeData

	Sub Class_Initialize
		nodeno = CLng(0)

		nodename = ""

		Set childNodes = New ArrayList
		childNodes.ItemClassName = "TreeNode"

		Set nodeData = Nothing

	End Sub

	Public Default Property Get Constructor(nodeno, nodename)
		Me.nodeno = nodeno
		Me.nodename = nodename
		Set Constructor = Me
	End Property

	Public Sub AddChild(child)
		childNodes.Add child
	End Sub

End Class
Function TreeNodePropNames
	TreeNodePropNames = Array("nodeno", "nodename", "childNodes", "nodeData")
End Function

Sub JSONValueTest

	Dim node
	Set node = (New TreeNode)(1, "node1")

	AssertEquals JSONValue(node), "{""nodeno"":1, ""nodename"":""node1"", ""childNodes"":[], ""nodeData"":null}"

	node.AddChild((New TreeNode)(2, "node2"))
	node.AddChild((New TreeNode)(3, "node3"))

	AssertEquals JSONValue(node), "{""nodeno"":1, ""nodename"":""node1"", ""childNodes"":[{""nodeno"":2, ""nodename"":""node2"", ""childNodes"":[], ""nodeData"":null}, {""nodeno"":3, ""nodename"":""node3"", ""childNodes"":[], ""nodeData"":null}], ""nodeData"":null}"

	Set node = (New TreeNode)(1, "node1")
	node.nodeData = #2021/12/31 23:05:06#
	AssertEquals JSONValue(node), "{""nodeno"":1, ""nodename"":""node1"", ""childNodes"":[], ""nodeData"":""2021-12-31 23:05:06""}"

	node.nodeData = 123.45678
	AssertEquals JSONValue(node), "{""nodeno"":1, ""nodename"":""node1"", ""childNodes"":[], ""nodeData"":123.45678}"

	Set node.nodeData = (New TreeNode)(2, "node2")
	AssertEquals JSONValue(node), "{""nodeno"":1, ""nodename"":""node1"", ""childNodes"":[], ""nodeData"":{""nodeno"":2, ""nodename"":""node2"", ""childNodes"":[], ""nodeData"":null}}"

	node.nodeData = "漢字"
	AssertEquals JSONValue(node), "{""nodeno"":1, ""nodename"":""node1"", ""childNodes"":[], ""nodeData"":""\u6F22\u5B57""}"

	node.nodeData = """abc"",""123"""
	AssertEquals JSONValue(node), "{""nodeno"":1, ""nodename"":""node1"", ""childNodes"":[], ""nodeData"":""\""abc\"",\""123\""""}"

End Sub

Sub css_displayTest
	AssertEquals css_display(True), ""
	AssertEquals css_display(False), "display:none;"
End Sub

Sub css_visibilityTest
	AssertEquals css_visibility(True), "visibility:visible;"
	AssertEquals css_visibility(False), "visibility:hidden;"
End Sub

Sub write_ifTest
	AssertEquals write_if("active", True), "active"
	AssertEquals write_if("active", False), ""
End Sub

Sub html_optionsTest
	Dim arrValues, arrCaptions, selectedValue
	arrValues = Array("1", "2", "3")
	arrCaptions = Array("田中", "佐藤", "斎藤")
	selectedValue = "2"

	AssertEquals html_options(arrValues, arrCaptions, selectedValue), "<option value=""1"">田中</option><option value=""2"" selected>佐藤</option><option value=""3"">斎藤</option>"
End Sub

%>
<%
'--------------------------------------------------------------------
' ASPVBSLib Test
' Imports aspvbs-core
'--------------------------------------------------------------------

'==========================================================
' ユニットテスト用ライブラリ
'   ※AssertEqualsしかない簡易的なものです。
'==========================================================

Dim UnitTest_Position
UnitTest_Position = 1
Class UnitTest
	Public Asserts

	Sub Class_Initialize
		Set Asserts = New ArrayList
		Position = 0
	End Sub

	Public Sub Add(funcName)
		Dim assert 
		Set assert = New AssertInfo
		assert.FuncName = funcName
		Asserts.Add assert
	End Sub

	Public Sub RunTest()
		Dim assert
		For Each assert In Asserts.Items
			UnitTest_Position = 1
			On Error Resume Next
			Execute(assert.FuncName)
			assert.ErrNumber = Err.Number
			assert.ErrPosition = UnitTest_Position
			assert.ErrSource = Err.Source
			assert.ErrMessage = Err.Description
			On Error Goto 0
		Next
	End Sub

	Public Sub ResultHtml()
%>
<style>
table.testresults {
  border-collapse: collapse;
  border: 1px solid #ccc;
  line-height: 1.5;
}

table.testresults th {
  padding: 10px;
  font-weight: bold;
  vertical-align: top;
}

table.testresults td {
  padding: 10px;
  vertical-align: top;
}

table.testresults tr:nth-child(even) {
  background: #d9d9d9;
}

</style>
<table class="testresults">
	<tr>
		<th>Test</th>
		<th>Result</th>
		<th>Type</th>
		<th>Position</th>
		<th>Message</th>
	</tr>
	<% Dim assert:	For Each assert In Me.Asserts.Items %>
	<tr>
		<td><%=assert.FuncName%></td>
		<td><%=IIf( assert.ErrNumber = 0, "OK", "NG")%></td>
		<td><%=assert.ErrSource%></td>
		<td><%=assert.ErrPosition%></td>
		<td><%=Server.HTMLEncode(assert.ErrMessage)%></td>
	</tr>
	<% Next %> 
</table>
<%
	End Sub
End Class

Class AssertInfo
	Public FuncName
	Public ErrNumber
	Public ErrPosition
	Public ErrSource
	Public ErrMessage
End Class


Sub AssertEquals( test, v )
	Dim isOk
	isOk = True

	If IsObject(test) And IsObject(v) Then
		If test Is v Then
		Else
			isOk = False
		End If
	ElseIf Not IsObject(test) And Not IsObject(v) Then
			If IsNull(test) And IsNull(v) Then
				'どちらもNullならばOk
			ElseIf IsNull(test) Or IsNull(v) Then
				'どちらかだけがNullならばNG
				isOk = False
			ElseIf test <> v Then
				'どちらもNullじゃない状態で比較して不一致ならばNG(Nullが混じっている状態で比較すると、この条件が常に満たされなくなる)
				isOk = False
			End If
		Else
			isOk = False
		End If

		If Err.Number <> 0 Then
			Err.Raise vbObjectError + 513 + 2, "AssertEquals", "AssertEqualsでエラーが発生しました。(test=" & CStr(test) & ")" 
		Exit Sub
		End If

		If Not isOk Then
			Dim testString
			If IsObject(test) Then
				testString = TypeName(test)
			Else
				testString = CStr(test)
			End If
			Err.Raise vbObjectError + 513 + 1, "AssertEquals", "結果が一致しません。(test=" & testString & ")" 
		Exit Sub
	End If

	UnitTest_Position = UnitTest_Position + 1

End Sub
%>
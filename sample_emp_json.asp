<% 
Option Explicit
%>
<!--#include file="./aspvbslib/adovbs.inc"-->
<!--#include file="./aspvbslib/aspvbslib-core.inc"-->
<!--#include file="./aspvbslib/aspvbslib-data.inc"-->
<!--#include file="./aspvbslib/aspvbslib-web.inc"-->
<%
'RequestModelクラス
Class RequestModel
    Public deptno
End Class
' RequestModelクラスのPropNames関数(LoadRequest関数用の定義)
Function RequestModelPropNames()
    RequestModelPropNames = Array("deptno")
End Function

'ViewModelクラス
Class ViewModel
    Public employees
End Class
' ViewModelクラスのPropNames関数(JSONValue関数用の定義)
Function ViewModelPropNames()
    ViewModelPropNames = Array("employees")
End Function

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
' EmpクラスのPropNames関数(JSONValue関数用の定義)
Function EmpPropNames()
    EmpPropNames = Array("empno", "empname", "deptno", "description")
End Function

Dim req, vm
Call Main

' メイン処理
Sub Main
    Set req = New RequestModel
    Call LoadRequest(req)

    Set vm = New ViewModel
    Call LoadViewModel(vm, req)

    Call RenderJSON(vm)
End Sub

' ViewModelの設定
Function LoadViewModel(vm, req)

    'データベース接続
    Dim DbConn
    Set DbConn = Server.CreateObject("ADODB.Connection")
    Dim dbfilepath :dbfilepath = Server.MapPath("aspvbslib_sample.mdb")
    dim dbconnString :dbconnString = "Driver={Microsoft Access Driver (*.mdb)}; DBQ=" & dbfilepath & ";"
    DbConn.Open dbconnString

    '社員リストの取得（部署指定）
    Dim cmd: Set cmd = Server.CreateObject("ADODB.Command")
    cmd.CommandText = "SELECT * FROM emp WHERE deptno=?"
    cmd.parameters.Append cmd.CreateParameter("@deptno", adInteger, adParamInput, , CInt(req.deptno))
    Set vm.employees = SqlQuery(DbConn, "Emp", cmd)

    DbConn.Close
    Set DbConn = Nothing

End Function

' ViewModelをJSON形式で返却
Function RenderJSON(vm)
    Response.Write JSONValue(vm)
End Function
%>
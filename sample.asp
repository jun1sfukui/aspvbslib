<% 
Option Explicit
%>
<!--#include file="./aspvbslib/adovbs.inc"-->
<!--#include file="./aspvbslib/aspvbslib-core.inc"-->
<!--#include file="./aspvbslib/aspvbslib-data.inc"-->
<!--#include file="./aspvbslib/aspvbslib-web.inc"-->
<%
' ViewModelクラス
Class ViewModel
    'タイトル
    Public title
    '部署リスト
    Public departments
    '社員リスト
    Public employees
    '選択中の社員番号
    Public empno
    '選択中の部署番号
    public deptno
    '選択中の社員明細
    Public empDetail
End Class

'RequestModelクラス
Class RequestModel
    Public deptno
    Public empno
End Class
' RequestModelクラスのPropNames関数(LoadRequest関数用の定義)
Function RequestModelPropNames
    RequestModelPropNames = Array("deptno", "empno")
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

'部署クラス
Class Dept
    '部署番号
    Public deptno
    '部署名
    Public deptname
End Class

Dim vm
Call Main

Sub Main
    Set vm = New ViewModel

    Call LoadViewModel(vm)

    Call RenderHTML(vm)
End Sub

Function LoadViewModel(vm)
    'タイトルの設定
    vm.title = "ArrayListとSqlQueryを使ったサンプル"

    'リクエストの取得
    Dim req: Set req = New RequestModel
    LoadRequest(req)

    vm.deptno = req.deptno
    If Len(vm.deptno) > 0 And ChkNum(vm.deptno) Then
        vm.deptno = CInt(vm.deptno)
    Else
        vm.deptno = ""
    End If

    vm.empno = req.empno
    If Len(vm.empno) > 0 And ChkNum(vm.empno) Then
        vm.empno = CInt(vm.empno)
    Else
        vm.empno = ""
    End If

    'データベース接続
    Dim DbConn
    Set DbConn = Server.CreateObject("ADODB.Connection")
    Dim dbfilepath :dbfilepath = Server.MapPath("aspvbslib_sample.mdb")
    dim dbconnString :dbconnString = "Driver={Microsoft Access Driver (*.mdb)}; DBQ=" & dbfilepath & ";"
    DbConn.Open dbconnString

    '部署リストの取得
    Set vm.departments = SqlQuery(DbConn, "Dept", "SELECT * FROM dept")
    if IsNullOrEmpty(vm.deptno) And vm.departments.Count > 0 Then
        vm.deptno = vm.departments(0).deptno
    End If
    
    '社員リストの取得（部署指定）
    Dim cmd: Set cmd = Server.CreateObject("ADODB.Command")
    cmd.CommandText = "SELECT * FROM emp WHERE deptno=?"
    cmd.parameters.Append cmd.CreateParameter("@deptno", adInteger, adParamInput, , vm.deptno)
    Set vm.employees = SqlQuery(DbConn, "Emp", cmd)

    '選択中のempnoが存在しなければ、現在のリストの先頭を選択する。
    If vm.employees.Where("item.empno = p", vm.empno).Count = 0 Then
        If vm.employees.Count > 0 Then
            vm.empno = vm.employees(0).empno
        Else
            vm.empno = ""
        End If
    End If

    '選択中の社員情報明細を設定
    If vm.empno = "" Then
        '選択されていなければ初期化
        Set vm.empDetail = New Emp
    Else
        Set vm.empDetail = vm.employees.FindFirst("item.empno = p", vm.empno)(0)
    End If

    DbConn.Close
    Set DbConn = Nothing

End Function

Function RenderHTML(vm)
%>
<!doctype html>
<html lang="jp">
<head>
    <meta charset="shift-jis">
    <title><%=vm.title%></title>
    <meta name="description" content="Sample 01">
    <!-- Bootstrap CSS -->
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.1.3/css/bootstrap.min.css" integrity="sha384-MCw98/SFnGE8fJT3GXwEOngsV7Zt27NXFoaoApmYm81iuXoPkFOJwJ8ERdknLPMO" crossorigin="anonymous">  
    <style>
    </style>
</head>
<body>
<div class="container">
    <h4><%=vm.title%></h4>
    <form method="get">
        <div class="form-group">
            <label for="deptno">部署</label>
            <select class="form-control" name="deptno" id="deptno" onchange="this.form.submit()">
                <% Dim dept: For Each dept In vm.departments.Items %>
                    <option value="<%=dept.deptno%>" <%=write_if("selected", dept.deptno=vm.deptno)%> data="<%=dept.deptno & ":" & vm.deptno & "(" & dept.deptno=vm.deptno & ")"%>"><%=dept.deptname%></option>
                <% Next %>
            </select>
        </div>

        <div class="form-group">
            <label for="empno">社員</label>
            <select class="form-control" name="empno" id="empno" onchange="this.form.submit()">
                <% Dim emp: For Each emp In vm.employees.Items %>
                    <option value="<%=emp.empno%>" <%=write_if("selected", emp.empno=vm.empno)%>><%=emp.empname%></option>
                <% Next %>
            </select>
        </div>

        <div class="form-group">
            <label>詳細</label>
            <div class="card p-3" style="width:18rem">
                <h5 class="card-title"><%=vm.empDetail.empname%></h5>
                <p class="card-text"><%=vm.empDetail.description%>
                </p>
            </div>
        </div>
    </form>
</div>
</body>
</html>
<%
End Function
%>
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
    '選択中の部署リスト
    Public selectedDepartments
    '選択中の社員no
    Public selectedEmp

End Class

'RequestModelクラス
Class RequestModel
    '選択中の部署Noのリスト
    Public selectedDepartments

    '選択中の社員No
    Public selectedEmp

    Sub Class_Initialize
        Set selectedDepartments = New ArrayList
        selectedDEpartments.ItemType = vbInteger    'リスト要素の型をIntegrで指定する（LoadRequestでこの型に自動変換される）

        selectedEmp = CLng(0)   'プロパティの型をLongで指定する（LoadRequestでこの型に自動変換される）
    End Sub
End Class
' RequestModelクラスのPropNames関数(LoadRequest関数用の定義)
Function RequestModelPropNames
    RequestModelPropNames = Array("selectedDepartments", "selectedEmp")
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
    vm.title = "複数のcheckboxの値を受け取るサンプル"

    'リクエストの取得
    Dim req: Set req = New RequestModel
    LoadForm(req)

    Set vm.selectedDepartments = req.selectedDepartments
    vm.selectedEmp = req.selectedEmp

    'データベース接続
    Dim DbConn
    Set DbConn = Server.CreateObject("ADODB.Connection")
    Dim dbfilepath :dbfilepath = Server.MapPath("aspvbslib_sample.mdb")
    dim dbconnString :dbconnString = "Driver={Microsoft Access Driver (*.mdb)}; DBQ=" & dbfilepath & ";"
    DbConn.Open dbconnString

    '部署リストの取得
    Set vm.departments = SqlQuery(DbConn, "Dept", "SELECT * FROM dept")
    
    '社員リストの取得(選択中の部署リスト（複数）に該当する社員のみ)
    Set vm.employees = SqlQuery(DbConn, "Emp", "SELECT * FROM emp").Where("InArray(item.deptno, p)", req.selectedDepartments.Items)

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
    <form name="frm" method="post">
        <div class="btn-group btn-group-toggle" data-toggle="buttons">
            <% Dim dept: For Each dept In vm.departments.Items %>
            <label class="btn btn-info <%=write_if("active", InArray(dept.deptno, vm.selectedDepartments.Items))%>">
                <input type="checkbox" name="selectedDepartments" value="<%=dept.deptno%>" <%=write_if("checked", InArray(dept.deptno, vm.selectedDepartments.Items))%> onchange="this.form.submit()" class="btn-check" autocomplete="off">
                <%=dept.deptname%>
            </label>
            <% Next %>
        </div>
        <input type="hidden" name="selectedEmp" value="" />

        <table class="table" style="width:500px">
            <thead style="<%=css_display(vm.employees.Count > 0)%>">
                <tr>
                    <th>empno</th>
                    <th>empname</th>
                    <th>deptname</th>
                </tr>
            </thead>
            <tbody>
                <% Dim emp: For Each emp In vm.employees.Items %>
                <tr class="<%=write_if("table-info", vm.selectedEmp = emp.empno)%>" onclick="document.frm.selectedEmp.value=<%=emp.empno%>; document.frm.submit()">
                    <td><%=emp.empno %></td>
                    <td><%=emp.empname %></td>
                    <td><%=vm.departments.FindFirst("item.deptno = p", emp.deptno).Map("item.deptname", Nothing).FirstOrDefault("") %></td>
                </tr>
                <% Next %>
            </tbody>
        </table>

    </form>

</div>
</body>
</html>
<%
End Function
%>
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
    Public deptno

    '空の社員詳細情報
    Public emptyEmp

End Class
' ViewModelクラスのPropNames関数(JSONValue関数用の定義)
Function ViewModelPropNames()
    ViewModelPropNames = Array("title", "departments", "employees", "empno", "deptno", "emptyEmp")
End Function

'部署クラス
Class Dept
    '部署番号
    Public deptno
    '部署名
    Public deptname
End Class
' DeptクラスのPropNames関数(JSONValue関数用の定義)
Function DeptPropNames()
    DeptPropNames = Array("deptno", "deptname")
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

Dim vm
Call Main

Sub Main
    Set vm = New ViewModel

    Call LoadViewModel(vm)

    Call RenderHTML(vm)
End Sub

Function LoadViewModel(vm)
    'タイトルの設定
    vm.title = "ArrayListとSqlQueryを使ったサンプル(Vue+Ajax版)"

    'リクエストの取得
    vm.deptno = NullValue(Request("deptno"), "")
    If Len(vm.deptno) > 0 And ChkNum(vm.deptno) Then
        vm.deptno = CInt(vm.deptno)
    Else
        vm.deptno = ""
    End If

    vm.empno = NullValue(Request("empno"), "")
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

    '社員リストの初期化
    Set vm.employees = New ArrayList

    '空の社員詳細情報の初期化
    Set vm.emptyEmp = New Emp

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
        [v-cloak] {
            display: none;
        }
    </style>
</head>
<body>
<div id="app" class="container" v-cloak>
    <h4>{{ title }}</h4>
    <form>
        <div class="form-group">
            <label for="deptno">部署</label>
            <select class="form-control" name="deptno" id="deptno" v-model="deptno" v-on:change="loadEmployees">
                <option v-for="dept in departments" v-bind:value="dept.deptno">{{dept.deptname}}</option>
            </select>
        </div>

        <div class="form-group">
            <label for="empno">社員</label>
            <select class="form-control" name="empno" id="empno" v-model="empno">
                <option v-for="emp in employees" v-bind:value="emp.empno">{{emp.empname}}</option>
            </select>
        </div>

        <div class="form-group">
            <label>詳細</label>
            <div class="card p-3" style="width:18rem">
                <h5 class="card-title">{{empDetail.empname}}</h5>
                <p class="card-text">{{empDetail.description}}
                </p>
            </div>
        </div>
    </form>
</div>

<script src="https://unpkg.com/vue@next"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/axios/0.18.0/axios.js"></script>
<script type="text/javascript">
    
    var baseVM = {
        data : function() {
            return {
            }
        },
        methods: {
        }
    }
    
    var model = <%=JSONValue(vm)%>;
    
    var vm = Vue.createApp({
        mixins: [baseVM],
        data() {
            return model;
        },
        computed : {
            empDetail(){
                let emp = this.employees.length > 0 ? this.employees.filter(item => item.empno == vm.empno )[0] : null;
                if (!emp) emp = this.emptyEmp;
                return emp;
            }
        },
        methods : {
            loadEmployees(event) {
                axios
                    .get("sample_emp_json.asp?deptno=" + this.deptno)
                    .then(function(response){
                        vm.empno = response.data.employees ? response.data.employees[0].empno : null;
                        vm.employees = response.data.employees;
                    })
                    .catch(function(error){
                        alert()
                    });
            },
        },

        mounted(){
            this.loadEmployees();
        },

    }).mount("#app");

</script>

</body>
</html>
<%
End Function
%>
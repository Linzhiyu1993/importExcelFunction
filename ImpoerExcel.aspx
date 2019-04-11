<%@ Page Language="C#" AutoEventWireup="true" CodeFile="ImpoerExcel.aspx.cs" Inherits="ImpoerExcel" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>import excel</title>
    <a class="easyui-linkbutton" iconcls="icon-import" plain="true" onclick="EquInfo_bt_imp()">excel导入</a>
    <div id="importExcel" class="easyui-dialog" data-options="
	        modal:true,
	        maximizable:false,
	        minimizable:false,
            resizable:false,
	        collapsible:false,
            closable:true,
            closed:true, 
            title:'导入excel',
            width: 660,
            height: 560,
            buttons: [
            { text: '关闭窗口', iconCls: 'icon-clear',handler:function(){imExcelClos();}}],">
    <div style="margin-top: 5px">
            <span>上&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;传：</span>
            <span class="input">
                <input type="file" id="upfile" name="upfil" onchange="getLocalExcel()" placeholder="" style="width:340px" data-options="buttonText:'选择文件' 
                	accept:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'"/>
            </span>
            <button id="saveButton" onclick="importExp();">导入</button>
            <span>格式：.xlxs</span>
        </div>
    </div> 
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
    </div>
    </form>
    <script>
        $('#importExcel').dialog('open');
        function importExp()
        {
            var formData = new FormData();
            var name = $("#upfile").val();
            formData.append("file", $("#upfile")[0].files[0]);
            formData.append("name", name);
            if (isNull(formData)) {
                $.messager.alert("系统提示", "未选择导入文件！", "warning");
                return;
            }
            $.ajax({
                type: "POST",
                url: "ashx/importExce.ashx",
                data: formData,
                processData: false,  // 不处理数据
                contentType: false,
                success: function (data) {
                    console.log(data);
                    if (data == '{"errcode":0,"message":"import excel success"}') {
                        00
                        $.messager.alert("success", "Excel import success!");
                        $("#importExcel").dialog("close");
                        loadDatagrid();
                    } else {
                        $.messager.alert("system error", "please confirm your data is correct！", "warning");
                    }
                },
                error: function () {
                    $.messager.alert(JSON.stringify(data.error));
                }
            })
        }
    </script>
</body>
</html>

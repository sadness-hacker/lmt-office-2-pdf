<%--
  Created by IntelliJ IDEA.
  User: ducx
  Date: 2019/7/17
  Time: 13:37
  To change this template use File | Settings | File Templates.
--%>
<%@ page contentType="text/html;charset=UTF-8" %>
<html>
<head>
    <title>pdf预览</title>
    <script src="http://code.jquery.com/jquery-1.12.4.min.js" type="text/javascript"></script>
    <script type="text/javascript">
        $(function () {
            var purl = 'download?file=' + encodeURI('${fileName}');//要展示的文件路径
            var flag = false;
            // 下面代码都是处理IE浏览器的情况
            if (window.ActiveXObject || "ActiveXObject" in window) {
                //判断是否为IE浏览器，"ActiveXObject" in window判断是否为IE11
                //判断是否安装了adobe Reader
                for (var x = 2; x < 10; x++) {
                    try {
                        var oAcro = eval("new ActiveXObject('PDF.PdfCtrl." + x + "');");
                        if (oAcro) {
                            flag = true;
                        }
                    } catch (e) {
                        flag = false;
                    }
                }
                try {
                    var oAcro4 = new ActiveXObject('PDF.PdfCtrl.1');
                    if (oAcro4) {
                        flag = true;
                    }
                } catch (e) {
                    flag = false;
                }
                try {
                    var oAcro7 = new ActiveXObject('AcroPDF.PDF.1');
                    if (oAcro7) {
                        flag = true;
                    }
                } catch (e) {
                    flag = false;
                }
                if (flag) {//支持
                    pdfShow(purl);//调用显示的方法
                } else {//不支持
                    $("#pdfContent").append("对不起,您还没有安装PDF阅读器软件呢,为了方便预览PDF文档,请选择安装！<a href='http://ardownload.adobe.com/pub/adobe/reader/win/9.x/9.3/chs/AdbeRdr930_zh_CN.exe'>下载PDF阅读器</a><br/><a href='download?file=${fileName}'>直接下载PDF文件</a>");
                    alert("对不起,您还没有安装PDF阅读器软件呢,为了方便预览PDF文档,请选择安装！");
                    location = "http://ardownload.adobe.com/pub/adobe/reader/win/9.x/9.3/chs/AdbeRdr930_zh_CN.exe";
                }

            } else {
                pdfShow(purl);//调用显示的方法
            }
        });


        //显示文件方法，就是将文件展示到div中
        function pdfShow(url) {
            // $("#pdfContent").append('<iframe style="height:100%;width:100%;" src="' + url + '"></iframe>');
            window.location.href = url;
        }
    </script>
</head>
<body>
<div id="pdfContent"></div>
</body>
</html>

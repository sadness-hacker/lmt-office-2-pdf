#Word/Excel/PPT转PDF
### 1.环境要求
```
主意application.properties文件里的配置

基于Jacob实现，需要运行在Windows下，Windows服务器需要安装Office 2007套件，JDK1.8以上，Tomcat 8以上
Tomcat需要配置java.library.path
拷贝WEB-INF/lib下的jar包到Tomcat根目录下的ext-lib目录，修改catalina文件,添加：
set "JAVA_OPTS=%JAVA_OPTS% -Djava.library.path=%CATALINA_HOME%/ext-lib"

Linux下需要安装LibreOffice-6.2.5:
https://mirrors.tuna.tsinghua.edu.cn/libreoffice/libreoffice/stable/6.2.5/rpm/x86_64/

LibreOffice_6.2.5_Linux_x86-64_rpm_langpack_zh-CN.tar.gz
LibreOffice_6.2.5_Linux_x86-64_rpm_sdk.tar.gz
LibreOffice_6.2.5_Linux_x86-64_rpm.tar.gz


需要安装simsum.ttc字体:
http://www.pc6.com/softview/SoftView_100415.html

# cp simsun.ttc /usr/share/fonts
# fc-cache -fv
```

### 2.文件上传接口/upload
```
get请求可以获取上传页面进行测试
post请求进行文件上传并转换,只包含文件字段即可：如<input type="file" name="file" />
```
#### 返回值
```
{
  "code":1,
  "msg":"success",
  "data":{
    "pdfFileName":"20190717191407-龙湖统一认证IAM项目.pdf",
    "pdfFileUrl":"http://localhost:9999/jacom/download?file=20190717191407-%E9%BE%99%E6%B9%96%E7%BB%9F%E4%B8%80%E8%AE%A4%E8%AF%81IAM%E9%A1%B9%E7%9B%AE.pdf",
    "previewUrl":"http://localhost:9999/jacom/preview?file=20190717191407-%E9%BE%99%E6%B9%96%E7%BB%9F%E4%B8%80%E8%AE%A4%E8%AF%81IAM%E9%A1%B9%E7%9B%AE.pdf"
   }
}
```

#### 返回值说明
```
code  返回状态码 1-成功，其他值失败
msg   返回信息
data  返回数据字段
  pdfFileName 存储的pdf文件名
  pdfFileUrl  pdf文件下载路径
  previewUrl  pdf预览路径
```

### 3.文件下载接口/download
```
get请求：/download?file=xxxx.pdf
注意：中文文件名时请进行urlencode编码
```

### 4.pdf文件预览接口/preview
```
get请求:/preview?file=xxx.pdf
注意：中文文件名时请进行urlencode编码
IE浏览器用户未安装pdf插件时会提示用户下载并安装pdf阅读器,Chrome/Firefox浏览器不用安装
```




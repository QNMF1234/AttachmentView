powerapps/dynamics365适用的注释预览/批量下载组件

原生预览支持的文件类型:图像,zip,音频,pdf

使用浏览器预览支持:音频,视频,图像,pdf,文本,xml,json等,理论上只需要浏览器支持打开的文件类型,均可预览

使用方法:

  1.导入解决方案zip
  
  2.在经典窗体选择文本类型字段->点击属性->控件->添加控件,选择AttachmentView
  
   ![Image text](https://github.com/QNMF1234/AttachmentView/blob/master/%E6%95%99%E7%A8%8B%E5%9B%BE%E5%83%8F/1.png)
  
   属性1:为绑定的字段
  
   属性2:指示是否优先(True)使用浏览器来查看附件,当浏览器有不支持的类型附件再尝试内置库预览
  
   属性3:填写预览pdf的必须配置文件,需要手动上传至CRM填写web资源url
  
  3.保存后发布,建议强制刷新窗体页浏览器缓存
  
  ![Image text](https://github.com/QNMF1234/AttachmentView/blob/master/%E6%95%99%E7%A8%8B%E5%9B%BE%E5%83%8F/2.png)
  
  4.通过顶部搜索栏筛选注释,支持使用|分割多个关键词
  
  ![Image text](https://github.com/QNMF1234/AttachmentView/blob/master/%E6%95%99%E7%A8%8B%E5%9B%BE%E5%83%8F/3.png)
  
  5.点击record左边按钮选中为下载状态
  
  6.通过下载按钮下载一个或多个注释文件,输出文件为zip
  
  ![Image text](https://github.com/QNMF1234/AttachmentView/blob/master/%E6%95%99%E7%A8%8B%E5%9B%BE%E5%83%8F/4.png)

使用的库:pdfjs-dist,jszip,file-saver

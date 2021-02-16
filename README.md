# word 转 pdf

> 注：仅适用于 windows 平台，同时，需要至少 java8 运行环境，至少要有 Microsoft Office 2013 以上。

[下载所需依赖（github）](https://github.com/freemansoft/jacob-project/releases/tag/Root_B-1_20)

* jacob.jar
* jacob-1.20-x64.dll

# 构建

将 jacob-1.20-x64.dll 复制到 jdk的jre下bin中

# 运行

```
java -jar xxx.jar <xxx.docx> <xxx.pdf>
```
xxx.docx：需要进行转换的 word 文档的绝对路径
xxx.pdf：保存转换后的 pdf 的保存路径（绝对路径）

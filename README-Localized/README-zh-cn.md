# Word  JavaScript Redact 外接程序

了解如何创建对文本进行搜索、突出显示和修订的外接程序。    

## 目录
* [更改历史记录](#change-history)
* [重要信息注释](#important-note)
* [先决条件](#prerequisites)
* [配置项目](#configure-the-project)
* [运行项目](#run-the-project)
* [问题和意见](#questions-and-comments)
* [其他资源](#additional-resources)

## 更改历史记录

2016 年 7 月 25 日：
* 初始示例版本。

##重要信息注释

此示例通过使用星号替换每个字母来修订文本。因此不保留布局。确定每个字母的字体、大小或格式时并未进行测量和计算。

## 先决条件

* Word 2016 for Windows，内部版本 16.0.6912.1000 或更高版本。
* [Node 与 npm](https://nodejs.org/en/)
* [Git Bash](https://git-scm.com/downloads) - 应使用较高的内部版本，因为早期内部版本在生成证书时可能会显示错误。

## 配置项目

在此项目的根目录处从你的 Bash shell 中运行以下命令：

1. 将此存储库克隆到你的本地计算机中。
2. ```npm install``` 可安装 package.json 中的所有依赖项。
3. ```bash gen-cert.sh``` 可创建运行此示例所需的证书。 
* 然后在本地计算机的存储库中，双击 ca.crt，然后选择“**安装证书**”。 
* 选择“**本地计算机**”，然后选择“**下一步**”以继续。 
* 选择“**将所有证书放入下列存储**”，然后选择“**浏览**”。  
* 选择“**受信任的根证书颁发机构**”，然后选择“**确定**”。 
* 选择“**下一步**”，然后选择“**完成**”，现在你的证书颁发机构证书已添加到你的证书存储中。
4. ```npm start``` 可启动节点 Web 服务器服务。

现在，需要让 Microsoft Word 知道在哪里可以找到该外接程序。

1. 创建网络共享，或[将文件夹共享到网络](https://technet.microsoft.com/zh-cn/library/cc770880.aspx)，并将 [word-add-in-javascript-speckit-manifest.xml](word-add-in-javascript-speckit-manifest.xml) 清单文件放入该文件夹中。
3. 启动 Word，然后打开一个文档。
4. 选择**文件**选项卡，然后选择**选项**。
5. 选择**信任中心**，然后选择**信任中心设置**按钮。
6. 选择**受信任的外接程序目录**。
7. 在“**目录 URL**”字段中，输入包含 word-add-in-js-redact-manifest.xml 的文件夹共享的网络路径，然后选择“**添加目录**”。
8. 选中**显示在菜单中**复选框，然后单击**确定**。
9. 随后会出现一条消息，告知你下次启动 Microsoft Office 时将应用你的设置。关闭并重新启动 Word。

## 运行项目

1. 打开一个 Word 文档。
2. 在 Word 2016 中的**插入**选项卡上，选择**我的外接程序**。
3. 选择“**共享文件夹**”选项卡。
4. 选择“**Word Redact 外接程序**”，然后选择“**确定**”。
5. 如果你的 Word 版本支持外接程序命令，UI 将通知你加载了外接程序。

### 功能区用户界面

在功能区上：
* 选择“**审查**”选项卡，然后选择“**显示 修订任务窗格**”以启动任务窗格。

 > 注意：如果你的 Word 版本不支持外接程序命令，则外接程序将在任务窗格中加载。

### 任务窗格 UI

在任务窗格上，你可以：
* 通过在文本框中输入词并选择“**搜索和突出显示**”按钮搜索和突出显示文档中的文本。
  
> 注意：搜索区分大小写。若要撤消操作，请在键盘上按 Ctrl+Z。

* 通过在文本框中输入词并选择“**修订**”按钮修订文档中的文本。
  
> 注意：修订区分大小写。   

* 通过选择“**修订所选文本**”按钮修订文档中所选的文本。
  
> 注意：修订区分大小写。       
  
### 外接程序命令

此示例演示如何在功能区中创建外接程序命令。外接程序命令声明位于 word-add-in-js-redact-manifest.xml 中。 

## 问题和意见

我们希望倾听你对 Word Redact 示例的相关反馈。你可以在该存储库中的“*问题*”部分将你的反馈发送给我们。

与 Microsoft Office 365 开发相关的问题一般应发布到 [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API)。确保你的问题使用了 [office-js] 和 [API] 标记。

## 其他资源

* [Office 外接程序文档](https://msdn.microsoft.com/zh-cn/library/office/jj220060.aspx)
* [Office 开发人员中心](http://dev.office.com/)
* [Office 365 API 初学者项目和代码示例](http://msdn.microsoft.com/en-us/office/office365/howto/starter-projects-and-code-samples)

## 版权
版权所有 © 2016 Microsoft Corporation。保留所有权利。



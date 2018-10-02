# <a name="word--javascript-redact-add-in"></a>Word  JavaScript Redact 外接程序

了解如何创建对文本进行搜索、突出显示和修订的外接程序。    

## <a name="table-of-contents"></a>目录
* [修订记录](#change-history)
* [重要说明](#important-note)
* [先决条件](#prerequisites)
* [配置项目](#configure-the-project)
* [运行项目](#run-the-project)
* [问题和意见](#questions-and-comments)
* [其他资源](#additional-resources)

## <a name="change-history"></a>修订记录

2016 年 7 月 25 日：
* 初始示例版本。

## <a name="important-note"></a>重要说明

此示例通过使用星号替换每个字母来修订文本。因此不保留布局。确定每个字母的字体、大小或格式时并未进行测量和计算。

## <a name="prerequisites"></a>先决条件

* Word 2016 for Windows（生成号 16.0.6912.1000 或更高版本）。
* [Node 和 npm](https://nodejs.org/en/)
* [Git Bash](https://git-scm.com/downloads) - 应使用较高的内部版本，因为较低的内部版本可能会在生成证书时显示错误。

## <a name="configure-the-project"></a>配置项目

在此项目的根目录处从你的 Bash shell 中运行以下命令：

1. 将此存储库克隆到本地计算机中。
2. ```npm install``` 可安装 package.json 中的所有依赖项。
3. ```bash gen-cert.sh``` 可创建运行此示例所需的证书。 
* 然后，在本地计算机的存储库中，双击“ca.crt”，再选择“安装证书”****。 
* 依次选择“本地计算机”****和“下一步”****，以继续操作。 
* 选择“**将所有证书放入下列存储**”，然后选择“**浏览**”。  
* 依次选择“受信任的根证书颁发机构”****和“确定”****。 
* 依次选择“下一步”****和“完成”****。此时，证书颁发机构证书已添加到证书存储中。
4. ```npm start``` 可启动节点 Web 服务器服务。

现在需要让 Microsoft Word 知道在哪里可以找到加载项。

1. 创建网络共享，或[将文件夹共享到网络](https://technet.microsoft.com/en-us/library/cc770880.aspx)，并将 [word-add-in-javascript-speckit-manifest.xml](word-add-in-javascript-speckit-manifest.xml) 清单文件放入其中。
3. 启动 Word，然后打开一个文档。
4. 选择**文件**选项卡，然后选择**选项**。
5. 选择**信任中心**，然后选择**信任中心设置**按钮。
6. 选择“受信任的加载项目录”****。
7. 在“目录 URL”****字段中，输入包含 word-add-in-js-redact-manifest.xml 的文件夹共享的网络路径，再选择“添加目录”****。
8. 选中“在菜单中显示”****复选框，再选择“确定”****。
9. 随后会出现一条消息，告知你下次启动 Microsoft Office 时将应用你的设置。关闭并重新启动 Word。

## <a name="run-the-project"></a>运行项目

1. 打开一个 Word 文档。
2. 在 Word 2016 的“插入”****选项卡中，选择“我的加载项”****。
3. 选择“共享文件夹”****选项卡。
4. 依次选择“Word Redact 加载项”****和“确定”****。
5. 如果你的 Word 版本支持外接程序命令，UI 将通知你加载了外接程序。

### <a name="ribbon-ui"></a>功能区用户界面

在功能区上：
* 依次选择“审阅”****选项卡和“显示编修任务窗格”****，启动任务窗格。

 > 注意：如果你的 Word 版本不支持外接程序命令，则外接程序将在任务窗格中加载。

### <a name="task-pane-ui"></a>任务窗格 UI

在任务窗格上，你可以：
* 在文档中搜索和突出显示文本，具体方法为在文本框中输入字词，并选择“搜索和突出显示”****按钮。
  
> 注意：搜索区分大小写。若要撤消操作，请在键盘上按 Ctrl+Z。

* 在文档中编修文本，具体方法为在文本框中输入字词，并选择“编修”****按钮。
  
> 注意：修订区分大小写。   

* 选择“编修选定文本”****按钮，编修文档中的选定文本。
  
> 注意：修订区分大小写。       
  
### <a name="add-in-commands"></a>外接程序命令

此示例演示如何在功能区中创建外接程序命令。外接程序命令声明位于 word-add-in-js-redact-manifest.xml 中。 

## <a name="questions-and-comments"></a>问题和意见

我们希望倾听你对 Word Redact 示例的相关反馈。你可以在该存储库中的“*问题*”部分将你的反馈发送给我们。

与 Microsoft Office 365 开发相关的问题一般应发布到 [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API)。确保你的问题使用了 [office-js] 和 [API] 标记。

## <a name="additional-resources"></a>其他资源

* 

  [Office 外接程序文档](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Office 开发人员中心](http://dev.office.com/)
* [Office 365 API 入门项目和代码示例](http://msdn.microsoft.com/en-us/office/office365/howto/starter-projects-and-code-samples)

## <a name="copyright"></a>版权
版权所有 (c) 2016 Microsoft Corporation。保留所有权利。



此项目已采用 [Microsoft 开放源代码行为准则](https://opensource.microsoft.com/codeofconduct/)。有关详细信息，请参阅[行为准则 FAQ](https://opensource.microsoft.com/codeofconduct/faq/)。如有其他任何问题或意见，也可联系 [opencode@microsoft.com](mailto:opencode@microsoft.com)。

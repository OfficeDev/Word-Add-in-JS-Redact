# <a name="word--javascript-redact-add-in"></a>Word  JavaScript Redact 增益集

了解如何建立可搜尋、反白及編寫文字的增益集。    

## <a name="table-of-contents"></a>目錄
* [變更歷程記錄](#change-history)
* [重要事項](#important-note)
* [必要條件](#prerequisites)
* [設定專案](#configure-the-project)
* [執行專案](#run-the-project)
* [問題和建議](#questions-and-comments)
* [其他資源](#additional-resources)

## <a name="change-history"></a>變更歷程記錄

2016 年 7 月 25 日：
* 初始範例版本。

## <a name="important-note"></a>重要事項

此編寫文字範例是以星號取代每個字母。版面配置不會完全保留。沒有進行測量和計算來判斷字型、大小或每個字母的格式。

## <a name="prerequisites"></a>必要條件

* Word 2016 for Windows，組建 16.0.6912.1000 或更新版本。
* [Node 和 npm](https://nodejs.org/en/)
* [Git Bash](https://git-scm.com/downloads) - 您應該使用較新的組建，因為較早的組建在建立憑證時會顯示錯誤。

## <a name="configure-the-project"></a>設定專案

在這個專案的根目錄，從您的 Bash 殼層執行下列命令︰

1. 複製此儲存機制到本機電腦。
2. ```npm install``` 可安裝 package.json 中的所有相依性。
3. ```bash gen-cert.sh``` 可建立執行這個範例所需的憑證。 
* 然後在您本機電腦上的儲存機制中，連按兩下 ca.crt，然後選取 [安裝憑證]****。 
* 選取 [本機電腦]****，然後選取 [下一步]**** 以繼續。 
* 選取 [將所有憑證放入以下的存放區]****，然後選取 [瀏覽]****。  
* 選取 [信任的根憑證授權]****，然後選取 [確定]****。 
* 選取 [下一步]****，然後選取 [完成]****，現在您的憑證授權單位憑證已新增至您的憑證存放區。
4. ```npm start``` 可開始節點網頁伺服器服務。

現在，您需要讓 Microsoft Word 知道哪裡可以找到此增益集。

1. 建立網路共用，或[共用網路資料夾](https://technet.microsoft.com/zh-tw/library/cc770880.aspx)，並將 [word-add-in-javascript-speckit-manifest.xml](word-add-in-javascript-speckit-manifest.xml) 資訊清單檔案放置在其中。
3. 啟動 Word 並開啟一個文件。
4. 選擇 [檔案]**** 索引標籤，然後選擇 [選項]****。
5. 選擇 [信任中心]****，然後選擇 [信任中心設定]**** 按鈕。
6. 選擇 [受信任的增益集目錄]****。
7. 在 [目錄 URL]**** 欄位中，輸入包含 word-add-in-js-redact-manifest.xml 的資料夾共用的網路路徑，然後選擇 [新增目錄]****。
8. 選取 [顯示於功能表中]**** 核取方塊，然後選擇 [確定]****。
9. 接著會顯示訊息，通知您下次啟動 Office 時就會套用您的設定。關閉並重新啟動 Microsoft Word。

## <a name="run-the-project"></a>執行專案

1. 開啟 Word 文件。
2. 在 Word 2016 的 [插入]**** 索引標籤上，選擇 [我的增益集]****。
3. 選取 [共用資料夾]**** 索引標籤。
4. 選擇 [Word Redact 增益集]****，然後選取 [確定]****。
5. 如果您的 Word 版本支援增益集命令，UI 會通知您已載入增益集。

### <a name="ribbon-ui"></a>功能區 UI

在功能區上：
* 選取 [檢閱]**** 索引標籤，然後選擇 [顯示 Redaction 工作窗格]**** 來啟動工作窗格。

 > 附註：如果您的 Word 版本不支援增益集命令，增益集會載入工作窗格。

### <a name="task-pane-ui"></a>工作窗格 UI

在工作窗格上，您可以：
* 要在文件中搜尋並反白文字，請在 [文字] 方塊中輸入文字，然後選取 [搜尋並反白]**** 按鈕。
  
> 附註：搜尋會區分大小寫。若要復原動作，請按鍵盤上的 Ctrl + Z。

* 要在文件中編寫文字，請在 [文字] 方塊中輸入文字，然後選取 [Redact]**** 按鈕。
  
> 附註：Redact 會區分大小寫。   

* 要在文件中編寫選取的文字，請選擇 [Redact 選取的文字]**** 按鈕。
  
> 附註：Redact 會區分大小寫。       
  
### <a name="add-in-commands"></a>增益集命令

這個範例會示範如何在功能區中建立增益集命令。增益集命令宣告位於 word-add-in-js-redact-manifest.xml。 

## <a name="questions-and-comments"></a>問題和建議

我們很樂於收到您對於 Word Redact 範例的意見反應。您可以在此儲存機制的*問題*區段中，將您的意見反應傳送給我們。

請在 [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API) 提出有關 Microsoft Office 365 開發的一般問題。務必以 [office-js] 和 [API] 標記您的問題。

## <a name="additional-resources"></a>其他資源

* [Office 增益集文件](https://msdn.microsoft.com/zh-tw/library/office/jj220060.aspx)
* [Office 開發人員中心](http://dev.office.com/)
* [Office 365 API 入門專案和程式碼範例](http://msdn.microsoft.com/zh-tw/office/office365/howto/starter-projects-and-code-samples)

## <a name="copyright"></a>著作權
Copyright (c) 2016 Microsoft Corporation 著作權所有，並保留一切權利。



此專案已採用 [Microsoft 開放原始碼管理辦法](https://opensource.microsoft.com/codeofconduct/)。如需詳細資訊，請參閱[管理辦法常見問題集](https://opensource.microsoft.com/codeofconduct/faq/)，如果有其他問題或意見，請連絡 [opencode@microsoft.com](mailto:opencode@microsoft.com)。

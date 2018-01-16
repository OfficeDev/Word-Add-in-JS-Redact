# <a name="word--javascript-redact-add-in"></a>Word JavaScript Redact アドイン

テキストの検索、強調表示、編集を行うアドインを作成する方法について説明します。    

## <a name="table-of-contents"></a>目次
* [変更履歴](#change-history)
* [重要なメモ](#important-note)
* [前提条件](#prerequisites)
* [プロジェクトを構成する](#configure-the-project)
* [プロジェクトを実行する](#run-the-project)
* [質問とコメント](#questions-and-comments)
* [その他のリソース](#additional-resources)

## <a name="change-history"></a>変更履歴

2016 年 7 月 25 日
* サンプルの初期バージョン。

##<a name="important-note"></a>重要なメモ

このサンプルでは、各文字をアスタリスクに置き換えてテキストを編集します。レイアウトは保持されません。各文字のフォント、サイズ、または書式設定を判別するための測定や計算は行われません。

## <a name="prerequisites"></a>前提条件

* Word 2016 for Windows (16.0.6912.1000 以降のビルド)。
* [Node と npm](https://nodejs.org/en/)
* [Git Bash](https://git-scm.com/downloads) -証明書を生成するときに、以前のビルドではエラーが発生する可能性があるため、最新のビルドを使用する必要があります。

## <a name="configure-the-project"></a>プロジェクトを構成する

このプロジェクトのルートで Bash シェルから以下のコマンドを実行します。

1. ローカル コンピューターにこのリポジトリのクローンを作成します。
2. ```npm install``` で package.json にすべての依存関係をインストールします。
3. ```bash gen-cert.sh``` でこのサンプルを実行するために必要な証明書を作成します。 
* ローカル コンピューターにあるリポジトリで、ca.crt をダブルクリックし、**[証明書のインストール]** を選択します。 
* **[ローカル コンピューター]** を選択して、**[次へ]** を選択して続行します。 
* **[証明書をすべて次のストアに配置する]** を選択してから **[参照]** を選択します。  
* **[信頼されたルート証明機関]** を選択して、**[OK]** を選択します。 
* **[次へ]**、**[完了]** の順に選択すると、証明書ストアに証明書機関の証明書が追加されます。
4. ```npm start``` でノード Web サーバー サービスを開始します。

次に、Microsoft Word がアドインを検索する場所を認識できるようにする必要があります。

1. ネットワーク共有を作成するか、[ネットワークでフォルダーを共有し](https://technet.microsoft.com/en-us/library/cc770880.aspx)、そのフォルダーに [word-add-in-javascript-speckit-manifest.xml](word-add-in-javascript-speckit-manifest.xml) マニフェスト ファイルを配置します。
3. Word を起動し、ドキュメントを開きます。
4. [**ファイル**] タブを選択し、[**オプション**] を選択します。
5. [**セキュリティ センター**] を選択し、[**セキュリティ センターの設定**] ボタンを選択します。
6. **[信頼されているアドイン カタログ]** を選択します。
7. **[カタログの URL]** フィールドに、word-add-in-js-redact-manifest.xml があるフォルダー共有へのネットワーク パスを入力して、**[カタログの追加]** を選択します。
8. **[メニューに表示する]** チェック ボックスをオンにして、**[OK]** を選択します。
9. これらの設定が Microsoft Office を次回起動したときに適用されることを示すメッセージが表示されます。Word を終了して、再起動します。

## <a name="run-the-project"></a>プロジェクトを実行する

1. Word 文書を開きます。
2. Word 2016 の**[挿入]** タブで、**[マイ アドイン]** を選択します。
3. **[共有フォルダー]** タブを選択します。
4. **[Word Redact アドイン]** を選択して、**[OK]** を選択します。
5. ご使用の Word バージョンでアドイン コマンドがサポートされている場合、UI によってアドインが読み込まれたことが通知されます。

### <a name="ribbon-ui"></a>リボン UI

リボンで、次の操作を実行します。
* **[校閲]** タブを選択し、**[Show Redaction Task Pane]** を選択して作業ウィンドウを起動します。

 > 注:アドイン コマンドが Word バージョンによってサポートされていない場合は、アドインが作業ウィンドウに読み込まれます。

### <a name="task-pane-ui"></a>作業ウィンドウの UI

作業ウィンドウで次の操作を実行できます。
* 文書内でテキストを検索して強調表示するには、テキスト ボックスに単語を入力して**[Search and Highlight]** ボタンを選択します。
  
> 注:検索では大文字と小文字が区別されます。操作を元に戻すには、キーボードの Ctrl キーを押しながら Z キーを押します。

* 文書内のテキストを編集するには、テキスト ボックスに単語を入力して**[Redact]** ボタンを選択します。
  
> 注:Redact では大文字と小文字が区別されません。   

* 文書内の選択したテキストを編集するには、**[Redact selected text]** ボタンを選択します。
  
> 注:Redact では大文字と小文字が区別されません。       
  
### <a name="add-in-commands"></a>アドイン コマンド

このサンプルでは、リボンにアドイン コマンドを作成する方法を示します。アドイン コマンドの宣言は、word-add-in-js-redact-manifest.xml にあります。 

## <a name="questions-and-comments"></a>質問とコメント

Word Redact サンプルについて、Microsoft にフィードバックをお寄せください。このリポジトリの *[問題]* セクションにフィードバックを送信できます。

Microsoft Office 365 開発全般の質問につきましては、「[スタック オーバーフロー](http://stackoverflow.com/questions/tagged/office-js+API)」に投稿してください。質問には、必ず [office-js] と [API] のタグを付けてください。

## <a name="additional-resources"></a>その他の技術情報

* 
  [Office アドインのドキュメント](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Office デベロッパー センター](http://dev.office.com/)
* [Office 365 API スタート プロジェクトとコード サンプル](http://msdn.microsoft.com/en-us/office/office365/howto/starter-projects-and-code-samples)

## <a name="copyright"></a>著作権
Copyright (c) 2016 Microsoft Corporation.All rights reserved.



このプロジェクトでは、[Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/) が採用されています。詳細については、「[Code of Conduct の FAQ](https://opensource.microsoft.com/codeofconduct/faq/)」を参照してください。また、その他の質問やコメントがあれば、[opencode@microsoft.com](mailto:opencode@microsoft.com) までお問い合わせください。

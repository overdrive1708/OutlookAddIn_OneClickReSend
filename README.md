# OutlookAddIn_OneClickReSend

## 【簡単な説明】

Outlookでメールをワンクリックで再送するためのアドインです。 

## 【開発の背景】

メールを再利用する場合、ThunderbirdとOutlookで以下のような違いがあります。  

Thunderbird：メールを一覧から右クリックして、「新しいメッセージとして編集」をクリックする。  
Outlook：メールを一覧からダブルクリックして新しい画面で開いてから、リボンの「メッセージ」タブの「移動」グループの「その他の移動アクション」をクリックし、「このメッセージを再送」をクリックする。その後送信したら、元のメールの画面を閉じる。  

使用頻度が高い操作にも関わらず、Outlookでは手順が複雑なので、ワンクリックで再送画面が開けるアドインを開発しました。

## 【必要要件】

- Microsoft .NET Framework 4.7.2
- Microsoft Visual Studio 2010 Tools for Office Runtime

不足している場合は、インストール時に自動的にインストールされます。

## 【インストール方法】

1. 「OutlookAddIn_OneClickReSend/installer/setup.exe」を実行します。
1. インストーラの指示に従ってインストールしてください。

## 【バージョンアップ方法】

一度アンインストールしてからインストールしなおしてください。

## 【アンインストール方法】

1. スタートメニューより、「設定」→「アプリ」→「アプリと機能」を開きます。
1. 「OutlookAddIn_OneClickReSend」をクリックして、「アンインストール」をクリックしてください。

## 【使用方法】

1. Outlookを起動します。
1. 再送したいメールを選択して、ホームタブのワンクリック再送にある、「再送」ボタンをクリックしてください。
1. メールの編集画面が開くので、再編集して送信します。

## 【注意事項】

再利用する項目は以下の通りです。

- To
- CC
- BCC
- Subject
- 本文の形式(テキスト形式/HTML形式)
- 本文

差出人の情報や、添付ファイル、開封確認の要求などの細かい項目は再利用されないのでご注意ください。  
また、リッチテキスト形式はサポートしていません。

## 【開発環境】
Microsoft Visual Studio Community 2019  
Version 16.8.2  

VisualStudio.16.Release/16.8.2+30717.126  
Microsoft .NET Framework  
Version 4.8.03752  

Office Developer Tools for Visual Studio   16.0.30502.00  
Microsoft Office Developer Tools for Visual Studio  

## 【ライセンス】

このプロジェクトはMITライセンスです。
詳細は [LICENSE](LICENSE) を参照してください。

## 【作者】

[overdrive1708](https://github.com/overdrive1708)

## 【変更履歴】

詳細は [CHANGELOG](CHANGELOG.md) を参照してください。
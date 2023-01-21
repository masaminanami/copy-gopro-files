# Copy-GoPro.ps1

PCに接続したGoProのMP4ファイルをローカルフォルダ（ローカルネットワーク上のフォルダ含む）と OneDrive にコピーします。
GoProのファイルは GX<サブシーケンスNo><シーケンスNo>.MP4 となっており、管理しにくいため、<日付>-GX<シーケンスNo><サブシーケンスNo>.MP4 形式の名前に変更します。

OneDrive は個人用、法人用（OneDrive for Business) の両方に対応しています。

# Installation

1. アプリをダウンロード
1. OneDrive 個人用へのアップロードを行う場合
    1. [アプリの登録](https://learn.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)を参考にアプリを登録
        1. この時に「個人用アカウントからのアプリケーション」タブを選択した状態で「＋新規登録」をクリックして登録を開始する必要があるようです
        1. さらに「このアプリケーションをどこに登録しますか？」で「個人アカウントのみに紐付ける」と進まないと期待する場所に登録されないようです。ここはまりポイント
    1. ダウンロードしたファイルの中の AppConfig sample.json を AppConfig.json としてコピーする
    1. AppConfig.json ファイル内の ```AppClientId = "xxxx"``` の xxx 部分を登録したアプリのクライアントIDに変更

1. OneDrive for Business へのアップロードを行う場合
    1. [アプリの登録](https://learn.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)を参考にアプリを登録。管理者権限が必要です。
        1. アプリの登録の際にはシングルテナントとして登録するのが良いでしょう
    1. ダウンロードしたファイルの中の AppConfig sample.json を AppConfig.json としてコピーする
    1. AppConfig.json ファイルを編集
        1. ```AppClientIdForBusiness = "xxxx"``` の xxx 部分を登録したアプリのクライアントIDに変更
        1. ```TenantId = "xxx"``` の xxx 部分を自組織のテナントIDに変更

OneDrive個人用もOneDrive for Business も RedirectUri は https://localhost などがよいでしょう。

やっぱり無理ゲーですかね。

# USAGE

```Windows Command prompt
> pwsh .\Copy-GoPro.ps1 -LocalDestination \\fileserver\Videos
```

PCに接続されている GoPro を検出し、MP4ファイルを _fileserver_ 上の _Videos_ フォルダにコピーします。

```Windows Command Prompt
> pwsh .\Copy-GoPro.ps1 -Name "HERO11 Black" -LocalDestination \\fileserver\Videos -OneDriveDestination Videos -date 4/5
```

PCに接続されている HERO11 Back の MP4ファイルを _fileserver\\Videos_ フォルダと OneDrive 上の _ドキュメント\\Videos_ フォルダにコピーします。4/5以降に撮影された動画のみを処理します。

# コピー先フォルダ、ファイル名

## ローカルフォルダ

-LocalDestination で指定したフォルダ配下に動画ファイルの撮影日時をもとに "*yyyy\\mm\\yyyymmdd コメント*" フォルダを作成します。

## OneDrive / OneDrive for Business

実行したユーザの OneDrive 上のドキュメントフォルダ以下の -OneDriveDestination または -OneDriveForBusinessDestination で指定したフォルダ配下に動画ファイルの撮影日時をもとに "*yyyy\\mm\\yyyymmdd コメント*" フォルダを作成します。

## ファイル名

GoProのファイルは GX\<サブシーケンス番号\>\<シーケンス番号\>.MP4 となって管理がしにくいため、\<yyyymmdd-HHmm\>-GX\<シーケンス番号\>-\<サブシーケンス番号\>.MP4 という形式のファイル名でコピーを作成します。

# PARAMETERS

## [ -LocalDestination \<folder-path\> ]

動画ファイルをコピーするローカルフォルダを指定します。UNCパス形式（例 *\\\fileserver\\folder*）を設定することも可能です。
このフォルダにコピーされたファイルは OneDrive へのアップロードの際のコピー元ファイルとして使用されます。

## [ -OneDriveDestination \<folder\> ]

別名: OD

OneDrive 上のフォルダを指定します。ドキュメントフォルダ以下のパスを指定します。
省略した場合、OneDrive へのアップロードは行われません。

注意！ OneDrive個人用に対応したことにより、-OneDriveDestination パラメータは OneDrive個人用に変更されました。

## [ -OneDriveForBusinessDestination \<folder\> ]

別名: ODB

OneDrive for Business 上のフォルダを指定します。ドキュメントフォルダ以下のパスを指定します。
省略した場合、OneDrive for Business へのアップロードは行われません。

注意！ OneDrive個人用に対応したことにより、OneDrive for Business上へのアップロード指定はこちらを使うように変更されました。

## [ -Remark \<コメント\> ]

日付フォルダのコメントを指定します。複数の日の動画ファイルをコピーする場合、すべての日付フォルダで同じコメントが付加されます。
このコメント部分はコピー後に変更しても次回のコピー実施時に検出されるので、動画ファイルが重複してコピーされることはありません。

## [ -Date \<日付\> ]

動画ファイルが多数ある場合、コピー済ファイルの検出に時間がかかることああります。そのような場合には -Date <日付> 設定を行うことで、指定日以降の動画ファイルのみをコピーすることができます。

## [ -NAME \<device-name\> ]

GoProのデバイス名を指定します。Explorer で表示される名前を指定してください。
省略した場合、種類が "ポータブル デバイス" で "GoPro MTP" で始まるサブフォルダを持つデバイスを検出して使用します。

## [ -BufferSize \<size\> ]

OneDrive にアップロードする際のバッファサイズを指定します。既定は 100 で、 100*320KiB 単位でのデータアップロードを行います。
基本的に大きな値を指定するとアップロード速度が向上します。Surface Pro8で300を指定した場合、ネットワーク転送速度が 100-200Mbps 程度で使用した回線の上限に達しています。

## [ -LogFile \<ファイル名\> ]

既定で log.txt にログを出力します。不要な場合は "" を指定してください。

# 特記事項

~~現状 100GoPro ファイルのみが対象となっています。(自分のGoProに101GOPROができたらやる！~~

かっこ悪いので対応しました。101GOPROフォルダがないので試せませんが、100GOPROフォルダは探してきているので多分大丈夫、たぶん。

|製品|-Nameで指定する名前|
|---|---|
|Hero11 Black|HERO11 Black|

まんまじゃん。自動検出するようにしたので -Name オプションはもう要らないかも。


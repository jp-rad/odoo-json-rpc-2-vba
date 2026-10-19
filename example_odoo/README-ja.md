# odoo-addons-dev

VSCodeとDockerを使ったOdooカスタムモジュール開発用のテンプレートです。

## 概要

このテンプレートを使うことで、VSCodeとDockerでOdooカスタムモジュールの開発環境を素早く構築できます。標準のOdooアプリケーションの動作確認も可能です。

### 必要なツール

- [git](https://git-scm.com/)
- [VSCode](https://code.visualstudio.com/download)
- [Docker](https://www.docker.com/)

## はじめ方

### テンプレートの利用

1. GitHubでリポジトリを開く  
   https://github.com/jp-rad/odoo-addons-dev
2. `Use this template`をクリック  
   [新しいリポジトリを作成](https://github.com/jp-rad/odoo-addons-dev/generate)

### リポジトリのクローン

コマンドプロンプトで次のコマンドを実行：

```cmd
mkdir c:\workgit
cd c:\workgit
git clone <GitHubリポジトリのURL>
```

### VSCodeコンテナで開く

1. クローンしたフォルダをVSCodeで開く
2. 「Remote - Containers」拡張機能をインストール
3. 左下の「Remote - Containers」アイコンをクリック
4. 「Remote-Containers: Reopen in Container」を選択
5. Dockerコンテナが起動し、VSCodeからリモート開発が可能になります

### ローカル環境への Root CA 証明書の登録

この開発環境では、nginx が自己署名の Root CA を使って HTTPS を提供します。  
ブラウザで安全にアクセスするため、Root CA をローカル OS に登録します。

1. ブラウザで次の URL にアクセス  
   **https://localhost/oreore**
2. `dev-root-certificate.crt` がダウンロードされます  
3. OS の「信頼されたルート証明機関」に `dev-root-certificate.crt` を登録します  

#### Windows の場合（例）

1. `dev-root-certificate.crt` を右クリック  
2. 「証明書のインストール」を選択  
3. 「ローカルコンピューター」を選択  
4. ストアを「信頼されたルート証明機関」に指定  
5. インストール完了後、ブラウザを再起動  

これで **https://localhost/** に安全にアクセスできるようになります。

### custom_addons

`custom_addons`フォルダで独自アドオンを開発できます。  
[Odoo開発チュートリアル](https://www.odoo.com/documentation/master/developer/howtos/backend.html)も参照してください。

新しいモジュールを作成するには、次のコマンドを実行します：

```
/opt/odoo/repo/odoo/odoo-bin scaffold <モジュール名>
```

### Odooへのログイン

1. VSCodeの「Run and Debug」アイコンをクリック
2. 「Odoo addons」を選択し、「Start Debugging」でOdooを起動
3. ブラウザで [http://localhost/](http://localhost/) にアクセス
4. `admin` / `admin` でログイン

## ライセンス

### odoo-addons-dev

このリポジトリは [MITライセンス](https://opensource.org/licenses/MIT) です。  
商用・非商用問わず、利用・改変・再配布が可能です。

### 独自モジュールのライセンス

`custom_addons`配下のモジュールには、[LGPLv3](https://www.gnu.org/licenses/lgpl-3.0.html) または [AGPLv3](https://www.gnu.org/licenses/agpl-3.0.html) を選択してください。

- **LGPLv3**: Odoo本体の標準。商用利用や閉源連携が可能ですが、LGPL部分の改変は公開義務があります。
- **AGPLv3**: OCAモジュールの標準。サーバー利用でもソースコード公開が必要です。

配布方針や互換性に応じてライセンスを選択してください。

### OdooおよびOCAのライセンス

Odoo本体は主にLGPLv3、OCAモジュールはAGPLv3またはLGPLv3です。  
各プロジェクトのライセンス条件を必ず確認してください。

詳細はこちら:  
- [OCA / Odoo Meeting on Licenses](https://odoo-community.org/blog/news-updates-1/oca-odoo-meeting-on-licenses-21)

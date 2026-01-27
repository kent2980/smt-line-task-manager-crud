# Windowsで定期実行する方法

このプロジェクトをWindowsで定期実行する方法を説明します。

## 方法1: タスクスケジューラを使用（推奨）

Windowsのタスクスケジューラを使用して、PowerShellスクリプトを定期実行します。

### 手順

#### 1. タスクスケジューラを開く

1. Windowsキーを押して「タスクスケジューラ」と入力
2. 「タスクスケジューラ」を選択

または

1. `Win + R` を押す
2. `taskschd.msc` と入力してEnter

#### 2. 基本タスクの作成

1. 右側の「基本タスクの作成」をクリック
2. **名前**: `Excel処理スクリプト`（任意の名前）
3. **説明**: `Excelファイルを処理してAPIに送信するスクリプト`（任意）
4. 「次へ」をクリック

#### 3. トリガーの設定

1. **タスクの開始**: 実行頻度を選択
   - 毎日
   - 毎週
   - 毎月
   - コンピューターの起動時
   - ログオン時
   - など
2. 「次へ」をクリック
3. 開始日時と実行時間を設定
4. 「次へ」をクリック

#### 4. 操作の設定

1. **操作**: 「プログラムの開始」を選択
2. 「次へ」をクリック
3. **プログラム/スクリプト**: PowerShellのパスを指定

   ```
   C:\Program Files\PowerShell\7\pwsh.exe
   ```

   または（Windows PowerShell 5.1の場合）

   ```
   C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
   ```

4. **引数の追加（オプション）**: スクリプトのパスとオプションを指定

   ```
   -ExecutionPolicy Bypass -File "C:\Users\kentaroyoshida\vscode\powershell\smt-line-task-manager-crud\Main.ps1"
   ```

   または、詳細ログを表示する場合

   ```
   -ExecutionPolicy Bypass -File "C:\Users\kentaroyoshida\vscode\powershell\smt-line-task-manager-crud\Main.ps1" -Verbose
   ```

5. **開始場所（オプション）**: スクリプトのディレクトリを指定

   ```
   C:\Users\kentaroyoshida\vscode\powershell\smt-line-task-manager-crud
   ```

6. 「次へ」をクリック

#### 5. 完了

1. 設定内容を確認
2. 「完了」をクリック

#### 6. 詳細設定の確認

1. 作成したタスクをダブルクリック
2. **全般**タブ:
   - 「ユーザーがログオンしているかどうかにかかわらず実行する」を選択（推奨）
   - 「最上位の特権で実行する」にチェック（必要に応じて）
3. **条件**タブ:
   - 「コンピューターをAC電源で使用している場合のみタスクを開始する」のチェックを外す（必要に応じて）
4. **設定**タブ:
   - 「タスクをすぐに実行できるようにスケジュールする」にチェック
   - 「要求時にタスクを実行する」にチェック
5. **OK**をクリック

### テスト実行

1. タスクスケジューラで作成したタスクを選択
2. 右側の「実行」をクリック
3. タスクが正常に実行されるか確認

### ログの確認

- **タスクスケジューラの履歴**: タスクを選択して「履歴」タブを確認
- **スクリプトのログ**: `Config.ps1`で設定した`LogDirectory`内のログファイルを確認

## 方法2: PowerShellスケジュールジョブを使用

PowerShellの`Register-ScheduledJob`コマンドレットを使用する方法です。

### 手順

PowerShellを管理者権限で開き、以下のコマンドを実行：

```powershell
# スケジュールジョブを登録
$scriptPath = "C:\Users\kentaroyoshida\vscode\powershell\smt-line-task-manager-crud\Main.ps1"
$trigger = New-JobTrigger -Daily -At "09:00"  # 毎日9時に実行

Register-ScheduledJob `
    -Name "Excel処理スクリプト" `
    -FilePath $scriptPath `
    -Trigger $trigger `
    -RunNow
```

### スケジュールジョブの管理

```powershell
# 登録済みのスケジュールジョブを確認
Get-ScheduledJob

# スケジュールジョブを削除
Unregister-ScheduledJob -Name "Excel処理スクリプト"

# スケジュールジョブの実行履歴を確認
Get-Job -Name "Excel処理スクリプト"
```

## 方法3: バッチファイル + タスクスケジューラ

バッチファイルを作成して、タスクスケジューラから実行する方法です。

### バッチファイルの作成

`run-script.bat`を作成（非表示実行対応）：

```batch
@echo off
REM スクリプトのディレクトリに移動
cd /d "C:\Users\kentaroyoshida\vscode\powershell\smt-line-task-manager-crud"

REM VBScriptラッパーを使用して非表示で実行
cscript //nologo "run-script-hidden.vbs"
```

または、直接PowerShellを非表示で実行する場合：

```batch
@echo off
cd /d "C:\Users\kentaroyoshida\vscode\powershell\smt-line-task-manager-crud"
start /min "" "C:\Program Files\PowerShell\7\pwsh.exe" -ExecutionPolicy Bypass -WindowStyle Hidden -File "Main.ps1"
```

### タスクスケジューラでの設定

1. 操作の設定で、**プログラム/スクリプト**にバッチファイルのパスを指定

   ```
   C:\Users\kentaroyoshida\vscode\powershell\smt-line-task-manager-crud\run-script.bat
   ```

2. **全般**タブで「ユーザーがログオンしているかどうかにかかわらず実行する」を選択すると、ウィンドウは表示されません

## トラブルシューティング

### エラー: スクリプトが実行されない

1. **実行ポリシーの確認**:

   ```powershell
   Get-ExecutionPolicy
   ```

   `Restricted`の場合は、実行ポリシーを変更：

   ```powershell
   Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

2. **パスの確認**: スクリプトのパスとPowerShellのパスが正しいか確認

3. **権限の確認**: タスクスケジューラで「最上位の特権で実行する」にチェック

### エラー: 環境変数や.envファイルが読み込まれない

1. **開始場所の設定**: タスクスケジューラの「開始場所」にスクリプトのディレクトリを指定

2. **作業ディレクトリの確認**: スクリプト内で`.env`ファイルのパスが正しく解決されているか確認

### エラー: ログファイルが作成されない

1. **ディレクトリの権限**: ログディレクトリへの書き込み権限があるか確認

2. **パスの確認**: `Config.ps1`で設定した`LogDirectory`のパスが正しいか確認

## 推奨設定

- **実行頻度**: 業務要件に応じて設定（例: 毎日1回、毎時間など）
- **実行時間**: 業務時間外を推奨（サーバー負荷を考慮）
- **ログ保持**: ログファイルのローテーション設定を検討
- **エラー通知**: メール通知が正常に動作することを確認

## 参考リンク

- [タスクスケジューラの概要](https://learn.microsoft.com/ja-jp/windows/win32/taskschd/task-scheduler-start-page)
- [PowerShellスケジュールジョブ](https://learn.microsoft.com/ja-jp/powershell/module/psscheduledjob/)

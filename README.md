# Get-MIMDiagData 実行手順
## 概要

`Get-MIMDiagData` は、Microsoft Identity Manager (MIM) / Forefront Identity Manager (FIM) のトラブルシューティングに必要な診断情報を一括で取得するための PowerShell スクリプトです。

---

## 実行環境

本スクリプトは、**FIMService / MIM Service がインストールされているサーバー上**で実行してください。

また、PowerShell は **管理者権限**で起動してください。

```text
実行サーバー:
  FIMService / MIM Service がインストールされているサーバー

実行ユーザー:
  MIM Service 構成、MIM Sync サーバー、AD、SQL Server、PCNS DC の情報を参照できる管理者アカウント
```

---

## 事前準備 / 事前確認

FIMService / MIM Service サーバー上で、管理者権限の Windows PowerShell を起動し、必要な PowerShell スナップイン / モジュールおよびコマンドが利用できることを確認します。

### 実行ポリシー

現在の実行ポリシーを確認します。

```powershell
Get-ExecutionPolicy -List
```

次に、現在の Process スコープの設定を控えたうえで、現在の PowerShell セッションのみ `Bypass` に変更します。

```powershell
# 現在の Process スコープの設定を控えます
$OriginalProcessPolicy = Get-ExecutionPolicy -Scope Process

# 現在の PowerShell セッションのみ Bypass にします
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
```

---

### FIMAutomation

FIMService / MIM Service サーバー上で以下を実行し、`Export-FIMConfig` が利用できることを確認します。

```powershell
Add-PSSnapin FIMAutomation -ErrorAction Stop
Get-Command Export-FIMConfig
```

何も表示されない、または `FIMAutomation` がインストールされていない旨のエラーが出る場合、そのサーバーでは MIM Service 側の構成情報を取得できません。実行サーバーが FIMService / MIM Service サーバーであることを確認してください。

---

### Active Directory

AD オブジェクト情報やドメイン コントローラー情報を取得するために使用します。

```powershell
Import-Module ActiveDirectory -ErrorAction SilentlyContinue
Get-Command Get-ADUser -ErrorAction SilentlyContinue
Get-Command Get-ADDomainController -ErrorAction SilentlyContinue
```

---

### LithnetMIISAutomation


```powershell
Import-Module LithnetMIISAutomation -Force -ErrorAction SilentlyContinue
Get-Command -Module LithnetMIISAutomation -ErrorAction SilentlyContinue
```

---

### SharePoint PowerShell (Optional)

必須ではありません。利用できない場合、該当情報はスキップされます。

```powershell
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
Get-Command Get-SPSolution -ErrorAction SilentlyContinue
```

---

## 実行コマンド

**FIMService サーバー上の管理者権限 PowerShell** で実行してください。

### パターン 1：通常診断 / PCNS あり

FIMService / MIM Service サーバー上で実行し、MIM Sync、MIM Service、PCNS、CONFIG、EVENTLOG などをまとめて取得する場合に使用します。
MIM ポータルを利用していない環境の場合は、-MimServiceUri は不要です。

```powershell
.\Get-MIMDiagData.ps1 `
  -Logpath "<診断ログの出力先フォルダー>" `
  -SyncServer "<FIMSynchronizationService がインストールされているサーバー名>" `
  -MimServiceUri "http://<FIMService / MIM Service サーバー名または DNS 名>:5725/ResourceManagementService" `
  -PcnSServer "<PCNS がインストールされているドメイン コントローラー名>" `
  -NoTranscript
```

---

### パターン 2：通常診断 / PCNS なし

PCNS がインストールされていない環境、または PCNS 情報を取得しない場合に使用します。
MIM ポータルを利用していない環境の場合は、-MimServiceUri は不要です。

```powershell
.\Get-MIMDiagData.ps1 `
  -Logpath "<診断ログの出力先フォルダー>" `
  -SyncServer "<FIMSynchronizationService がインストールされているサーバー名>" `
  -MimServiceUri "http://<FIMService / MIM Service サーバー名または DNS 名>:5725/ResourceManagementService" `
  -SkipPCNS `
  -NoTranscript
```

---

### パターン 3：特定 AD オブジェクトの診断

特定の AD オブジェクトについて、AD / Connector Space / Metaverse 関連情報を取得する場合に使用します。
MIM ポータルを利用していない環境の場合は、-MimServiceUri は不要です。
このパターンでは、`CONFIG` と `EVENTLOG` フォルダーは作成されません。


```powershell
.\Get-MIMDiagData.ps1 `
  -Logpath "<診断ログの出力先フォルダー>" `
  -SyncServer "<FIMSynchronizationService がインストールされているサーバー名>" `
  -MimServiceUri "http://<FIMService / MIM Service サーバー名または DNS 名>:5725/ResourceManagementService" `
  -PcnSServer "<PCNS がインストールされているドメイン コントローラー名>" `
  -GetObjDomainName "<対象 AD オブジェクトが存在する AD ドメイン名>" `
  -GetObjADdn "<取得対象 AD オブジェクトの Distinguished Name>" `
  -DomainAdminName "<対象 AD ドメインの情報を参照できる管理者アカウント>" `
  -NoTranscript
```

PCNS がインストールされていない環境で特定 AD オブジェクト診断を行う場合は、`-PcnSServer` の代わりに `-SkipPCNS` を指定します。
MIM ポータルを利用していない環境の場合は、-MimServiceUri は不要です。

```powershell
.\Get-MIMDiagData.ps1 `
  -Logpath "<診断ログの出力先フォルダー>" `
  -SyncServer "<FIMSynchronizationService がインストールされているサーバー名>" `
  -MimServiceUri "http://<FIMService / MIM Service サーバー名または DNS 名>:5725/ResourceManagementService" `
  -SkipPCNS `
  -GetObjDomainName "<対象 AD オブジェクトが存在する AD ドメイン名>" `
  -GetObjADdn "<取得対象 AD オブジェクトの Distinguished Name>" `
  -DomainAdminName "<対象 AD ドメインの情報を参照できる管理者アカウント>" `
  -NoTranscript
```

---

## パラメーター確認方法

実行コマンド内の各パラメーターに何を指定すればよいか分からない場合は、以下の方法で確認します。

### `-Logpath`

診断ログの出力先フォルダーを指定します。
任意のフォルダーで問題ありません。例では `C:\Temp` を使用します。

```powershell
New-Item -ItemType Directory -Path C:\Temp -Force
```

指定例です。

```powershell
-Logpath C:\Temp
```

---

### `-SyncServer`

`FIMSynchronizationService` がインストールされている MIM Sync サーバー名を指定します。
スクリプト自体は FIMService / MIM Service サーバー上で実行しますが、`-SyncServer` には MIM Sync サーバー名を指定します。

MIM Sync サーバー上、またはリモート確認できる端末で以下を実行し、サービスが存在することを確認します。

```powershell
Get-Service FIMSynchronizationService
```

サーバー名は以下で確認できます。

```powershell
$env:COMPUTERNAME
[System.Net.Dns]::GetHostEntry($env:COMPUTERNAME).HostName
```

指定例です。

```powershell
-SyncServer "MIMSYNC01"
```

---

### `-MimServiceUri`

FIMService / MIM Service の `ResourceManagementService` エンドポイントを指定します。

`-MimServiceUri` には、物理サーバー名、FQDN、または `mim.contoso.com` のような MIM Service 用に構成された DNS 名 / 別名を指定できます。

```powershell
-MimServiceUri "http://<FIMService / MIM Service サーバー名または DNS 名>:5725/ResourceManagementService"
```

> 注意: MIM Portal / SharePoint の URL ではなく、`ResourceManagementService` に到達できる URI を指定します。通常は `:5725/ResourceManagementService` を付けて指定します。

#### FIMService がインストールされているサーバー名を確認する

FIMService がインストールされているサーバー上で以下を実行し、サービスが存在することを確認します。

```powershell
Get-Service FIMService
```

そのサーバーのホスト名 / FQDN は以下で確認できます。

```powershell
$env:COMPUTERNAME
[System.Net.Dns]::GetHostEntry($env:COMPUTERNAME).HostName
```

この場合の指定例です。

```powershell
-MimServiceUri "http://MIMSERVICE01.contoso.com:5725/ResourceManagementService"
```


---

### `-PcnSServer`

PCNS がインストールされているドメイン コントローラー名を指定します。
PCNS を使用している環境では、対象 DC 上に `PCNSSVC` サービスが存在します。

対象 DC が分かっている場合は、以下で確認します。

```powershell
Get-Service -ComputerName <PCNS がインストールされている DC 名> -Name PCNSSVC
```

対象 DC が分からない場合は、ドメイン コントローラー一覧を取得し、各 DC の `PCNSSVC` を確認します。

```powershell
Import-Module ActiveDirectory

Get-ADDomainController -Filter * |
Select-Object HostName, Site, IPv4Address
```

PCNS がインストールされていない環境、または PCNS 情報を取得しない場合は、`-PcnSServer` ではなく `-SkipPCNS` を指定します。

```powershell
-SkipPCNS
```

---

### `-GetObjDomainName`

特定 AD オブジェクト診断で使用します。
対象 AD オブジェクトが存在する AD ドメイン名を指定します。

```powershell
Import-Module ActiveDirectory
Get-ADDomain | Select-Object DNSRoot, NetBIOSName
```

指定例です。

```powershell
-GetObjDomainName "contoso.com"
```

---

### `-GetObjADdn`

特定 AD オブジェクト診断で使用します。
取得対象 AD オブジェクトの Distinguished Name を指定します。

ユーザーの DN を確認する例です。

```powershell
Get-ADUser <SamAccountName> | Select-Object DistinguishedName
```

グループの DN を確認する例です。

```powershell
Get-ADGroup <GroupName> | Select-Object DistinguishedName
```

汎用的に AD オブジェクトを検索する例です。

```powershell
Get-ADObject -LDAPFilter "(name=<オブジェクト名>)" -Properties distinguishedName |
Select-Object Name, ObjectClass, DistinguishedName
```

指定例です。

```powershell
-GetObjADdn "CN=u01,OU=SyncUsers,DC=contoso,DC=com"
```

---

### `-DomainAdminName`

特定 AD オブジェクト診断で使用します。
対象 AD ドメインの情報を参照できるアカウントを指定します。

指定例です。

```powershell
-DomainAdminName "CONTOSO\Administrator"
```

通常診断では、このパラメーターは不要です。

---

### `-NoTranscript`

PowerShell Transcript の取得を無効化します。
通常の診断取得や共有用ログを作成する場合は、`-NoTranscript` を指定することを推奨します。

```powershell
-NoTranscript
```

---

## 事後作業

診断ログの取得が完了したら、必要に応じて以下を実施します。

### 実行ポリシーを元に戻す

事前準備で Process スコープの実行ポリシーを `Bypass` に変更している場合は、スクリプト実行後に変更前の設定へ戻します。

```powershell
# スクリプト実行後、変更前の Process スコープの設定に戻します
Set-ExecutionPolicy -Scope Process -ExecutionPolicy $OriginalProcessPolicy -Force

# 戻ったことを確認します
Get-ExecutionPolicy -List
```

この手順は、`$OriginalProcessPolicy` を取得したものと同じ PowerShell セッションで実行してください。
`-Scope Process` の設定は現在の PowerShell セッションのみで有効です。PowerShell を閉じた場合も、Process スコープの設定は破棄されます。

### 不要になったモジュールの削除

診断完了後、環境に不要なモジュールを残したくない場合は、追加で導入したモジュールのみ削除します。
MIM / SharePoint の製品コンポーネントとして導入されているスナップインは、個別削除ではなく製品側の変更操作で管理してください。

#### LithnetMIISAutomation を削除する場合

PowerShell Gallery から `Install-Module` で導入した場合は、以下で削除します。

```powershell
Get-InstalledModule LithnetMIISAutomation -ErrorAction SilentlyContinue
Uninstall-Module LithnetMIISAutomation -AllVersions -Force
```

#### Active Directory PowerShell モジュールを削除する場合

`ActiveDirectory` モジュールは RSAT / AD DS 管理ツールの一部です。
他の管理作業で使用する可能性があるため、不要であることを確認してから削除してください。

Windows Server の場合です。

```powershell
Remove-WindowsFeature RSAT-AD-PowerShell
```
#### FIMAutomation / SharePoint PowerShell について

`FIMAutomation` は MIM Service 関連コンポーネント、`Microsoft.SharePoint.PowerShell` は SharePoint 関連コンポーネントとして導入されます。
通常、診断スクリプトのためだけに個別削除するものではありません。削除が必要な場合は、MIM または SharePoint のセットアップ / 役割変更手順に従ってください。

---

## 主な出力フォルダー

通常診断では、指定した `-Logpath` 配下に以下のようなフォルダーが作成されます。

```text
MIMLOG_yyyyMMdd_HHmmss
├─ CONFIG
├─ DIAGNOSTIC
├─ EVENTLOG
├─ SYNC DATA
└─ MIM_Diagnostic_Report.html
```

特定 AD オブジェクト診断では、`CONFIG` と `EVENTLOG` フォルダーは作成されません。

---

## 主な出力ファイル

### HTML レポート

主要な診断情報を 1 つの HTML で確認できます。

```text
MIMLOG_yyyyMMdd_HHmmss\MIM_Diagnostic_Report.html
```

通常診断では、以下にも同じ HTML レポートがコピーされます。

```text
CONFIG\MIM_Diagnostic_Report.html
```

特定 AD オブジェクト診断では、以下にも同じ HTML レポートがコピーされます。

```text
SYNC DATA\MIM_Diagnostic_Report.html
```

通常診断の HTML レポートには、主に以下の情報が含まれます。

- Management Agent 一覧
- AD MA の OU スコープ
- SynchronizationRule 一覧
- SynchronizationRule / Workflow / MPR / Set 関連表
- Attribute Flow
- PCNS 診断結果
- MIM / FIM ビルド情報
- SharePoint バージョン情報
- Metaverse 拡張 DLL 情報
- DC / ネットワーク確認結果

特定 AD オブジェクト診断の HTML レポートには、主に以下の情報が含まれます。

- AD DS の対象オブジェクト属性と値
- Connector Space の対象オブジェクト属性と値（Lithnet の Attributes 辞書を展開）
- Metaverse の対象オブジェクト属性と値（Lithnet の Attributes 辞書を展開）
- Connector Space / Metaverse の検索サマリー

---

### MIM / FIM バージョン情報

```text
DIAGNOSTIC\MIM_Sync_Build.csv
DIAGNOSTIC\MIM_Service_Status.csv
DIAGNOSTIC\MIM_Service_Portal_Build.csv
DIAGNOSTIC\MIM_InstalledProducts.csv
DIAGNOSTIC\MIM_BuildVersion.csv
```

---

### SharePoint バージョン情報

SharePoint PowerShell または SharePoint 関連 DLL / インストール情報から取得できる範囲で、SharePoint のバージョン情報を出力します。

```text
DIAGNOSTIC\SharePoint_Build.csv
```

このファイルには、取得できる場合に以下の情報が含まれます。

- SharePoint Farm BuildVersion
- Microsoft.SharePoint.dll の ProductVersion / FileVersion
- インストール済み SharePoint 製品情報

SharePoint PowerShell が利用できない場合でも、取得可能な範囲のみ出力し、スクリプトは停止しません。

---

### Management Agent 情報

```text
CONFIG\ManagementAgents
CONFIG\ManagementAgent_List_From_SyncServer.csv
```

主に以下の情報を取得します。

- MA XML
- MA 一覧
- AD MA の OU スコープ
- MA 接続先情報

---

### SynchronizationRule / Workflow / MPR / Set 関連情報

```text
CONFIG\SynchronizationRules
CONFIG\All_SyncRule_Workflow_MPR_Set_Map.csv
CONFIG\SynchronizationRules\All_SyncRule_Workflow_MPR_Set_Map.csv
```

全 SynchronizationRule を対象に、各同期ルールに紐づく Workflow / MPR / Set の関連を一覧化します。

---

### Metaverse 拡張 DLL 情報

Metaverse Rules Extension の設定と DLL 実体情報を出力します。

```text
CONFIG\MetaverseExtension\Metaverse_Extension_Config.csv
CONFIG\MetaverseExtension\Metaverse_Extension_DllInfo.csv
CONFIG\MetaverseExtension\Metaverse_Extension_RegisteredDlls.csv
CONFIG\MetaverseExtension\Metaverse_Extension_CollectionStatus.csv
CONFIG\MetaverseExtension\Metaverse_Extension_CollectionStatus.txt
```

主に以下の情報を取得します。

- Metaverse Rules Extension の有効 / 無効
- 設定されている DLL 名
- Application Protection 設定
- DLL の存在有無
- DLL のパス、サイズ、更新日時、FileVersion、ProductVersion、SHA256
- Sync DB に登録されている Extension DLL 一覧

---

### PCNS 情報

PCNS ありの場合、以下に PCNS 関連情報を出力します。

```text
CONFIG\PCNS
CONFIG\PCNS\PCNS_CollectionStatus.csv
CONFIG\PCNS\PCNS_CollectionStatus.txt
```

PCNS がない場合、または `-SkipPCNS` を指定した場合も、PCNS 収集をスキップしたことが分かるようにステータスファイルを出力します。

---

### イベントログ

通常診断では以下にイベントログを出力します。

```text
EVENTLOG
```

特定 AD オブジェクト診断では、`EVENTLOG` フォルダーは作成されません。

---

## パラメーター説明

| パラメーター | 説明 |
| --- | --- |
| `-Logpath` | 診断ログの出力先フォルダーを指定します。 |
| `-SyncServer` | FIMSynchronizationService がインストールされている MIM Sync サーバーを指定します。 |
| `-MimServiceUri` | FIMService / MIM Service の ResourceManagementService エンドポイントを指定します。 |
| `-PcnSServer` | PCNS がインストールされているドメイン コントローラーを指定します。 |
| `-SkipPCNS` | PCNS 情報を取得しない場合に指定します。 |
| `-GetObjDomainName` | 特定 AD オブジェクト取得時に、対象オブジェクトが存在する AD ドメイン名を指定します。 |
| `-GetObjADdn` | 特定 AD オブジェクト取得時に、対象オブジェクトの Distinguished Name を指定します。 |
| `-DomainAdminName` | 対象 AD ドメインの情報を参照できる管理者アカウントを指定します。 |
| `-NoTranscript` | PowerShell Transcript の取得を無効化します。通常の診断取得では指定を推奨します。 |

---

## 補足

`-NoTranscript` を指定しない場合、PowerShell Transcript ログが作成されます。

Transcript にはコンソール出力や実行状況が記録されるため、スクリプトの詳細調査には有用ですが、共有前に機微情報が含まれていないか確認してください。

通常の診断取得や共有用ログを作成する場合は、`-NoTranscript` を指定することを推奨します。

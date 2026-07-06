# Get-MIMDiagData 実行手順
## 概要

`Get-MIMDiagData` は、Microsoft Identity Manager (MIM) / Forefront Identity Manager (FIM) のトラブルシューティングに必要な診断情報を一括で取得するための PowerShell スクリプトです。

---

## 実行環境

本スクリプトは、**FIMService がインストールされているサーバー上**で実行してください。

また、PowerShell は **管理者権限**で起動してください。

```text
実行サーバー:
  FIMService がインストールされているサーバー

実行ユーザー:
  FIMService 構成、FIMSynchronizationService サーバー、AD、SQL Server、PCNS DC の情報を参照できる管理者アカウント
```


---

## 事前準備

FIMService サーバー上で、管理者権限の Windows PowerShell を起動し、必要な PowerShell スナップイン / モジュールおよびコマンドが利用できることを確認します。

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

### FIMSynchronizationService サーバーへのリモート接続

FIMService サーバーと FIMSynchronizationService サーバーが異なる場合は、FIMService サーバー上で以下を実行し、FIMSynchronizationService サーバーへ WinRM 接続できることを確認します。

```powershell
$SyncServer = "<FIMSynchronizationService がインストールされているサーバー名>"
Test-NetConnection $SyncServer -Port 5985
Test-WSMan $SyncServer
```

`Test-NetConnection` の `TcpTestSucceeded` が `False`、または `Test-WSMan` が失敗する場合、FIMSynchronizationService サーバー側の情報を取得できません。

接続できない場合は、FIMSynchronizationService サーバー上で管理者権限の Windows PowerShell を起動し、WinRM / PowerShell Remoting を有効化します。

```powershell
# FIMSynchronizationService サーバー上で実行します
Enable-PSRemoting -Force
Set-Service WinRM -StartupType Automatic
Start-Service WinRM
Enable-NetFirewallRule -DisplayGroup "Windows Remote Management"
```

有効化後、FIMService サーバーに戻り、再度 FIMService サーバー上から、FIMSynchronizationService サーバーへ WinRM 接続できることを確認します。


---

### FIMAutomation

FIMService サーバー上で以下を実行し、`Export-FIMConfig` が利用できることを確認します。

```powershell
Add-PSSnapin FIMAutomation -ErrorAction Stop
Get-Command Export-FIMConfig
```

何も表示されない、または `FIMAutomation` がインストールされていない旨のエラーが出る場合、そのサーバーでは FIMService 側の構成情報を取得できません。実行サーバーが FIMService サーバーであることを確認してください。

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

FIMSynchronizationService サーバー上で以下を実行し、`LithnetMIISAutomation` が利用できることを確認します。

```powershell
Install-Module -Name LithnetMIISAutomation -Scope AllUsers -Force
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

### 通常診断

FIMService サーバー上で実行し、FIMSynchronizationService、FIMService、CONFIG、EVENTLOG などをまとめて取得する場合に使用します。
MIM ポータルを利用していない環境の場合は、-MimServiceUri は不要です。

PCNS 情報を取得する場合は `-PcnSServer` を指定します。
PCNS 情報を取得しない場合は、`-PcnSServer` の代わりに `-SkipPCNS` を指定します。

```powershell
.\Get-MIMDiagData.ps1 `
  -Logpath "<診断ログの出力先フォルダー>" `
  -SyncServer "<FIMSynchronizationService がインストールされているサーバー名>" `
  -MimServiceUri "http://<FIMService サーバー名または DNS 名>:5725/ResourceManagementService" `
  -PcnSServer "<PCNS がインストールされているドメイン コントローラー名>" 
```

---

### 特定 AD オブジェクトの診断

特定の AD オブジェクトについて、AD / Connector Space / Metaverse 関連情報を取得する場合に使用します。
MIM ポータルを利用していない環境の場合は、-MimServiceUri は不要です。
この診断では、対象 AD オブジェクトが存在するドメイン名、DN、および参照用アカウントを指定します。

```powershell
.\Get-MIMDiagData.ps1 `
  -Logpath "<診断ログの出力先フォルダー>" `
  -SyncServer "<FIMSynchronizationService がインストールされているサーバー名>" `
  -MimServiceUri "http://<FIMService サーバー名または DNS 名>:5725/ResourceManagementService" `
  -GetObjDomainName "<対象 AD オブジェクトが存在する AD ドメイン名>" `
  -GetObjADdn "<取得対象 AD オブジェクトの Distinguished Name>" `
  -DomainAdminName "<対象 AD ドメインの情報を参照できる管理者アカウント>" 
```

---

## パラメーター確認方法

実行コマンド内の各パラメーターに何を指定すればよいか分からない場合は、以下の方法で確認します。

### `-Logpath`

診断ログの出力先フォルダーを指定します。
任意のフォルダーで問題ありません。例では `C:\Temp` を使用します。

```powershell
-Logpath C:\Temp
```

---

### `-SyncServer`

`FIMSynchronizationService` がインストールされているサーバー名を指定します。
スクリプト自体は FIMService サーバー上で実行しますが、`-SyncServer` には FIMSynchronizationService サーバー名を指定します。

FIMSynchronizationService サーバー上、またはリモート確認できる端末で以下を実行し、サービスが存在することを確認します。

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
-SyncServer "WIN-U74XXXXXXC"
```

---

### `-MimServiceUri`

FIMService の `ResourceManagementService` エンドポイントを指定します。

`-MimServiceUri` には、物理サーバー名、FQDN、または `mim.contoso.com` のような FIMService 用に構成された DNS 名 / 別名を指定できます。

```powershell
-MimServiceUri "http://<FIMService サーバー名または DNS 名>:5725/ResourceManagementService"
```

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
-MimServiceUri "http://WIN-M6F0SXXXXXXG.contoso.com:5725/ResourceManagementService"
```


---

### `-PcnSServer`

PCNS がインストールされているドメイン コントローラー名を指定します。
PCNS を使用している環境では、対象 DC 上に `PCNSSVC` サービスが存在します。

対象 DC が分かっている場合は、以下で確認します。

```powershell
Get-Service -ComputerName <PCNS がインストールされている DC 名> -Name PCNSSVC
```

対象 DC が分からない場合は、ドメイン コントローラー一覧を取得し `PCNSSVC` を確認します。

```powershell
Import-Module ActiveDirectory

Get-ADDomainController -Filter * |
Select-Object HostName, Site, IPv4Address
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

### FIMSynchronizationService サーバーへのリモート接続を無効化する

**診断のために一時的に WinRM / PowerShell Remoting を有効化した場合のみ**実施します。
既に運用管理で WinRM / PowerShell Remoting を使用している環境では、無効化しないでください。

FIMSynchronizationService サーバー上で、管理者権限の Windows PowerShell を起動し、以下を実行します。

```powershell
# FIMSynchronizationService サーバー上で実行します
Disable-PSRemoting -Force
Stop-Service WinRM
Set-Service WinRM -StartupType Disabled
```

### 不要になったモジュールの削除

診断完了後、環境に不要なモジュールを残したくない場合は、追加で導入したモジュールのみ削除します。
MIM / SharePoint の製品コンポーネントとして導入されているスナップインは、個別削除ではなく製品側の変更操作で管理してください。

#### LithnetMIISAutomation を削除する場合

PowerShell Gallery から `Install-Module` で導入した場合は、以下で削除します。

```powershell
Get-InstalledModule LithnetMIISAutomation -ErrorAction SilentlyContinue
Uninstall-Module LithnetMIISAutomation -AllVersions -Force
```

---

## 出力について

既定では、指定した `-Logpath` 配下に `MIMLOG_yyyyMMdd_HHmmss.zip` のみを出力します。

---

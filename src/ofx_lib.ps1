<#
    発注書.xlsmランチャー・インストーラ　ライブラリ
    copyright 2024 C.Nagata
    2024.6.16   Initial writing.
#>

# 最低限必要なファイルのキー
$REQUIRED_FILES = @("launch-bat", "install-ps1", "lib-ps1")
#   Excelへ設定パスを渡す環境変数
$ENVVAR_NAME = "OFX_BOOK_CONFIG"
#   各ファイルへのパス
$CONF_PATH = $PSScriptRoot + "\ofx_config.json"
$BOOK_CONF_PATH = $PSScriptRoot + "\ofx_book_config.json"
#   設定の取得
$CONFIG = (Get-Content $CONF_PATH -Raw | ConvertFrom-Json)
<#    環境の取得    #>
function getEnv($path="") {
    if($path -eq "") {
        $path = $script:PSScriptRoot
    }
    $drv = $path.Substring(0, 2)
    if($path.Contains("src")) {
        return "debug"
    } elseif( $drv -eq "\\" ) {
        return "online"
    } elseif ($drv -eq "C:") {
        return "mobile"
    }
    return "USB"
}
#このファイルの環境（online/mobile/USB）
$THIS_ENV = getEnv


<#  オンライン（ネットワークフォルダがある）か調べる  #>
function IsOnline {
    $path = $CONFIG.env.online.dirs.'base-dir'
    return (Test-Path -LiteralPath $path)
}

<#  ソースの環境を取得する  #>
function getSourceEnv {
    if(IsOnline) {
        # onlineなら、ソースもonline
        return "online"
    } elseif($THIS_ENV -eq "USB") {
        # 自分（このファイル）がUSBにあるなら、ソースなし
        return "none"
    }
    # 自分の環境（mobile）
    return $THIS_ENV
}

function echoVars {
    Write-Host "必須ファイル:${REQUIRED_FILES}"
    Write-Host "環境変数名:${ENVVAR_NAME}"
    Write-Host "設定ファイル:${CONF_PATH}"
    Write-Host "発注書Excel設定ファイル:${BOOK_CONF_PATH}"
    Write-Host "ホストの環境:${THIS_ENV}"
}

<#  パスのための辞書を作る
#>
function makePathDict($obj, $drive_letter="") {
    if($null -eq $obj) {
        Write-Host "makePathDict に渡された PSObject は null です。"
        exit
    }
    # ドライブが指定されていなかったら、スクリプトと同じ場所
    if( $drive_letter -eq "") {
        $drive_letter = $script:PSScriptRoot.Substring(0, 2)
    }
    #   辞書の変数。初期メンバーはドライブレター部
    $dict = @{"Drive" = $drive_letter; "AppDir" = $script:PSScriptRoot}
    #   特殊フォルダ
    foreach($key in "Desktop","MyDocuments","StartMenu","Templates") {
        $dict.Add($key, [System.Environment]::GetFolderPath($key))
    }
    #   オブジェクトのメンバ
    foreach($key in $obj.psobject.properties.name) {
        $s = [regex]::replace($obj.$key, "\%([\w\-]+)\%", { $dict[$args.groups[1].value] })
        $dict.Add($key, $s)
    }
    return $dict
}

<#  パスを展開する  #>
function expandPath($path, $dict, $add_base) {
    $s = [regex]::replace($path, "\%([\w\-]+)\%", { $dict[$args.groups[1].value] })
    if($add_base) {
        $paths = $s.Split(";")
        for ($i = 0; $i -lt $paths.Length; $i++) {
            if(($paths[$i] -eq "") -or `
                ($paths[$i].Substring(0,1) -ne "\" -and $paths[$i].Substring(1,1) -ne ":")) {
                $paths[$i] = $dict."base-dir" + "\" + $paths[$i]
            }
        }
        $s = $paths -join ";"
    }
    $s = $s.replace("/", "\")
    return $s
}

<#  辞書で、データ中の変数を展開する    #>
function convertVars($data, $dict, $add_base=$false) {
    if($null -eq $data) {
        Write-Host "convertVars に渡された PSObject は null です。"
        exit
    }
    foreach($key in $data.psobject.properties.name) {
        $data.$key = expandPath $data.$key $dict $add_base
    }
    return $data
}


<#  インストールのソースにする環境の優先順位
    若い方が優先順位が高い   #>
function envPriority($env) {
    $ary = @("debug", "online", "mobile", "USB", "none")
    return [Array]::IndexOf($ary, $env)
}

<#  インストール／アップデート可能か？    #>
function canInstall($drv_inf) {
    $src_env = getSourceEnv
    return ((envPriority $drv_inf["env"]) -gt (envPriority $src_env))
}
    
<#   ドライブ情報の取得
    @param $drive ドライブレター
    @return ドライブ情報のハッシュテーブル
            letter: (string) ドライブレター
            env: (string) 環境（mobile/USB）
            exists: (boolean) インストールされているか？
            can-install: (boolean) インストール／アップデート可能か？
#>
function getDriveInfo($drive) {
    # 環境の取得
    $env = "USB"
    if($drive -eq "C:") {
        $env = "mobile"
    } elseif($drive -eq "\\") {
        $env = "online"
    }
    # 辞書の作成
    $dict = makePathDict $CONFIG.env.$env.dirs $drive
    # 初期情報（ドライブレターとmobile/USB）
    $info = @{ "letter" = $drive; "env" = $env }
    # 必要なファイルは全て揃っているか？
    $exists = $true
    foreach($fkey in $REQUIRED_FILES) {
        # キーに対応するパスを得る
        $path = expandPath $CONFIG.install.files.$fkey $dict $true
        if( -not (Test-Path -LiteralPath $path) ) {
            $exists = $false
            break
        }
    }
    $info.Add("exists", $exists)
    $info.Add("can-install", (canInstall $info))
    return $info
}

<#  ドライブの検出し、対象のドライブ全てに付いて情報を取得する  #>
function getDrives($only_usb=$false) {
    # Cドライブの情報。初期メンバーは、PCのドライブ（C:、mobile）
    $drives = @{}
    if($only_usb) {
        $drives.Add("C:", (getDriveInfo "C:"))
    }
    # USB（リムーバブルメディア）のドライブを取得
    $usbs = (Get-WmiObject CIM_LogicalDisk | Where-Object DriveType -eq 2).DeviceID
    foreach( $usb in $usbs) {
        # ここのドライブの情報を取得
        $drives.Add($usb, (getDriveInfo $usb))
    }
    return $drives
}


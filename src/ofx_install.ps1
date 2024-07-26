<#
    発注書ランチャ―インストーラー
    copyright 2024 C.Nagata
    2024.6.16   Initial writing.
#>

Set-StrictMode -Version 1.0

# ライブラリのインポート
. "$($PSScriptRoot)\ofx_lib.ps1"

# インストーラーのファイルのキー
$INSTALLER_FILES = @("install-ps1", "lib-ps1")

<#  ディレクトリの作成
    CONFIGの　install > dirs にあるディレクトリを作成する
#>
function makeDirs($dict, $to_make=$true) {
    #　ディレクトリ配列を得、展開する
    $dirs = $CONFIG.install.dirs
    for($i=0; $i -lt $dirs.Length; $i++) {
        $dirs[$i] = expandPath $dirs[$i] $dict $true
    }
    # 要素ごと、「；」で分解し、ディレクトリを作成する
    foreach($dir in $dirs) {
        $paths = $dir -split ";"
        foreach($path in $paths) {
            if($to_make) {
                New-Item $path -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
                Write-Host $path
            } elseif( Test-Path -LiteralPath $path) {
                Remove-Item -LiteralPath $path -Recurse -Force
                Write-Host $path
            }
        }
    }
}

<#  ファイルをコピーする    
    CONFIGの　install > files にあるファイルが、ないか古ければコピーする
#>
function copyFiles($src_dict, $dest_dict, $is_update=$false) {
    $copied = 0
    #   転送元のファイル
    $src_files = $CONFIG.install.files.psobject.copy()
    $src_files = convertVars $src_files $src_dict $true
    #   転送先のファイルパス
    $dest_files = $CONFIG.install.files.psobject.copy()
    $dest_files = convertVars $dest_files $dest_dict $true
    #   ループ
    foreach($key in $src_files.psobject.properties.name) {
        # アップデートのとき、キー値がインストーラーのファイルならコピーしない
        if($is_update -and $INSTALLER_FILES.Contains($key) ) {
            continue
        }
        $msg = ""
        #   転送元・転送先ファイルパス
        $src = $src_files.$key
        $dest = $dest_files.$key
        #   転送先のファイルがなければ、転送する
        $do = -not (Test-Path -LiteralPath $dest)
        if(-not $do) {
            #   ファイルが既にあるときは、タイムスタンプを比較して本が新しければコピーする
            $src_date = [datetime](Get-ItemProperty $src).LastWriteTime
            $dest_date = [datetime](Get-ItemProperty $dest).LastWriteTime
            $do = ( $src_date.CompareTo($dest_date) -gt 0)
        }
        if($do) {
            # ディレクトリがなければ作成する
            $dir = Split-Path -Parent $dest
            if(-not (Test-Path -LiteralPath $dir)) {
                New-Item $dir -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
                Write-Host "Making a directory ... ${dir}"
            }
            if(Test-Path -LiteralPath $src) {
                Copy-Item $src $dest -Force | Out-Null
                Unblock-File $dest
                $copied += 1
                $msg = "Copying from " + $src + "`r`n to "
            } else {
                Write-Host "Not found. ${src}"
            }
        } else {
            $msg = "Latest. "
        }
        if($msg -ne "") {
            Write-Host "${msg} ${dest}"
        }
    }
    return $copied
}

<#  ショートカットのファイルパスの取得#>
function getShortCutPath($name) {
    return [System.Environment]::GetFolderPath("Desktop") + "\" + $name + ".lnk"
}

<#  ショートカットの作成    #>
function createShortCut($env, $dect) {
    $sc_inf = $CONFIG.env.$env.shortcut
    $dir = expandPath $sc_inf.dir $dect $true
    $path = $dir + '\' + $sc_inf.name + '.lnk'
    $target = $CONFIG.install.files.($sc_inf.target)
    $target = expandPath $target $dect $true
    $icon = $CONFIG.install.files.($sc_inf.icon)
    $icon = expandPath $icon $dect $true
    $wd = expandPath $sc_inf."working-dir" $dect $true

    $WsShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WsShell.CreateShortcut($path)
    $Shortcut.TargetPath = $target
    $Shortcut.IconLocation = $icon
    $Shortcut.WorkingDirectory = $wd
    $Shortcut.Save()
}
<#  インストール／アップデート 
#>
function install($drv_info, $SRC_ENV) {
    #   転送元のディレクトリ辞書（Driveは、中で自分のパスから取得される）
    $src_dict = makePathDict $CONFIG.env.$SRC_ENV.dirs
    # 転送先のパスの辞書の作成
    $letter = $drv_info["letter"]
    $dest_dict = makePathDict $CONFIG.env.($drv_info["env"]).dirs $letter
    # ドライブにインストールされていたらアップデート。その他はインストール
    if( $drv_info["exists"] ) {
        Write-Host "${letter}ドライブのアプリが最新か確認します…"
        $copied = copyFiles $src_dict $dest_dict $true
        if($copied -eq 0) {
            Write-Host "${letter}ドライブのアプリは最新です。"
        } else {
            Write-Host "${letter}ドライブのアプリを更新しました。"
        }
    } else {
        Write-Host "${letter}ドライブにアプリをインストールします…"
        Write-Host "ディレクトリを作成しています…"
        makeDirs $dest_dict
        Write-Host "ファイルをコピーしています…"
        $copied = copyFiles $src_dict $dest_dict $false
        if($drv_info["env"] -ne "USB") {
            Write-Host "ショートカットを作成しています…"
            createShortCut $drv_info["env"] $dest_dict
        }
        Write-Host "${letter}ドライブへのインストールが完了しました"
    }
}

<#  ファイルの削除  #>
function removeFiles($dict) {
    # 削除するファイル
    $files = $CONFIG.install.files.psobject.copy()
    $files = convertVars $files $dict $true
    # 対象ドライブがC:なら、ショートカットも削除
    if($dict['Drive'] -eq "C:") {
        $fn = getShortCutPath $CONFIG.install.shortcut.name
        $files | Add-Member -MemberType NoteProperty -Name "shortcut" -Value $fn
    }

    foreach($key in $files.psobject.properties.name) {
        if(Test-Path -LiteralPath $files.$key) {
            Write-Host $files.$key
            Remove-Item -LiteralPath $files.$key -Force
        }
    }
}

<#   アンインストール   #>
function uninstall($drv_info) {
    $letter = $drv_info["letter"]
    Write-Host "${letter}ドライブのアプリを削除します…"
    # 転送先のパスの辞書の作成
    $dict = makePathDict $CONFIG.env.($drv_info["env"]).dirs $letter
    Write-Host "ディレクトリを削除しています…"
    makeDirs $dict $false
    Write-Host "ファイルを削除しています…"
    removeFiles $dict
    Write-Host "${letter}ドライブのアプリを削除しました。"
}

<#  インストール／アンインストールをするドライブを選択させる
    @param $drive (hashtable) ドライブ情報のハッシュテーブル
    @param $is_install (boolean) $true:インストール／$false:アンインストール
#>
function askDrive($drives, $is_install) {
    # 対象のドライブのリストを作成する。
    $lst = @()
    foreach( $letter in $drives.Keys) {
        if(($is_install -and -not $drives[$letter]["exists"] `
                -and $drives[$letter]["can-install"])`
            -or (-not $is_install -and $drives[$letter]["exists"])) {
            $lst = $lst + $letter.Substring(0,1)
        }
    }
    $proc = "アプリをインストール"
    if(-not $is_install) {
        $proc = "アプリを削除"
    }
    if($lst.Length -eq 0) {
        Write-Host "${proc}できるドライブはありません。"
        return ""
    }
    $drv_letter = Read-Host ("${proc}するドライブを選択してください。（"`
                    + ($lst -join ", ") + ", その他：取り消し）")
    
    $drv_letter = $drv_letter.ToUpper()
    if( $lst.Contains($drv_letter) ) {
        return ($drv_letter + ":")
    }
    return ""
}

<#   ドライブ情報のハッシュテーブルの中身を表示する   #>
function HashDsp($hash) {
    $lst = @()
    foreach($key in $hash.Keys) {
        $lst += $hash[$key]
    }
    Write-Host
    ($lst | ForEach-Object { New-Object PSCustomObject -Property $_ } `
         | Out-String).trim() | Write-Host
}

<#  アップデートの処理  #>
function updateProc($arg) {
    Write-Host "ドライブを検出しています…"
    $sec_env = getSourceEnv
    $drives = getDrives $arg["media-only"]
    $cnt = 0
    HashDsp $drives
        foreach($drv in $drives.Keys) {
        # ドライブに既にインストールされていて、アップデート可能なら、アップデート（確認）
        if($drives[$drv]["exists"] -and $drives[$drv]["can-install"]) {
            install $drives[$drv] $sec_env
            $cnt++
        }
    }
    if($cnt -eq 0) {
        Write-Host "アップデート可能なドライブはありません。"
    }
}

<#  インストール処理    #>
function installProc($arg) {
    $sec_env = getSourceEnv
    while($true) {
        $drives = getDrives $arg["media-only"]
        HashDsp $drives
        $prc = Read-Host "処理を選択してください。（I:インストール／U:アプリを削除／その他：終了）"
        $prc = $prc.ToLower()
        if($prc -eq "i") {
            $drv_letter = askDrive $drives $true
            if($drv_letter -ne "") {
                install $drives[$drv_letter] $sec_env
            }
        } elseif ( $prc -eq "u") {
            $drv_letter = askDrive $drives $false
            if($drv_letter -ne "") {
                uninstall $drives[$drv_letter]
            }
        } else {
            break
        }
    }
}

<#  引数の取得  #>
function getArguments {
    # 引数のデフォルト。無指定の場合はアップデートのみ
    $arg = @{
        "update" = $true
        "install" = $false
        "media-only" = $false
    }
    # もし引数があったら、アップデートもしない
    if($Script:Args.length -gt 0) {
        $arg["update"] = $false
    }

    # オプションの解析
    for($i=0; $i -lt $Script:Args.length; $i++) {
        $opt = $Script:Args[$i].Replace("-","")
        switch($opt) {
            "full" {   $arg["update"] = $true; $arg["install"] = $true;    }
            "update" {   $arg["update"] = $true;     }
            "install" {   $arg["install"] = $true;     }
            "mediaonly" {   $arg["media-only"] = $true;     }
            Default { Write-Host "不正なオプション ${opt} が指定されました。" }
        }
    }
    return $arg
}

$arg = getArguments

# アップデート
if($arg["update"]) {
    Write-Host "アップデート処理を呼び出します"
    updateProc $arg
}

# インストール/アンインストール
if($arg["install"]) {
    Write-Host "インストール処理を呼び出します"
    installProc $arg
}

<#
    発注書.xlsmランチャー
    copyright 2024 C.Nagata
    2024.6.16   Initial writing.
#>

Set-StrictMode -Version 1.0

# アセンブリ
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# PowerShellのパス
$POWER_SHELL = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
$MAIN_WINDOW_TITLE = "発注書ランチャ"
# ライブラリのインポート
# 設定ファイルも読み込まれ、グローバル変数に設定される
. "$($PSScriptRoot)\ofx_lib.ps1"

<#
    コンソールウィンドウを最小化するためのコード
    https://qiita.com/AWtnb/items/34fe77fda53820a8546e
#>
function Get-ConsoleWindowHandle {
    $p = Get-Process -Id $PID
    $i = 0
    while ($p.MainWindowHandle -eq 0) {
        if ($i++ -gt 10) {
            return $null
        }
        if("Parent" -in $p.PSobject.Properties.Name) {
            $p = $p.Parent
        } else {
            break
        }
    }
    if(0 -eq $p.MainWindowHandle) {
        $procs = (Get-Process | Where-Object {$_.MainWindowTitle -match $MAIN_WINDOW_TITLE})
        if($null -eq $procs) {
            return $null
        }
        $p = $procs[0]
    }
    return $p.MainWindowHandle
}
$Global:CONSOLE_HWND = Get-ConsoleWindowHandle

if(-not ('Console.Window' -as [type])) {
    Add-Type -Name Window -Namespace Console -MemberDefinition `
@'
[DllImport("user32.dll")]
private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, int wParam, int lParam);

public static void Minimize(IntPtr hwnd) {
    SendMessage(hwnd, 0x0112, 0xF020, 0);
}
'@
}
function Hide-ConsoleWindow {
    if ($Global:CONSOLE_HWND -and ($env:TERM_PROGRAM -ne "vscode")) {
        [Console.Window]::Minimize($Global:CONSOLE_HWND)
    }
}

<#   Excelファイルに渡す設定ファイルを作成し、保存する   #>
function makeBookConf($dict) {
    # CONFIGファイルのenv > xxx > excel-book > settings　を切り出し
    $settings = $CONFIG.env.$THIS_ENV."excel-book".settings
    if($null -eq $settings) {
        Write-Host "${CONF_PATH} の中に env>${THIS_ENV}>excel-book セクションがありません。"
        exit
    }
    # 辞書で変数をデコードする
    $settings = convertVars $settings $dict
    # $settings | Add-Member -MemberType NoteProperty -Name 'env' -Value $THIS_ENV
    # JSONテキストに変換して保存する
    $settings | ConvertTo-Json -Depth 32 | Out-File $BOOK_CONF_PATH -Encoding default
}

<#  ファイルのパスの取得    #>
function getFilePath($key) {
    return expandPath $CONFIG.install.files.$key $THIS_DICT $true
}

<#  ディレクトリのパスの取得    #>
function getDirPath($key, $env=$null, $dict=$null) {
    if($null -eq $env) {
        $env = $THIS_ENV
    }
    if($null -eq $dict) {
        $dict = $THIS_DICT
    }
    return expandPath $CONFIG.env.$env.dirs.$key $dict $true
}

<#  初期化処理
#>
function init {
    #   ないか、或いはUSBなら、ブック用の設定ファイルを作る
    if( -not (Test-Path -LiteralPath $BOOK_CONF_PATH) -or $THIS_ENV -eq "USB" ) {
        Write-Host "Excel Bookの設定ファイルを作成します…"
        $dict = makePathDict $CONFIG.env.$THIS_ENV.dirs
        makeBookConf $dict
    }
    #   環境変数の設定
    [Environment]::SetEnvironmentVariable($ENVVAR_NAME, $BOOK_CONF_PATH, 'User')    
}

<#  終了処理
#>
function exitProc {
    #   環境変数の削除
    [Environment]::SetEnvironmentVariable($ENVVAR_NAME, "", 'User')    
}

<#  オンラインのチェック    #>
function checkOnline {
    $new_status = isOnline
    # オンラインの状態が変わらなければ終了
    if($new_status -eq $STATUS.is_online) {
        return $false
    }
    # 状態の更新
    $STATUS.is_online = $new_status
    return $true
}

<#  リムーバブルメディアのチェック    #>
function checkMedia( $first_time = $false ) {
    $changed = $false
    $STATUS.usb_exists = $true
    $new_drives = @("C:")
    $usbs = (Get-WmiObject CIM_LogicalDisk | Where-Object DriveType -eq 2).DeviceID
    if($null -eq $usbs) {
        $STATUS.usb_exists = $false
    } else {
        $new_drives += $usbs
    }
    # 既存のメディア情報の更新
    $keys = $STATUS.drives.Keys.clone()
    foreach($key in $keys) {
        if($new_drives.Contains($key)) {
            # 更新
            $STATUS.drives[$key] = newDriveInfo $key
        } else {
            # なくなっていたら削除
            $STATUS.drives.Remove($key)
            $changed = $true
        }
    }
    # 新規メディアの情報を取得
    $is_new = $false
    foreach($key in $new_drives) {
        if(-not $STATUS.drives.ContainsKey($key)) {
            $drv_inf = newDriveInfo $key
            $STATUS.drives.Add($key, $drv_inf)
            $is_new = $true
        }
    }
    # 新規メディアがあれば、アップデートする
    if($is_new -and (-not $first_time)) {
        callInstaller "--update --mediaonly" | Out-Null
        $changed = $true
    }
    # 表示の切り替え
    $STATUS.file_exists = $false
    $keys = @()
    $keys += ($STATUS.drives.Keys | Sort-Object)
    for($i=0; $i -lt $keys.length -and $i -lt 3; $i++) {
        $res = switchDriveLabel $i $STATUS.drives[$keys[$i]]
        if($res -eq 2) {
            $STATUS.file_exists = $true
        }
    }
    for(; $i -lt 3; $i++) {
        switchDriveLabel $i $null
    }
    # 戻り値が配列になる…
    return $changed
}

function newDriveInfo($drv_letter) {
    $drv_inf = getDriveInfo $drv_letter
    $dirs = makeDirList $drv_inf["env"] $drv_letter
    $file_exists = @()
    if($drv_inf["exists"]) {
        foreach($dir in $dirs) {
            $files = Get-ChildItem -Path $dir -Name
            if($null -eq $files) {
                $file_exists += $false
            } else {
                $file_exists += $true
            }
        }
    }
    $drv_inf.Add("file_exists", $file_exists)
    return $drv_inf
}


<#  ファイルの移動  #>
function moveFiles($drv_inf, $net_dirs) {
    $src_dirs = makeDirList $drv_inf["env"] $drv_inf["letter"]
    for($i=0; $i -lt $src_dirs.length; $i++) {
        $src = $src_dirs[$i]
        $dest = $net_dirs[$i]
        Write-Host "#${i} moving files in ${src}"
        Write-Host "to ${dest}"
        if((Split-Path -Leaf $src) -ne (Split-Path -Leaf $src)){
            Write-Host "Wrong folder name." -ForegroundColor "Red"
            break
        } elseif( -not (Test-Path -LiteralPath $src)) {
            Write-Host "A source folder does not exist." -ForegroundColor "Yellow"
        } elseif( -not (Test-Path -LiteralPath $dest)) {
            Write-Host "A destination folder does not exist." -ForegroundColor "Yellow"
        } else {
            Move-Item -Path "${src}\*.*" -Destination $dest
        }
    }
}

<#  コピーするディレクトリのリストを作る    #>
function makeDirList($env, $drive_letter) {
    $dict = makePathDict $CONFIG.env.$env.dirs $drive_letter
    $lst = (getDirPath "save-dirs" $env $dict) -split ";"
    $lst += (getDirPath "letter-dirs" $env $dict) -split ";"
    return $lst
}

<#  ネットワークフォルダにアップロードする    #>
function moveFilesToNet {
    $drv_infs = getDrives
    $net_dirs = makeDirList "online" "\\"
    foreach($key in $drv_infs.Keys) {
        if($drv_infs[$key]["exists"]) {
            moveFiles $drv_infs[$key] $net_dirs
        }
    }
}

function callInstaller($str_arg) {
    # 呼び出すスクリプトを指定
    $script = getFilePath "install-ps1"
    $Argument   = "-Command `"${script}`" ${str_arg}"
    Start-Process -FilePath $POWER_SHELL -ArgumentList $Argument

}

# $FONT_FAMILY = "游ゴシック Medium"
# $FONT_FAMILY = "MSPゴシック"
$FONT_FAMILY = "メイリオ"
$FONT_SIZE = 10

<#  フォームのための定数    #>
$MARGIN_W = 20;   $PAD_COL = 16;   $COL_W = 120;
$MARGIN_H = 10;   $PAD_ROW = 12;   $ROW_H = 50;

<#  描画位置の計算  #>
function getX($idx) {
    return ($MARGIN_W + ($COL_W + $PAD_COL) * $idx  )
}
function getY($idx) {
    return ($MARGIN_H +  ($ROW_H + $PAD_ROW) * $idx  )
}
function getWidth($num) {
    return ( ($COL_W + $PAD_COL) * $num - $PAD_COL)
}
function getHeight($num) {
    if($num -lt 1.0) {
        return ($ROW_H * $num)
    }
    return ( ($ROW_H + $PAD_ROW) * $num - $PAD_ROW)
  
}

<#  フォームの作成    #>
function newForm($dlg, $name, $x, $y, $width, $height, $caption, $objIcon) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $caption
    $form.ClientSize=New-Object System.Drawing.Size($width,$height)
    $form.StartPosition = "Manual"
    $form.Location = "${x},${y}"
    $form.MinimizeBox = $true
    $form.MaximizeBox = $false
    $form.Icon = $objIcon
    foreach($key in $dlg.psobject.properties.name) {
        $form.Controls.Add($dlg.$key)
    }
    $dlg | Add-Member -MemberType NoteProperty -Name $name -Value $form
}

<#  ボタンの作成    #>
function newButton($dlg, $name, $x, $y, $width, $height, $text) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Location = "${x},${y}"
    $btn.Size = New-Object System.Drawing.Size($width, $height)
    $btn.Font = New-Object System.Drawing.Font($FONT_FAMILY, $FONT_SIZE)
    $btn.Text = $text
    $btn.Enabled = $false
    $dlg | Add-Member -MemberType NoteProperty -Name $name -Value $btn
}

<#  ラベルの作成    #>
function newLabel($dlg, $name, $x, $y, $width, $height, $text) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = "${x},${y}"
    $lbl.Size = New-Object System.Drawing.Size($width, $height)
    $lbl.Font = New-Object System.Drawing.Font($FONT_FAMILY, $FONT_SIZE)
    $lbl.BackColor = "#F8F8F8"
    $lbl.Text = $text
    $dlg | Add-Member -MemberType NoteProperty -Name $name -Value $lbl
}

<#  ラベルのON/OFF  #>
function switchDriveLabel($idx, $drv_inf=$null) {
    $name = "lblDrive" + ([string]$idx).Trim()
    $col = "#080808"
    $txt = ""
    $res = 0
    if($null -ne $drv_inf) {
        $txt = $drv_inf["letter"] + " "
        if(-not $drv_inf["exists"]) {
            $txt += "‐"
        } else {
            $res = 1
            $txt += "●"
            $col = "#000080"
            foreach($exists in $drv_inf["file_exists"]) {
                if($exists) {
                    $res = 2
                    $txt += "■"
                    $col = "#800000"
                } else {
                    $txt += "□"
                }
            }
        }
    }
    $DIALOG.$name.foreColor = $col
    $DIALOG.$name.Text = $txt
    return $res
}

<#  タイマーの作成  #>
function newTimer($dlg, $name, $interval_ms) {
    $timer = New-Object Windows.Forms.Timer
    $timer.Interval = $interval_ms
    $timer.Enabled = $false
    $dlg | Add-Member -MemberType NoteProperty -Name $name -Value $timer
}

<#  フォームを作成する
#>
function makeDialog {
    $dlg =  New-Object PSCustomObject

    #   フォーム上のパーツ
    # ラベル
    newLabel $dlg "lblDrive0" `
        (getX 0) (getY 2) (getWidth 1) (getHeight 0.5) ""
    newLabel $dlg "lblDrive1" `
        (getX 0) (getY 2.5) (getWidth 1) (getHeight 0.5) ""
    newLabel $dlg "lblDrive2" `
        (getX 0) (getY 3) (getWidth 1) (getHeight 0.5) ""
    # 新規作成ボタン
    newButton $dlg "btnNew" `
        (getX 0) (getY 0) (getWidth 1) (getHeight 1) "発注書の`r`n新規作成"
    # アップロードボタン
    newButton $dlg "btnUpload" `
        (getX 1) (getY 0) (getWidth 1) (getHeight 1) "共有フォルダに`r`nアップロード"
    # ローカルフォルダボタン
    newButton $dlg "btnFolder" `
        (getX 0) (getY 1) (getWidth 1) (getHeight 1) "ローカル`r`nフォルダを開く"
    # 共有フォルダボタン
    newButton $dlg "btnShareFolder" `
        (getX 1) (getY 1) (getWidth 1) (getHeight 1) "共有フォルダを`r`n開く"
    # インストールボタン
    newButton $dlg "btnInstall" `
        (getX 1) (getY 2) (getWidth 1) (getHeight 1) "USBにｲﾝｽﾄｰﾙ`r`n/ ｱﾝｲﾝｽﾄｰﾙ"

    # フォーム
    $w = (getX 2) - $PAD_COL + $MARGIN_W
    $h = (getY 3) + (getHeight 0.5) + $MARGIN_H
    $scr = [System.Windows.Forms.SystemInformation]::WorkingArea.Size
    $x = $scr.Width - $w
    $y = 0
    $icon = new-object System.Drawing.Icon ($script:PSScriptRoot + "\ofx.ico")
    newForm $dlg "form" $x $y $w $h "発注書ランチャー" $icon
    # タイマー
    newTimer $dlg "timer" 2500

    return $dlg
}

<#  タイマーによる処理  #>
function timerProc($first_time = $false) {
    if(checkOnline) {
        if($STATUS.is_online) {
            Write-Host "network folder connected."
        } else {
            Write-Host "network folder disconnected."
        }
    }
    $changed = checkMedia $first_time 
    if($changed) {
        $drives = $STATUS.drives.Keys -join ","
        Write-Host  "media changed: ${drives}"
    }
    # 表示とボタンの切り替え
    $DIALOG.btnInstall.Enabled = ($STATUS.usb_exists -and ($THIS_ENV -ne "USB"))
    $DIALOG.btnUpload.Enabled = ($STATUS.is_online -and $STATUS.file_exists)
}

Write-Host "Script starts."

# この環境の辞書を作成
$THIS_DICT = makePathDict $CONFIG.env.$THIS_ENV.dirs
# 状態管理
$STATUS = [PSCustomObject]@{
    is_online = $false
    drives = @{}
    usb_exists = $false
    file_exists = $false
}
# フォームを作成
$DIALOG = makeDialog

# イベント処理

# フォームが表示されたときの処理
$DIALOG.form.Add_Shown({
    Write-Host "Initiarizing."
    $DIALOG.form.text = "起動しています…"
    init
    $DIALOG.btnNew.Enabled = $true
    $DIALOG.btnFolder.Enabled = $true
    $DIALOG.timer.Enabled = $true
    # タイマー処理の最初の一回（アップデートなし）
    timerProc $true
    $DIALOG.timer.Start()
    $DIALOG.form.text = "発注書ランチャー"
    Write-Host "Start."
})
# フォームを閉じるときの処理
$DIALOG.form.Add_Closing({
    $DIALOG.form.text = "終了しています…"
    $DIALOG.btnNew.Enabled = $false
    $DIALOG.btnFolder.Enabled = $false
    $DIALOG.btnUpload.Enabled = $false
    $DIALOG.btnInstall.Enabled = $false
    $DIALOG.timer.Stop()
    $DIALOG.timer.Enabled = $false
    exitProc
    Write-Host "Finished."
})
# 新規作成ボタンのクリック
$DIALOG.btnNew.Add_Click({
    Invoke-Item (getFilePath "excel-book")
})
# ローカルフォルダボタンのクリック
$DIALOG.btnFolder.Add_Click({
    $paths = (getDirPath "save-dirs") -split ";"
    Invoke-Item $paths[0]
})
# アップロードボタンのクリック
$DIALOG.btnUpload.Add_Click({
    moveFilesToNet
})
# インストールボタンのクリック
$DIALOG.btnInstall.Add_Click({
    callInstaller "--install --mediaonly"
})
# タイマー処理
$DIALOG.timer.Add_Tick({
    timerProc
})

# フォームを表示
Hide-ConsoleWindow
$DIALOG.form.ShowDialog()
# 終了
Remove-Variable -Name DIALOG
Remove-Variable -Name THIS_DICT

Write-Host "Totaly finished."

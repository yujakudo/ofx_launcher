<#
    発注書.xlsmランチャー
    copyright 2024 C.Nagata
    2024.6.16   Initial writing.
#>

# ライブラリのインポート
# 設定ファイルも読み込まれ、グローバル変数に設定されている
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
        $p = $p.Parent
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
    # 辞書で変数をデコードする
    $settings = convertVars $CONFIG.env.$THIS_ENV."excel-book".settings $dict
    # $settings | Add-Member -MemberType NoteProperty -Name 'env' -Value $THIS_ENV
    # JSONテキストに変換して保存する
    $settings | ConvertTo-Json -Depth 32 | Out-File $BOOK_CONF_PATH -Encoding default
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

<#  フォームのための定数    #>
$MARGIN_W = 20;   $PAD_COL = 16;   $COL_W = 120;
$MARGIN_H = 10;   $PAD_ROW = 12;   $ROW_H = 54;

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

<#  ボタンの作成    #>
function newButton($x, $y, $width, $height, $text) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Location = "${x},${y}"
    $btn.Size = New-Object System.Drawing.Size($width, $height)
    $btn.Font = New-Object System.Drawing.Font("游ゴシック Medium", 12)
    $btn.Text = $text
    return $btn
}
<#  ラベルの作成    #>
function newLabel($x, $y, $width, $height, $text) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = "${x},${y}"
    $lbl.Size = New-Object System.Drawing.Size($width, $height)
    $lbl.Font = New-Object System.Drawing.Font("游ゴシック Medium", 12)
    $lbl.BackColor = "#F8F8F8"
    $lbl.Text = $text
    return $lbl
}

<#  フォーム
#>
function makeForm {

    # アセンブリ
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $scr = [System.Windows.Forms.SystemInformation]::WorkingArea.Size

    # フォーム
    $w = (getX 2) - $PAD_COL + $MARGIN_W
    $h = (getY 2) + (getHeight 0.5) + $MARGIN_H
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "発注書ランチャー"
    # $form.Size = New-Object System.Drawing.Size($w,$h)
    $form.ClientSize=New-Object System.Drawing.Size($w,$h)
    $form.MinimizeBox = $true
    $scr = [System.Windows.Forms.SystemInformation]::WorkingArea.Size
    $x = $scr.Width - $w;   $y = 0;
    $form.StartPosition = "Manual"
    $form.Location = "${x},${y}"
    $icon = new-object System.Drawing.Icon ($script:PSScriptRoot + "\ofx.ico")
    $form.Icon = $icon

    # $form.Opacity = 0.2
    $form.Add_Shown({
        Write-Host "Initiarizing."
        init
        Write-Host "Start."
    })
    $form.Add_Closing({
        exitProc
        Write-Host "Done."
    })
    # ラベル
    $lblOnline = newLabel `
        (getX 0) (getY 2) (getWidth 1) (getHeight 0.5) "Offline"
    $label.forecolor = "#080808"
    $lblUSB = newLabel `
        (getX 1) (getY 2) (getWidth 1) (getHeight 0.5) "USB None"
    $lblUSB.forecolor = "#080808"
        # 新規作成ボタン
    $btnNew = newButton `
        (getX 0) (getY 0) (getWidth 1) (getHeight 1) "新規作成"
    $btnNew.Add_Click({

    })
    $btnNew.Enabled = $false
    # ローカルフォルダボタン
    $btnFolder = newButton `
        (getX 1) (getY 0) (getWidth 1) (getHeight 1) "ローカル`r`nフォルダ"
    $btnFolder.Add_Click({

    })
    $btnFolder.Enabled = $false
    # アップロードボタン
    $btnUpload = newButton `
        (getX 0) (getY 1) (getWidth 1) (getHeight 1) "アップロード"
    $btnUpload.Add_Click({

    })
    $btnUpload.Enabled = $false
    # インストールボタン
    $btnInstall = newButton `
        (getX 1) (getY 1) (getWidth 1) (getHeight 1) "インストール"
    $btnInstall.Add_Click({

    })
    $btnInstall.Enabled = $false

    $form.Controls.Add($lblOnline)
    $form.Controls.Add($lblUSB)
    $form.Controls.Add($btnNew)
    $form.Controls.Add($btnFolder)
    $form.Controls.Add($btnUpload)
    $form.Controls.Add($btnInstall)

    # タイマー
    $timer = New-Object Windows.Forms.Timer
    $timer.Add_Tick({
        
    })
    $timer.Interval = 1000
    $timer.Enabled = $true
    $timer.Start()

    return $form
}

$form = makeForm
$form.ShowDialog()

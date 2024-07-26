<#
    ������.xlsm�����`���[
    copyright 2024 C.Nagata
    2024.6.16   Initial writing.
#>

Set-StrictMode -Version 1.0

# �A�Z���u��
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# PowerShell�̃p�X
$POWER_SHELL = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
$MAIN_WINDOW_TITLE = "�����������`��"
# ���C�u�����̃C���|�[�g
# �ݒ�t�@�C�����ǂݍ��܂�A�O���[�o���ϐ��ɐݒ肳���
. "$($PSScriptRoot)\ofx_lib.ps1"

<#
    �R���\�[���E�B���h�E���ŏ������邽�߂̃R�[�h
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

<#   Excel�t�@�C���ɓn���ݒ�t�@�C�����쐬���A�ۑ�����   #>
function makeBookConf($dict) {
    # CONFIG�t�@�C����env > xxx > excel-book > settings�@��؂�o��
    $settings = $CONFIG.env.$THIS_ENV."excel-book".settings
    if($null -eq $settings) {
        Write-Host "${CONF_PATH} �̒��� env>${THIS_ENV}>excel-book �Z�N�V����������܂���B"
        exit
    }
    # �����ŕϐ����f�R�[�h����
    $settings = convertVars $settings $dict
    # $settings | Add-Member -MemberType NoteProperty -Name 'env' -Value $THIS_ENV
    # JSON�e�L�X�g�ɕϊ����ĕۑ�����
    $settings | ConvertTo-Json -Depth 32 | Out-File $BOOK_CONF_PATH -Encoding default
}

<#  �t�@�C���̃p�X�̎擾    #>
function getFilePath($key) {
    return expandPath $CONFIG.install.files.$key $THIS_DICT $true
}

<#  �f�B���N�g���̃p�X�̎擾    #>
function getDirPath($key, $env=$null, $dict=$null) {
    if($null -eq $env) {
        $env = $THIS_ENV
    }
    if($null -eq $dict) {
        $dict = $THIS_DICT
    }
    return expandPath $CONFIG.env.$env.dirs.$key $dict $true
}

<#  ����������
#>
function init {
    #   �Ȃ����A������USB�Ȃ�A�u�b�N�p�̐ݒ�t�@�C�������
    if( -not (Test-Path -LiteralPath $BOOK_CONF_PATH) -or $THIS_ENV -eq "USB" ) {
        Write-Host "Excel Book�̐ݒ�t�@�C�����쐬���܂��c"
        $dict = makePathDict $CONFIG.env.$THIS_ENV.dirs
        makeBookConf $dict
    }
    #   ���ϐ��̐ݒ�
    [Environment]::SetEnvironmentVariable($ENVVAR_NAME, $BOOK_CONF_PATH, 'User')    
}

<#  �I������
#>
function exitProc {
    #   ���ϐ��̍폜
    [Environment]::SetEnvironmentVariable($ENVVAR_NAME, "", 'User')    
}

<#  �I�����C���̃`�F�b�N    #>
function checkOnline {
    $new_status = isOnline
    # �I�����C���̏�Ԃ��ς��Ȃ���ΏI��
    if($new_status -eq $STATUS.is_online) {
        return $false
    }
    # ��Ԃ̍X�V
    $STATUS.is_online = $new_status
    return $true
}

<#  �����[�o�u�����f�B�A�̃`�F�b�N    #>
function checkMedia {
    $changed = $false
    $STATUS.usb_exists = $true
    $new_drives = @("C:")
    $usbs = (Get-WmiObject CIM_LogicalDisk | Where-Object DriveType -eq 2).DeviceID
    if($null -eq $usbs) {
        $STATUS.usb_exists = $false
    }
    $new_drives += $usbs
    # �����̃��f�B�A���̍X�V
    foreach($key in $STATUS.drives.Keys) {
        if($new_drives.Contains($key)) {
            # �Ȃ��Ȃ��Ă�����폜
            $STATUS.drives.Remove($key)
            $changed = $true
        } else {
            # �X�V
            $STATUS.drives[$key] = newDriveInfo $key
        }
    }
    # �V�K���f�B�A�̏����擾
    $is_new = $false
    foreach($key in $new_drives) {
        if(-not $STATUS.drives.ContainsKey($key)) {
            $drv_inf = newDriveInfo $key
            $STATUS.drives.Add($key, $drv_inf)
            $is_new = $true
        }
    }
    # �V�K���f�B�A������΁A�A�b�v�f�[�g����
    if($is_new) {
        callInstaller "--update --mediaonly"
        $changed = $true
    }
    # �\���̐؂�ւ�
    $STATUS.file_exists = $false
    $keys = $STATUS.drives.Keys | Sort-Object
    for($i=0; $i -lt $keys.length; $i++) {
        $res = switchDriveLabel $i $STATUS.drives[$keys[$i]]
        if($res -eq 2) {
            $STATUS.file_exists = $true
        }
    }
    for(; $i -lt 4; $i++) {
        switchDriveLabel $i $null
    }
    return $changed
}

function newDriveInfo($drv_letter) {
    $drv_inf = getDriveInfo $drv_letter
    $dirs = makeDirList $drv_inf["env"] $drv_letter
    $file_exists = @()
    if($drv_inf["exists"]) {
        foreach($dir in $dirs) {
            $files = Get-ChildItem -Parent $dir -Name
            if($null -eq $files) {
                $file_exists += $true
            } else {
                $file_exists += $false
            }
        }
    }
    $drv_inf.Add("file_exists", $file_exists)
    return $drv_inf
}


<#  �t�@�C���̈ړ�  #>
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

<#  �R�s�[����f�B���N�g���̃��X�g�����    #>
function makeDirList($env, $drive_letter) {
    $dict = makePathDict $CONFIG.env.$env.dirs $drive_letter
    $lst = (getDirPath "save-dirs" $env $dict) -split ";"
    $lst += (getDirPath "letter-dirs" $env $dict) -split ";"
    return $lst
}

<#  �l�b�g���[�N�t�H���_�ɃA�b�v���[�h����    #>
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
    # �Ăяo���X�N���v�g���w��
    $script = getFilePath "install-ps1"
    $Argument   = "-Command `"${script}`" ${str_arg}"
    Start-Process -FilePath $POWER_SHELL -ArgumentList $Argument

}

# $FONT_FAMILY = "���S�V�b�N Medium"
# $FONT_FAMILY = "MSP�S�V�b�N"
$FONT_FAMILY = "���C���I"
$FONT_SIZE = 10

<#  �t�H�[���̂��߂̒萔    #>
$MARGIN_W = 20;   $PAD_COL = 16;   $COL_W = 120;
$MARGIN_H = 10;   $PAD_ROW = 12;   $ROW_H = 50;

<#  �`��ʒu�̌v�Z  #>
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

<#  �t�H�[���̍쐬    #>
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

<#  �{�^���̍쐬    #>
function newButton($dlg, $name, $x, $y, $width, $height, $text) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Location = "${x},${y}"
    $btn.Size = New-Object System.Drawing.Size($width, $height)
    $btn.Font = New-Object System.Drawing.Font($FONT_FAMILY, $FONT_SIZE)
    $btn.Text = $text
    $btn.Enabled = $false
    $dlg | Add-Member -MemberType NoteProperty -Name $name -Value $btn
}

<#  ���x���̍쐬    #>
function newLabel($dlg, $name, $x, $y, $width, $height, $text) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = "${x},${y}"
    $lbl.Size = New-Object System.Drawing.Size($width, $height)
    $lbl.Font = New-Object System.Drawing.Font($FONT_FAMILY, $FONT_SIZE)
    $lbl.BackColor = "#F8F8F8"
    $lbl.Text = $text
    $dlg | Add-Member -MemberType NoteProperty -Name $name -Value $lbl
}

<#  ���x����ON/OFF  #>
function switchDriveLabel($idx, $drv_inf=$null) {
    $name = "lblDrive" + ([string]$idx).Trim()
    $col = "#080808"
    $txt = ""
    $res = 0
    if($null -ne $drv_inf) {
        $txt = $drv_inf["letter"] + " "
        if($drv_inf["exists"]) {
            $txt += "�]"
        } else {
            $res = 1
            $txt += "��"
            $col = "#00FF00"
            foreach($exists in $drv_inf["file_exists"]) {
                if($exists) {
                    $res = 2
                    $txt += "��"
                    $col = "#FF0000"
                } else {
                    $txt += "��"
                }
            }
        }
    }
    $DIALOG.$name.foreColor = $col
    $DIALOG.$name.Text = $txt
    return $res
}

<#  �^�C�}�[�̍쐬  #>
function newTimer($dlg, $name, $interval_ms) {
    $timer = New-Object Windows.Forms.Timer
    $timer.Interval = $interval_ms
    $timer.Enabled = $false
    $dlg | Add-Member -MemberType NoteProperty -Name $name -Value $timer
}

<#  �t�H�[�����쐬����
#>
function makeDialog {
    $dlg =  New-Object PSCustomObject

    #   �t�H�[����̃p�[�c
    # ���x��
    newLabel $dlg "lblDrive0" `
        (getX 0) (getY 2) (getWidth 1) (getHeight 0.5) ""
    newLabel $dlg "lblDrive1" `
        (getX 0) (getY 2.5) (getWidth 1) (getHeight 0.5) ""
    newLabel $dlg "lblDrive2" `
        (getX 0) (getY 3) (getWidth 1) (getHeight 0.5) ""
    # �V�K�쐬�{�^��
    newButton $dlg "btnNew" `
        (getX 0) (getY 0) (getWidth 1) (getHeight 1) "��������`r`n�V�K�쐬"
    # �A�b�v���[�h�{�^��
    newButton $dlg "btnUpload" `
        (getX 1) (getY 0) (getWidth 1) (getHeight 1) "���L�t�H���_��`r`n�A�b�v���[�h"
    # ���[�J���t�H���_�{�^��
    newButton $dlg "btnFolder" `
        (getX 0) (getY 1) (getWidth 1) (getHeight 1) "���[�J��`r`n�t�H���_���J��"
    # ���L�t�H���_�{�^��
    newButton $dlg "btnShareFolder" `
        (getX 1) (getY 1) (getWidth 1) (getHeight 1) "���L�t�H���_��`r`n�J��"
    # �C���X�g�[���{�^��
    newButton $dlg "btnInstall" `
        (getX 1) (getY 2) (getWidth 1) (getHeight 1) "USB�ɲݽİ�`r`n/ �ݲݽİ�"

    # �t�H�[��
    $w = (getX 2) - $PAD_COL + $MARGIN_W
    $h = (getY 3) + (getHeight 0.5) + $MARGIN_H
    $scr = [System.Windows.Forms.SystemInformation]::WorkingArea.Size
    $x = $scr.Width - $w
    $y = 0
    $icon = new-object System.Drawing.Icon ($script:PSScriptRoot + "\ofx.ico")
    newForm $dlg "form" $x $y $w $h "�����������`���[" $icon
    # �^�C�}�[
    newTimer $dlg "timer" 2500

    return $dlg
}

Write-Host "Script starts."

# ���̊��̎������쐬
$THIS_DICT = makePathDict $CONFIG.env.$THIS_ENV.dirs
# ��ԊǗ�
$STATUS = [PSCustomObject]@{
    is_online = $false
    drives = @{}
    usb_exists = $false
    file_exists = $false
}
# �t�H�[�����쐬
$DIALOG = makeDialog

# �C�x���g����

# �t�H�[�����\�����ꂽ�Ƃ��̏���
$DIALOG.form.Add_Shown({
    Write-Host "Initiarizing."
    $DIALOG.form.text = "�N�����Ă��܂��c"
    init
    $DIALOG.btnNew.Enabled = $true
    $DIALOG.btnFolder.Enabled = $true
    $DIALOG.timer.Enabled = $true
    $DIALOG.timer.Start()
    $DIALOG.form.text = "�����������`���["
    Write-Host "Start."
})
# �t�H�[�������Ƃ��̏���
$DIALOG.form.Add_Closing({
    $DIALOG.form.text = "�I�����Ă��܂��c"
    $DIALOG.btnNew.Enabled = $false
    $DIALOG.btnFolder.Enabled = $false
    $DIALOG.btnUpload.Enabled = $false
    $DIALOG.btnInstall.Enabled = $false
    $DIALOG.timer.Stop()
    $DIALOG.timer.Enabled = $false
    exitProc
    Write-Host "Finished."
})
# �V�K�쐬�{�^���̃N���b�N
$DIALOG.btnNew.Add_Click({
    Invoke-Item (getFilePath "excel-book")
})
# ���[�J���t�H���_�{�^���̃N���b�N
$DIALOG.btnFolder.Add_Click({
    $paths = (getDirPath "save-dirs") -split ";"
    Invoke-Item $paths[0]
})
# �A�b�v���[�h�{�^���̃N���b�N
$DIALOG.btnUpload.Add_Click({
    moveFilesToNet
})
# �C���X�g�[���{�^���̃N���b�N
$DIALOG.btnInstall.Add_Click({
    callInstaller "--install --mediaonly"
})
# �^�C�}�[����
$DIALOG.timer.Add_Tick({
    if(checkOnline) {
        Write-Host "network folder: ${STATUS.is_online}"
    }
    if(checkMedia) {
        $drives = $STATUS.drives.Keys -join ","
        Write-Host  "removable media: ${drives}"
    }
    # �\���ƃ{�^���̐؂�ւ�
    $DIALOG.btnInstall.Enabled = ($STATUS.usb_exists -and ($THIS_ENV -ne "USB"))
    $DIALOG.btnUpload.Enabled = ($STATUS.is_online -and $STATUS.file_exists)
})
# �t�H�[����\��
Hide-ConsoleWindow
$DIALOG.form.ShowDialog()
# �I��
Remove-Variable -Name DIALOG
Remove-Variable -Name THIS_DICT

Write-Host "Totaly finished."

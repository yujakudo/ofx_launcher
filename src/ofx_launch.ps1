<#
    ������.xlsm�����`���[
    copyright 2024 C.Nagata
    2024.6.16   Initial writing.
#>

# ���C�u�����̃C���|�[�g
# �ݒ�t�@�C�����ǂݍ��܂�A�O���[�o���ϐ��ɐݒ肳��Ă���
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

<#   Excel�t�@�C���ɓn���ݒ�t�@�C�����쐬���A�ۑ�����   #>
function makeBookConf($dict) {
    # CONFIG�t�@�C����env > xxx > excel-book > settings�@��؂�o��
    # �����ŕϐ����f�R�[�h����
    $settings = convertVars $CONFIG.env.$THIS_ENV."excel-book".settings $dict
    # $settings | Add-Member -MemberType NoteProperty -Name 'env' -Value $THIS_ENV
    # JSON�e�L�X�g�ɕϊ����ĕۑ�����
    $settings | ConvertTo-Json -Depth 32 | Out-File $BOOK_CONF_PATH -Encoding default
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

<#  �t�H�[���̂��߂̒萔    #>
$MARGIN_W = 20;   $PAD_COL = 16;   $COL_W = 120;
$MARGIN_H = 10;   $PAD_ROW = 12;   $ROW_H = 54;

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

<#  �{�^���̍쐬    #>
function newButton($x, $y, $width, $height, $text) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Location = "${x},${y}"
    $btn.Size = New-Object System.Drawing.Size($width, $height)
    $btn.Font = New-Object System.Drawing.Font("���S�V�b�N Medium", 12)
    $btn.Text = $text
    return $btn
}
<#  ���x���̍쐬    #>
function newLabel($x, $y, $width, $height, $text) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = "${x},${y}"
    $lbl.Size = New-Object System.Drawing.Size($width, $height)
    $lbl.Font = New-Object System.Drawing.Font("���S�V�b�N Medium", 12)
    $lbl.BackColor = "#F8F8F8"
    $lbl.Text = $text
    return $lbl
}

<#  �t�H�[��
#>
function makeForm {

    # �A�Z���u��
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $scr = [System.Windows.Forms.SystemInformation]::WorkingArea.Size

    # �t�H�[��
    $w = (getX 2) - $PAD_COL + $MARGIN_W
    $h = (getY 2) + (getHeight 0.5) + $MARGIN_H
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "�����������`���["
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
    # ���x��
    $lblOnline = newLabel `
        (getX 0) (getY 2) (getWidth 1) (getHeight 0.5) "Offline"
    $label.forecolor = "#080808"
    $lblUSB = newLabel `
        (getX 1) (getY 2) (getWidth 1) (getHeight 0.5) "USB None"
    $lblUSB.forecolor = "#080808"
        # �V�K�쐬�{�^��
    $btnNew = newButton `
        (getX 0) (getY 0) (getWidth 1) (getHeight 1) "�V�K�쐬"
    $btnNew.Add_Click({

    })
    $btnNew.Enabled = $false
    # ���[�J���t�H���_�{�^��
    $btnFolder = newButton `
        (getX 1) (getY 0) (getWidth 1) (getHeight 1) "���[�J��`r`n�t�H���_"
    $btnFolder.Add_Click({

    })
    $btnFolder.Enabled = $false
    # �A�b�v���[�h�{�^��
    $btnUpload = newButton `
        (getX 0) (getY 1) (getWidth 1) (getHeight 1) "�A�b�v���[�h"
    $btnUpload.Add_Click({

    })
    $btnUpload.Enabled = $false
    # �C���X�g�[���{�^��
    $btnInstall = newButton `
        (getX 1) (getY 1) (getWidth 1) (getHeight 1) "�C���X�g�[��"
    $btnInstall.Add_Click({

    })
    $btnInstall.Enabled = $false

    $form.Controls.Add($lblOnline)
    $form.Controls.Add($lblUSB)
    $form.Controls.Add($btnNew)
    $form.Controls.Add($btnFolder)
    $form.Controls.Add($btnUpload)
    $form.Controls.Add($btnInstall)

    # �^�C�}�[
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

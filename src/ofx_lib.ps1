<#
    ������.xlsm�����`���[�E�C���X�g�[���@���C�u����
    copyright 2024 C.Nagata
    2024.6.16   Initial writing.
#>

# �Œ���K�v�ȃt�@�C���̃L�[
$REQUIRED_FILES = @("launch-bat", "install-ps1", "lib-ps1")
#   Excel�֐ݒ�p�X��n�����ϐ�
$ENVVAR_NAME = "OFX_BOOK_CONFIG"
#   �e�t�@�C���ւ̃p�X
$CONF_PATH = $PSScriptRoot + "\ofx_config.json"
$BOOK_CONF_PATH = $PSScriptRoot + "\ofx_book_config.json"
#   �ݒ�̎擾
$CONFIG = (Get-Content $CONF_PATH -Raw | ConvertFrom-Json)
<#    ���̎擾    #>
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
#���̃t�@�C���̊��ionline/mobile/USB�j
$THIS_ENV = getEnv


<#  �I�����C���i�l�b�g���[�N�t�H���_������j�����ׂ�  #>
function IsOnline {
    $path = $CONFIG.env.online.dirs.'base-dir'
    return (Test-Path -LiteralPath $path)
}

<#  �\�[�X�̊����擾����  #>
function getSourceEnv {
    if(IsOnline) {
        # online�Ȃ�A�\�[�X��online
        return "online"
    } elseif($THIS_ENV -eq "USB") {
        # �����i���̃t�@�C���j��USB�ɂ���Ȃ�A�\�[�X�Ȃ�
        return "none"
    }
    # �����̊��imobile�j
    return $THIS_ENV
}

function echoVars {
    Write-Host "�K�{�t�@�C��:${REQUIRED_FILES}"
    Write-Host "���ϐ���:${ENVVAR_NAME}"
    Write-Host "�ݒ�t�@�C��:${CONF_PATH}"
    Write-Host "������Excel�ݒ�t�@�C��:${BOOK_CONF_PATH}"
    Write-Host "�z�X�g�̊�:${THIS_ENV}"
}

<#  �p�X�̂��߂̎��������
#>
function makePathDict($obj, $drive_letter="") {
    if($null -eq $obj) {
        Write-Host "makePathDict �ɓn���ꂽ PSObject �� null �ł��B"
        exit
    }
    # �h���C�u���w�肳��Ă��Ȃ�������A�X�N���v�g�Ɠ����ꏊ
    if( $drive_letter -eq "") {
        $drive_letter = $script:PSScriptRoot.Substring(0, 2)
    }
    #   �����̕ϐ��B���������o�[�̓h���C�u���^�[��
    $dict = @{"Drive" = $drive_letter; "AppDir" = $script:PSScriptRoot}
    #   ����t�H���_
    foreach($key in "Desktop","MyDocuments","StartMenu","Templates") {
        $dict.Add($key, [System.Environment]::GetFolderPath($key))
    }
    #   �I�u�W�F�N�g�̃����o
    foreach($key in $obj.psobject.properties.name) {
        $s = [regex]::replace($obj.$key, "\%([\w\-]+)\%", { $dict[$args.groups[1].value] })
        $dict.Add($key, $s)
    }
    return $dict
}

<#  �p�X��W�J����  #>
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

<#  �����ŁA�f�[�^���̕ϐ���W�J����    #>
function convertVars($data, $dict, $add_base=$false) {
    if($null -eq $data) {
        Write-Host "convertVars �ɓn���ꂽ PSObject �� null �ł��B"
        exit
    }
    foreach($key in $data.psobject.properties.name) {
        $data.$key = expandPath $data.$key $dict $add_base
    }
    return $data
}


<#  �C���X�g�[���̃\�[�X�ɂ�����̗D�揇��
    �Ⴂ�����D�揇�ʂ�����   #>
function envPriority($env) {
    $ary = @("debug", "online", "mobile", "USB", "none")
    return [Array]::IndexOf($ary, $env)
}

<#  �C���X�g�[���^�A�b�v�f�[�g�\���H    #>
function canInstall($drv_inf) {
    $src_env = getSourceEnv
    return ((envPriority $drv_inf["env"]) -gt (envPriority $src_env))
}
    
<#   �h���C�u���̎擾
    @param $drive �h���C�u���^�[
    @return �h���C�u���̃n�b�V���e�[�u��
            letter: (string) �h���C�u���^�[
            env: (string) ���imobile/USB�j
            exists: (boolean) �C���X�g�[������Ă��邩�H
            can-install: (boolean) �C���X�g�[���^�A�b�v�f�[�g�\���H
#>
function getDriveInfo($drive) {
    # ���̎擾
    $env = "USB"
    if($drive -eq "C:") {
        $env = "mobile"
    } elseif($drive -eq "\\") {
        $env = "online"
    }
    # �����̍쐬
    $dict = makePathDict $CONFIG.env.$env.dirs $drive
    # �������i�h���C�u���^�[��mobile/USB�j
    $info = @{ "letter" = $drive; "env" = $env }
    # �K�v�ȃt�@�C���͑S�đ����Ă��邩�H
    $exists = $true
    foreach($fkey in $REQUIRED_FILES) {
        # �L�[�ɑΉ�����p�X�𓾂�
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

<#  �h���C�u�̌��o���A�Ώۂ̃h���C�u�S�Ăɕt���ď����擾����  #>
function getDrives($only_usb=$false) {
    # C�h���C�u�̏��B���������o�[�́APC�̃h���C�u�iC:�Amobile�j
    $drives = @{}
    if($only_usb) {
        $drives.Add("C:", (getDriveInfo "C:"))
    }
    # USB�i�����[�o�u�����f�B�A�j�̃h���C�u���擾
    $usbs = (Get-WmiObject CIM_LogicalDisk | Where-Object DriveType -eq 2).DeviceID
    foreach( $usb in $usbs) {
        # �����̃h���C�u�̏����擾
        $drives.Add($usb, (getDriveInfo $usb))
    }
    return $drives
}


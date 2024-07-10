<#
    �����������`���\�C���X�g�[���[
    copyright 2024 C.Nagata
    2024.6.16   Initial writing.
#>

Set-StrictMode -Version 2.0

# ���C�u�����̃C���|�[�g
. "$($PSScriptRoot)\ofx_lib.ps1"


<#  �f�B���N�g���̍쐬
    CONFIG�́@install > dirs �ɂ���f�B���N�g�����쐬����
#>
function makeDirs($dict, $to_make=$true) {
    #�@�f�B���N�g���z��𓾁A�W�J����
    $dirs = $CONFIG.install.dirs
    for($i=0; $i -lt $dirs.Length; $i++) {
        $dirs[$i] = expandPath $dirs[$i] $dict $true
    }
    # �v�f���ƁA�u�G�v�ŕ������A�f�B���N�g�����쐬����
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

<#  �t�@�C�����R�s�[����    
    CONFIG�́@install > files �ɂ���t�@�C�����A�Ȃ����Â���΃R�s�[����
#>
function copyFiles($src_dict, $dest_dict, $is_update=$false) {
    $copied = 0
    #   �]�����̃t�@�C��
    $src_files = $CONFIG.install.files.psobject.copy()
    $src_files = convertVars $src_files $src_dict $true
    #   �]����̃t�@�C���p�X
    $dest_files = $CONFIG.install.files.psobject.copy()
    $dest_files = convertVars $dest_files $dest_dict $true
    #   ���[�v
    foreach($key in $src_files.psobject.properties.name) {
        # �A�b�v�f�[�g�̂Ƃ��A�L�[�l���K�{�t�@�C���Ȃ�R�s�[���Ȃ�
        if($is_update -and $REQUIRED_FILES.Contains($key) ) {
            continue
        }
        $msg = ""
        #   �]�����E�]����t�@�C���p�X
        $src = $src_files.$key
        $dest = $dest_files.$key
        #   �]����̃t�@�C�����Ȃ���΁A�]������
        $do = -not (Test-Path -LiteralPath $dest)
        if(-not $do) {
            #   �t�@�C�������ɂ���Ƃ��́A�^�C���X�^���v���r���Ė{���V������΃R�s�[����
            $src_date = [datetime](Get-ItemProperty $src).LastWriteTime
            $dest_date = [datetime](Get-ItemProperty $dest).LastWriteTime
            $do = ( $src_date.CompareTo($dest_date) > 0)
        }
        if($do) {
            # �f�B���N�g�����Ȃ���΍쐬����
            $dir = Split-Path -Parent $dest
            if(-not (Test-Path -LiteralPath $dir)) {
                New-Item $dir -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
            }
            if(Test-Path -LiteralPath $src) {
                Copy-Item $src $dest -Force | Out-Null
                Unblock-File $dest
                $copied += 1
                $msg = "Copied..."
            } else {
                Write-Host "Not found... ${src}"
            }
        } else {
            $msg = "Latest..."
        }
        if($msg -ne "") {
            Write-Host "${msg} ${dest}"
        }
    }
    # Drive��C:�Ȃ�A�V���[�g�J�b�g���쐬
    if($dest_dict['Drive'] -eq "C:") {
        $sc_inf = $CONFIG.install.shortcut
        $target = $dest_files.($sc_inf.target)
        $icon = $dest_files.($sc_inf.icon)
        createShortCut $sc_inf.name $target $icon
    }
    return $copied
}

<#  �V���[�g�J�b�g�̃t�@�C���p�X�̎擾#>
function getShortCutPath($name) {
    return [System.Environment]::GetFolderPath("Desktop") + "\" + $name + ".lnk"
}

<#  �V���[�g�J�b�g�̍쐬    #>
function createShortCut($name, $target, $icon, $wd="") {
    $fn = getShortCutPath $name
    $WsShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WsShell.CreateShortcut($fn)
    $Shortcut.TargetPath = $target
    $Shortcut.IconLocation = $icon
    if($wd -eq "") {
        $wd = Split-Path -Parent $target
    }
    $Shortcut.WorkingDirectory = $wd
    $Shortcut.Save()
}

<#  �C���X�g�[���^�A�b�v�f�[�g 
#>
function install($drv_info, $SRC_ENV) {
    #   �]�����̃f�B���N�g�������iDrive�́A���Ŏ����̃p�X����擾�����j
    $src_dict = makePathDict $CONFIG.env.$SRC_ENV.dirs
    # �]����̃p�X�̎����̍쐬
    $letter = $drv_info["letter"]
    $dest_dict = makePathDict $CONFIG.env.($drv_info["env"]).dirs $letter
    # �h���C�u�ɃC���X�g�[������Ă�����A�b�v�f�[�g�B���̑��̓C���X�g�[��
    if( $drv_info["exists"] ) {
        Write-Host "${letter}�h���C�u�̃A�v�����ŐV���m�F���܂��c"
        $copied = copyFiles $src_dict $dest_dict $true
        if($copied -eq 0) {
            Write-Host "${letter}�h���C�u�̃A�v���͍ŐV�ł��B"
        } else {
            Write-Host "${letter}�h���C�u�̃A�v�����X�V���܂����B"
        }
    } else {
        Write-Host "${letter}�h���C�u�ɃA�v�����C���X�g�[�����܂��c"
        Write-Host "�f�B���N�g�����쐬���Ă��܂��c"
        makeDirs $dest_dict
        Write-Host "�t�@�C�����R�s�[���Ă��܂��c"
        $copied = copyFiles $src_dict $dest_dict $false
        Write-Host "${letter}�h���C�u�ւ̃C���X�g�[�����������܂���"
    }
}

<#  �t�@�C���̍폜  #>
function removeFiles($dict) {
    # �폜����t�@�C��
    $files = $CONFIG.install.files.psobject.copy()
    $files = convertVars $files $dict $true
    # �Ώۃh���C�u��C:�Ȃ�A�V���[�g�J�b�g���폜
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

<#   �A���C���X�g�[��   #>
function uninstall($drv_info) {
    $letter = $drv_info["letter"]
    Write-Host "${letter}�h���C�u�̃A�v�����폜���܂��c"
    # �]����̃p�X�̎����̍쐬
    $dict = makePathDict $CONFIG.env.($drv_info["env"]).dirs $letter
    Write-Host "�f�B���N�g�����폜���Ă��܂��c"
    makeDirs $dict $false
    Write-Host "�t�@�C�����폜���Ă��܂��c"
    removeFiles $dict
    Write-Host "${letter}�h���C�u�̃A�v�����폜���܂����B"
}

<#  �C���X�g�[���^�A���C���X�g�[��������h���C�u��I��������
    @param $drive (hashtable) �h���C�u���̃n�b�V���e�[�u��
    @param $is_install (boolean) $true:�C���X�g�[���^$false:�A���C���X�g�[��
#>
function askDrive($drives, $is_install) {
    # �Ώۂ̃h���C�u�̃��X�g���쐬����B
    $lst = @()
    foreach( $letter in $drives.Keys) {
        if(($is_install -and -not $drives[$letter]["exists"] `
                -and $drives[$letter]["can-install"])`
            -or (-not $is_install -and $drives[$letter]["exists"])) {
            $lst = $lst + $letter.Substring(0,1)
        }
    }
    $proc = "�A�v�����C���X�g�[��"
    if(-not $is_install) {
        $proc = "�A�v�����폜"
    }
    if($lst.Length -eq 0) {
        Write-Host "${proc}�ł���h���C�u�͂���܂���B"
        return ""
    }
    $drv_letter = Read-Host ("${proc}����h���C�u��I�����Ă��������B�i"`
                    + ($lst -join ", ") + ", ���̑��F�������j")
    
    $drv_letter = $drv_letter.ToUpper()
    if( $lst.Contains($drv_letter) ) {
        return ($drv_letter + ":")
    }
    return ""
}

<#   �h���C�u���̃n�b�V���e�[�u���̒��g��\������   #>
function HashDsp($hash) {
    $lst = @()
    foreach($key in $hash.Keys) {
        $lst += $hash[$key]
    }
    ($lst | ForEach-Object { New-Object PSCustomObject -Property $_ } `
         | Out-String).trim() | Write-Host
}

<#  �A�b�v�f�[�g�̏���  #>
updateProc($arg) {
    Write-Host "�h���C�u�����o���Ă��܂��c"
    $sec_env = getSourceEnv
    $drives = getDrives $arg["media-only"]
    $cnt = 0
    HashDsp $drives
        foreach($drv in $drives.Keys) {
        # �h���C�u�Ɋ��ɃC���X�g�[������Ă��āA�A�b�v�f�[�g�\�Ȃ�A�A�b�v�f�[�g�i�m�F�j
        if($drives[$drv]["exists"] -and $drives[$drv]["can-install"]) {
            install $drives[$drv] $sec_env
            $cnt++
        }
    }
    if($cnt -eq 0) {
        Write-Host "�A�b�v�f�[�g�\�ȃh���C�u�͂���܂���B"
    }
}

<#  �C���X�g�[������    #>
function installProc($arg) {
    $sec_env = getSourceEnv
    while($true) {
        $drives = getDrives $arg["media-only"]
        HashDsp $drives
        $prc = Read-Host "������I�����Ă��������B�iI:�C���X�g�[���^U:�A�v�����폜�^���̑��F�I���j"
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

# �����̃f�t�H���g�B���w��̏ꍇ�̓A�b�v�f�[�g�̂�
$arg = @{
    "update" = $true
    "istall" = $false
    "media-only" = $false
}
# ������������������A�A�b�v�f�[�g�����Ȃ�
if($Args.length -gt 0) {
    $arg["update"] = $false
}

# �I�v�V�����̉��
for($i=0; $i -lt $Args.length; $i++) {
    $opt = $Arg[$i].Replace("-","")
    switch($opt) {
        "full" {   $arg["update"] = $true; $arg["istall"] = $true;    }
        "update" {   $arg["update"] = $true;     }
        "install" {   $arg["install"] = $true;     }
        "mediaonly" {   $arg["media-only"] = $true;     }
    }
}

# �A�b�v�f�[�g
if($arg["update"]) {
    updateProc $arg
}

# �C���X�g�[��/�A���C���X�g�[��
if($arg["install"]) {
    installProc $arg
}

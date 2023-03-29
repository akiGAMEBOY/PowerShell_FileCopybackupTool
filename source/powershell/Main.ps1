#################################################################################
# 処理名　｜FileCopybackupTool（メイン処理）
# 機能　　｜ファイルをコピーバックアップするツール
#--------------------------------------------------------------------------------
# 戻り値　｜下記の通り。
# 　　　　｜   0: 正常終了
# 　　　　｜-101: エラー メイン - 引数1なし
# 　　　　｜-101: エラー メイン - 既定以外の値
# 　　　　｜-201: エラー メイン - 設定ファイル読み込み
# 　　　　｜-301: エラー 共有フォルダ―への接続 - 接続時、失敗
# 　　　　｜-302: エラー 共有フォルダ―への接続 - 接続後、失敗
# 　　　　｜-303: エラー 共有フォルダ―への切断 - 切断時、失敗
# 　　　　｜-401: エラー バックアップローテーション - フォルダの削除
# 　　　　｜-501: エラー コピーバックアップ - フォルダの削除
# 　　　　｜-502: エラー コピーバックアップ - フォルダのコピー
# 　　　　｜-999: エラー メイン - 処理中断
# 引数　　｜-Args[0]: Rotation or Copy
#################################################################################
# 設定
# DEBUG用
[System.Boolean]$IS_VALID_DEBUG = $false
[System.String]$MODE_DEBUG = 'Rotation'
# [System.String]$MODE_DEBUG = 'Copy'

# 定義されていない変数があった場合にエラーとする
Set-StrictMode -Version Latest
# アセンブリ読み込み（フォーム用）
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# try-catchの際、例外時にcatchの処理を実行する
$ErrorActionPreference = "Stop"
# 定数
[System.String]$c_config_file = "setup.ini"
[System.String]$c_backupdrive = "COPYBACKUP"
[System.String[]]$c_dateformats = @('yyyyMMdd')
[System.String]$c_mode_rotation = 'Rotation'
[System.String]$c_mode_copy = 'Copy'
# Function
#################################################################################
# 処理名　｜ExpandString
# 機能　　｜文字列を展開（先頭桁と最終桁にあるダブルクォーテーションを削除）
#--------------------------------------------------------------------------------
# 戻り値　｜String（展開後の文字列）
# 引数　　｜target_str: 対象文字列
#################################################################################
Function ExpandString($target_str) {
    [System.String]$expand_str = $target_str
    
    If ($target_str.Length -ge 2) {
        if (($target_str.Substring(0, 1) -eq "`"") -and
                ($target_str.Substring($target_str.Length - 1, 1) -eq "`"")) {
            # ダブルクォーテーション削除
            $expand_str = $target_str.Substring(1, $target_str.Length - 2)
        }
    }

    return $expand_str
}

#################################################################################
# 処理名　｜ConfirmYesno_winform
# 機能　　｜YesNo入力（Windowsフォーム）
#--------------------------------------------------------------------------------
# 戻り値　｜Boolean（True: 正常終了, False: 処理中断）
# 引数　　｜prompt_message: 入力応答待ち時のメッセージ内容
#################################################################################
Function ConfirmYesno_winform([System.String]$prompt_message) {
    [System.Boolean]$return = $false

    # フォームの作成
    [System.Windows.Forms.Form]$form = New-Object System.Windows.Forms.Form
    $form.Text = "実行前の確認"
    $form.Size = New-Object System.Drawing.Size(460,210)
    $form.StartPosition = "CenterScreen"
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon("${root_dir}\source\icon\shell32-296.ico")
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    # ピクチャボックス作成
    [System.Windows.Forms.PictureBox]$pic = New-Object System.Windows.Forms.PictureBox
    $pic.Size = New-Object System.Drawing.Size(32, 32)
    $pic.Image = [System.Drawing.Image]::FromFile("${root_dir}\source\icon\shell32-296.ico")
    $pic.Location = New-Object System.Drawing.Point(30,30)
    $pic.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
    # ラベル作成
    [System.Windows.Forms.Label]$label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(85,30)
    $label.Size = New-Object System.Drawing.Size(350,80)
    $label.Text = $prompt_message
    $font = New-Object System.Drawing.Font("ＭＳ ゴシック",12)
    $label.Font = $font
    # OKボタンの作成
    [System.Windows.Forms.Button]$btnOkay = New-Object System.Windows.Forms.Button
    $btnOkay.Location = New-Object System.Drawing.Point(255,120)
    $btnOkay.Size = New-Object System.Drawing.Size(75,30)
    $btnOkay.Text = "OK"
    $btnOkay.DialogResult = [System.Windows.Forms.DialogResult]::OK
    # Cancelボタンの作成
    [System.Windows.Forms.Button]$btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(345,120)
    $btnCancel.Size = New-Object System.Drawing.Size(75,30)
    $btnCancel.Text = "キャンセル"
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    # ボタンの紐づけ
    $form.AcceptButton = $btnOkay
    $form.CancelButton = $btnCancel
    # フォームに紐づけ
    $form.Controls.Add($pic)
    $form.Controls.Add($label)
    $form.Controls.Add($btnOkay)
    $form.Controls.Add($btnCancel)
    # フォーム表示
    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $return = $true
    } else {
        $return = $false
    }
    $pic.Image.Dispose()
    $pic.Image = $null
    $form = $null

    return $return
}

#################################################################################
# 処理名　｜RotationBackupfile
# 機能　　｜バックアップローテーション
#--------------------------------------------------------------------------------
# 戻り値　｜Int
# 　　　　｜       0: 正常終了
# 　　　　｜    -301: エラー 共有フォルダ―への接続 - 接続時、失敗
# 　　　　｜    -302: エラー 共有フォルダ―への接続 - 接続後、失敗
# 　　　　｜    -303: エラー 共有フォルダ―への切断 - 切断時、失敗
# 　　　　｜    -401: エラー バックアップローテーション - フォルダの削除
# 引数　　｜バックアップ先
# 　　　　｜BackuptoHost: ホスト名、またはIP, BackuptoId: ユーザ名,
# 　　　　｜BackuptoPass: パスワード, BackuptoPath: パス, BackuptoGene: 保持世代数
#################################################################################
Function RotationBackupfile([System.String]$BackuptoHost, [System.String]$BackuptoId,
                            [System.String]$BackuptoPass, [System.String]$BackuptoPath, [System.Int32]$BackuptoGene) {
    [System.Int32]$result = 0    

    # 共有フォルダ接続
    ## 接続先の設定
    [SecureString]$securepass = ConvertTo-SecureString $BackuptoPass -AsPlainText -Force
    [PSCredential]$cred = New-Object System.Management.Automation.PSCredential "${BackuptoHost}\${BackuptoId}", $securepass
    try {
        New-PSDrive -Name $c_backupdrive -PSProvider FileSystem -Root $BackuptoPath -Credential $cred 2>&1>$null
    } catch {
        $result = -301
    }
    ## 接続確認
    if ($result -eq 0) {
        [System.Management.Automation.PSDriveInfo]$psdrive = $null
        $psdrive = Get-PSDrive $c_backupdrive 2> $null
        if ($null -eq $psdrive) {
            $result = -302
        }
    }

    # バックアップローテーション
    if ($result -eq 0) {
        [System.Int32]$generation = 0
        [Object[]]$itemlist = Get-ChildItem "${c_backupdrive}`:\" | Sort-Object -Descending {$_.Name}
        [System.DateTime]$parseddate = [System.DateTime]::MinValue
        foreach($item in $itemlist) {
            [System.Boolean]$parseresult = [System.DateTime]::TryParseExact(
                $item.Name,
                $c_dateformats,
                [Globalization.DateTimeFormatInfo]::CurrentInfo,
                [Globalization.DateTimeStyles]::AllowWhiteSpaces,
                [ref]$parseddate
            )

            # コピー対象外はスキップする
            ## 日付フォルダ以外の場合
            ## または、フォルダ名が翌日以降の場合、
            ## または、フォルダ以外の場合
            if ((-not($parseresult)) -or `
                    ($item.Name -gt $today) -or `
                    (-not($item.PSIsContainer))) {
                continue
            }

            # バックアップ対象をカウント
            $generation = $generation + 1
            # 指定フォルダー配下にあるサブフォルダの数をカウントする場合
            # $generation += [System.Int32]((Get-ChildItem "${c_backupdrive}`:\$($item.Name)" | Where-Object { $_.PSIsContainer } | Measure-Object).Count)

            # コピーされたフォルダ数が既定の世代数を超える場合
            if ($generation -gt $BackuptoGene) {
                # 削除するフォルダを設定
                $deldate = [DateTime]::ParseExact($item.Name,"yyyyMMdd", $null)
            }

            # フォルダ削除
            if ($deldate -ne [System.DateTime]::MaxValue) {
                try {
                    # Forceオプションをつけるとアクセス拒否エラー
                    # Remove-Item "${c_backupdrive}`:\$($item.Name)" -Recurse -Force
                    Remove-Item "${c_backupdrive}`:\$($item.Name)" -Recurse

                    $sbprompt=New-Object System.Text.StringBuilder
                    @("通知　　　: バックアップローテーション処理`r`n",`
                    "　　　　　　フォルダを削除しました。`r`n",`
                    "　　　　　　対象 [$($item.FullName)]`r`n")|
                    ForEach-Object{[void]$sbprompt.Append($_)}
                    $prompt_message = $sbprompt.ToString()
                    Write-Host $prompt_message
                } catch {
                    $result = -401
                    break
                }
            }
        }
    }

    # 共有フォルダ切断
    If (-not($result -in @(-301,-302))) {
        [System.Management.Automation.PSDriveInfo]$psdrive = $null
        $psdrive = Get-PSDrive $c_backupdrive 2> $null
        if ($null -ne $psdrive) {
            try {
                Remove-PSDrive -Name $c_backupdrive
            } catch {
                $result = -303
            }
        }
    }

    return $result
}

#################################################################################
# 処理名　｜CopyBackupfile
# 機能　　｜コピーバックアップ
#--------------------------------------------------------------------------------
# 戻り値　｜Int
# 　　　　｜       0: 正常終了
# 　　　　｜    -301: エラー 共有フォルダ―への接続 - 接続時、失敗
# 　　　　｜    -302: エラー 共有フォルダ―への接続 - 接続後、失敗
# 　　　　｜    -303: エラー 共有フォルダ―への切断 - 切断時、失敗
# 　　　　｜    -501: エラー コピーバックアップ - フォルダの削除
# 　　　　｜    -502: エラー コピーバックアップ - フォルダのコピー
# 引数　　｜バックアップ先
# 　　　　｜BackuptoHost: ホスト名、またはIP, BackuptoId: ユーザ名, $BackuptoPass: パスワード, BackuptoPath: パス, BackuptoGene: 保持世代数
# 　　　　｜バックアップ元
# 　　　　｜BackupfmPath: パス
#################################################################################
Function CopyBackupfile([System.String]$BackuptoHost, [System.String]$BackuptoId,
                        [System.String]$BackuptoPass, [System.String]$BackuptoPath, [System.String]$BackupfmPath) {
    [System.Int32]$result = 0    

    # 共有フォルダ接続
    ## 接続先の設定
    [SecureString]$securepass = ConvertTo-SecureString $BackuptoPass -AsPlainText -Force
    [PSCredential]$cred = New-Object System.Management.Automation.PSCredential "${BackuptoHost}\${BackuptoId}", $securepass
    try {
        New-PSDrive -Name $c_backupdrive -PSProvider FileSystem -Root $BackuptoPath -Credential $cred 2>&1>$null
    } catch {
        $result = -301
    }
    ## 接続確認
    if ($result -eq 0) {
        [System.Management.Automation.PSDriveInfo]$psdrive = $null
        $psdrive = Get-PSDrive $c_backupdrive 2> $null
        if ($null -eq $psdrive) {
            $result = -302
        }
    }

    # コピーバックアップ
    if ($result -eq 0) {
        # フォルダがある場合は削除
        if (Test-Path "${c_backupdrive}`:\${today}") {
            try {
                Remove-Item "${c_backupdrive}`:\${today}" -Recurse -Force
                $sbprompt=New-Object System.Text.StringBuilder
                @("通知　　　: コピーバックアップ処理`r`n",`
                "　　　　　　フォルダを削除しました。`r`n",`
                "　　　　　　対象 [${BackuptoPath}\${today}]`r`n")|
                ForEach-Object{[void]$sbprompt.Append($_)}
                $prompt_message = $sbprompt.ToString()
                Write-Host $prompt_message
            } catch {
                $result = -501
            }
        }

        # コピー
        if ($result -eq 0) {
            try {
                Copy-Item "${BackupfmPath}" -Recurse "${c_backupdrive}`:\${today}"
                $sbprompt=New-Object System.Text.StringBuilder
                @("通知　　　: コピーバックアップ処理`r`n",`
                  "　　　　　　コピーバックアップが完了しました。`r`n",`
                  "　　　　　　対象 [${BackuptoPath}\${today}]`r`n")|
                ForEach-Object{[void]$sbprompt.Append($_)}
                $prompt_message = $sbprompt.ToString()
                Write-Host $prompt_message
            } catch {
                $result = -502
            }
        }
    }

    # 共有フォルダ切断
    If (-not($result -in @(-301,-302))) {
        [System.Management.Automation.PSDriveInfo]$psdrive = $null
        $psdrive = Get-PSDrive $c_backupdrive 2> $null
        if ($null -ne $psdrive) {
            try {
                Remove-PSDrive -Name $c_backupdrive
            } catch {
                $result = -303
            }
        }
    }

    return $result
}

#################################################################################
# 処理名　｜メイン処理
# 機能　　｜同上
#--------------------------------------------------------------------------------
# 　　　　｜-
#################################################################################
[System.Int32]$result = 0
[System.String]$prompt_message = ''
[System.String]$result_message = ''
[System.Text.StringBuilder]$sbtemp=New-Object System.Text.StringBuilder

# 引数チェック
## 引数の有無
If ($IS_VALID_DEBUG) {
    [System.String]$mode = $MODE_DEBUG
} else {
    if ($Args.Length -eq 0) {
        $result = -101
        $sbresult=New-Object System.Text.StringBuilder
        @("エラー　　: 引数1がなし`r`n",`
          "　　　　　　正しい引数で起動しているかご確認ください。`r`n")|
        ForEach-Object{[void]$sbresult.Append($_)}
        $result_message = $sbresult.ToString()
    } elseif ([System.String]::IsNullOrEmpty($Args[0])) {
        $result = -101
        $sbresult=New-Object System.Text.StringBuilder
        @("エラー　　: 引数1がなし`r`n",`
          "　　　　　　正しい引数で起動しているかご確認ください。`r`n")|
        ForEach-Object{[void]$sbresult.Append($_)}
        $result_message = $sbresult.ToString()
    } else {
        [System.String]$mode = $Args[0]
    }
}

## 引数の値
if ($result -eq 0) {
    if (-not ($mode -in @($c_mode_rotation, $c_mode_copy))) {
        $result = -102
        $sbresult=New-Object System.Text.StringBuilder
        @("エラー　　: 引数1が既定以外の値`r`n",`
          "　　　　　　正しい引数で起動しているかご確認ください。`r`n")|
        ForEach-Object{[void]$sbresult.Append($_)}
        $result_message = $sbresult.ToString()
    }
}

# 初期設定
if ($result -eq 0) {
    ## 日付取得
    [System.String]$today=Get-Date -Format yyyyMMdd
    [System.DateTime]$deldate = [System.DateTime]::MaxValue
    ## ディレクトリの取得
    [System.String]$current_dir=Split-Path ( & { $myInvocation.ScriptName } ) -parent
    Set-Location $current_dir"\..\.."
    [System.String]$root_dir = (Convert-Path .)
    ## 設定ファイル読み込み
    $sbtemp=New-Object System.Text.StringBuilder
    @("$current_dir",`
    "\",`
    "$c_config_file")|
    ForEach-Object{[void]$sbtemp.Append($_)}
    [System.String]$config_fullpath = $sbtemp.ToString()
    try {
        [System.Collections.Hashtable]$param = Get-Content $config_fullpath -Raw | ConvertFrom-StringData
        # バックアップ先 - ホスト名、またはIP
        [System.String]$BackuptoHost=ExpandString($param.BackuptoHost)
        # バックアップ先 - ユーザ名
        [System.String]$BackuptoId=ExpandString($param.BackuptoId)
        # バックアップ先 - パスワード
        [System.String]$BackuptoPass=ExpandString($param.BackuptoPass)
        # バックアップ先 - パス
        [System.String]$BackuptoPath=ExpandString($param.BackuptoPath)
        # バックアップ先 - 世代数
        [System.Int32]$BackuptoGene=ExpandString($param.BackuptoGene)
        # バックアップ対象
        [System.String]$BackupfmPath=ExpandString($param.BackupfmPath)

        $sbtemp=New-Object System.Text.StringBuilder
        @("通知　　　: 設定ファイル読み込み`r`n",`
        "　　　　　　設定ファイルの読み込みが正常終了しました。`r`n",`
        "　　　　　　対象: [${config_fullpath}]`r`n")|
        ForEach-Object{[void]$sbtemp.Append($_)}
        $prompt_message = $sbtemp.ToString()
        Write-Host $prompt_message
    }
    catch {
        $result = -201
        $sbtemp=New-Object System.Text.StringBuilder
        @("エラー　　: 設定ファイル読み込み`r`n",`
        "　　　　　　設定ファイルの読み込みが異常終了しました。`r`n",`
        "　　　　　　エラー内容: [${config_fullpath}",`
        "$($_.Exception.Message)]`r`n")|
        ForEach-Object{[void]$sbtemp.Append($_)}
        $result_message = $sbtemp.ToString()
    }
}

# ローテーション or コピーバックアップ
if ($result -eq 0) {
    # ローテーション
    if ($mode -eq $c_mode_rotation) {
        $sbtemp=New-Object System.Text.StringBuilder
        # 実行有無の確認
        @("バックアップローテーションを実行します。`r`n",`
          "処理を続行しますか？`r`n")|
        ForEach-Object{[void]$sbtemp.Append($_)}
        $prompt_message = $sbtemp.ToString()
        If (ConfirmYesno_winform $prompt_message) {
            $result = RotationBackupfile $BackuptoHost $BackuptoId $BackuptoPass $BackuptoPath $BackuptoGene
            if ($result -ne 0) {
                $sbtemp=New-Object System.Text.StringBuilder
                @("エラー　　: バックアップローテーション`r`n",`
                  "　　　　　　バックアップローテーションが異常終了しました。`r`n")|
                ForEach-Object{[void]$sbtemp.Append($_)}
                $result_message = $sbtemp.ToString()
            }
        } else {
            $result = -999
            $sbtemp=New-Object System.Text.StringBuilder
            @("エラー　　: コピーバックアップの実行`r`n",`
              "　　　　　　キャンセルしました。`r`n")|
            ForEach-Object{[void]$sbtemp.Append($_)}
            $result_message = $sbtemp.ToString()
        }
    # コピーバックアップ
    } else {
        $sbtemp=New-Object System.Text.StringBuilder
        # 実行有無の確認
        @("コピーバックアップを実行します。`r`n",`
          "処理を続行しますか？`r`n")|
        ForEach-Object{[void]$sbtemp.Append($_)}
        $prompt_message = $sbtemp.ToString()
        If (ConfirmYesno_winform $prompt_message) {
            $result = CopyBackupfile $BackuptoHost $BackuptoId $BackuptoPass $BackuptoPath $BackupfmPath
            if ($result -ne 0) {
                $sbtemp=New-Object System.Text.StringBuilder
                @("エラー　　: コピーバックアップの実行`r`n",`
                  "　　　　　　コピーバックアップの実行が異常終了しました。`r`n")|
                ForEach-Object{[void]$sbtemp.Append($_)}
                $result_message = $sbtemp.ToString()
            }
        } else {
            $result = -999
            $sbtemp=New-Object System.Text.StringBuilder
            @("エラー　　: コピーバックアップの実行`r`n",`
              "　　　　　　キャンセルしました。`r`n")|
            ForEach-Object{[void]$sbtemp.Append($_)}
            $result_message = $sbtemp.ToString()
        }
    }
}

# 処理結果の表示
$sbtemp=New-Object System.Text.StringBuilder
if ($result -eq 0) {
    @("処理結果　: 正常終了`r`n",`
      "　　　　　　メッセージコード: [${result}]`r`n")|
    ForEach-Object{[void]$sbtemp.Append($_)}
    $result_message = $sbtemp.ToString()
    Write-Host $result_message
}
else {
    @("処理結果　: 異常終了`r`n",`
      "　　　　　　メッセージコード: [${result}]`r`n",`
      $result_message)|
    ForEach-Object{[void]$sbtemp.Append($_)}
    $result_message = $sbtemp.ToString()
    Write-Host $result_message -ForegroundColor DarkRed
}

# 終了
exit $result

param(
    [Parameter(
        Mandatory = $false,
        HelpMessage = "開始日時を指定してください。[""yyyy-MM-dd hh:mm""]`r`n指定しない場合はスクリプト実行時間＋20秒を開始日時に指定します"
    )]
    [datetime] $StartDatetime,

    [Parameter(
        Mandatory = $false,
        HelpMessage = "終了日時を指定してください。[""yyyy-MM-dd hh:mm""]`r`n指定しない場合はスクリプト実行当日の23:59を終了日時に指定します"
    )]
    [datetime] $EndDatetime,

    [Parameter(
        Mandatory = $false,
        HelpMessage = "ユーザー名の配列を指定してください。"
    )][ValidateNotNull()]
    [string[]] $Users,

    [Parameter(
        Mandatory = $false,
        HelpMessage = "ユーザー名が1行に1つずつ記載されたテキストファイルのパスを指定してください。"
    )][ValidateNotNull()]
    [string] $UserListFile
)

function Validate-Input {
    param(
        $StartDatetime,
        $EndDatetime,
        $Users
    )

    
    # ユーザーが1人も見つからない場合のエラーチェック
    if ($Users.Count -eq 0) {
        Write-Host "ユーザーが1人も見つかりませんでした。"
        return $false
    }
    
    # 開始日時が指定されていない場合、現在時刻を設定
    if ([System.String]::IsNullOrEmpty($StartDatetime) -or $StartDatetime -lt (Get-Date)) {
        $StartDatetime = (Get-Date)
        Write-Host "StartDatetimeが過去日または未入力です。現在時刻($($StartDatetime))を利用します。"
    }

    # 終了日時が指定されていない場合、当日の23:59に設定
    if ([System.String]::IsNullOrEmpty($EndDatetime)) {
        $EndDatetime = (Get-Date).Date.AddDays(1).AddSeconds(-1)
        Write-Host "EndDatetimeが未入力です。当日の23:59を利用します。"
    }
    
    $minEndDate = $StartDatetime.AddMinutes(10).ToString("yyyy-MM-dd hh:mm")
    $maxEndDate = $StartDatetime.AddDays(30).ToString("yyyy-MM-dd hh:mm")
    $lifeTimeMinutes = [math]::Round(($EndDatetime - $StartDatetime).TotalMinutes)
    if ($lifeTimeMinutes -lt 60 -or $lifeTimeMinutes -ge 43200) {
        Write-Warning "開始日時と終了日時の差は60分以上かつ30日未満である必要があります。"
        Write-Host "StartDatetimeが$($StartDatetime)の場合、$($minEndDate)～$($maxEndDate)の間の日時をEndDatetimeに指定してください"
    }
    
    $validInput = @{
        'StartDatetime' = $StartDatetime
        'EndDatetime'   = $EndDatetime
        'Users'         = $Users
    }
    
    return $validInput
}


function Add-TemporaryAccessPass {
    param(
        $StartDatetime,
        $EndDatetime,
        $Users
    )
    
    $TAPInfo = @()
    foreach ($User in $Users) {
        # 処理実行中に指定された開始時間を超過した場合はコマンド実行時間の20秒後を指定する
        if ($StartDatetime -lt (Get-Date)) {
            $StartDatetime = (Get-Date).AddSeconds(20)
        }
        # TAPの期間が60分未満となった場合、終了時間に開始時間の60分後を指定する
        if ([math]::Round(($EndDatetime - $StartDatetime).TotalMinutes) -lt 60) {
            $EndDatetime = $StartDatetime.AddMinutes(60)
        }

        # パラメーターをまとめて指定（スプラット演算子）
        $TAPParams = @{
            UserId            = $User
            LifetimeInMinutes = [math]::Round(($EndDatetime - $StartDatetime).TotalMinutes)
            IsUsable          = $true
            IsUsableOnce      = $false
            StartDatetime     = $StartDatetime.ToString("yyyy-MM-dd HH:mm")
        }

        try {
            $TAP = New-MgUserAuthenticationTemporaryAccessPassMethod @TAPParams
            #アウトプットの成型
            if (![System.String]::IsNullOrEmpty($TAP)) {
                $TAPInfo += [PSCustomObject]@{
                    UPN                 = $User
                    StartDatetime       = $TAP.StartDatetime.AddHours(9)
                    EndDatetime         = $TAP.StartDatetime.AddHours(9).AddMinutes($TAP.LifetimeInMinutes)
                    TemporaryAccessPass = $TAP.TemporaryAccessPass
                }
            }
        }
        catch {
            Write-Warning "ユーザー $_ の処理中にエラーが発生しました: $($_.Exception.Message)"
            $TAP = $null
        }
    }
    return $TAPInfo
}

# メイン処理
$actionDate = Get-Date -Format "yyyy-MM-dd-HH-mm-ss"

# ファイルからユーザー一覧を取得
if (![System.String]::IsNullOrEmpty($UserListFile)) {
    try {
        $Users = Get-Content $UserListFile
    }
    catch {
        Write-Warning "UserListFileに指定されたファイルに情報が入っていません。指定するファイルを見直してください:`r`n $($_.Exception.Message)"
    }
}

$params = @{
    StartDatetime = $StartDatetime
    EndDatetime   = $EndDatetime
    Users         = $Users
}

# モジュール導入
if (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication) {
}
else {
    Install-Module Microsoft.Graph.Authentication -Force
}
if (Get-Module -ListAvailable -Name Microsoft.Graph.Identity.SignIns) {
    Import-Module Microsoft.Graph.Identity.SignIns
}
else {
    Install-Module Microsoft.Graph.Identity.SignIns -Force
    Import-Module Microsoft.Graph.Identity.SignIns
}

# Graphに接続
Connect-MgGraph -Scopes UserAuthenticationMethod.ReadWrite.All | Out-Null

# 入力値チェック
$isValidInput = Validate-Input @params
# 処理実行
if ($isValidInput) {
    $TAPInfo = Add-TemporaryAccessPass @isValidInput
}
else {
    Write-Host "入力が無効です。"
}

#結果出力
if ($TAPInfo.count) {
    $TAPInfo | Export-Csv -Path ".\TAPList_$($actionDate).csv" -Encoding UTF8 -NoTypeInformation
    "$($TAPInfo.count)件のCSVが生成されました。.\TAPList_$($actionDate).csvを確認してください"
    $TAPInfo | Out-GridView
}
else {
    Write-Warning "結果が見つかりません。処理エラーが発生している可能性があります"
}
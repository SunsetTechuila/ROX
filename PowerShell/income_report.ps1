[Console]::Title = "Создание отчета по поступлениям"
$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#region Functions
function Get-Credentials {
    if (-not(Test-Path -Path $env:AppData\income_report -PathType Leaf)) {
        $login = Read-Host -Prompt "Введите логин для сайта placeholder.salesdoc.io" -AsSecureString
        $login | ConvertFrom-SecureString | Set-Content -Path $env:AppData\income_report
        
        $password = Read-Host -Prompt "Введите пароль для сайта placeholder.salesdoc.io" -AsSecureString
        $password | ConvertFrom-SecureString | Add-Content -Path $env:AppData\income_report
        
		(Get-Item -Path $env:AppData\income_report).Attributes = "Hidden"
    }
    else {
        $login = (Get-Content -Path $env:AppData\income_report | Select-Object -First 1) | ConvertTo-SecureString
        $password = (Get-Content -Path $env:AppData\income_report | Select-Object -Last 1) | ConvertTo-SecureString
    }
    
    return @{
        login    = $login
        password = $password
    }
}

function ValidateCredentials {
    param (
        [Parameter(Mandatory)]
        [hashtable]$Credentials
    )
    $loggedSession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    $body = @{
        "LoginForm[username]" = [System.Net.NetworkCredential]::new("", $Credentials.login).Password
        "LoginForm[password]" = [System.Net.NetworkCredential]::new("", $Credentials.password).Password
    }
    $Parameters = @{
        Uri             = "https://placeholder.salesdoc.io/site/login"
        UseBasicParsing = $true
        WebSession      = $loggedSession
        Body            = $body
        Method          = "POST"
    }
    
    if ((Invoke-WebRequest @Parameters).Content -match "Sales Doctor - Sales Doctor - Login") {
        return $false
    }
    else {
        return $loggedSession
    }
}

function CheckDate {
    param (
        [Parameter(Mandatory)]
        [DateTime]$FirstDate, 
        
        [Parameter(Mandatory)]
        [DateTime]$SecondDate
    )
    switch ($FirstDate) {
        { $PSItem -lt $SecondDate } { return "lt" }
        { $PSItem -gt $SecondDate } { return "gt" }
        { $PSItem -eq $SecondDate } { return "eq" }
    }
}

function CheckArguments {
    param (
        [array]$Arguments
    )
    if ($Arguments.Count -eq 2) {
        try {
            $dateOne = Get-Date -Date $Arguments[0]
            $dateTwo = Get-Date -Date $Arguments[1]
            
            if ((CheckDate -FirstDate $dateOne -SecondDate $dateTwo) -eq "lt") {
                return $true
            }
            else {
                return $false
            }
        }
        catch { return $false }
    }
    else { return $false }
}
#endregion Functions

#region Main

#region Input credentials
do {
    $userCredentials = Get-Credentials
    $session = ValidateCredentials -Credentials $userCredentials
    if (-not($session)) {
        Remove-Item -Path $env:AppData\income_report -Force
        Write-Host -Object "Учетные данные не верны!" -ForegroundColor Red
        Read-Host -Prompt "Нажмите Enter и повторите ввод"
        Clear-Host
    }
}
while (-not($session))
#endregion Input credentials

#region Get agents with id
$agents = @()

$Parameters = @{
    Uri             = "https://placeholder.salesdoc.io/agents/agent"
    UseBasicParsing = $true
    WebSession      = $session 
}
$res = (Invoke-WebRequest @Parameters).RawContent

$res.Split("`n") | ForEach-Object {
    if ($PSItem -match '<td><div style="border-radius: 50%;width: 10px;height: 10px;background-color: green;float: left;margin-right: 5px;margin-top: 5px"><\/div>(.+)<\/td>') {
        $name = $matches[1]
    }
    elseif ($PSItem -match '<a class="btn btn-danger delete_agent" agent="(.+)"><i') {
        $id = $matches[1]
        
        if ($name -ne "Бурхон") {
            $agents += @{
                name = $name
                id   = $id 
            }
        }
    }	
}
#endregion Get agents with id

#region Arguments
if (-not(CheckArguments -Arguments $args)) {
    do {
        Clear-Host
        
        $dateStart = Get-Date -Date (Read-Host -Prompt "Введите начальную дату в формате день месяц год(опционально)")
        $dateEnd = Get-Date -Date (Read-Host -Prompt "Введите конечную дату в формате день месяц год(опционально)")
        
        $CheckDateResult = CheckDate -FirstDate $dateStart -SecondDate $dateEnd
        
        if ($CheckDateResult -eq "gt") {
            Write-Host -Object "Начальная дата больше конечной! Возможно, вы забыли указать годы" -ForegroundColor Red
            Read-Host -Prompt "Нажмите Enter и повторите ввод"
        }
        elseif ($CheckDateResult -eq "eq") {
            Write-Host -Object "Начальная дата равна конечной! Возможно, вы забыли указать годы" -ForegroundColor Red
            Read-Host -Prompt "Нажмите Enter и повторите ввод"
        }
    }
    until ($CheckDateResult -eq "lt")
}
else {
    $dateStart = Get-Date -Date $args[0]
    $dateEnd = Get-Date -Date $args[1]
}
#endregion Arguments

#region Excel
$appExcel = New-Object -ComObject excel.application
$appExcel.Visible = $true
$appExcel.Interactive = $false
$workbook = $appExcel.Workbooks.Add(1)
$worksheet = $workbook.Worksheets.Item(1)

#region Add agents
$row = $collumn = 2
$agents | ForEach-Object {
    $worksheet.Cells.Item($row, 1) = $PSItem.Name
    $worksheet.Cells.Item($row, 1).Font.Size = 12
    $worksheet.Cells.Item($row, 1).Font.Bold = $true
    $worksheet.Cells.Item($row, 1).HorizontalAlignment = -4131
    $worksheet.Cells.Item($row, 1).VerticalAlignment = -4108
    $row++
}
#endregion Add agents

#region Add income by date
$currentDate = $dateStart
while ($currentDate -le $dateEnd) {
    $worksheet.Cells.Item(1, $collumn) = $currentDate
    $worksheet.Cells.Item(1, $collumn).NumberFormat = "[$-419]d mmm;@"
    $worksheet.Cells.Item(1, $collumn).NumberFormatLocal = "[$-419]Д МММ;@"
    $worksheet.Cells.Item(1, $collumn).Font.Size = 12
    $worksheet.Cells.Item(1, $collumn).Font.Italic = $true
    $worksheet.Cells.Item(1, $collumn).HorizontalAlignment = -4108
    $worksheet.Cells.Item(1, $collumn).VerticalAlignment = -4108
    
    $row = 2
    $agents | ForEach-Object {
        $date = Get-Date -Date $currentDate -Format "yyyy-MM-dd"
        $Parameters.Uri = "https://placeholder.salesdoc.io/dashboard/kassaIncome?agent%5B%5D=$($PSItem.id)`&bydate=DATE`&datestart=$date`&endstart=$date"
        $res = (Invoke-WebRequest @Parameters).RawContent
        $res -match '<div class="col-md-3 pull-right"><div role="button" class="btn btn-danger" style="width:100%;"><h3><b>(.+)<\/b><\/h3> <span>Общий \(сум\)<\/span>' | Out-Null
        $income = $matches[1] -Replace (",", "")
        $worksheet.Cells.Item($row, $collumn) = $income
        $worksheet.Cells.Item($row, $collumn).NumberFormat = "#,##0"
        $worksheet.Cells.Item($row, $collumn).NumberFormatLocal = "# ##0"
        $worksheet.Cells.Item($row, $collumn).HorizontalAlignment = -4108
        $worksheet.Cells.Item($row, $collumn).VerticalAlignment = -4108
        $row++
    }
    
    if ((Get-Date -Date $currentDate -Format dddd) -ne "суббота") {
        $currentDate = $currentDate.AddDays(1)
    }
    else {
        $currentDate = $currentDate.AddDays(2)
    }
    
    $collumn++
}
#endregion Add income by date

$worksheet.Cells.EntireColumn.AutoFit() | Out-Null
$worksheet.Cells.EntireRow.AutoFit() | Out-Null
$appExcel.Interactive = $true
#endregion Excel

#endregion Main

<#

Microsoft Exchange Attachments Killer
Removes attachments from Microsoft Exchange email messages

author: RiG87

#>

# ----- begin of script ----- #

param (
    # filter database names
    [Parameter(Mandatory=$false,Position=6)][string[]]$Databases,
    
    # attachments time to live in days
    [Parameter(Mandatory=$false,Position=5)][int]$TimeToLive = 360,
    
    # log directory
    [Parameter(Mandatory=$false,Position=4)][string]$LogDir = ".\logs",
    
    # portion limit - maximum size of processed accounts for this runtime-exemplar
    [Parameter(Mandatory=$false,Position=3)][int]$Limit = -1,
    
    # portion start (private) - processing start position
    [Parameter(Mandatory=$false,Position=2)][int]$Start = -1,
    
    # child counter
    [Parameter(Mandatory=$false,Position=1)][int]$ChildCounter = 0,

    # flag means "using all databases for processing"
    [Parameter(Mandatory=$false)][switch]$AllAccounts,

    # flag means "using all accounts for processing"
    [Parameter(Mandatory=$false)][switch]$AllDatabases,

    # flag means "debug mode"
    [Parameter(Mandatory=$false)][switch]$DebugMode,
    
    # one account for processing
    [Parameter(Mandatory=$false)][string]$Account,
    
    # file path
    [Parameter(Mandatory=$false)][string]$AccountsFile,

    # path to exchange web services dll 
    [Parameter(Mandatory=$false)]
    [string]$WebServicesDllPath="C:\Program Files\Microsoft\Exchange Server\V15\Bin\Microsoft.Exchange.WebServices.dll"
)

#region script init

$ErrorActionPreference = "Stop";
$Programm = "Microsoft Exchange Attachments Killer 1.1";
$selfFile = "$PSScriptRoot\$($MyInvocation.MyCommand.Name)";

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.Support

[void][Reflection.Assembly]::LoadFile($WebServicesDllPath);

#endregion

#region utils

function now {
    param([Parameter(Mandatory=$false)][string]$format="yyyy.MM.dd HH:mm:ss.fffff")
    return Get-Date -Format $format;
}

# log file init
function getLog {
    param(
        [Parameter(Mandatory=$false)][string]$LogDir="",
        [Parameter(Mandatory=$false)][int]$ChildCounter=""
    )

    # init log dir
    if (! (Test-Path $LogDir -PathType Container)){
        New-Item -ItemType Directory -Force -Path $LogDir
    }
    $LogDir = (Get-Item -Path $LogDir -Verbose).FullName;

    # init log file name
    $LogFileName = ""
    if ($ChildCounter -eq 0){
        $LogFileName = "meak_log.txt";
    } else {
        $LogFileName = "meak_" + $ChildCounter + "_log.txt";
    }

    # init full path
    $logFile = "$LogDir\$LogFileName";
    if (! (Test-Path $logFile -PathType Leaf)){
        $logFile = (New-Item -Path $LogDir -Name $LogFileName -ItemType "file").FullName
    }

    return $logFile;
}

$logFile = $(getLog -LogDir $LogDir -ChildCounter $ChildCounter);

# main logger
function log([string]$msg) {
    Add-Content -Path $logFile "$(now) $msg"
}

# nothing to do
function ntd {
    $msg = "$(Now) Nothing to do. Exit";
    Write-Output $msg
    log $msg
    Exit
}

# end of work
function eow {
    log " `n"
    log "End Work"
    Exit
}

# errors language
function SetCulture {
    param([Parameter(Mandatory=$false)][string]$name="en-US")
    $culture = New-Object system.globalization.cultureinfo($name);
    [System.Threading.Thread]::CurrentThread.CurrentUICulture = $culture;
    [System.Threading.Thread]::CurrentThread.CurrentCulture = $culture;
}

# Extracts Emails from file or from string
function ExtractEmails {
    param(
        [Parameter(Mandatory=$false)]
        [string]$Content="",
        
        [Parameter(Mandatory=$false)]
        [string]$FilePath="",
        
        [Parameter(Mandatory=$false)]
        [string]$regex = �\b[A-Za-z0-9._%-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,4}\b�
    ) 
    
    if (![string]::IsNullOrEmpty($FilePath) -and 
        (Test-Path -Path $FilePath -PathType Leaf)) 
    {
        $content = [IO.File]::ReadAllText($FilePath);
    } else {
        $content = $Content;
    }
    
    return (Select-String -InputObject $content -Pattern $regex -AllMatches).Matches
}

# returns name like 'Exchange2013_SP1'
function GetExchangeVersionName {
    $version = $(GCM Exsetup.exe | % {$_.FileVersionInfo}).ProductVersion;
    $versionName = "Exchange";  
    $versionDetails = $version.Split('.');
    
    $nv = [convert]::ToInt32($versionDetails[0], 10);
    if ($nv -eq 15){
        $versionName += "2013";
    } else {
        $versionName += "2010";
    }

    $sp = [convert]::ToInt32($versionDetails[1], 10);
    if ($sp -ne 0){
        $versionName += "_SP$sp";
    }

    return $versionName;
}

#endregion

#region code base

# returns the current DateTime minus $days
function GetTargetDateTimeReceived([int]$days){
    $shift = $(-1 * $days);
    $trgDt = $((Get-Date).AddDays($shift));
    return $trgDt;
}

# returns SearchFilterCollection where HasAttachments AND DateTimeReceived IsLessThan (NOW - TimeToLive)
function GetSearchFilterCollection($TargetDateTimeReceived){
    $filterHasAttachments = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo(
        [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::HasAttachments, 
        $true
    );

    $filterDateTimeReceived = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThan(
        [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::DateTimeReceived, 
        $TargetDateTimeReceived
    );

    return New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection(
        [Microsoft.Exchange.WebServices.Data.LogicalOperator]::And, 
        @($filterHasAttachments, $filterDateTimeReceived)
    );
}

# returns FolderId object by email address
function GetFolderId([string]$email) {
    $folder = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root;
    $mailbox = New-Object Microsoft.Exchange.WebServices.Data.Mailbox($email);
    $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId($folder, $mailbox);
    return $folderId;
}

# returns inited ExchangeService object by $email
function GetWebService([string]$email){
    $versionEnum = [Microsoft.Exchange.WebServices.Data.ExchangeVersion];
    $exchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Parse($versionEnum, $(GetExchangeVersionName));
    $connectingType = [Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress;

    $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($exchangeVersion);
    $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId($connectingType, $email);
    $service.HttpHeaders.Add("X-AnchorMailbox", $email);
    $service.AutodiscoverUrl($email);

    return $service;
}

# returns FolderView object with base PropertySet
function GetFolderView(){
    $folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView([int]::MaxValue);
    $folderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow; # Deep
    $folderView.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet(
        [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly,
        [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName
    );
    return $folderView;
}

# returns ItemView object with base PropertySet
function GetItemView(){
    $itemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView([int]::MaxValue);
    $itemView.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet(
        [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly,
        [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::HasAttachments,
        [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::DateTimeReceived
    );
    return $itemView;
}

# removes all attachments older than $ttl for each message in each account from $accounts
function RemoveAttachments([string[]]$accounts, [int]$ttl) {
    $targetDate = GetTargetDateTimeReceived -days $ttl;
    $filters = $(GetSearchFilterCollection -TargetDateTimeReceived $targetDate);

    log "Process Start Accounts Count: $($accounts.count) TTL: $ttl TargetDateTime: $targetDate";

    foreach ($EmailAddress in $accounts){
        try {
            $service  = $(GetWebService -email $EmailAddress);
            log "Process Account $EmailAddress"
        }
        catch [Exception] {
            log "Can't init WebService for '$EmailAddress' Message: $($_.Exception.Message)"
            continue
        }

        foreach($folder in $service.FindFolders($(GetFolderId -email $EmailAddress), $(GetFolderView))){
            if (([string]($folder.DisplayName)).ToUpper() -eq "SYSTEM"){
                continue
            }

            try {
                $res = $folder.FindItems($filters, $(GetItemView));
            } 
            catch [Exception] {
                log "Can't process Folder: '$($folder.DisplayName)' Message: $($_.Exception.Message)"
                continue
            }

            foreach ($item in $res.Items){
                # Redundancy Check but very important
                if (($item -is [Microsoft.Exchange.WebServices.Data.EmailMessage]) -and 
                    ($targetDateTimeReceived -lt $item.DateTimeReceived) -and 
                    $item.HasAttachments) 
                    {

                    $_item = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($service, $item.Id);

                    $must_rm = @();
                    foreach($attach in $_item.Attachments) {
                        log "Attachment. Size: $($attach.Size) Folder: '$($folder.DisplayName)' Name: $($attach.Name) IsInline: $($attach.IsInline) ContentType: $($attach.ContentType)"
                        # can't remove here. broken continuity of collection
                        $must_rm += $attach;
                    }

                    foreach ($attach in $must_rm){
                        $status = $_item.Attachments.Remove($attach);
                        if (! $status){
                            log "Can't delete $($attach.Name)"
                        }
                    }

                    if ($_item.IsDirty){
                        
                        if (! $DebugMode.IsPresent){
                            $_item.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite);
                        
                        } else {
                            log "Debug mode. Cancel deleting attachment $($attach.Name)"
                        }
                    }
                } else {
                    log "Futile message. Type: $($item.GetType()) DateTime: $($item.DateTimeReceived) Attachments: $($item.HasAttachments)"
                }
            }
        }
    }
}

<# 
Returns list of PrimarySmtpAddress from $Databases.
If $limit or/and $start redefined - returns top or portion of list
#>
function GetAccounts {
    param(
        [Parameter(Mandatory=$true)][string[]]$Databases,
        [Parameter(Mandatory=$false)][int]$start = -1,
        [Parameter(Mandatory=$false)][int]$limit = -1
    )

    $res = @();
    foreach ($db in $Databases) {
        foreach ($account in $(Get-Mailbox -Database $db -ResultSize unlimited | Select PrimarySmtpAddress)) {
            $res += $account.PrimarySmtpAddress;
        }
    }

    if ($limit -ne -1){
        if ($start -ne -1){
            return $res[$start..($start + ($limit - 1))]; # portion
        } else {
            return $res[0..($limit - 1)]; # top
        }
    } else {
        return $res; # all
    }
}

#endregion

#region main processing

SetCulture 

# Database choice
if ($AllDatabases.IsPresent) {
    log "AllDatabases IsPresent"
    $Databases = $(Get-MailboxDatabase)
} elseif ($Databases.Count -gt 0) {
    log "Databases: $($Databases -join ", ")"
} else {
    ntd
}

# main
if ($ChildCounter -gt 0) {
    log "Child counter: $ChildCounter"
    RemoveAttachments -accounts $(GetAccounts -Databases $Databases -start $Start -limit $Limit) -ttl $TimeToLive

} elseif ($AllAccounts.IsPresent){
    log "AllAccounts IsPresent"
    $accounts = $(GetAccounts -Databases $Databases);
    
    if ($Limit -le 0){
        $Limit = 1000;
    }

    $max_index = $($accounts.count - 1);
    if ($Limit -ge $max_index){
        log "Limit: $Limit less max_index: $max_index"
        RemoveAttachments -accounts $accounts -ttl $TimeToLive
        eow
    }

    $ChildCounter = 0;
    for ($i=0; $i -lt $max_index; $i += $Limit) {
        $ChildCounter ++;
        $cmd = "$selfFile $ChildCounter $i $Limit $LogDir $TimeToLive $($Databases -join ',')"
        log "cmd '$cmd'"
        Start-Process -WindowStyle Hidden "powershell.exe" $cmd
    }
    eow

} elseif (![string]::IsNullOrEmpty($AccountsFile) -and (Test-Path $AccountsFile -PathType Leaf)){    
    log "Accounts File exists: $AccountsFile"
    RemoveAttachments -accounts $(ExtractEmails -FilePath $AccountsFile) -ttl $TimeToLive

} elseif ($Account -match "@") {
    log "Process account $Account"
    RemoveAttachments -accounts $(ExtractEmails -Content $Account) -ttl $TimeToLive

} else {
    ntd
}

#endregion

eow

# ----- end of script ----- #
# log start
$fileName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Path)
Write-Host "Starting $fileName..."

# define outputCsv Path
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$scriptFolder = Split-Path -Parent $MyInvocation.MyCommand.Path
$outputCsv = Join-Path $scriptFolder "${fileName}_results_${timestamp}.csv"

# setup outlook application
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

# Choose mailbox in outlook 
Write-Host ""
Write-Host "Available Mailboxes:"
Write-Host "--------------------"

for ($i = 1; $i -le $namespace.Folders.Count; $i++)
{
    $folder = $namespace.Folders.Item($i)
    Write-Host "$i : $($folder.Name)"
}

$mailboxIndex = Read-Host "Select mailbox number"

if (-not ($mailboxIndex -as [int]) -or
    [int]$mailboxIndex -lt 1 -or
    [int]$mailboxIndex -gt $namespace.Folders.Count)
{
    Write-Host "Invalid selection. Using mailbox 1."
    $mailboxIndex = 1
}

$mailbox = $namespace.Folders.Item([int]$mailboxIndex)

Write-Host ""
Write-Host "Selected mailbox: $($mailbox.Name)"

# assign sent folder location
$sentFolder = $mailbox.Folders["Sent Items"]

# get folder location for received emails
$userFolder = Read-Host "Enter name of folder to compare received emails (default: Inbox)"
if ([string]::IsNullOrWhiteSpace($userFolder))
	{
		$userFolder = "Inbox"
	}
try
	{
		$receivedFolder = $mailbox.Folders[$userFolder]
	}
catch
	{
		Write-Host "Folder '$userFolder' not found. Using Inbox."
		$receivedFolder = $mailbox.Folders["Inbox"]
	}

# cutoff date range based on weeksback
$weeksback = Read-Host "Enter number of weeks to analyze (default: 4)"
if ("" -eq $weeksback)
{
    $weeksback = 4
}
elseif (-not ($weeksback -as [int]))
{
    Write-Host "Invalid entry. Defaulting to 4 weeks."
    $weeksback = 4
}
$cutoff = (Get-Date).AddDays(-([int]$weeksBack * 7))

# commitment identification patterns
$commitmentPatterns = @(
    '\b(i|we)\s+(will|can|plan\s+to|expect\s+to)\b',
    "\b(i'll|we'll)\b",
    '\bfollow\s*up\b',
    '\bneed\s+to\b'
)

# closed commitment identification patterns
$closurePatterns = @(
    '\battached\b',
    '\bsee attached\b',
    '\bprovided\b',
    '\bcompleted\b',
    '\bfinished\b',
    '\bresolved\b',
    '\bconfirmed\b',
    '\bdelivered\b',
    '\bupdated\b',
    '\bdone\b'
)

# empty messages list
$messages = New-Object System.Collections.Generic.List[Object]

# ---
# SENT ITEMS
# ---
$items = $sentFolder.Items
$items.Sort("[SentOn]", $true)
Write-Host "gathering sent items..."
foreach($item in $items)
{
	try
    {
        if($item.MessageClass -ne "IPM.Note"){ continue }
        if($item.SentOn -lt $cutoff){ break }

        $body = ""

        if($null -ne $item.Body)
        {
			$body = $item.Body
			$body = $body -replace '(?is)Principal Engineer.*$', ''
            $body = ($body -replace '\s+',' ')
			#Write-Host $body
        }

        $conversationId = [string]$item.ConversationID
        if (($null -eq $conversationId) -or ($conversationId -eq "") -or ([string]($conversationId).Trim().Length -eq 0))
		{
			$conversationId = "NOID_" + $item.Subject
		}

        $attachmentCount = 0
        try
        {
            $attachmentCount = $item.Attachments.Count
        }
        catch {}

        $messages.Add(
            [PSCustomObject]@{
                ConversationID = $conversationId
                Subject = $item.Subject
                Date = $item.SentOn
                Direction = "Sent"
                MessageDate = $item.SentOn
                Body = $body
                HasAttachments = ($attachmentCount -gt 0)
            }
        )
    }
    catch {}
}

# ---
# RECEIVED ITEMS
# ---
$items = $receivedFolder.Items
$items.Sort("[ReceivedTime]", $true)
Write-Host "gathering received items..."
foreach($item in $items)
{
  	try
    {
        if($item.MessageClass -ne "IPM.Note"){ continue }
        if($item.ReceivedTime -lt $cutoff){ break }
		
		$conversationId = [string]$item.ConversationID

        if ([string]::IsNullOrWhiteSpace($conversationId)){ continue }

        $messages.Add(
            [PSCustomObject]@{
                ConversationID = $conversationId
                Subject = $item.Subject
                Date = $item.ReceivedTime
                Direction = "Received"
                MessageDate = $item.ReceivedTime
                Body = ""
                HasAttachments = $false
            }
        )
    }
    catch {}
}

# ---
# RESULTS
# ---
Write-Host "building conversation results..."
$conversationResults =
foreach ($conversation in ($messages | Group-Object ConversationID))
{
	$timeline =
        $conversation.Group |
        Sort-Object Date

    $msgs =
        $timeline |
        Where-Object { $_.Direction -eq "Sent" }
    if ($msgs.Count -eq 0){ continue }
	
	$msg =
	$msgs |
	Sort-Object SentOn -Descending |
	Select-Object -First 1
	if ($null -eq $msg){ continue }
	
	$commitmentCount = 0
	$closedCount = 0
	$outstandingCount = 0
	$score = 0
	
	$text = ($msg.Subject + " " + $msg.Body).ToLower()
    
	$lastMessage =
        $timeline |
        Sort-Object Date -Descending |
        Select-Object -First 1

    $latestSent =
        $msgs |
        Sort-Object MessageDate -Descending |
        Select-Object -First 1
	
	$daysOpen =
	[Math]::Round(
		((Get-Date) - $lastMessage.Date).TotalDays
	)

	foreach($commitmentRegex in $commitmentPatterns)
	{
		if($text -match $commitmentRegex)
		{
			# Write-Host ""
			# Write-Host "MATCHED COMMITMENT"
			# Write-Host "Conversation:" $conversation.Name
			# Write-Host "Subject:" $msg.Subject
			# Write-Host "Pattern:" $commitmentRegex
			# Write-Host "Text:" $text.Substring(0,[Math]::Min(200,$text.Length))
			$commitmentCount = 1
			$outstandingCount = 1
			$score = 5
			if ($daysOpen -gt 21 -lt 31)
			{
				$score += 6
			}
			elseif ($daysOpen -gt 14 -lt 31)
			{
				$score += 4
			}
			elseif ($daysOpen -gt 7 -lt 31)
			{
				$score += 2
			}
			break
		}
	}

    if ($outstandingCount -gt 0 -and $lastMessage.Direction -eq "Received")
	{
		$status = "WAITING"
		$score = 0
	}
	elseif ($outstandingCount -gt 0)
	{
		$status = "OPEN"
	}
	else
	{
		$status = "CLOSED"
	}

    [PSCustomObject]@{
        DaysSinceLastMessage    = $daysOpen
        OutstandingCommitments  = $outstandingCount
        Subject                 = $latestSent.Subject
        LastMessageDate         = $lastMessage.Date
        LastDirection           = $lastMessage.Direction
        CommitmentCount         = $commitmentCount
        ClosedCommitmentCount   = $closedCount
        ActionScore             = [Math]::Max(0,$score)
        ConversationID          = $conversation.Name
        Status                  = $status
		Body					= $lastMessage.Body			
    }
	
	# Write-Host ""
	# Write-Host "Subject:" $latestSent.Subject
	# Write-Host "Commitments:" $commitmentCount
	# Write-Host "Outstanding:" $outstandingCount
	# Write-Host "Days Open:" $daysOpen
	# Write-Host "Final Score:" $score
}

$conversationResults |
Sort-Object ActionScore -Descending |
Export-Csv $outputCsv -NoTypeInformation -Encoding UTF8
Write-Host ""
Write-Host "Output written to:"
Write-Host $outputCsv


Write-Host ""
Write-Host "Top Open Conversations"
$conversationResults |
Where-Object {$_.Status -eq "OPEN"} |
Sort-Object ActionScore -Descending |
Select-Object -First 5 |
Format-Table -AutoSize
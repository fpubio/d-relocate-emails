param(
    [Parameter()]
    [Switch]$print_config = $false
)

### VARIABLES defined by end-user

# $config_path specifies the location of the configuration file.

# Consider using an absolute path:
# $config_path='c:/Users/John.Smith/my/path/to/d-relocate-emails.ini'

# Default to looking in parent directory of the script:
$config_path = Join-Path (Split-Path -Path $PSScriptRoot) 'd-relocate-emails.ini'

# Number of days to filter over (filter over emails from the previous
# NN days)
$previous_days = 90             # past 90 days

$verbosep = $true               # increase the chatter

### Read config file

Get-Content $config_path | foreach-object -begin {$config=@{}} -process { $k = [regex]::split($_,'='); if(($k[0].CompareTo("") -ne 0) -and ($k[0].StartsWith("[") -ne $True)) { $config.Add($k[0], $k[1]) } }

### VARIABLES defined via the configuration file

# The name of the top-level folder. A string.
# Example: 'username@somedomain.org'
# It may be shown in the Outlook client interface as the name of the
# folder which contains 'Inbox', 'Drafts', 'Sent Items', etc.
$mailbox_name = $config.mailbox_name

# The full path for the CSV file specifying the rules for filtering. A
# string.
$rules_path = $config.rules_path

# The full path for logging processed emails. A string.
$record_path = $config.record_path

# Option to print configuration and exit
if ($print_config) {
    Write-Host "Configuration:"
    Write-Host $mailbox_name
    Write-Host $rules_path
    Write-Host $record_path
    exit
}

### VARIABLES (internal)

# Import rules
try {
    $instruction = Import-Csv $rules_path
}
catch {
    Write-Host "Couldn't read rules CSV"
    Write-Host $rules_path
}

# for logging processed emails
$matched_Emails = @()

# Access Outlook application
Add-Type -Assembly "Microsoft.Office.Interop.Outlook"
$Outlook = New-Object -ComObject Outlook.Application
$namespace = $Outlook.GetNameSpace("MAPI")

# The namespace folders collection -- all folders in the namespace
$folders = $namespace.Folders

# Access the mailbox
$current_mailbox = $namespace.Folders($mailbox_name)

$inbox = $current_mailbox.Folders("Inbox")

# FIXME: Preferable to fall back to this if $mailbox_name isn't defined by the ini file
# Default inbox for the account
#$inbox = $namespace.GetDefaultFolder(6)

$inbox_items = $inbox.items

$inbox_size = $inbox_items.Count

#
# filter items
#
$ourdate = (Get-Date).AddDays(-1*$previous_days).ToShortDateString() + " 00:00"

$ourfilter = "[SentOn] > '$ourdate'"
#$ourfilter = "[UnRead] = True"

$newEmails = $inbox_items.Restrict($ourfilter)

if ($verbosep) { "{0} emails under consideration" -f $newEmails.count | Write-Host }

$from_rules = New-Object System.Collections.Generic.List[System.Object]
$subject_rules = New-Object System.Collections.Generic.List[System.Object]
foreach ($row in $instruction){
    $rowsubject = $row.Subject
    $rowfrom = $row.From
    $destination_folder_name = $row.DestinationFolder
    if ($rowsubject) {
        $subject_rule = @($rowsubject,$destination_folder_name)
        $subject_rules.Add($subject_rule)
    }
    if ($rowfrom) {
        $from_rule = @($rowfrom,$destination_folder_name)
        $from_rules.Add($from_rule)
    }
}

# if show_rules_only_p ... {
# Write-Output "subject_rules"
# Write-Output $subject_rules
# Write-Output "from_rules"
# Write-Output $from_rules
#exit
# ... }

if ($verbosep) { Write-Host "Beginning matching" }

#$newEmails = @('Zero') # debug

# Matching email on rules
foreach ($email in $newEmails){
    # Preferable:
    # - a single set of rules and a generic test function -- rather than looping through separate subject and from rules
    # - single move function rather than duplicating code
    foreach ($subject_rule in $subject_rules) {
        $rule_subject = $subject_rule[0]
        $destination_folder_name = $subject_rule[1]
        $emailsubject = $email.Subject
        if ($emailsubject -like "*$rule_subject*"){
	    if ($verbosep) { "matched on {0}" -f $emailsubject | Write-Host }
            # Only support target folders which are directly
            # accessible via the top-level folder
            $destination_folder = $current_mailbox.Folders($destination_folder_name)
            # Move emails
            $email.move($destination_folder) | Out-Null
            # Add the moved email to list to record later
            $email| add-member -MemberType NoteProperty -Name "Folderpath" -Value $destination_folder
            $timestamp = Get-Date
            $email| add-member -MemberType NoteProperty -Name "MovedOn" -Value $timestamp
            $matched_Emails += $email
        }
    }
    # test against From
    foreach ($from_rule in $from_rules) {
        $rule_from = $from_rule[0]
        $destination_folder_name = $from_rule[1]
        # SenderEmailAddress may or may not be defined
        # SenderName, Sender
        $from_senderemailaddress = $email.SenderEmailAddress
        $from_sendername = $email.SenderName
        # FIXME: we should almost certaily be using an exact match here, not -like w/wildcards
        if (($from_sendername -like "*$rule_from*") -or
            ($from_senderemailaddress -like "*$rule_from*")) {
		if ($verbosep) { "matched on {0} or {1}" -f $from_sendername,$from_senderemailaddress | Write-Host  }
                $destination_folder = $current_mailbox.Folders($destination_folder_name)
                # Move email
                $email.move($destination_folder) | Out-Null
                # Add the moved email to list to record later
                $email| add-member -MemberType NoteProperty -Name "Folderpath" -Value $destination_folder
                $timestamp = Get-Date
                $email| add-member -MemberType NoteProperty -Name "MovedOn" -Value $timestamp
                $matched_Emails += $email
            }
    }
}

### Export records of matched emails

$matched_Emails|Select-Object -Property Subject, SenderName, SentOn, Folderpath, MovedOn|
    Export-Csv -path $record_path -NoTypeInformation -Append

if ($verbosep) { Write-Host "Logged to " $record_path }

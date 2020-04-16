Clear-Host

#1) First thing’s first, import the csv as a variable.  Use this variable for all subsequent tasks.
Write-Host "Importing data into a variable" -ForegroundColor Green
$userData = Import-Csv "C:\temp\Powershell Assessment\Users.csv" -Delimiter ","
Write-Output $userData |  Out-GridView

#2) How many users are there?
Write-Host ""
Write-Host "Total number of users: " $userData.Count -ForegroundColor Yellow

#3) What is the total size of all mailboxes?
Write-Host ""
Write-Host "Total size of all Mailboxes:" -ForegroundColor DarkMagenta
($userData.MailboxSizeGB | Measure-Object -Sum).Sum

#5) Same as question 3, but limited only to Site: NYC
Write-Host ""
Write-Host "Same as question 3, but limited only to Site: NYC" -ForegroundColor Yellow
(($userData | ?{$_.Site -eq "NYC"}).MailboxSizeGB | Measure-Object -Sum).Sum

#4) How many accounts exist with non-identical EmailAddress/UserPrincipalName? Be mindful of case sensitivity.
Write-Host ""
Write-Host "How many accounts exist with non-identical EmailAddress/UserPrincipalName? Be mindful of case sensitivity" -ForegroundColor DarkCyan
(Compare-Object -ReferenceObject $userData.EmailAddress -DifferenceObject $userData.UserPrincipalName -CaseSensitive).count

#6) How many Employees (AccountType: Employee) have mailboxes larger than 10 GB? (remember MailboxSizeGB is already in GB.)
Write-Host ""
Write-Host "How many Employees (AccountType: Employee) have mailboxes larger than 10 GB" -ForegroundColor Green
($userData | ?{($_.AccountType -eq "Employee") -and ([int]$_.MailboxSizeGB -gt 10)}).count

<#7) Provide a list of the top 10 users with EmailAddress @domain2.com in Site: NYC by mailbox
size, descending.
a. The boss already knows that they’re @domain2.com; he wants to only know their
usernames, that is, the part of the EmailAddress before the “@” symbol.  There is
suspicion that IT Admins managing domain2.com are a quirky bunch and are encoding
hidden messages in their directory via email addresses.  Parse out these usernames (in
the expected order) and place them in a single string, separated by spaces – should look
like: “user1 user2 … user10”#>
Write-Host ""
Write-Host "Provide a list of the top 10 users with EmailAddress @domain2.com in Site: NYC by mailbox size, descending."
$userNames = $userData | ?{($_.Site -eq "NYC") -and ($_.EmailAddress -like "*@domain2.com")} |`
Select-Object -property @{Label=”User Name”;Expression={($_.EmailAddress).Replace("@domain2.com","")}} | `
Sort-Object -Property MailBoxSizeGB -Descending

<#8) Create a new CSV file that summarizes Sites, using the following headers: Site, TotalUserCount,
EmployeeCount, ContractorCount, TotalMailboxSizeGB, AverageMailboxSizeGB
a. Create this CSV file based off of the original Users.csv.  Note that the boss is picky when
it comes to formatting – make sure that AverageMailboxSizeGB is formatted to the
nearest tenth of a GB (e.g. 50.124124 is formatted as 50.1).  You must use PowerShell to
format this because Excel is down for maintenance.#>
Write-Host ""
Write-Host "Create a new CSV file that summarizes Sites, using the following headers: Site, TotalUserCount,
EmployeeCount, ContractorCount, TotalMailboxSizeGB, AverageMailboxSizeGB"

$sites = $userData.Site | select -Unique
$resultArray = @()
$newCSVPath = "C:\temp\Powershell Assessment\newUser.csv"

try
{
    foreach($site in $sites)
    {
        $resultArray += New-Object -TypeName psobject -Property @{Site= $site; 
                                                                    TotalUserCount = ($userData | ?{$_.Site -eq $site}).count;
                                                                    EmployeeCount = GetTotalNumberOf -siteinQuestion $site -AccountToLookFor "Employee";
                                                                    ContractorCount = GetTotalNumberOf -siteinQuestion $site -AccountToLookFor "Contractor";
                                                                    TotalMailboxSizeGB = (($userData | ?{$_.Site -eq $site}).MailboxSizeGB | Measure-Object -Sum).Sum;
                                                                    AverageMailboxSizeGB = (($userData | ?{$_.Site -eq $site}).MailboxSizeGB | Measure-Object -Average).Average.tostring("#.#")}
    }
    
    $resultArray | Export-Csv $newCSVPath -NoTypeInformation

    if(Test-Path($newCSVPath))
    {
        Write-Output "New CSV created"
    }
    else
    {
        throw "File not created."
    }
}
catch
{
    Write-Output $Error[0].ErrorDetails   
}
finally
{
    $resultArray = $null
}


function GetTotalNumberOf
{
    Param(
        [parameter(Mandatory=$true)]
        [String]$siteinQuestion,

        [parameter(Mandatory=$true)]
        [String]$AccountToLookFor
    )
   return ($userData | ?{($_.Site -eq $siteinQuestion) -and ($_.AccountType -eq $AccountToLookFor.ToLower())}).count;
}
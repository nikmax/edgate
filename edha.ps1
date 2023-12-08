function Select-CsvFile {
    
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null

    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
	$FileBrowser.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    $FileBrowser.ShowHelp = $true
	$FileBrowser.filter = "Txt (*.txt)|*.txt|Csv (*.csv)| *.csv"
	[void]$FileBrowser.ShowDialog()
    $FileBrowser.FileName
}
function Validate-CsvFile ($file, $delimiter){
    $read = New-Object System.IO.StreamReader($file)

    $CsvHeaders = $read.ReadLine()
    $validHeaders = "Gender","GivenName","Surname","vorname","nachname"  
    $validCount = $validHeaders.Count -1
    $isValid = $true

    # Validate CSV Headers
    #    
    [System.Environment]::NewLine
    foreach ($item in $validHeaders) { 
        if ($CsvHeaders -notmatch $item) { 
            Write-Error "Die Header der CSV Datei sind fehlerhaft. "
            $isValid = $false 
            exit
        }
    }
    #Validate CSV Rows
    #    
    $counter = 0
    while (($line = $read.ReadLine()) -ne $null) {
        $counter ++
        $total = $line.Split($delimiter).Length - 1;
        if ($total -ne $validCount) {
            Write-Error "CSV Datei ist fehlerhaft in Zeile $counter."
            $isValid = $false
            exit
        }
    }
    [System.Environment]::NewLine
    if ($isValid -eq $false) {
        Write-Warning "Fehler, Datei nicht gefunden"
        exit
    }else{
      #$Import | Format-Table
      Write-Warning "eingelesen: $counter."
    }
    $read.Close()
    $read.Dispose()
    
}
function User-Login {
  $cookie = New-Object System.Net.Cookie
  $login = read-host -prompt "Enter login"
  $securedValue = Read-Host "Enter password" -AsSecureString
  $sess = New-Object Microsoft.PowerShell.Commands.WebRequestSession
  $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securedValue)
  $psw = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)


  $post = @{rm='form_login';login=$login;kennwort=$psw}


  $res = Invoke-WebRequest "http://$url$script" -Method POST -Body $post -SessionVariable $sess
  #Write-Host $res.RawContent
  if ($res.RawContent -match 'type=password'){
    Write-Host "Sorry! Falsche Daten."
    exit
   }else{
    if( $res.RawContent -match  '(Set-Cookie: )(\w+)=(.+);(.+)' ) {
             $cookie.Domain = $url
             $cookie.Name = $Matches[2]
             $cookie.Value = $Matches[3]
             $session.Cookies.Add($cookie)
             #Write-Host $cookie.Name "=" $cookie.Value
    }
  }
}
######################## main ####################

Add-Type -AssemblyName System.Web

$logfile = "create_user_log.log"

$url="edga.ncc.eurodata.de"
$script="/cgi-bin/index.pl"

if($session -eq $null){ $global:session = New-Object Microsoft.PowerShell.Commands.WebRequestSession }
if($session.Cookies.Count -eq 0){ User-Login }
$session2= $session
### 1. ### 
# frage nach Email/Suchbegriff
###
$suchbegriff = Read-Host -Prompt "Email/Suchbegriff"
$suchbegriff = [System.Web.HttpUtility]::UrlEncode($suchbegriff)
### 2. ###
# liste alle entsprechend dem Suchmuster  gefundenen Email-IDs  
###
$request = "http://{0}/cgi-bin/index.pl?rm=email_main&stichwort={1}" -f $url,$suchbegriff
$response = Invoke-WebRequest $request -WebSession $session
$pattern  = '(.+email_id=)(\d+)([^>]+>)([^<]+)(.+)'
$r_emails = Select-String -InputObject $response -Pattern $pattern -AllMatches
"{0} Email-Adresse(n) gefunden zu diesem Suchbegriff" -f $r_emails.Matches.Count
Write-Host 'Davon sind folgende Accounts verfügbar:'
Write-Host '-----------------------------------------'
#
# es können mehrere Email eines Account gefunden worden sein - alle durchsuchen
#
$n = 1
$accounts = @()
foreach ($i in $r_emails.Matches) {
  $email_id = $i.Groups[2].Value
  $email    = $i.Groups[4].Value

  ### 3. ###
  # liste alle gefundenen Account-IDs auf  
  ###
  $pattern = '(/show_emails_by_person\.asp\?id=\d*)'
  $request="http://{0}/show_email.asp?email_id={1}" -f $url,$email_id
  $response = Invoke-WebRequest $request -WebSession $session
  $r_accounts = Select-String -InputObject $response -Pattern $pattern -AllMatches
  
  #
  if ($r_accounts.Matches.Count -ne 0) {
    if( -not $accounts.Contains($r_accounts.Matches.Groups[1].Value))
    {$accounts+= $r_accounts.Matches.Groups[1].Value
    #if($n -gt 9) { $s = ''} else {$s = ' '}
    #"{0}{1}) Account HauptEmailID: {3} # {4}" -f $s, $n, $email_id, $email
    $n = $n + 1
    }
  }
}  
$n = $n - 1
if ($n -eq 0) { Write-Host 'Keine Postfächer gefunden' }
else {
  "{0} Accounts gefunden" -f $accounts.Length
  $pattern = '(.+pf_id.+value=[^\d]+)(\d+)(.+)'
  foreach( $cmd in $accounts){
    ### 4. ###
    # Suche nach den Postfächern  
    ###
    $request  ="http://{0}{1}" -f $url,$cmd
    $response = Invoke-WebRequest $request -WebSession $session
    $mmm = Select-String -InputObject $response -Pattern $pattern -AllMatches 
    #"E-Mail-Ad in diesem Account: {0}" -f $mmm.Matches.Count
    foreach ($iii in $mmm.Matches) {write-host "#########PF IDs : "$iii.Groups[2].Value}
  }
}


exit


#if( $i -match '(.+email_id=)(?<id>\d+)(.+)' ) {  write-host $Matches.id} # else { Write-Host "nix gefunden"}
# anlegen # $cmd="?rm=insert_person"
          # $params="&login=&pass=&domain_id1=100002&edgates1=12615544&squid_select=&squid_edgates=12615544&oe_id=4514232&b_login_id="
#suchen
          $cmd="?rm=email_main"
          $params=""


$CsvPath = Select-CsvFile
Validate-CsvFile -file $CsvPath -delimiter ","
$Import = Import-CSV $CsvPath -Encoding ASCII

foreach ($user in $Import) {
    $Anrede = $user.Gender
    $GivenName = [System.Web.HttpUtility]::UrlEncode($user.GivenName)
    $Surname = [System.Web.HttpUtility]::UrlEncode($user.Surname)
	$vorname = $user.vorname
	$nachname = $user.nachname

    $person="&anrede=$Anrede&vname=$GivenName&nname=$Surname&email1=$vorname.$nachname"

    $email= "&stichwort=$vorname.$nachname%40etl.de"

    $date = Get-Date -Format "yyyy-MM-dd_HH:mm:ss"

    try {
        #create 
            # $req = "http://$url$script$cmd$params$person"
        #search #
            $req = "http://$url$script$cmd$params$email"
            #write-host $req
         $response = Invoke-WebRequest $req -WebSession $session
      
         Add-content $logfile -value $date" Success "$Surname" "$Response
         Add-content $logfile -value ""
         if( $Response -match '(.+email_id=)(\d+)(.+)' ) {
           $id = $Matches[2]
           $Surname = [System.Web.HttpUtility]::UrlDecode($Surname)
           $GivenName = [System.Web.HttpUtility]::UrlDecode($GivenName)
           Write-Host  $GivenName" "$Surname": "$id
         }else {
           Write-Host  $GivenName" "$Surname": nicht gefunden"
         }
    }
    catch {
      Add-content $logfile -value $date" Error "$Surname" "$_.Exception.Message
      Add-content $logfile -value ""
    }
}

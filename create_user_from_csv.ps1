
cls
Add-Type -AssemblyName System.Web

$logfile = "create_user_log.log"

$url="edga.ncc.eurodata.de"
$script="/cgi-bin/index.pl"


$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$cookie = New-Object System.Net.Cookie

User-Login


# anlegen # $cmd="?rm=insert_person"
          # $params="&login=&pass=&domain_id1=100002&edgates1=12615544&squid_select=&squid_edgates=12615544&oe_id=4514232&b_login_id="
#suchen
          $cmd="?rm=email_main"
          $params=""

### main loop ###

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
            write-host $req
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
  $login = read-host -prompt "Enter login"
  $securedValue = Read-Host "Enter password" -AsSecureString
  $sess = New-Object Microsoft.PowerShell.Commands.WebRequestSession
  $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securedValue)
  $psw = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)


  $post = @{rm='form_login';login=$login;kennwort=$psw}


  #$post = @{rm='form_login';login='ansa038';kennwort='8533Eg7739!'}
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

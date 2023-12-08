function main_menu  {

Write-Host '+----------------------------------------------+'
Write-Host '|                  Hauptmenü                   |'
Write-Host '+----------------------------------------------+'
Write-Host ' z: Postfächer anzeigen'
Write-Host ' e: EDGA User'
#Write-Host ' i: Postfächer aus Datei erstellen/aktivieren'
Write-Host ' a: deaktivierte(s) Postfach(fächer) aktivieren'
Write-Host ' q: Script beenden'
Write-Host '+----------------------------------------------+'
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
    if($validHeaders -eq '') {$validHeaders = "Gender","GivenName","Surname","vorname","nachname" }
    $CsvHeaders = $read.ReadLine() 
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
function edga-login {
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
function edga {
  Add-Type -AssemblyName System.Web
  $url="edga.ncc.eurodata.de"
  $script="/cgi-bin/index.pl"
  if($session -eq $null){ $global:session = New-Object Microsoft.PowerShell.Commands.WebRequestSession }
  if($session.Cookies.Count -eq 0){ edga-login }
  $session2= $session
}
function mb_zeigen {
  Write-Host '+----------------------------------------------+'
  Write-Host '|          Postfacher auflisten                |'
  Write-Host '+----------------------------------------------+'
  Write-Host ''
  $EtlNr = Read-Host -Prompt 'ETL Nr. eingeben'
  $data = @()

  Write-Host "+----------------------------------------------------+------------+-------------+-------------+-------------+"
  "| {0,-50} | {1,-10} | {2,-11} | {3,-11} | {4,-11} |" -f "Email","erstellt","   390", "   699", "deaktiviert"
  Write-Host "+----------------------------------------------------+------------+-------------+-------------+-------------+"

  foreach ($a in (Get-Mailbox -Filter "customAttribute15 -eq $EtlNr" )){
    $data += $a
    "| {0:d3} {1,-46} | {2,-10} | {3,-11} | {4,-11} | {5,-11} |" -f $data.length ,$a.WindowsEmailAddress ,$a.customAttribute5,$a.customAttribute1,$a.customAttribute2,$a.AccountDisabled
  }
  if ($data.length -eq 0) { Write-Host "+  es gibt keine Postfächer zu der ETL-Nummer      +            +             +             +             +"}
  Write-Host "+----------------------------------------------------+------------+-------------+-------------+-------------+"

}
function mb_aktivieren {
  Write-Host '+----------------------------------------------+'
  Write-Host '|           Aktivieren der Postfacher          |'
  Write-Host '+----------------------------------------------+'
  Write-Host 'Es werden alle deaktivierten Postfächer einer Kanzlei'
  Write-Host 'wieder aktiviert (alle am Stück oder einzeln bestätigt)'
  Write-Host ''
  $EtlNr = Read-Host -Prompt 'ETL Nr. eingeben'
  $data = @()

  Write-Host "+----------------------------------------------------+------------+-------------+-------------+-------------+"
  "| {0,-50} | {1,-10} | {2,-11} | {3,-11} | {4,-11} |" -f "Email","deakt.am","   390", "   699", "deaktiviert"
  Write-Host "+----------------------------------------------------+------------+-------------+-------------+-------------+"


  foreach ($a in (Get-Mailbox -Filter "customAttribute15 -eq $EtlNr" )){
   $t = $false
   if ($a.customAttribute1 -eq "deaktiviert") { $t = $true }
   if ($a.customAttribute2 -eq "deaktiviert") { $t = $true }
   if ($a.customAttribute6 -ne '' ){ $t = $true }
   if ($a.AccountDisabled -eq $true ){ $t = $true }
   if ($a.WindowsEmailAddress.Address.indexOf($EtlNr) -gt 0) { if ($a.WindowsEmailAddress.Address -ne $a.UserPrincipalName) { $t = $true } }
   if ($t){
     $data += $a.Guid
     "| {0:d3} {1,-46} | {2,-10} | {3,-11} | {4,-11} | {5,-11} |" -f $data.length ,$a.WindowsEmailAddress ,$a.customAttribute6,$a.customAttribute1,$a.customAttribute2,$a.AccountDisabled
   }
  }
  if ($data.length -eq 0) { Write-Host "+  es gibt keine deaktivierten Postfächer            +            +             +             +             +"}
  Write-Host "+----------------------------------------------------+------------+-------------+-------------+-------------+"

  if ($data.length -ne 0) {
    $items = @()
    $b = Read-Host -Prompt '(a)lle aktivieren oder (e)inzeln oder Nummer eingeben oder a(b)brechen'

    if ($b -eq 'b') { return }
    if ( ($b -ne "a") -and ($b -ne "e") ) { if($b -lt 1) { break } $items += $data[$b-1] }
    else { $items = $data }
    foreach($Guid in $items){
      if ($Guid) {

          

          if( $b -eq 'e' ){
              Write-Host (Get-Mailbox -Identity "$Guid" | select WindowsEmailAddress).WindowsEmailAddress.Address -NoNewline
              $c = Read-Host -Prompt " aktivieren n/[J]"
              if ($c -eq 'n') { continue }
          }
              
          $mb = Get-Mailbox -Identity "$Guid"

          if ($mb.HiddenFromAddressListsEnabled -eq $true){ Set-Mailbox -Identity "$Guid" -HiddenFromAddressListsEnabled $false }
          if ($mb.CustomAttribute6 -ne '')                { Set-Mailbox -Identity "$Guid" -CustomAttribute6 $false }
          if ($mb.CustomAttribute1 -eq "deaktiviert")     { Set-Mailbox -Identity "$Guid" -CustomAttribute1 "390" }
          if ($mb.CustomAttribute2 -eq "deaktiviert")     { Set-Mailbox -Identity "$Guid" -CustomAttribute2 "699" }
          if ($mb.CustomAttribute5 -eq '')                { Set-Mailbox -Identity "$Guid" -CustomAttribute5 (Get-Date).ToShortDateString() }
          
          $mails = @()
          ForEach ($smtp in $mb.EmailAddresses) {
              if( $smtp.IsPrimaryAddress ) { $mails += "SMTP:" +  $smtp.SmtpAddress.Replace($EtlNr,'') }
              else { $mails += $smtp.SmtpAddress.Replace($EtlNr,'') }
          }
          Set-Mailbox -Identity "$Guid" -EmailAddresses $mails
          Set-Mailbox -Identity "$Guid" -AccountDisabled $false
          Write-Host -ForegroundColor Green $mb.UserPrincipalName "aktiviert"
      

      
      } # if $Guid
    }
  }
}



###### main loop ########


Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

do{
    main_menu
    $cmd = Read-Host "Was möchtest du tun"
    Write-Host ''
    switch ($cmd) {
      #'e'  { edga }
      'z'  { mb_zeigen }
      'a'  { mb_aktivieren }
      'q'  { exit }
    }
}until($cmd -eq 'q')
    

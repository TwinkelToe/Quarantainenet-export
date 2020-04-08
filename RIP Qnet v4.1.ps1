#########################
##       Setup VAR     ##
#########################

$QnetUsername = ""
$QnetPassword = ""
$QmangeURL = ""
$DesFile = "c:\temp\DefinitieveExportQnet.csv"

# Aanmaken Array #
$QnetExport = @()

class QnetComputer {
    
    [string]$MAC
    [string]$AccessGroup
    [string]$RealName
    [string]$Description
    [string]$RegBy
    [string]$TimeReg
    [string]$Expire


    QnetComputer([string]$MACaddr, $AccessGroup, $RN, $Desc, $RegBy) {
        $this.MAC = $MACaddr
        $this.RealName = $RN
        $this.Description = $Desc
        $this.AccessGroup = $AccessGroup
        $this.RegBy = $RegBy
    }

}

"Starten Qnet"
$r=Invoke-WebRequest -Uri $($QmangeURL + "/admin/login") -Method GET -SessionVariable td

# Use the session variable that you created in Example 1. Output displays values for Headers, Cookies, Credentials, etc. $td
# Gets the first form in the Forms property of the HTTP response object in the $r variable, and saves it in the $form variable. 
$form = $r.Forms[0]
# Pipes the form properties that are stored in the $forms variable into the Format-List cmdlet, to display those properties in a list. 
$dump = $form | Format-List
# Displays the keys and values in the hash table (dictionary) object in the Fields property of the form.
#$form.fields
# The next two commands populate the values of the "email" and "pass" keys of the hash table in the Fields property of the form. Of course, you can replace the email and password with values that you want to use. 
$form.Fields["username"] = $QnetUsername
$form.Fields["password"] = $QnetPassword
# The final command uses the Invoke-WebRequest cmdlet to sign in to the Facebook web service. 
$r=Invoke-WebRequest -Uri $($QmangeURL + $form.Action) -WebSession $td -Method POST -Body $form.Fields
# We heben nu een session cookie onder $td

$QnetResult = Invoke-RestMethod -Uri $($QmangeURL + "/admin/registrations/macs?current_search=&page=1&button=%3E") -WebSession $td -Method Get

#Totaal aantal pagina's opzoeken. Op dit moment 480
$a = $QnetResult.IndexOf('paging_browser_total')
$b = $QnetResult.IndexOf("</span", $a)
$TotalPages = $QnetResult.Substring($a + 22, $b - $a - 22)
$TotalPages = [convert]::ToInt32($TotalPages, 10)
"Totaal paginas: " + $TotalPages

#Beginnen bij 0 met tellen
#$TotalPages = $TotalPages - 1
#Debug aantal pagina's
#$TotalPages = 1

#Download pagina 1 tot x
for($c=0; $c -le $TotalPages; $c++) {

    #Per pagina het begin en einde van de tabel opzoeken.
    $TableBodyStart = $QnetResult.IndexOf("tbody")
    $TableBodyEnd = $QnetResult.IndexOf("/tbody")

    #Reset Positie op pagina
    $a = 0
    $b = 0
    
    #De maximaal 20 records per pagina doorlopen 
    for($i=0; $i -le 19; $i++) {
        
        #Als we zoeken buiten de tabel stoppen (alleen bij laaste pagina van toepassing)
        if ($b -gt $TableBodyEnd) {
            break
        }

        #Als dit de eerste ronde is begin bij de tabel. Anders vergaan waar gebleven.
        if ($i -eq 0) {
            $a = $TableBodyStart
        }

        #Knip het MAC adres
        $a = $QnetResult.IndexOf("tooltip_mac", $a)
        $b = $QnetResult.IndexOf("span", $a)
        $MacAddress = $QnetResult.Substring($a + 13, $b - $a - 15)

        #Knip de accessgroup
        $a = $QnetResult.IndexOf("vlan_name", $b)
        $a = $QnetResult.IndexOf('">', $a)
        $b = $QnetResult.IndexOf("</a></td>", $a)
        $AccessGroup = $QnetResult.Substring($a + 2, $b - $a - 2)

        #Knip de realname
        $a = $QnetResult.IndexOf('"field">', $b)
        $b = $QnetResult.IndexOf("</td>", $a)
        $RealName = $QnetResult.Substring($a + 8, $b - $a - 8)

        #Reg by
        $a = $QnetResult.IndexOf('email', $b)
        $a = $QnetResult.IndexOf('field', $a + 3)
        $a = $QnetResult.IndexOf('field', $a + 3)
        $b = $QnetResult.IndexOf("</td>", $a)
        $RegisteredBy = $QnetResult.Substring($a + 7, $b - $a -7)

        #Knip de description
        $a = $QnetResult.IndexOf('field', $a + 3)
        $b = $QnetResult.IndexOf("</td>", $a)
        $Description = $QnetResult.Substring($a + 7, $b - $a - 7)
        
        #Time Registration

        #Time Expire
        #Ter info
        $MacAddress + " | " + $AccessGroup + " | " + $RealName + " | " + $Description + " | " + $RegisteredBy    

        #Wegschrijven in object.
        $QnetExport += New-Object QnetComputer($MacAddress, $AccessGroup, $RealName, $Description, $RegisteredBy)
    }

    #Download volgende pagina. 
    $t = $c + 2
    Start-Sleep -Milliseconds 500
    "Download Pagina: " + $t
    $QnetResult = Invoke-RestMethod -Uri $($QmangeURL + "/admin/registrations/macs?current_search=&page=" + $t + "&button=%3E") -WebSession $td -Method Get
}

$QnetExport | Export-Csv -Path $DesFile
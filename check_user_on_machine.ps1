 
                # ------------------Global Variables------------------- #
                $LugaroRedazione = "OU=Redazione,OU=Piano0,OU=Lugaro,OU=LaStampa,OU=Struttura AD 2008,DC=lastampa,DC=prv"
                $LugaroUffici = "OU=Piano1,OU=Lugaro, OU=LaStampa, OU=Struttura AD 2008, DC=lastampa, DC=prv"
                $Abbonamenti = "OU=Abbonamenti,OU=Piano0,OU=Lugaro,OU=LaStampa,OU=Struttura AD 2008,DC=lastampa,DC=prv"
                $Altro = "OU=Altro,OU=Piano0,OU=Lugaro,OU=LaStampa,OU=Struttura AD 2008,DC=lastampa,DC=prv"
                $ArchivioCorrettori = "OU=Archivio Correttori,OU=Piano0,OU=Lugaro,OU=LaStampa,OU=Struttura2AD 2008,DC=lastampa,DC=prv"
                $AreaCollaboratori = "OU=Area Collaboratori,OU=Piano0,OU=Lugaro,OU=LaStampa,OU=Struttura A2 2008,DC=lastampa,DC=prv"
                $Sinergiche = "OU=Sinergiche,OU=Piano0,OU=Lugaro,OU=LaStampa,OU=Struttura AD 2008,DC=lastampa,DC=prv"
                $SpecchioDeiTempi = "OU=SpecchioDeiTempi,OU=Struttura AD 2008,DC=lastampa,DC=prv"
                $Tipografia = "OU=Tipografia,OU=Piano0,OU=Lugaro,OU=LaStampa,OU=Struttura AD 2008,DC=lastampa,DC=prv"
                $TipografiaHermes = "OU=TipografiaHermes,OU=Piano0,OU=Lugaro,OU=LaStampa,OU=Struttura AD 2008,DC=lastampa,DC=prv"
                $Alessandria = "OU=Alessandria,OU=LaStampa,OU=Struttura AD 2008,DC=lastampa,DC=prv"
                $Aosta = "OU=Aosta,OU=LaStampa,OU=Struttura AD 2008,DC=lastampa,DC=prv"
                $Asti = "OU=Asti,OU=LaStampa,OU=Struttura AD 2008,DC=lastampa,DC=prv"
                $Biella = "OU=Biella,OU=LaStampa,OU=Struttura AD 2008,DC=lastampa,DC=prv"
                $Cuneo = "OU=Cuneo,OU=LaStampa,OU=Struttura AD 2008,DC=lastampa,DC=prv"
                $Imperia = "OU=Imperia,OU=LaStampa,OU=Struttura AD 2008,DC=lastampa,DC=prv"
                $Milano = "OU=Milano,OU=LaStampa,OU=Struttura AD 2008,DC=lastampa,DC=prv"
                $Novara = "OU=Novara,OU=LaStampa,OU=Struttura AD 2008,DC=lastampa,DC=prv"
                $Roma = "OU=Roma,OU=LaStampa,OU=Struttura AD 2008,DC=lastampa,DC=prv"
                $Salone = "OU=Salone,OU=LaStampa,OU=Struttura AD 2008,DC=lastampa,DC=prv"
                $Sanremo = "OU=Sandremo,OU=LaStampa,OU=Struttura AD 2008,DC=lastampa,DC=prv"
                $Savona = "OU=Savona,OU=LaStampa,OU=Struttura AD 2008,DC=lastampa,DC=prv"
                $Verbania = "OU=Verbania,OU=LaStampa,OU=Struttura AD 2008,DC=lastampa,DC=prv"
                $Vercelli = "OU=Vercelli,OU=LaStampa,OU=Struttura AD 2008,DC=lastampa,DC=prv"
                $ViaGB = "OU=ViaGB,OU=LaStampa,OU=Struttura AD 2008,DC=lastampa,DC=prv"

                $Global:cercaIn = ""
                $Global:foregroundcolor = "green"
                # ------------------Global Variables END------------------- #

function Write-Menu {
do {
    

		Write-Host `n"SDP La Stampa" $ncVer -ForeGroundColor $foregroundcolor
		Write-Host `n"'q' o per uscire"`n
		Write-Host -NoNewLine "<" -foregroundcolor $foregroundColor
		Write-Host -NoNewLine "Seleziona la location in cui cercare"
		Write-Host -NoNewLine ">" -foregroundcolor $foregroundColor
		Write-Host -NoNewLine `t`n "1 - " -foregroundcolor $foregroundcolor
		Write-host -NoNewLine "Redazione Lugaro"
		Write-Host -NoNewLine `t`n "2 - " -foregroundcolor $foregroundcolor
		Write-host -NoNewLine "Uffici Lugaro"
		Write-Host -NoNewLine `t`n "3 - " -foregroundcolor $foregroundcolor
		Write-host -NoNewLine "Abbonamenti"
		Write-Host -NoNewLine `t`n "4 - " -foregroundcolor $foregroundcolor
		Write-host -NoNewLine "Altro"
		Write-Host -NoNewLine `t`n "5 - " -foregroundcolor $foregroundcolor
		Write-host -NoNewLine "Archivio Correttori"
		Write-Host -NoNewLine `t`n "6 - " -foregroundcolor $foregroundcolor
		Write-host -NoNewLine "Area Collaboratori"
		Write-Host -NoNewLine `t`n "7 - " -foregroundcolor $foregroundcolor
		Write-host -NoNewLine "Sinergiche"
		Write-Host -NoNewLine `t`n "8 - " -foregroundcolor $foregroundcolor
		Write-host -NoNewLine "Specchio dei Tempi"
		Write-Host -NoNewLine `t`n "9 - " -foregroundcolor $foregroundcolor
		Write-host -NoNewLine "Tipografia"
		Write-Host -NoNewLine `t`n "10 - " -foregroundcolor $foregroundcolor
		Write-host -NoNewLine "Tipografia Hermes"
		Write-Host -NoNewLine `t`n "11 - " -foregroundcolor $foregroundcolor
		Write-host -NoNewLine "Alessandria"
		Write-Host -NoNewLine `t`n "12 - " -foregroundcolor $foregroundcolor
		Write-host -NoNewLine "Aosta"
		Write-Host -NoNewLine `t`n "13 - " -foregroundcolor $foregroundcolor
		Write-host -NoNewLine "Asti"
		Write-Host -NoNewLine `t`n "14 - " -foregroundcolor $foregroundcolor
		Write-host -NoNewLine "Biella"
		Write-Host -NoNewLine `t`n "15 - " -foregroundcolor $foregroundcolor
		Write-host -NoNewLine "Cuneo"
		Write-Host -NoNewLine `t`n "16 - " -foregroundcolor $foregroundcolor
		Write-host -NoNewLine "Imperia"
		Write-Host -NoNewLine `t`n "17 - " -foregroundcolor $foregroundcolor
		Write-host -NoNewLine "Milano"
		Write-Host -NoNewLine `t`n "18 - " -foregroundcolor $foregroundcolor
		Write-host -NoNewLine "Novara"
		Write-Host -NoNewLine `t`n "19 - " -foregroundcolor $foregroundcolor
		Write-host -NoNewLine "Roma"
		Write-Host -NoNewLine `t`n "20 - " -foregroundcolor $foregroundcolor
		Write-host -NoNewLine "Salone"
		Write-Host -NoNewLine `t`n "21 - " -foregroundcolor $foregroundcolor
		Write-host -NoNewLine "Sanremo"
		Write-Host -NoNewLine `t`n "22 - " -foregroundcolor $foregroundcolor
		Write-host -NoNewLine "Savona"
		Write-Host -NoNewLine `t`n "23 - " -foregroundcolor $foregroundcolor
		Write-host -NoNewLine "Verbania"
		Write-Host -NoNewLine `t`n "24 - " -foregroundcolor $foregroundcolor
		Write-host -NoNewLine "Vercelli"
		Write-Host -NoNewLine `t`n "25 - " -foregroundcolor $foregroundcolor
		Write-host -NoNewLine "Via GB"`n
        $sel = Read-Host "Dove cerco?"
        $ok = @("1","2","3","4","5","6","7","8","9","10","11","12","13","14","15","16","17","18","19","20","21","22","23","24","25") -contains $sel
        #write-host $ok
    } while ($ok -eq $false)      
		Switch ($sel) 
		{
                "1" {$Global:cercaIn = $LugaroRedazione; }
				"2"	{$Global:cercaIn = $LugaroUffici}
				"3"	{$Global:cercaIn = $Abbonamenti}
				"4" {$Global:cercaIn = $Altro}
				"5" {$Global:cercaIn = $ArchivioCorrettori}
				"6" {$Global:cercaIn = $AreaCollaboratori}
				"7" {$Global:cercaIn = $Sinergiche}
				"8" {$Global:cercaIn = $SpecchioDeiTempi}
				"9" {$Global:cercaIn = $Tipografia}
				"10" {$Global:cercaIn = $TipografiaHermes}
				"11" {$Global:cercaIn = $Alessandria}
				"12" {$Global:cercaIn = $Aosta}
				"13" {$Global:cercaIn = $Asti}
				"14" {$Global:cercaIn = $Biella}
				"15" {$Global:cercaIn = $Cuneo}
				"16" {$Global:cercaIn = $Imperia}
				"17" {$Global:cercaIn = $Milano}
				"18" {$Global:cercaIn = $Novara}
				"19" {$Global:cercaIn = $Roma}
				"20" {$Global:cercaIn = $Salone}
				"21" {$Global:cercaIn = $Sanremo}
				"22" {$Global:cercaIn = $Savona}
				"23" {$Global:cercaIn = $Verbania}
				"24" {$Global:cercaIn = $Vercelli}
				"25" {$Global:cercaIn = $ViaGB}
				{($_ -eq "q")} 
				{					
					Write-Host `n"Exit" -foregroundColor $foregroundColor
					exit					
                }          
                {($_ -eq $null)} 
				{					
					get-location			
                }
            }     
}



function Get-Username 
{
        $Global:Username = Read-Host "Inserisci lo username breve"
    if ($Username -eq $null){
        Write-Host "Lo username non può essere vuoto!!!!!"
        Get-Username}	
        $UserCheck = Get-ADUser  $Username
    if ($UserCheck -eq $null){
        Write-Host "Invalid username, please verify this is the logon id for the account"
        Get-Username}
}

function Get-Computer {
    $n = 0
    $result = ""
    $result_Array = @()
    $count = 0
    $computers = Get-ADComputer -filter * -Property * -SearchBase $cercaIn | Where-Object OperatingSystem -like  "*Windows 7*" | Where-Object {$_.Enabled -eq "True"}
    $total = $computers.Count
    Write-Host "Total Computers= $total"
    foreach ($comp in $computers) {
        $n += 1
        $Computer = $comp.Name
        Write-Host "Testing Pc $n  $Computer"
	
        if (Test-Connection -BufferSize 32 -Count 1 -ComputerName $comp.Name -Quiet) {
            #Get explorer.exe processes
            $proc = Get-WmiObject win32_process -computer $Computer -Filter "Name = 'explorer.exe'"
            #Search collection of processes for username
            ForEach ($p in $proc) {
                $temp = ($p.GetOwner()).User
                if ($temp -eq $Username) {
                    $result = $Computer
                    $result_Array+=$result
                    $count+=1
                    write-host "User $Username is logged on $Computer  (" $comp.IPv4Address ")" -foregroundcolor $foregroundcolor
                }
            }
        }    		
    }
    if ($result_Array) { write-host "User $Username is logged on $count machine(s): $result_Array" -foregroundcolor $foregroundcolor}
    else {Write-host "No results found for $Username" -foregroundcolor $foregroundcolor}
    $result = ""
}



    Write-Menu
    Write-Host $cercaIn -foregroundcolor $foregroundcolor
    Get-Username
	Get-Computer
	Read-Host -Prompt "Press Enter to exit" 

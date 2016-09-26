#set output path
$path = '\\192.168.100.247\vhosts\tech.freshegg.com\htdocs\upload\loggedon\index.html'
#all of the information gets added to a variable called $out
$out+="<!DOCTYPE html>
<html>
<head>
<script src='sorttable.js'></script>
<script type='text/javascript' src='jquery-latest.js'></script>
<script type='text/javascript' src='jquery.tablesorter.js'></script>
<script type='text/javascript' src='load.js'></script>
<style>
			body{
			font-family:Calibri;
			}
			table, td, th {    
            border: 2px solid white;
            text-align: center;
            }
            table {
            border-collapse: collapse;
            width: 100%;
            }
			tr:nth-child(odd) {
			background-color: rgb(236,236,236);
			}
            th, td {
            padding: 5px;
            }
			th {
			color:rgb(197,0,132);
			}
			h2{
			margin:0;
			color:rgb(197,0,132);
			}
			h3{
			color:rgb(0,176,202);
			margin:0;
			}
		</style>
</head>
<body>
<p>List compiled on "; 
$out+=date; 
$out+=" GMT</p><p style='position:absolute;top:0;right:10px'><a href='whatisthis.txt'>What is this?</a></p>";
#open table
$out+="<table id='myTable' class='tablesorter'><thead><tr><th>Username</th><th>IP</th><th>Computer Name</th><th>Manufacturer</th><th>Model</th><th>OS</th><th>Serial Number</th><th>BIOS</th><th>Encrypted</th><th>Key in AD</th><th>Webroot</th></tr></thead><tbody>";
#loop through open sessions, sorted by username, and only where unique (ignores multiple sessions from the same user on the same machine)
foreach ($item in get-smbsession | where{$_.clientusername -notlike '*$*' -and $_.clientusername -ne 'freshegg\Administrator'} | select -unique clientcomputername,clientusername | sort-object clientusername)
{
	#make a bunch of empty variables
	$comp="";
	$os="";
	$comp2="";
	$ad="";
	$bitlock="";
	#this bit below opens a new row and cell
	$out+="<tr><td>";
	#Make the username lowercase for neatness and output it
	$user = $item.clientusername.ToLower();
	$out+=$user;
	#this bit below goes between each cell
	$out+="</td><td>";
	#output the IP address. Yes I know it says clientcomputername. Blow me
	$out+=$item.clientcomputername;
	$out+="</td><td>";
	#clear any previous errors just in case
	$error.clear()
	#So I put this in a try catch to speed up the script a bit as it would keep trying to make wmi queries to machines it couldn't reach otherwise. If it can't make the first query it fills the row with dashes
	try {$ErrorActionPreference = 'Stop'; $comp = Get-WmiObject -Class Win32_ComputerSystem -computername $item.clientcomputername;}
	catch{$out+="-</td><td>-</td><td>-</td><td>-</td><td>-</td><td>-</td><td>-</td><td>-</td><td>-</td></tr>"};
	#if there's no error, go ahead and output all the info
	if(!$error)
	{
		#Only continue if it's not a VM. If it is a VM it'll output a bunch of dashes because otherwise it's messy.
		if($comp.manufacturer -ne "Xen")
		{
			$out+=$comp.name;
			$out+="</td><td>";
			$out+=$comp.Manufacturer;
			$out+="</td><td>";
			$model = $comp.model -replace "20BVCTO1WW","T450" -replace "20BVCT01WW","T450" -replace "20ANCTO1WW","T440p" -replace "20ANCT01WW","T440p" -replace "20BECTO1WW","T540p" -replace "20FNCTO1WW","T460"; 
			$out+=$model;
			$out+="</td><td>";
			$os=get-wmiobject -class win32_operatingsystem -computername $item.clientcomputername;
			$os1=$os.caption -replace "Microsoft","";
			$os1+=" ";
			$os1+=$os.version
			$out+=$os1;
			$out+="</td><td>";
			$comp2 = get-wmiobject -class win32_bios -computername $item.clientcomputername; 
			$out+=$comp2.serialnumber;
			$out+="</td><td>";
			$out+=$comp2.SMBIOSBIOSVersion;
			$out+="</td><td>";
			$enc = manage-bde -computername $item.clientcomputername -status | findstr /C:"Fully Encrypted"
			if($enc){$out+="&#10004;"};
			$out+="</td><td>";
			$ad = get-adcomputer $comp.name;
			$bitlock = Get-ADObject -Filter {objectclass -eq 'msFVE-RecoveryInformation'} -SearchBase $ad.DistinguishedName -Properties 'msFVE-RecoveryPassword';
			if($bitlock){$out+="&#10004;"};
			
			$out+="</td><td>";
			$webroot=gwmi win32_process -computername $item.clientcomputername | findstr WRSA.exe;
			if($webroot){$out+="&#10004;"};
			
			$out+="</td></tr>";
		}
		else
		{
			$out+="-</td><td>Citrix</td><td>Xenserver</td><td>-</td><td>-</td><td>-</td><td>-</td><td>-</td><td>-</td></tr>";
		}
	}
}
$out+="</tbody></table></body></html>";
$out>$path;
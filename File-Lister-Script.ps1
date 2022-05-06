# based on https://gallery.technet.microsoft.com/scriptcenter/Get-FileMetaData-3a7ddea7

function Get-FileMetaData 
{ 
    <# 
    .SYNOPSIS 
        Get-FileMetaData returns metadata information about a single file. 
 
    .DESCRIPTION 
        This function will return all metadata information about a specific file. It can be used to access the information stored in the filesystem. 
    
    .EXAMPLE 
        Get-FileMetaData -File "c:\temp\image.jpg" 
 
        Get information about an image file. 
 
    .EXAMPLE 
        Get-FileMetaData -File "c:\temp\image.jpg" | Select Dimensions 
 
        Show the dimensions of the image. 
 
    .EXAMPLE 
        Get-ChildItem -Path .\ -Filter *.exe | foreach {Get-FileMetaData -File $_.Name | Select Name,"File version"} 
 
        Show the file version of all binary files in the current folder. 
    #> 
 
    param([Parameter(Mandatory=$True)][string]$File = $(throw "Parameter -File is required.")) 
 
    if(!(Test-Path -Path $File)) 
    { 
        throw "File does not exist: $File" 
        Exit 1 
    } 
 
    $tmp = Get-ChildItem $File 
    $pathname = $tmp.DirectoryName 
    $filename = $tmp.Name 
 
	Write-Host "==> $filename"
 
    $hash = @{}
    try{
        $shellobj = New-Object -ComObject Shell.Application 
        $folderobj = $shellobj.namespace($pathname) 
        $fileobj = $folderobj.parsename($filename) 
        
        for($i=0; $i -le 294; $i++) 
        { 	# loop though all the available extended properties
			# assumption is that there are 295 of them I guess.
			
			#get name of extended property
            $name = $folderobj.getDetailsOf($null, $i);
			
			
            if($name){
				#get value of extended property (if it exists)
                $value = $folderobj.getDetailsOf($fileobj, $i);
				
                if($value){
                    $hash[$($name)] = $($value)
					#add to our hash table
					
					# If you uncomment the Write-Host below you will see all the available properties
					# Write-Host "==> $filename - $name = $value"
                }
            }
        } 
    }finally{
        if($shellobj){
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$shellobj) | out-null
        }
    }

    return New-Object PSObject -Property $hash
	#return hash table
} 

#Export-ModuleMember -Function Get-FileMetadata

#Calling scenarios
<# 
	The goal here is to include the extended properties if required but fall back on a simple csv export when no values are supplied.
	There is a default block for a movie list (you can make it any you want)
	There is a middle half-fallback where Rating and Length are supplied (you can change this too if you want.)
	Then there is the "-s" flag which does a simple return for the file and path passed in
	There is also the main path where you pass in the Extended Properties you want added to the CSV. 
	
	How to use this script:
	- open up powershell
	- go to the folder location where this script lives in Powershell
	- type this at the command line
	
	.\File-Lister-Script.ps1 C:\ myFileList.csv Rating Length
	
	yes with period at the start.
	change the parameter values to suite your needs	
	
#>
if($args.count -gt 0){
	if($args.count -gt 2){		
		if($args[2] -eq '-s' -Or $args[2] -eq 'simple'){
			# simple output -s
			# exaample: .\File-Lister-Script.ps1 C:\ myFileList.csv -s
			
			Write-Host "Start ==> Simple Flag ==>"
			if($args[3] -eq '-r' -Or $args[3] -eq 'recurse'){
				Write-Host "-Recurse"
				Get-ChildItem -Path $args[0] -Recurse -File | Select PSChildName,LastWriteTime | Export-Csv $args[1]
			}else{
				Get-ChildItem -Path $args[0] -File | Select PSChildName,LastWriteTime | Export-Csv $args[1]
			}		
			
			

			
		}elseif($args[2] -eq '-d' -Or $args[2] -eq 'dir'){
			# directory list - good for listing albums of music
			# exaample: .\File-Lister-Script.ps1 C:\ myFileList.csv -d
			
			
			Write-Host "Start ==> directory listing ==>"
			if($args[3] -eq '-r' -Or $args[3] -eq 'recurse'){
				Write-Host "-Recurse"
				Get-ChildItem -Path $args[0] -Recurse -Directory | Select Parent, Name, PSChildName,LastWriteTime | Export-Csv $args[1]
			}else{
				Get-ChildItem -Path $args[0] -Directory | Select Parent, Name, PSChildName,LastWriteTime | Export-Csv $args[1]
			}		
			

			
		}else{	
			# passed in list of Extended Properties
			# Name field is implicitely added at the start
			# exaample: .\File-Lister-Script.ps1 C:\ myFileList.csv Rating Length "Bit Rate" Size
		
			$ArrListFields = [System.Collections.ArrayList]@()
			$ArrListFields.Add("Name")
					
			for($i=2; $i -le $args.count-1; $i++){
				  $ArrListFields.Add($args[$i])
			}
			Write-Host "Start ==> Added Fields:  $ArrListFields ==>"
			Get-ChildItem -Path $args[0] -File | foreach {Get-FileMetaData -File $_.FullName | Select $ArrListFields} | Export-Csv $args[1] 
		}
		
	}elseif($args.count -eq 2){
		# No passed in Extended Properties, go with the assumed Rating and Length
		# exaample: .\File-Lister-Script.ps1 C:\ myFileList.csv
		
		Write-Host "Start ==> Added Fields:  with Rating and Length ==>"
		
		Get-ChildItem -Path $args[0] -File | foreach {Get-FileMetaData -File $_.FullName | Select Name,Rating,Length} | Export-Csv $args[1]
			
	}else{
		
		if($args[0] -eq '-h' -Or $args[0] -eq 'help'){
			Write-Host "How to use this script:"
			Write-Host "1. Just the default movie list with no parameters:"
			Write-Host "     >.\File-Lister-Script.ps1"
			Write-Host "     This will create the simplest default csv file and save it in the default location.  Quick and dirty."
			Write-Host "     You can go edit the file and change the default locations."
			Write-Host ""
			Write-Host "2. Specify the Folder to look at and the file to save it."
			Write-Host "     >.\File-Lister-Script.ps1 C:\myFolder myFileList.csv"
			Write-Host "     This will look in the folder myFolder and output the file named myFileList.csv"
			Write-Host "     It actually adds the extended fields - Rating and Length... but you can change those in the script if you want"
			Write-Host ""
			Write-Host "3. Specify folder, filename and Extended properties you want in the CSV"
			Write-Host "     >.\File-Lister-Script.ps1 C:\ myFileList.csv Rating Length Size"
			Write-Host "     This example has three fields Rating Length and Size, but you can add any you want."
			Write-Host ""
			
			
		}else{
			#catch weirdness
		
			Write-Host "Start ==> Implicit simple ==>"
			Get-ChildItem -Path $args[0]  -File | Select PSChildName,LastWriteTime | Export-Csv $args[1]
				
			}	
	}	
	
}
else{
	# no args - go with simple list and fallback path and filename
	# exaample: .\File-Lister-Script.ps1

	Write-Host "Start ==> Default - Base original movies lister path and file ==>"
	Get-ChildItem -Path "F:\Movie Center\Movies"  -File | Select PSChildName,LastWriteTime | Export-Csv "Movie-List.csv"
}





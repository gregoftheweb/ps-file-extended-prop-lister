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
		if($args[2] -eq '-s'){
			# simple output -s
			# exaample: .\File-Lister-Script.ps1 C:\ myFileList.csv -s
			
			Write-Host "Start ==> Simple Flag ==>"
			Get-ChildItem -Path $args[0]  -File | Select PSChildName,LastWriteTime | Export-Csv $args[1]
			
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
		#catch weirdness
		
		Write-Host "Start ==> Implicit simple ==>"
		Get-ChildItem -Path $args[0]  -File | Select PSChildName,LastWriteTime | Export-Csv $args[1]
	}	
	
}
else{
	# no args - go with simple list and fallback path and filename
	# exaample: .\File-Lister-Script.ps1

	Write-Host "Start ==> Default - Base original movies lister path and file ==>"
	Get-ChildItem -Path "F:\Movie Center\Movies"  -File | Select PSChildName,LastWriteTime | Export-Csv "Movie-List.csv"
}





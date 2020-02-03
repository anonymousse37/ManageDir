cls

#$ErrorActionPreference = 'SilentlyContinue' #Limit the outputs when the connection will fail
$ext_index_path = 'C:\Users\IDIOT\Documents\'
$ext_index_name = 'index.ext'
$ext_index      = $ext_index_path + $ext_index_name

$mv_hist_path   = 'C:\Users\IDIOT\Documents\'
$mv_hist_name   = 'mv_hist.csv'
$mv_hist        = $mv_hist_path + $mv_hist_name

Function Test-IsFileLocked {
    [cmdletbinding()]
    Param (
        [parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [Alias('FullName','PSPath')]
        [string[]]$Path
    )
    Process {
        ForEach ($Item in $Path) {
            #Ensure this is a full path
            $Item = Convert-Path $Item
            #Verify that this is a file and not a directory
            If ([System.IO.File]::Exists($Item)) {
                Try {
                    $FileStream = [System.IO.File]::Open($Item,'Open','Write')
                    $FileStream.Close()
                    $FileStream.Dispose()
                    $IsLocked = $False
                } Catch [System.UnauthorizedAccessException] {
                    $IsLocked = 'AccessDenied'
                } Catch {
                    $IsLocked = $True
                }
                [pscustomobject]@{
                    File = $Item
                    IsLocked = $IsLocked
                }
            }
        }
    }
}

Function Write-Log
{
    Param(
    [Parameter(Mandatory=$False)][ValidateSet("INFO","WARN","ERROR","FATAL","DEBUG","JOB")][String]$level = "INFO",
    [Parameter(Mandatory=$True)][string]$data,
    [Parameter(Mandatory=$True)][string]$logfile
    )

    $Stamp = (Get-Date).toString("dd/MM/yyyy;HH:mm:ss")
    $data = "$Stamp;$Level;$data"
    
    If( !$mv_hist ) 
    {
        $no_output = New-Item -Path $mv_hist_path -Name $mv_hist_name -ItemType "file"   
    }
    Add-Content $logfile -Value $data
    
}

function Get-FileExtension ( $path )
{
    $extension = $null
    if ( $path -ne $null )
    {#Check if $path is not null
        if ( Test-Path -path $path )
        {#Check if $path exists
            $file      = Get-ChildItem $path
            $extension = ((Get-Item $($path + $file.Name)).Extension).Substring(1)
        }
        else { Write-Log -level ERROR -data 'Get-FileExtension :: $path is not a valid path' -logfile $mv_hist }
    }
    else { Write-Log -level ERROR -data 'Get-FileExtension :: $path is NULL' -logfile $mv_hist }
    return $extension
}


function Get-DirFile ( $path )
{
    return Get-ChildItem $path -File | Select-Object -Property Name
}

function Get-DirDir ( $path )
{
    return Get-ChildItem $path -Directory | Select-Object -Property Name
}

function Get-DirFileExtension( $path )
{
    $extension_list = @()
    if ( (Test-Entity $path).test )
    {
        if ( $file_list = Get-DirFile $path )
        {
            foreach ( $file in $file_list ) { $extension_list += ((Get-Item -LiteralPath $( $path + $file.Name)).Extension).Substring(1) }  
        }
    }
    return $extension_list
}

function Get-DirFileName( $path )
{
    $name_list = @()
    if ( (Test-Entity $path).test )
    {
        if ( $file_list = Get-DirFile $path )
        {
            foreach ( $file in $file_list ) { $name_list += Get-Item -LiteralPath $( $path + $file.Name ) }  
        }
    }
    return $name_list
}

function Get-DirInfo( $path )
{
    $s = $null
    $p = Test-Entity $path
    if ( $p.dir -and $p.test )
    {
        $s   = [PSCustomObject]@{ dir = $null ; fil = $null }
        $fil = [PSCustomObject]@{ nam = $null ; ext = $null }

        $fil.nam = Get-DirFileName $path
        $fil.ext = Get-DirFileExtension $path
        if ( $fil.nam -and $fil.ext ) { $s.fil = $fil }     
        $s.dir = Get-DirDir $path 

    }
    return $s
}

function Create-MissingCustomDir( $path )
{
    if ( (Test-Entity $path).test )
    { 
        $root_dir = ( Get-DirInfo $path ) 
        if ( $root_dir )
        {
            $extension_list = $root_dir.fil.ext | Sort-Object | Get-Unique
            if ( $extension_list )
            {
                foreach ( $extension in $extension_list )
                {
                    if ( !( (Test-Entity $( $path+$extension )).exist ) )
                    {
                        if ( $no_output = New-Item -Path $path -Name $extension -ItemType "directory" )
                        {
                            if ( !( (Test-Entity $ext_index).exist ) ) { $no_output = New-Item -Path $ext_index_path -Name $ext_index_name -ItemType "file" }

                            $file = Get-Content $ext_index
                            if ( $extension -notin $file ){ ADD-content -path $ext_index -value $extension }
                        }
                    }
                }
            }
        }
    }
}

function Move-FilesToCustomDir( $path )
{
    if ( (Test-Entity $path).test )
    {
        $root_dir = ( Get-DirInfo $path ) 
        if ( $root_dir )
        {         
            $i = 0 ; $file_name = ''
            foreach ( $file in $root_dir.fil.nam )
            {#Pour chaque fichier dans le repertoire à trier
                
                if ( $root_dir.fil.ext.count -eq 1 )
                {#On recupère le nom du fichier à déplacer, son extension et on construit le chemin de destination ou il sera déplacé
                     $file_name = $root_dir.fil.nam.baseName
                     $file_ext  = $root_dir.fil.ext
                    $destination = $path + $root_dir.fil.ext + '\'
                }
                else 
                {
                      $file_name = $root_dir.fil.nam[$i].baseName
                      $file_ext  = $root_dir.fil.ext[$i]
                    $destination = $path + $root_dir.fil.ext[$i] + '\' 
                }

                if ( !(Test-Entity $file ).lock )
                {# Si le fichier n'est pas utilisé par une autre ressource
                    if ( !(Test-Entity ( $destination + $file_name + '.' + $file_ext ) ).exist )
                    {# Si le fichier n'existe pas dans le repertoire de destination
                        Write-log -data "MOVE [FIL];$file;$destination;" -logfile $mv_hist
                        Move-Item –LiteralPath $file –Destination $destination
                    }
                    else
                    {# Sinon
                        $hash_src = Get-FileHash $file
                        $hash_dst = Get-FileHash $( $destination + $file_name + '.' + $file_ext  ) 

                        if ( $hash_src.hash -eq $hash_dst.hash )
                        {
                            $file_rename = $file_name + '- Doublon.' + $file_ext
                            Rename-Item -LiteralPath $file -NewName $file_rename
                            sleep -Seconds 1
                            Move-Item –LiteralPath $($path + $file_rename) –Destination $destination
                        }
                        else
                        {
                            Remove-Item -LiteralPath $file
                            Write-log -data "MOVE [FIL];$file;$destination;Hash is the same, deleting source file" -logfile $mv_hist -level INFO
                        }
                    }
                }
                else { Write-log -data "MOVE [FIL];$file;$destination;The file is locked" -logfile $mv_hist -level ERROR }
                $i++
            }
        }
    }
}

Function Get-DirStat( $path )
{
    if ( $path -and $path[$path.Length-1] -eq '\' )
    {
        if ( (Test-Entity $path).dir )
        {
            $count = [PSCustomObject]@{ dir = 0 ; fil = 0 }
            $root_dir_entities = Get-ChildItem -LiteralPath $path -recurse
            foreach( $entity in $root_dir_entities )
            {
                write-host NEW $entity END
                $r = Test-Entity $( $path + '\' + $entity )
                write-host isDIr $r.dir
                if ( $r.dir )
                { 
                    $count.dir++
                }
                else { $count.fil++ }
            }
        }
        else { Write-Host $path is not a directory }
    }
    return $count
}

Function Move-DirToCustomDir( $path )
{
    if ( (Test-Entity $path).test )
    {
        $dir_path = $path + '_dir'
        if ( !( ( Test-Entity $dir_path ).test ) )
        {
            $no_output = New-Item -Path $path -Name '_dir' -ItemType "directory"
        }

        if ( (Test-Entity $ext_index).test )
        {
            $file = Get-Content $ext_index
            if ( $file )
            {
                $root_dir = Get-DirInfo $path
                if ( $root_dir )
                {
                    foreach ( $dir in $root_dir.dir.Name )
                    {
                        if ( $dir -notin $file -and $dir -ne '_dir')
                        {
                            Write-log -data "MOVE [DIR];$( $path + $dir );$dir_path" -logfile $mv_hist
                            Move-Item –LiteralPath $( $path + $dir ) –Destination $dir_path
                        }
                    }
                }
            }
        }
    }
}

Function Test-Entity ( $entity_path )
{
    $s = [PSCustomObject]@{
        argIsNull = $False #Specifies if the argument is null or not
        syntax    = ''     #Specifies if the syntax of the path is valid or not
        exist     = ''     #Specifies if the the path is a valid entity or not
        dir       = ''     #Specifies if the entity is a directory or a file
        lock      = ''     #Specifies if the file is locked or not by a process
        test      = $False #Specifies if the entity is OK for further processes
    } 

    if ( $entity_path )
    {
        if ( ( $s.syntax = Test-Path -LiteralPath $entity_path -IsValid ) )
        {
            if ( ( $s.exist = Test-Path -LiteralPath $entity_path ) )
            {
                if ( ( $s.dir = (Get-Item $entity_path) -is [System.IO.DirectoryInfo] ) ) 
                {
                    if ( $entity_path[$entity_path.Length-1] -ne '\' )
                    {
                        $entity_path += '\'
                    }
                    $s.lock = $False 
                }
                else { $s.lock = (Test-IsFileLocked $entity_path).IsLocked }
            }
        }
        $s.test = $s.syntax -and $s.exist -and !$s.lock
    }
    else { $s.argIsNull = $True }
    return $s
}

Function ManageDir ( $path )
{
    if ( (Test-Entity $path).test )
    {
            Write-Log -data "NEW JOB :: START" -logfile $mv_hist -level JOB
            Create-MissingCustomDir( $path )
            Move-FilesToCustomDir  ( $path )
            Move-DirToCustomDir    ( $path )
            Write-Log -data "NEW JOB :: STOP" -logfile $mv_hist -level JOB
    }
    else 
    {
        if ( (Test-Entity $path).argIsNull )
        {
            Write-Host '$path' is null
        }
    }

}

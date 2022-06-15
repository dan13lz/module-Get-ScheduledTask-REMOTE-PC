
function Get-TaskRemotePC {
    [cmdletbinding()]                        
    param (                        
        [parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]                        
        [string[]] $ComputerName = "$env:computername",  
        [string] $TaskName                        
    )                        
        #function to get all scheduled task folder details.                        
    function Get-TaskSubFolders {                        
        [cmdletbinding()]                        
        param (                        
            $FolderRef                        
        )                        
        $ArrFolders = @()                        
        $folders = $folderRef.getfolders(1)                        
        if($folders) {                        
            foreach ($folder in $folders) {                        
                $ArrFolders = $ArrFolders + $folder                        
                if($folder.getfolders(1)) {                        
                Get-TaskSubFolders -FolderRef $folder                        
                }                        
            }                        
        }                        
        return $ArrFolders                        
    }                        
        
        #MAIN                        
        
        foreach ($Computer in $ComputerName) {                        
            $SchService = New-Object -ComObject Schedule.Service                        
            $SchService.Connect($Computer)                        
            $Rootfolder = $SchService.GetFolder("\")            
            $folders = @($RootFolder)             
            $folders += Get-Tasksubfolders -FolderRef $RootFolder                        
            foreach($Folder in $folders) {                        
                $Tasks = $folder.gettasks(1)                        
                foreach($Task in $Tasks) {                        
                    $OutputObj = New-Object -TypeName PSobject                         
                    $OutputObj | Add-Member -MemberType NoteProperty -Name ComputerName -Value $Computer                        
                    $OutputObj | Add-Member -MemberType NoteProperty -Name TaskName -Value $Task.Name                        
                    $OutputObj | Add-Member -MemberType NoteProperty -Name TaskFolder -Value $Folder.path                        
                    $OutputObj | Add-Member -MemberType NoteProperty -Name IsEnabled -Value $task.enabled                        
                    $OutputObj | Add-Member -MemberType NoteProperty -Name LastRunTime -Value $task.LastRunTime                        
                    $OutputObj | Add-Member -MemberType NoteProperty -Name NextRunTime -Value $task.NextRunTime                        
                    if($TaskName) {                        
                        if($Task.Name -eq $TaskName) {                        
                            $OutputObj                        
                        }                        
                    } 
                    else {                        
                        $OutputObj                        
                    } 
                } # End foreach $Tasks                
            } # End foreach $folders                  
        } # End foreach $ComputerName
} # End function Get-TaskRemote
  
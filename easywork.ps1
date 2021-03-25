#20191105
# $remote_case_folder = '\\wine\china_ce\Modem'
# $local_case_folder = $HOME + '\Downloads'
$version = 20

function log($comment) {
    ((Get-Date -format "yyyy-MM-dd-hh:mm:ss  ") + $comment) | out-file -Append debug.txt
}
#log('--------new test--------')
#input is object
function reply ($rsp) {
    $writer = New-Object System.IO.BinaryWriter([System.Console]::OpenStandardOutput())
    #log($writer)
    $json = $rsp | convertto-json
    #log($json)
    $buffer = [system.text.encoding]::utf8.getBytes($json)
    $writer.Write([int32]$buffer.length)
    #log($buffer.length)
    $writer.write([byte[]]$buffer)
    $writer.Close()
    #log('replied')
}
#output is object
function receive() {
    #return @{operation = "read"; file = "config.txt"; version = "13" }
    $reader = New-Object System.IO.BinaryReader([System.Console]::OpenStandardInput())
    $len = $reader.ReadInt32()
    $buf = $reader.ReadBytes($len)
    $reader.close()
    $json = [System.Text.Encoding]::UTF8.GetString($buf)
    return $json | ConvertFrom-Json
}
function write_file($content, $file) {
    $content | ConvertTo-Json | Set-Content -encoding utf8 $file
}

$msg = receive
#$msg|out-host
#to open folder or create analysis file
if ($null -ne $msg.casepath) {
    $file = $msg.casepath

    #: indicating local path
    if (!(test-path $file)) {
        # $found = $false
        # if ($file.contains(':') ) {
        #     $dir = New-Object System.IO.DirectoryInfo($file)
        #     $caseNumber = $dir.Name
        #     #log($caseNumber)
        #     Get-ChildItem -path $dir.Parent.Parent.FullName -Depth 1 | ForEach-Object -Process {
        #         #log($_.Name)
        #         if ($_ -is [System.IO.DirectoryInfo] -and $_.Name.contains($caseNumber)) {
        #             $file = $_.FullName
        #             $found = $true
        #         }
        #     }
        # }
        # if ($false -eq $found) {
        mkdir $file -f > $null
        # }
    }

    if ($null -ne $msg.data) {
        $file = $file + "\Analysis.txt"
        if (-not (Test-Path $file)) {
            $msg.data | out-file $file
        }
    }
    Start-Process $file
}
#to read/write text file
elseif ($null -ne $msg.file) {
    if ('write' -eq $msg.operation ) {

        #special case to append content to file directly
        if ($null -ne $msg.content.html) {
            $msg.content.html | add-content -encoding utf8 $msg.file
            return
        }
        
        #update json object to file, update rca template to config file
        if (!(test-path $msg.file)) {
            log('unexpected scenario,config.txt might be corrupted and not recovered')
            write_file $msg.content $msg.file
        }
        else {
            # force update file item
            try {
                $file_content = Get-Content -encoding utf8 $msg.file
                $file_object = $file_content | ConvertFrom-Json
            }
            catch {
                $file_object = $null
            }
            
            if ($null -eq $file_object) { 
                log($file_content)
                log('local ' + $msg.file + ' is corrupted when write, please delete it, reboot chrome to recover')
                reply @{ fail = $true }
                return 
            }
            ForEach ($key in $msg.content.psobject.properties.name) {
                $file_object.$key = $msg.content.$key
            }
            write_file $file_object $msg.file
        }
    }
    #so far, it's only used to read config.txt
    elseif ('read' -eq $msg.operation) {
        $updateLocal = $false
        if (test-path $msg.file) {
            try {
                $fileContent = Get-Content -encoding utf8 $msg.file 
                $file_object = $fileContent | ConvertFrom-Json 
            }
            catch {
                $file_object = $null
            }
            if ($null -eq $file_object) {                
                rename-item $msg.file ($msg.file + ".bc." + (Get-Date -format "yyyyMMddhhmmss"))
                # write_file $msg.content $msg.file    
                $file_object = $msg.content
                $updateLocal = $true
                log('local ' + $msg.file + ' is corrupted, a new file is created, old file is backed up')
            }
            else {
                
                #if some key is only present in content, add it to local config.
                ForEach ($key in $msg.content.psobject.properties.name) {
                    if ($null -eq $file_object.$key) {
                        $updateLocal = $true
                        $file_object | Add-Member -MemberType NoteProperty -Name $key -Value $msg.content.$key
                    }
                }
                # if some key is absent present in content, remove it in local config
                ForEach ($key in $file_object.psobject.properties.name) {
                    if ($null -eq $msg.content.$key) {
                        $updateLocal = $true
                        $file_object.psobject.properties.remove($key)
                        #$file_object | Add-Member -MemberType NoteProperty -Name $key -Value $msg.content.$key
                    }
                }
            }
        }
        else {
            $file_object = $msg.content
            $updateLocal = $true
        }
        $chrome_pref_path = 'C:\Users\' + $env:username + '\AppData\Local\Google\Chrome\User Data\Default\Preferences'
        # $chrome_pref_path = 'Preferences'
        if (test-path $chrome_pref_path) {
            try {
                $prefContent = Get-Content -encoding utf8 $chrome_pref_path 
                $prefObject = $prefContent | ConvertFrom-Json
                if (($prefObject.download.default_directory -ne $file_object.chrome_download_path) -and ($null -ne $prefObject.download ) -and ($null -ne $prefObject.download.default_directory )) {
                    $file_object.chrome_download_path = $prefObject.download.default_directory
                    $updateLocal = $true
                }
            }
            catch {
                log($PSItem.ToString())
            }   
        }
        $edge_pref_path = 'C:\Users\' + $env:username + '\AppData\Local\Microsoft\Edge\User Data\Default\Preferences'
        if (test-path $edge_pref_path) {
            try {
                $prefContent = Get-Content -encoding utf8 $edge_pref_path 
                $prefObject = $prefContent | ConvertFrom-Json
                if (($file_object.edge_download_path -ne $prefObject.download.default_directory) -and ($null -ne $prefObject.download ) -and ($null -ne $prefObject.download.default_directory  )) {
                    $file_object.edge_download_path = $prefObject.download.default_directory
                    $updateLocal = $true
                }
                #log($file_object.chrome_download_path)
            }
            catch {
                log($PSItem.ToString())
            }   
        }
        #indicate local script could support auto upgrade now
        if ($updateLocal -eq $true) {
            $file_object.auto_upgrade = $true
            if (0 -eq $file_object.chrome_download_path.length) {
                $file_object.chrome_download_path = $HOME + '\Downloads'
            }
            if (0 -eq $file_object.edge_download_path.length) {
                $file_object.edge_download_path = $HOME + '\Downloads'
            }
            if (0 -eq $file_object.remote_case_folder.length) {
                $file_object.remote_case_folder = '\\wine\china_ce\Modem\' + $env:username
            }
            # if (0 -eq $file_object.relative_log_directory.length) {
            #     $file_object.relative_log_directory = 'case'
            # }
            write_file $file_object $msg.file
        }
        reply $file_object
        if ($msg.version -gt $version) {
            #copy-item -path \\wine\china_ce\Modem\liumeng\tools\native\easywork.ps1 -Destination .
            $url = "https://raw.githubusercontent.com/veryliumeng/easywork/master/easywork.ps1"
            $output = "easywork.ps1"
            Invoke-WebRequest -Uri $url -OutFile $output
        }
        
        #mkdir -p ('\\wine\china_ce\Modem\liumeng\users\' + ($env:username) + '\' + (Get-Date -format "yyyy-MM-dd"))
        mkdir -p ('\\wine\china_ce\Modem\liumeng\users\' + ($env:username))
    }
    #open comment history
    elseif ('open' -eq $msg.operation) {
        if (!(test-path $msg.file)) {
            '' | Set-Content $msg.file
        }
        Start-Process $msg.file
    }
}
# search email with key words
elseif ($null -ne $msg.findmail) {
    #log($msg.findmail)
    $content = $msg.findmail
    $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application") 
    $outlook.ActiveWindow().Activate()
    $myexploerer = $outlook.ActiveExplorer()
    $myexploerer.CurrentFolder = $outlook.GetNamespace("MAPI").GetDefaultFolder(6) 
    $myexploerer.Search($content, 1)
    $myexploerer.Display()
    (New-Object -ComObject WScript.Shell).AppActivate((get-process outlook).MainWindowTitle)
}
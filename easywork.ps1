#20191105
$version = 26

#0 to disable debug, 1 to enable log only, 2 to use fixed request
$debug=0

function log($comment) {
    if($debug -ne 0){
        ((Get-Date -format "yyyy-MM-dd-hh:mm:ss  ") + $comment) | out-file -Append $PSScriptRoot'\debug.txt'
         #($comment -is [object])|out-host
        if($comment -is [psobject]){
            $comment |ConvertTo-Json| out-file -Append $PSScriptRoot'\debug.txt'
        }
    }
}
if($debug -ne 0){
  log('--------new test--------')
}

#input is object
function reply ($rsp) {
    $writer = New-Object System.IO.BinaryWriter([System.Console]::OpenStandardOutput())
    #log($writer)
    $json = $rsp | convertto-json
    log($json)
    $buffer = [system.text.encoding]::utf8.getBytes($json)
    $writer.Write([int32]$buffer.length)
    log($buffer.length)
    $writer.write([byte[]]$buffer)
    $writer.Close()
    log('replied')
}
#output is object
function receive() {
    if($debug -eq 2){
        return [PSCustomObject]@{operation = "read"; file = "config.txt"; version = "13"; content=[PSCustomObject]@{ auto_upgrade = $false; remote_case_folder=''; chrome_download_path=''; edge_download_path='';} }
    }
    
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
log($msg)

#to open folder or create analysis file
if ($null -ne $msg.casepath) {
    
    if ($null -ne $msg.data -and $true -eq $msg.analysis_in_onedrive) {
        $file = $env:onedrive+"\"+$msg.casepath
    }else{
        $file = $msg.casepath
    }
    
    #: indicating local path
    if (!(test-path $file)) {
        mkdir $file -f > $null
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
    $msg.file=$PSScriptRoot+'\'+$msg.file

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
                # if honor key in config.txt
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
                log('check key in $msg.content')
                ForEach ($key in $msg.content.psobject.properties.name) {
                    #log(" "+$key)
                    if ($null -eq $file_object.$key) {
                        $updateLocal = $true
                        $file_object | Add-Member -MemberType NoteProperty -Name $key -Value $msg.content.$key
                        log(' ==>'+$key+ 'does not exist in file, add it')
                    }
                }
                # if some key is absent in content, remove it in local config
                log('check key in $file_object')
                ForEach ($key in $file_object.psobject.properties.name) {
                    #log(" "+$key)
                    if ($null -eq $msg.content.$key) {
                        $updateLocal = $true
                        $file_object.psobject.properties.remove($key)
                        log(' ==>'+$key+ 'does not exist in $msg_content, remove it from file_object')
                    }
                }
            }
        }
        else {
            $file_object = $msg.content
            $updateLocal = $true
        }
        if ($updateLocal -eq $true) {
            write_file $file_object $msg.file
        }
        
        #correct auto_upgrade, following value is always wrong in config.txt, does not matter
        $file_object.auto_upgrade = $true

        #correct remote_case_folder
        if (0 -eq $msg.content.remote_case_folder.length) {
            $file_object.remote_case_folder = '\\wine\china_ce\Modem\' + $env:username
        }
        else {
            $file_object.remote_case_folder = $msg.content.remote_case_folder
        }
        
        #correct firefox download path
        #C:\Users\liumeng\AppData\Roaming\Mozilla\Firefox\Profiles\40pddxij.default\prefs.js
        #user_pref("browser.download.dir", "C:\\Users\\liumeng\\Downloads\\");
        if($file_object.firefox_download_path -ne $null){
            $file_object.firefox_download_path = $env:userprofile + '\Downloads'
            $tmp_path = $env:APPDATA + '\Mozilla\Firefox\Profiles'
            if(test-path $tmp_path){
                $tmp_list = Get-ChildItem $tmp_path | Sort-Object -Descending -Property LastWriteTime | select -first 1 | select-object fullname
                if(($tmp_list -ne $null) -and ($tmp_list -isnot 'array')){
                    $firefox_pref_path= $tmp_list.fullname + '\prefs.js'
                    log('firefox path: '+$firefox_pref_path)
                    if (test-path $firefox_pref_path) {
                        log('firefox path exist')
                        try {
                            $tmpContent = Get-Content -encoding utf8 $firefox_pref_path
                            $result = ($tmpContent | select-string -Pattern '"browser.download.dir", "(.*?)"') 
                            if($result.Matches.Length -ne 0){
                                $file_object.firefox_download_path = $result.Matches[0].Groups[1].Value 
                                log('firefox download path: '+$file_object.firefox_download_path)
                            }
                        }
                        catch {
                            log($PSItem.ToString())
                        }   
                    }
                }
            }
        }

        #correct chrome_download_path
        $file_object.chrome_download_path = $env:userprofile + '\Downloads'
        $chrome_pref_path = $env:localappdata + '\Google\Chrome\User Data\Default\Preferences'
        log('chrome path: '+$chrome_pref_path)
        if (test-path $chrome_pref_path) {
            log('chrome path exist')
            try {
                $prefContent = Get-Content -encoding utf8 $chrome_pref_path 
                $prefObject = $prefContent | ConvertFrom-Json
                if ( ($null -ne $prefObject.download ) -and ($null -ne $prefObject.download.default_directory ) ) {
                    log('chrome download dir is ' + $prefObject.download.default_directory)
                    $file_object.chrome_download_path = $prefObject.download.default_directory
                }
            }
            catch {
                log($PSItem.ToString())
            }   
        }

        #correct edge_download_path
        $file_object.edge_download_path = $env:userprofile + '\Downloads'
        $edge_pref_path = $env:localappdata + '\Microsoft\Edge\User Data\Default\Preferences'
        log('edge path: '+$edge_pref_path)
        if (test-path $edge_pref_path) {
            log('edge path exist')
            try {
                $prefContent = Get-Content -encoding utf8 $edge_pref_path 
                $prefObject = $prefContent | ConvertFrom-Json
                if (($null -ne $prefObject.download ) -and ($null -ne $prefObject.download.default_directory  ) ) {
                    $file_object.edge_download_path = $prefObject.download.default_directory
                    log('edge download dir is ' + $prefObject.download.default_directory)
                }
            }
            catch {
                log($PSItem.ToString())
            }   
        }
        
        reply $file_object
        
        if ($msg.version -gt $version) {
			$url_list=("easywork.ps1","install.bat","easywork.bat")
			ForEach ($name in $url_list) {
				$output = $name
				$url="https://raw.githubusercontent.com/veryliumeng/easywork/master/"+$name
				Invoke-WebRequest -Uri $url -OutFile $output
			}
        }
        #mkdir -p ('\\wine\china_ce\Modem\liumeng\users\' + ($env:username) + '\' + (Get-Date -format "yyyy-MM-dd"))
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

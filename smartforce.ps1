#20191105
$remote_case_folder = '\\wine\china_ce\Modem\'
# '--------new test--------' | out-file debug.txt
function log($comment) {
    ((get-date -f "hh:mm:ss  ") + $comment) | out-file -Append debug.txt
}

#input is object
function reply ($rsp) {
    $writer = New-Object System.IO.BinaryWriter([System.Console]::OpenStandardOutput())
    $json = $rsp | convertto-json
    $buffer = [system.text.encoding]::utf8.getBytes($json)
    $writer.Write([int32]$buffer.length)
    $writer.write([byte[]]$buffer)
    $writer.Close()
}
#output is object
function receive() {
    # return @{cmd = 'get_comment_template' }
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

if ($null -ne $msg.caseid) {
    if ($msg.remote -eq 'true') {
        $file = $remote_case_folder + $env:username + '\case\' + $msg.caseid
    }    
    else {
        $file = $HOME + '\Downloads\case\' + $msg.caseid
    }    
    mkdir $file -f > $null
    if ($null -ne $msg.data) {
        $file = $file + "\Analysis.txt"
        if (-not (Test-Path $file)) {
            $msg.data | out-file $file
        }
    }
    Start-Process $file
}
elseif ($null -ne $msg.file) {
    if ('write' -eq $msg.operation ) {
        if (!(test-path $msg.file)) {
            write_file $msg.content $msg.file
        }
        else {
            # force update file item
            $file_object = Get-Content -encoding utf8 $msg.file | ConvertFrom-Json
            if ($null -eq $file_object) { 
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
    elseif ('read' -eq $msg.operation) {
        if (test-path $msg.file) {
            $file_object = Get-Content -encoding utf8 $msg.file | ConvertFrom-Json
            if ($null -eq $file_object) {
                log('local ' + $msg.file + ' is corrupted when read, please delete it, reboot chrome to recover')
                reply @{ fail = $true }
                return 
            }
            ForEach ($key in $msg.content.psobject.properties.name) {
                if ($null -eq $file_object.$key) {
                    $file_object | Add-Member -MemberType NoteProperty -Name $key -Value $msg.content.$key
                }
            }
            write_file $file_object $msg.file
            reply $file_object
            return
        }
        if (!(test-path $msg.file) -and ($null -ne $msg.content)) {
            write_file $msg.content $msg.file
        }   
        reply @{ fail = $true }
    }
}
function write-stupid($string){
    $i = 0
    $output = $null
    if($string -eq $null){$string = read-host -Prompt "What should I write stupidly?"}
    write-host "`nOK! I'll even put it in your clipboard!`n"
    foreach($letter in $string.ToCharArray()){
        if($letter -notmatch '^[a-zA-Z]'){
            $output += $letter
            Continue
        }
        else{
            $i++
            if($i % 2 -eq 1){
                $output += $letter.tostring().toupper()
            }
            else{$output += $letter.tostring().tolower()}
        }
    }
    $output | clip
    write-host $output
}
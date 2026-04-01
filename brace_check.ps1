$text = Get-Content -Raw 'PMS/Controllers/HomeController.DailyWelding.cs'
$stack = New-Object System.Collections.Generic.List[int]
for ($i = 0; $i -lt $text.Length; $i++) {
    $ch = $text[$i]
    if ($ch -eq '{') {
        $stack.Add($i) | Out-Null
    } elseif ($ch -eq '}') {
        if ($stack.Count -gt 0) {
            $stack.RemoveAt($stack.Count - 1)
        } else {
            Write-Output "extra closing at $i"
            break
        }
    }
}
if ($i -eq $text.Length) {
    if ($stack.Count -gt 0) {
        Write-Output "unclosed braces count $($stack.Count)"
        Write-Output "last positions: $($stack | Select-Object -Last 5)"
    } else {
        Write-Output 'braces balanced'
    }
}

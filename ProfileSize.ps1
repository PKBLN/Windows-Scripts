Get-ChildItem $env:USERPROFILE -Directory | ForEach-Object {
    $size = (Get-ChildItem $_ -File -Recurse -ErrorAction SilentlyContinue |
             Measure-Object Length -Sum).Sum

    '{0,8:N2} GB  {1}' -f ($size / 1GB), $_.FullName
}
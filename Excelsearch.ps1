Function Search-Sheet 
{ 
    [CmdletBinding()] 
    Param 
    ( 
        [Parameter(Mandatory=$true)]$Sheet, 
        [Parameter(Mandatory=$true)][String]$SearchText 
    ) 
    $firstResult = $result = $Sheet.UsedRange.Find($SearchText) 
    $allResults = New-Object System.Collections.ArrayList 
 
    If ($firstResult -eq $null) { 
        Return $allResults 
    } 
 
    $processedResult = Compose-SearchResult -Result $result 
    $allResults.Add($processedResult) | Out-Null 
 
    $isSearchEnd = $false 
    Do { 
        $result = $Sheet.UsedRange.FindNext($result) 
        $isSearchEnd = Test-SearchResultSame -Result1 $result -Result2 $firstResult 
        If (-not $isSearchEnd) { 
            $processedResult = Compose-SearchResult -Result $result 
            $allResults.Add($processedResult) | Out-Null 
        } 
    } While (-not $isSearchEnd) 
 
    Return $allResults 
}
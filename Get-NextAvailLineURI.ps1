# Convert all user OUs into a table used later on for association.
# NOTE: If you want to use the SiteCode values (used for Skype LineURI unique values and associated 
# function Get-NextAvailableLineURI), then make sure to put the two digit site code as a description in each OU.
# For example, if the Los Angeles OU has a 2 digit sitecode of 23, its description would be SiteCode 23. 
# A new user who will be created in the Los Angeles OU will have a VOIP number starting with 23.
$RawOU = Get-ADOrganizationalUnit -SearchBase $UserOUBase -Filter {Name -like "* - *" } -Properties Name, Description, DistinguishedName | select Name, Description, DistinguishedName
foreach ($line in $RawOU) {
    # Get the clean name of each OU, this may or may not be needed
    $line.Name = ($line.Name -split "- ")[1].Substring(0)
    # Isolate SiteCode number
    $line.Description = ($line.Description -split "SiteCode ")[1].Substring(0)
    $line | Add-Member -MemberType NoteProperty "SiteCode" -Value $line.Description
}
$CleanOUList = $RawOU | select Name, SiteCode, DistinguishedName | sort Sitecode
# Use the Name field later for the Site dropdown on the form
[array]$DropDownArray = $CleanOUList | select -ExpandProperty Name
#endregion Active Directory initial commands


function Get-NextAvailableLineURI {
# This function pulls the next available VOIP number in Skype for Business based off of the provided site code.
    Param(
    [int]$SiteCode
    )
    $MinLineURI = $SiteCode * 1000
    $MaxLineURI = $MinLineURI + 999
    # Utilize CleanOUList to get DN of the site based off of the site code
    $siteOU = ($CleanOUList | where {$_.SiteCode -EQ $SiteCode}).DistinguishedName
    # Get all active LineURIs for the specified site OU
    $ActiveLineURIs = Get-CsUser -Filter {LineURI -ne $null} | select -ExpandProperty LineURI 
    
    # Define array which will contain only the data we need
    $NewActiveLineURIs = @()
    foreach ($line in $ActiveLineURIs) {
        # Remove "tel:" from LineURI columns"
        $line.LineURI = ($line.LineURI -replace 'tel:','')
        # Check to see if the LineURIs are within scope set by MinLineURI and MaxLineURI variables, if true then add to new array
        if ( ($line.LineURI -lt $MaxLineURI) -and ($line.LineURI -gt $MinLineURI) ) {
            $NewActiveLineURIs = $NewActiveLineURIs += $line.LineURI.ToString()
        }
    }
    # Sort list so we get the last LineURI, select the last item, convert to an integer then add 1
    $NextAvailableLineURI = (($NewActiveLineURIs |sort | select -last 1) -as [int]) +1
    
    # If no LineURIs are found, create the first one
    if ($NextAvailableLineURI -eq "1") {
        Write-Warning "This is the first LineURI for the given range"
        $SiteCode *= 1000
        $SiteCode += 50
        $NextAvailableLineURI = $SiteCode
    }
    return $NextAvailableLineURI
}

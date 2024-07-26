# Define the manager's distinguished name (DN) or sAMAccountName
$managerDN = "CN=John Doe,OU=Users,DC=yourdomain,DC=com"

# Function to get all subordinates recursively
function Get-Subordinates {
    param (
        [string]$ManagerDN
    )
    
    # Get all direct subordinates of the manager
    $subordinates = Get-ADUser -Filter "manager -eq '$ManagerDN'" -Properties distinguishedName

    foreach ($subordinate in $subordinates) {
        # Output the subordinate's details
        [PSCustomObject]@{
            Name             = $subordinate.Name
            DistinguishedName = $subordinate.DistinguishedName
            sAMAccountName   = $subordinate.sAMAccountName
        }

        # Recursively get subordinates of the current subordinate
        Get-Subordinates -ManagerDN $subordinate.DistinguishedName
    }
}

# Get the initial manager's user object
$manager = Get-ADUser -Filter "distinguishedName -eq '$managerDN'" -Properties distinguishedName

# Start the recursive function
Get-Subordinates -ManagerDN $manager.DistinguishedName

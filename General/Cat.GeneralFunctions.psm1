<# BCT.GeneralFunctions.psm1

This loads all general fucntions. They can also be loaded individually.

Sample Call:
    Import-Module "..\General\Cat.GeneralFunctions.psm1"

#>
# This module contains the Complete-CatScript function that is used by all scripts.
Import-Module "..\General\Functions\Cat.PS-CompleteScript.psm1"

# This module contains the Start-CatScript function that is used by all scripts.
Import-Module "..\General\Functions\Cat.PS-StartScript.psm1"

# This module contains the Wait-BackgroundJobs function that monitors the number of background jobs.
Import-Module "..\General\Functions\Cat.PS-WaitBackgroundJobs.psm1"

# This module contains the Wait-PSSession function that monitors the server to wait for a free powershell session.
Import-Module "..\General\Functions\Cat.PS-WaitPSSessions.psm1"

# This module contains the Wait-SiteCreation function that waits for a site to be created. (Potential Archival)
Import-Module "..\General\Functions\Cat.PS-WaitSiteCreation.psm1"

# This module contains the Wait-SystemResourcesFree function that waits for system resource to drop below a certain point.
Import-Module "..\General\Functions\Cat.PS-WaitSystemResources.psm1"

# This module contains the Write-CatErrorCSV function that logs errors to when looping a CSV file.
Import-Module "..\General\Functions\Cat.PS-WriteError.psm1"
<# Cat.SkySyncFunctions.psm1

This loads all SkySync fucntions. They can also be loaded individually.

Sample Call:
    Import-Module "..\SkySync\Cat.SkySyncFunctions.psm1"

#>
# This module contains the connectionExists function that is used by SkySync scripts.
Import-Module "..\SkySync\Functions\Cat.SkySync-ConnectionExists.psm1"

# This module contains the GetConnectionID function that is used by SkySync scripts.
Import-Module "..\SkySync\Functions\Cat.SkySync-GetConnectionID.psm1"

# This module contains the JobExists function that is used by SkySync scripts.
Import-Module "..\SkySync\Functions\Cat.SkySync-JobExists.psm1"

# This module contains the Start-SkySyncMigration function that is used by SkySync scripts.
Import-Module "..\SkySync\Functions\Cat.SkySync-StartMigration.psm1"

# This module contains the Get-SkySyncMapID function that is used by SkySync scripts.
Import-Module "..\SkySync\Functions\Cat.SkySync-GetMapID.psm1"

# This module contains the Get-SkySyncMapID function that is used by SkySync scripts.
Import-Module "..\SkySync\Functions\Cat.SkySync-IsNull.psm1"
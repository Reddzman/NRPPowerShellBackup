<# Cat.ReportingFunctions.psm1

    The following functions are providing various reporting capabilties.

#>

# This module contains the Get-LastID function.
Import-Module "..\Reporting\Functions\Cat.CSOM-GetLastID.psm1"

# This module contains the Add-CatMigrationRecord function.
Import-Module "..\Reporting\Functions\Cat.PnP-AddMigrationRecord.psm1"

# This module contains the Add-Get-CatADGroups function.
Import-Module "..\Reporting\Functions\Cat.PnP-GetADGroups.psm1"

# This module contains the Get-CatPnPAllWebs function that is used by all scripts.
Import-Module "..\Reporting\Functions\Cat.PnP-GetAllWebs.psm1"

# This module contains the Get-CatPostMigItemInfo function.
Import-Module "..\Reporting\Functions\Cat.PnP-GetPostMigItemInfo.psm1"

# This module contains the Get-CatPostMigWebListInfo function.
Import-Module "..\Reporting\Functions\Cat.PnP-GetPostMigWebListInfo.psm1"

# This module contains the Test-CatPnPSiteExists function.
Import-Module "..\Reporting\Functions\Cat.PnP-TestSiteExists.psm1"

# This module contains the Get-CatSP10ADGroups function.
Import-Module "..\Reporting\Functions\Cat.SP10-GetADGroups.psm1"

# This module contains the Get-CatSP10PowerUsers function.
Import-Module "..\Reporting\Functions\Cat.SP10-GetPowerUsers.psm1"

# This module contains the Test-CatSP10SiteExists function.
Import-Module "..\Reporting\Functions\Cat.SP10-TestSiteExists.psm1"

# This module contains the Get-CatPowerUsers function.
Import-Module "..\Reporting\Functions\Cat.SP13-GetPowerUsers.psm1"

# This module contains the Get-CatSPOAllSites function.
Import-Module "..\Reporting\Functions\Cat.SPO-GetAllSites.psm1"
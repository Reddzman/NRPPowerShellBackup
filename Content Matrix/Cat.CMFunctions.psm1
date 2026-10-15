<# Cat.CMFunctions.psm1

This loads all content matrix fucntions. They can also be loaded individually.

Sample Call:
    Import-Module "..\Content Matrix\Cat.CMFunctions.psm1"
#>

# This module contains the Migrate-CMAllChildLists function that is used by content matix scripts.
Import-Module "..\Content Matrix\Functions\Cat.CM-MigrateAllChildLists.psm1"

# This module contains the Create-CMListJob function that is used by content matix scripts.
Import-Module "..\Content Matrix\Functions\Cat.CM-NewListJob.psm1"

# This module contains the Create-CMListItemsJob function that is used by content matix scripts.
Import-Module "..\Content Matrix\Functions\Cat.CM-NewListItemsJob.psm1"
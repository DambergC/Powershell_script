Dashboard -Name 'Dashimo Test' -FilePath $PSScriptRoot\Dashboard.html {
    Tab -Name 'Forest' {
        Section -Name 'Forest Information' -Invisible {
            Section -Name 'Forest Information' {
                Table -HideFooter -DataTable $DataSetForest.ForestInformation
            }
            Section -Name 'FSMO Roles' {
                Table -HideFooter -DataTable $DataSetForest.ForestFSMO
            }

        }

    }
}
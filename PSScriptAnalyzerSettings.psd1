@{
    # PSScriptAnalyzer configuration for SharePoint-Theme-Studio
    IncludeRules = @(
        'PSAvoidDefaultValueSwitchParameter',
        'PSAvoidGlobalVars',
        'PSAvoidNullOrEmptyHelpMessageAttribute',
        'PSAvoidUsingEmptyCatchBlock',
        'PSAvoidUsingPlainTextForPassword',
        'PSAvoidUsingPositionalParameters',
        'PSAvoidUsingWMICmdlet',
        'PSAvoidTrailingWhitespace',
        'PSReservedCmdletChar',
        'PSReservedParams',
        'PSUseApprovedVerbs',
        'PSUseDeclaredVarsMoreThanAssignments',
        'PSUseSingularNouns',
        'PSUseToExportFieldsInManifest'
    )
    ExcludeRules = @(
        # Allow script-scope variables for GUI state management
        'PSAvoidGlobalVars'
    )
}

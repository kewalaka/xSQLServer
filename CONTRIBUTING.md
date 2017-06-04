# Contributing to xSQLServer

If you are keen to make xSQLServer better, why not consider contributing your work to the project? Every little change helps us make a better resource for everyone to use, and we would love to have contributions from the community.

## Core contribution guidelines

We follow all of the standard contribution guidelines for DSC resources [outlined in DscResources repository](https://github.com/PowerShell/DscResources/blob/master/CONTRIBUTING.md), so please review these as a baseline for contributing.

## xSQLServer specific guidelines

### SQL Server products supported by resources

Any resource should be able to target at least all SQL Server versions that are currently supported by Microsoft (also those in extended support).
Unless the functionality that the resource targets does not exist in a certain SQL Server version.
There can also be other limitations that restrict the resource from targeting all supported versions.

Those SQL Server products that are still supported can be listed at the [Microsoft life cycle site](https://support.microsoft.com/en-us/lifecycle/search?alpha=SQL%20Server).

### Naming convention

#### mof-based resource

All mof-based resource (with Get/Set/Test-TargetResource) should be prefixed with 'MSFT_xSQLServer'. I.e. MSFT_xSQLServerConfiguration

Please note that not all places should contain the prefix 'MSFT_'.

##### Folder and file structure

Please note that for the examples folder we don't use the 'MSFT_' prefix on the resource folders.
This is to make those folders more user friendly, to resemble the name the user would use in the configuration file.

```Text
DSCResources/MSFT_xSQLServerConfiguration/MSFT_xSQLServerConfiguration.psm1
DSCResources/MSFT_xSQLServerConfiguration/MSFT_xSQLServerConfiguration.schema.mof

Tests/Unit/MSFT_xSQLServerConfiguration.Tests.ps1

Examples/Resources/xSQLServerConfiguration/1-AddConfigurationOption.ps1
Examples/Resources/xSQLServerConfiguration/2-RemoveConfigurationOption.ps1
```

##### Schema mof file

Please note that the `FriendlyName` in the schema mof file should not contain the prefix `MSFT_`.

```powershell
[ClassVersion("1.0.0.0"), FriendlyName("xSQLServerConfiguration")]
class MSFT_xSQLServerConfiguration : OMI_BaseResource
{
    # Properties removed for readability.
};
```

#### Composite or class-based resource

Any composite (with a Configuration) or class-based resources should be prefixed with just 'xSQLServer'

### Localization

To be able to support localization we have added wrappers for the cmdlets `Write-Verbose` and `Write-Warning`, and also for the keyword `Throw`.
The localized strings are placed in a file named 'xSQLServer.strings.psd1' which are located in each language folder in the root of the module. For English language strings the folder is ['en-US'](https://github.com/PowerShell/xSQLServer/blob/dev/en-US).

|Function|Short description|
---|---|---
|`New-TerminatingError`| Throws a localized error message using `Throw`. The parameter `ErrorType` takes the message type for which it will get the localized message string. |
|`New-VerboseMessage`| Writes a localized verbose message using `Write-Verbose`. The parameter `ErrorType` takes the message type for which it will get the localized message string. |
|`New-WarningMessage`| Writes a localized warning message using `Write-Warning`. The parameter `ErrorType` takes the message type for which it will get the localized message string. |

### Helper functions

Helper functions or wrapper functions that are used by the resource can preferably be placed in the resource module file. If the functions are of a type that could be used by more than
one resource, then the functions can also be placed in the common [xSQLServerHelper.psm1](https://github.com/PowerShell/xSQLServer/blob/dev/xSQLServerHelper.psm1) module file.

### Tests

For a review of a Pull Request (PR) to start, all tests must pass without error. If you need help to figure why some test don't pass, just write a comment in the Pull Request (PR), or submit an issue, and somebody will come along and assist.

To run all tests manually run the following.

```powershell
Install-Module Pester
cd '<path to cloned repository>\Tests'
Invoke-Pester
```

#### Tests for style check of Markdown files

When sending in a Pull Request (PR) a style check will be performed on all Markdown files, and if the tests find any error the build will fail.
See the section [Documentation with Markdown](#documentation-with-markdown) how these errors kan be found before sending in the PR.

The Markdown tests can be run locally if the packet manager 'npm' is available. To have npm available you need to install [node.js](https://nodejs.org/en/download/).
If 'npm' is not available, a warning text will print and the rest of the tests will continue run.

#### Tests for examples files

When sending in a Pull Request (PR) all example files will be tested so they can be compiled to a .mof file. If the tests find any errors the build will fail.
Before the test runs in AppVeyor the module will be copied to the first path of `$env:PSModulePath`.
To run this test locally, make sure you have the xSQLServer module deployed to a path where it can be used. See `$env:PSModulePath` to view the existing paths.

#### Using SMO stub classes

There are [stub classes](https://github.com/PowerShell/xSQLServer/blob/dev/Tests/Unit/Stubs/SMO.cs) for the SMO classes which can be used and improved on when creating tests where SMO classes are used in the code being tested.

#### AppVeyor

AppVeyor is the platform where the tests is run when sending in a Pull Request (PR). All tests are run on a clean AppVeyor build worker for each push to the Pull Request (PR).
The tests that are run on the build worker are common tests, unit tests and integration tests (with some limitations).

### Documentation with Markdown

If using Visual Studio Code to edit Markdown files it can be a good idea to install the markdownlint extension. It will help to do style checking.
The file [.markdownlint.json](/.markdownlint.json) is prepared with a default set of rules which will automatically be used by the extension.

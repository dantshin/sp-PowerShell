# sp-PowerShell

[PowerShellRef]: https://docs.microsoft.com/en-us/powershell/
[PnPPowerShell]: https://pnp.github.io/powershell/

`sp-PowerShell` is a collection of scripts that perform various operations on SharePoint sites.

### Requirements

* [PowerShell 5][PowerShellRef]
* [PnP-PowerShell][PnPPowerShell] (Automatically imported by the sp-PowerShell module)

### Note

A [previous issue where PowerShell 7 was not supported](https://github.com/pnp/PnP-PowerShell/issues/2595) has now been resolved as we moved from Windows PnP PowerShell to PnP-PowerShell.

### Installing

* Clone repository to corresponding [PowerShell Modules location](https://docs.microsoft.com/en-us/powershell/scripting/developer/module/installing-a-powershell-module).
* Run `Import-Module sp-PowerShell` and you are ready to go!
 * This might look more like `Import-Module .\sp-PowerShell.psd1`

## Authors

* Daniel Tshin - **[@dantshin](https://github.com/dantshin)**
* Giuseppe Campanelli - **[@themilanfan](https://github.com/themilanfan)**

## Contributing

Please see the [Contribution Guide](CONTRIBUTING.md) for information on how to develop and contribute.

## License

sp-PowerShell is licensed under the [MIT License](LICENSE.md).


# excelAzureVmPricing
This VBA project allows you to find the cheapst VM size for a given Core/Ram configuration in a specific datacenter.
You can then pull an hour price for the VM. 

The solution relies on a custom backend, that pulls data from https://azure.microsoft.com/api/v2/pricing

It uses a custom service for retrieving VM prices
![Demoimage](https://raw.githubusercontent.com/KillerFeature/excelAzureVmPricing/master/Capture3.PNG)


# Installation
1. Download the VM_Prices.bas (https://raw.githubusercontent.com/KillerFeature/excelAzureVmPricing/master/VM_Prices.bas?raw=true)
2. Open Excel
3. Press Alt-F11 to go to Macro Editor
4. Select "File" -> "Import Module"
5. Select the VM_Prices.bas file
6. Select "Tools" -> "References..."
7. Check "Microsoft XML, v6.0"
8. Press Alt-F11 to go back to Excel
9. Enter this in a cell =getVM(1;1;0;"europe-west";"EUR") the resulting VM should be linux-b1s-standard
10. Move to the next cell and type =getVMPriceHour("linux-b1s-standard";0;"europe-west";"EUR") the result should be something like 0,07929684

# Function syntax

=getVM([minimum cores];[minimum ram];[reserved instance years 0 or 1 or 3];[azure-region];[currency])

Optionally you can exclude certain vm's by using semicolon seperated tags

=getVM([minimum cores];[minimum ram];[reserved instance years 0 or 1 or 3];[azure-region];[currency];"linux-b")
This will exclude all burstable VM's

=getVMPriceHour([VM Name (result from getVM)];[reserved instance];[azure region];[currency])

=getVMData([VM Name];[Region];[Currency];[Parameter you want returned])

Example:
=getVMData("linux-b2s-standard";"us-east";"USD";"isVcpu")

Supported parameters : isVCPU cores ram (DO NOT USE getVMData to get prices)

How to calculate monthly fee?
=[Hour price]*730


See demovideo here : [video](https://github.com/KillerFeature/excelAzureVmPricing/blob/master/comp.mp4?raw=true)


# Known issues
## Changing currency doesn't work!
Try to keep your currency identifier in a single cell and reference it. When you change the currency save your workbook, exit excel and reopen the workbook. Refresh your cells.

## I want more that 10 azure regions in one workbook
NO!! ... Well yes now you can get 100 :)

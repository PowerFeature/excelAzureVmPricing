

# excelAzureVmPricing
This VBA project allows you to find the cheapst VM size for a given Core/Ram configuration in a specific datacenter.
You can then pull an hour price for the VM. 

The solution relies on a custom backend, that pulls data from https://azure.microsoft.com/api/v2/pricing

It uses a custom service for retrieving VM prices
![Demoimage](https://github.com/KillerFeature/excelAzureVmPricing/blob/master/Capture2.PNG?raw=true)

# Installation
1. Download the ![VM_Prices.bas](https://raw.githubusercontent.com/KillerFeature/excelAzureVmPricing/master/VM_Prices.bas?raw=true)
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

=getVMPriceHour([VM Name (result from getVM)];[reserved instance];[azure region];[currency])

See demovideo here : [video](https://github.com/KillerFeature/excelAzureVmPricing/blob/master/comp.mp4?raw=true)

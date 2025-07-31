param (
    [string]$site
)

if(($site.ToLower() -eq "rs") -or ($site.ToLower() -eq "riverside") ) {
    exit 0;
}

if(($site.ToLower() -eq "nb") -or ($site.ToLower() -eq "north bay") -or ($site.ToLower() -eq "northbay")) {
    Write-Host 
    Write-Host "Setting up North Bay Printers" -BackgroundColor "blue"
    Start-Process -Wait .\..\drivers\northbay_printer_setup.exe

    
    #Kyocera 3554ci driver printer(s)
    $k3354ci_driver = "Kyocera TASKalfa 3554ci KX"
    Add-PrinterDriver -Name $k3354ci_driver

    Write-Host "Adding Middle School Printer" -BackgroundColor "green"
    $ms_printer_ip = "10.11.10.200"
    $ms_port_name = "IP_" + $ms_printer_ip + "_KX"
    $ms_printer_name = "Middle School"
    Remove-Printer -ErrorAction SilentlyContinue -Name $ms_printer_name
    Add-PrinterPort -Name $ms_port_name -PrinterHostAddress $ms_printer_ip
    Add-Printer -Name $ms_printer_name -PortName $ms_port_name -DriverName $k3354ci_driver

    
    #Kyocera 358ci driver priner(s)
    $k358ci_driver = "Kyocera TASKalfa 358ci KX"
    Add-PrinterDriver -Name $k358ci_driver

    Write-Host "Adding Staff Area Printer" -BackgroundColor "green"
    $staff_printer_ip = "10.11.10.202"
    $staff_port_name = "IP_" + $staff_printer_ip + "_KX"
    $staff_printer_name = "Staff Area"
    Remove-Printer -ErrorAction SilentlyContinue -Name $staff_printer_name
    Add-PrinterPort -Name $staff_port_name -PrinterHostAddress $staff_printer_ip
    Add-Printer -Name $staff_printer_name -PortName $staff_port_name -DriverName $k358ci_driver

    Write-Host "Adding Suite A Printer" -BackgroundColor "green"
    $stea_printer_ip = "10.11.10.141"
    $stea_port_name = "IP_" + $stea_printer_ip + "_KX"
    $stea_printer_name = "Suite A"
    Remove-Printer -ErrorAction SilentlyContinue -Name $stea_printer_name
    Add-PrinterPort -Name $stea_port_name -PrinterHostAddress $stea_printer_ip
    Add-Printer -Name $stea_printer_name -PortName $stea_port_name -DriverName $k358ci_driver
}

if(($site.ToLower() -eq "nv") -or ($site.ToLower() -eq "north valley") -or ($site.ToLower() -eq "northvalley")) {
    Write-Host 
    Write-Host "Setting up North Valley Printers" -BackgroundColor "blue"
    pnputil.exe /add-driver ..\drivers\PCL6\KOBxxK__01.inf

    #konica minolta printer
    $km_driver = "KONICA MINOLTA Universal V4 PCL"
    Add-PrinterDriver -Name $km_driver

    Write-Host "Adding North Valley printer" -BackgroundColor "green"
    $km_printer_ip = "192.169.1.150"
    $km_port_name = "IP_" + $km_printer_ip
    $km_printer_name = "North Valley"
    Remove-Printer -ErrorAction SilentlyContinue -Name $km_port_name
    Add-PrinterPort -Name $km_printer_name -PrinterHostAddress $km_printer_ip
    Add-Printer -Name $km_port_name -PortName $km_printer_name -DriverName $k358ci_driver
}

if(($site.ToLower() -eq "SD") -or ($site.ToLower() -eq "san diego") -or ($site.ToLower() -eq "northvalley")) {
    Write-Host 
    Write-Host "Setting up San Diego Printers" -BackgroundColor "blue"
    $inf_file="..\drivers\cannon\CNS30MA64.INF"
    pnputil.exe /add-driver $inf_file

    #Cannon printer
    $driver_name = "Canon Generic Plus PS3"
    Add-PrinterDriver -Name $driver_name

    Write-Host "Adding San Diego printer" -BackgroundColor "green"
    $cannon_printer_ip = "192.168.128.45"
    $cannon_port_name = "IP_" + $cannon_printer_ip
    $cannon_printer_name = "Meeting Room San Diego"
    Remove-Printer -ErrorAction SilentlyContinue -Name $cannon_printer_name
    Remove-PrinterPort -ErrorAction SilentlyContinue -Name $cannon_port_name
    Add-PrinterPort -Name $cannon_port_name -PrinterHostAddress $cannon_printer_ip
    Add-Printer -Name $cannon_printer_name -PortName $cannon_port_name -DriverName $driver_name
}
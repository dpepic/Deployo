##############################################################################
#.SYNOPSIS
# Creates a VM.
#
#.DESCRIPTION
# Just a placeholder for now, will plop down a VM in the default location
# trying for a VHD of the same name.
#
#.PARAMETER Name
# Name for the VM.
#
#.PARAMETER RAM
# Amount of RAM.
#
#.PARAMETER Generation
# Generation of VM.
#
#.EXAMPLE
# New-DepoloyoVM -Name testVM
# Will deploy a 2GB RAM gen1 VM named testVM, trying for test.vhd in the 
# default directory.
#
#.EXAMPLE
# New-DepoloyoVM -Name testVM -RAM 4GB -Generation 2
# Will deploy a 4GB RAM gen2 VM named testVM, trying for test.vhdx in the
# default directory.
#
#.EXAMPLE
# New-DeployoVHD -Name test -PartitionType MBR -Size 25GB -Dynamic | New-DeployoOS -WimLocation .\IMG\install.wim | New-DeployoVM -Name testVM
# Used in the full Deployo pipeline.
##############################################################################
function New-DeployoVM
{ 
    param
    (
        [parameter(Mandatory=$true, ParameterSetName='Pipeline')]
        [string] $Name,
        [parameter(Mandatory=$false, ParameterSetName='Pipeline')]
        [UInt64] $RAM = 2GB,
        [parameter(Mandatory=$false, ParameterSetName='Pipeline')]
        [string] $Generation = 1,
        [Parameter(Mandatory=$True, ValueFromPipeline=$True, ParameterSetName='Pipeline')]
        $VHDlocation
    )

    #Just a quick and dirty placeholder, more to come here at a later date
    New-VM -Name $Name -Generation $Generation -MemoryStartupBytes $RAM -BootDevice VHD -VHDPath $VHDlocation
}

##############################################################################
#.SYNOPSIS
# Creates a VHD or VHDX and deploys a OS on it.
#
#.DESCRIPTION
# NOTE: GPT work in PS is buggy at best ATM, once in a blue moon you will get 
# a bootable machine, most of the time not. You can see the sort of 
# workarounds I had to do to get it working below. I suggest sticking 
# with MBR and gen1 VMs for the moment.
#
# This function will create a VHD or VHDX, create a MBR or GPT partition
# table on it according to standards (GPT will have a fat32 sys partition)
# and prep it for wim deployment.
#
#.PARAMETER Name
# Name of the VHD/X.
#
#.PARAMETER PartitionType
# MBR or GPT (not recommended!).
#
#.PARAMETER Size
# Size of disk.
#
#.PARAMETER Dynamic
# Will the disk by dynamically allocated or static.
#
#.EXAMPLE
# New-DeployoVHD -Name test -PartitionType MBR -Size 25GB -Dynamic
# 
# Creates a MBR VHD named test that is dynamically allocated.
#
#.EXAMPLE
# New-DeployoVHD -Name test -PartitionType GPT -Size 25GB
# 
# Creates a GPT VHD named test that is statically allocated.
#
#.EXAMPLE
# New-DeployoVHD -Name test -PartitionType MBR -Size 25GB -Dynamic | New-DeployoOS -WimLocation .\IMG\install.wim | New-DeployoVM -Name testVM
# Used in the full Deployo pipeline.
##############################################################################
function New-DeployoVHD 
{    
    param
    (
        [parameter(Mandatory=$true)]
        [string] $Name,
        [parameter(Mandatory=$true)]
        [ValidateSet('MBR', 'GPT')]
        [string] $PartitionType,
        [parameter(Mandatory=$true)]
        [UInt64] $Size,
        [switch] $Dynamic
    )

    [hashtable] $diskData = @{'VHDName' = $Name}
    $diskData.Add('partitionType', $PartitionType)
    
    if ($PartitionType -eq 'MBR')
    { 
        $diskData.Add('VHDExtension', '.vhd')
    }
    else
    {
        $diskData.Add('VHDExtension', '.vhdx') #UEFI disks are for gen2 VMs so VHDX
    }

    $vmms = Get-Service vmms
    switch ((Measure-Object -InputObject $vmms).count)
    {
        0 
        {
            Write-Verbose 'VMMS service does not exist. You must have Hyper-V installed for this to work, aborting.'
            exit
        }
        1 
        {
            if ((Get-Service $vmms.Name).Status -ne 'Running')
            {
                Write-Verbose 'VMMS is here, but is not running, starting the service...'
                Start-Service $vmms   
            }
        }
    }

    Write-Verbose 'Creating the disk...'
    $diskData.Add('VHDPath', (Get-VMHost).VirtualHardDiskPath) 
    $VHDFullLocation = $diskData.VHDPath + '\' + $diskData.VHDName + $diskData.VHDExtension
    $VHD = Get-VHD -Path $VHDFullLocation -ErrorAction SilentlyContinue
    if ((Measure-Object -InputObject $VHD).count -lt 1)
    {
       if ($Dynamic)
        {
            $disk = New-VHD -Path $VHDFullLocation -Dynamic -SizeBytes $Size    
        }
        else
        {
            $disk = New-VHD -Path $VHDFullLocation -SizeBytes $Size 
        } 
        Mount-DiskImage -ImagePath $VHDFullLocation #Create the VHD/X and mount it
        $diskNumber = (Get-DiskImage -ImagePath $VHDFullLocation).Number
    }
    else
    {
        #Sanity check for existing VHD
        $title = "VHD already exists"
        $message = "If you proceed the VHD will be WIPED of all data."
        $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
            "Reinitializes the VDH."
        $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
             "Aborts the operation."
        $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
        $result = $host.ui.PromptForChoice($title, $message, $options, 0) 
        if ($result -eq 1)
        {
            exit
        }   

        $disk = $VHD
        Mount-DiskImage -ImagePath $VHDFullLocation #Create the VHD/X and mount it
        $diskNumber = (Get-DiskImage -ImagePath $VHDFullLocation).Number
        Clear-Disk -Number $diskNumber -RemoveData -RemoveOEM
    } 

    Write-Verbose 'Working on the partition table...'
    Initialize-Disk -Number $diskNumber -PartitionStyle $diskData.partitionType -ErrorAction SilentlyContinue

    Stop-Service -Name ShellHWDetection #We stop the service while working on the VHD/X to prevent dialog popups

    if ($diskData.partitionType -eq 'MBR')
    {
        Write-Verbose 'MBR setup...'
        $bootPartition = New-Partition -DiskNumber $diskNumber -UseMaximumSize -IsActive -AssignDriveLetter | Format-Volume -FileSystem NTFS  
        $diskData.Add('bootDrive', ($bootPartition | Get-Partition).DriveLetter)
        $diskData.Add('windowsDrive', $diskData.bootDrive )
        #Straightforward MBR setup, one partition covering the full disk, the boot is on the same partition as Windows
    }
    else
    {   
        Write-Verbose 'GPT setup...'
        #Creating an EFI system partition for boot data
        $bootPartition = New-Partition -DiskNumber $diskNumber -Size 100MB -GptType '{c12a7328-f81f-11d2-ba4b-00a0c93ec93b}' 
        $bootPartition | Add-PartitionAccessPath -AssignDriveLetter
        $diskData.Add('bootDrive', ($bootPartition | Get-Partition).DriveLetter)
        #Workaround I : Format-Volume cannot format an EFI system partition, it will fail trying, so we are dropping to 
        #diskpart for the formatting
        #Workaround II : Any assigned drive letters by PowerShell on EFI system partitions will work, but PS is not going
        #to be able to free them when dismounting the image and they will read as mounted but can not be unmounted
        #(as they are indeed mounted to nothing) nor do they exist in the registry hives, leaving them in limbo until reboot.
        #So the pro forma -AssignDriveLetter above is just to easily grab an available drive letter, we will depend on 
        #diskpart for actual assignment.
        "@
        select disk $diskNumber
        select partition $($bootPartition.PartitionNumber)
        format quick fs=fat32
        assign letter $diskData.bootDrive
        exit
        @" | diskpart | Out-Null 
        
        #This is where Windows lives
        $windowsPartition = New-Partition -DiskNumber $diskNumber -UseMaximumSize -GptType '{ebd0a0a2-b9e5-4433-87c0-68b6b72699c7}' -AssignDriveLetter | Format-Volume -FileSystem NTFS 
        $diskData.Add('windowsDrive', ($windowsPartition | Get-Partition).DriveLetter)    
    }

    Start-Service -Name ShellHWDetection #Done working with the disk so we can start the service back up
   
   
    #Workaround II contd. : As mentioned, PS cmdlets will not be able to free letters they assigned to EFI system
    #partitions so we are dropping to diskpart before we dismount to free up the drive letter.
    if ($diskData.partitionType -eq 'GPT')
    {
        echo 'Cleaning up EFI system drive letter...'
        "@
        select disk $diskNumber
        select partition $($bootPartition.PartitionNumber)
        remove letter $diskData.bootDrive
        exit
        @" | diskpart | Out-Null
    }

    Write-Verbose 'Done, dismounting...'
    Dismount-DiskImage -ImagePath $VHDFullLocation
    Write-Output ($diskData -as [hashtable])
}

##############################################################################
#.SYNOPSIS
# Deploys a WIM on a VHD.
#
#.DESCRIPTION
# Work in progress.
# TODO: Unattend.xml applying
# TODO: Additional software deployment
#
# This function deploys a WIM over a VHD. Currently in a very crude form,
# be mindful it expects the DISM tools in .\tools relative to from where 
# you are running the function.
#
#.PARAMETER windowsDrive
# Drive letter for the volume where \Windows is.
#
#.PARAMETER bootDrive
# Drive letter for the volume where the boot data should be.
#
#.PARAMETER BootType
# BIOS for MBR disks or UEFI for GPT based disks.
#
#.PARAMETER diskData
# A hashtable only used if a vhd is getting piped from New-DepolyoVHD.
#
#.EXAMPLE
# New-DeployoOS -windowsDrive Z -bootDrive S -WimLocation .\IMG\win.wim -bootType UEFI
# Will deploy win.wim image on index 1 over the Z drive and configure BCD on S:.
#
#.EXAMPLE
# New-DeployoOS -windowsDrive Z -WimLocation .\IMG\win.wim -WimIndex 4 -bootType BIOS
# Will deploy win.wim image on index 4 over the Z drive and configure BCD on 
# Z: as well.
#
#.EXAMPLE
# New-DeployoVHD -Name test -PartitionType MBR -Size 25GB -Dynamic | New-DeployoOS -WimLocation .\IMG\install.wim | New-DeployoVM -Name testVM
# Used in the full Deployo pipeline.
##############################################################################
function New-DeployoOS
{ 
    param
    (

        [parameter(Mandatory=$true, ParameterSetName='Standard in')]
        [string] $WindowsDrive,
        [parameter(Mandatory=$false, ParameterSetName='Standard in')]
        [string] $BootDrive = $WindowsDrive,
        [parameter(Mandatory=$true, ParameterSetName='Standard in')]
        [Parameter(Mandatory=$true, ParameterSetName='Pipeline')]
        [string] $WimLocation,
        [parameter(Mandatory=$false, ParameterSetName='Standard in')]
        [Parameter(Mandatory=$false, ParameterSetName='Pipeline')]
        [int] $WimIndex = 1,
        [parameter(Mandatory=$true, ParameterSetName='Standard in')]
        [ValidateSet('BIOS', 'UEFI')]
        [string] $BootType,
        [Parameter(Mandatory=$True, ValueFromPipeline=$True, ParameterSetName='Pipeline')]
        $diskData
    )

        if ($PSCmdlet.ParameterSetName -eq 'Pipeline')
        {
            $WindowsDrive = $diskData.windowsDrive
            $BootDrive = $diskData.bootDrive     
            if ($diskData.partitionType -eq 'MBR')
            {
                $BootType = 'BIOS'
            }
            else
            {
                $BootType = 'UEFI'
            }
            $VHDLocation = $diskData.VHDPath +'\' + $diskData.VHDName + $diskData.VHDExtension            
        }

        $WimLocation = 'C:\Apex\P\IMG\boot.wim'
        $extension = $WimLocation.Split('.')[$WimLocation.Split('.').Count-1]
        
        if ($extension -eq 'ESD')
        {
            Write-Verbose 'Found ESD, converting...'
            $wimRenamed = [string]::Concat('.', $WimLocation.Split('.')[$WimLocation.Split('.').Count-2], '.wim')
            if (Test-Path -Path $wimRenamed)
            {
                #Sanity check for existing WIM
                $title = "WIM already exists"
                $message = "If you proceed the WIM will get the ESD content appended to it, as it exists."
                $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
                    "Append the content of the ESD."
                $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
                    "Aborts the operation."
                $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
                $result = $host.ui.PromptForChoice($title, $message, $options, 0) 
                if ($result -eq 1)
                {
                    exit
                }
            } 
            Export-WindowsImage -SourceImagePath $WimLocation -SourceIndex $WimIndex -DestinationImagePath $wimRenamed -CheckIntegrity -CompressionType max
            $WimLocation = $wimRenamed
        }

        Mount-DiskImage -ImagePath $VHDLocation

        Write-Verbose "Applying WIM to ${WindowsDrive}:..."
        .\tools\DISM\imagex.exe /apply $WimLocation $WimIndex ${WindowsDrive}:
        Write-Verbose "Installing bootloader for ${WindowsDrive}: to ${BootDrive}:"
        .\tools\BCDBoot\bcdboot.exe ${WindowsDrive}:\Windows /s ${BootDrive}:  /f $BootType 

        Dismount-DiskImage -ImagePath $VHDLocation
        Write-Output $VHDLocation
}

##############################################################################
#.SYNOPSIS
# Gets info on all indexes in a WIM.
#
#.DESCRIPTION
# A simple function that will output nicely formatted information on all 
# editions present in a WIM file. be mindful it expects the DISM tools in 
# .\tools relative to from where you are running the function. 
#
#.PARAMETER WIM
# WIM file that we are interested in.
#
#.EXAMPLE
# Get-WimOSEditions -WIM .\IMG\install.wim
# Will list information for all indexes in .\IMG\install.wim.
##############################################################################
function Get-WimOSEditions
{
    param
    (
        [parameter(Mandatory=$true)]
        [string]$WIM
    )

    $dry = .\tools\DISM\imagex.exe /info $WIM
    foreach ($line in $dry)
    {
        if($line -ilike '*<IMAGE INDEX*')
        {
            $clean += 'Index ' + ($line -split '"')[1]
        }
        elseif ($line -ilike '*<ARCH>*')
        {
            if ($line -ilike '*9*')
            {
                $clean += ' is a x64 version of '
            } 
            else
            {
                $clean += ' is a x86 version of '
            }
        }
        elseif ($line -ilike '*<DESCRIPTION>*')
        {
            $clean += (($line -split '>')[1] -split '<')[0] + "`n"
        }
    }
    $clean 
}

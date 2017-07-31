##############################################################################
#.SYNOPSIS
# Creates a VM.
#
#.DESCRIPTION
# Just a placeholder for now, will plop down a VM in the default location
# trying for a VHD of the same name.
#
#.PARAMETER Name
# Name for the VM
#
#.PARAMETER RAM
# Amount of RAM
#
#.PARAMETER Generation
# Generation of VM.
#
#.EXAMPLE
# New-DepoloyoVM -Name testVM
# Will deploy a 2GB RAM gen1 VM named testVM, trying for test.vhd in the 
# default directory
#
#.EXAMPLE
# New-DepoloyoVM -Name testVM -RAM 4GB -Generation 2
# Will deploy a 4GB RAM gen2 VM named testVM, trying got test.vhdx in the
# default directory
##############################################################################
function New-DeployoVM
{ 
    param
    (
        [parameter(Mandatory=$true)]
        [string] $Name,
        [parameter(Mandatory=$false)]
        [UInt64] $RAM = 2GB,
        [parameter(Mandatory=$false)]
        [string] $Generation = 1
    )

    #Just a quick and dirty placeholder, more to come here at a later date
    $VHDpath = (Get-VMHost).VirtualHardDiskPath
    if ($Generation -eq 1)
    {
        $VHDextension = '.vhd'
    }
    else
    {
        $VHDextension = '.vhdx'
    }
    New-VM -Name $Name -Generation $Generation -MemoryStartupBytes $RAM -BootDevice VHD -VHDPath $VHDpath\$Name$VHDextension
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
# MBR or GPT (not recommended!)
#
#.PARAMETER Size
# Size of disk
#
#.PARAMETER Dynamic
# Will the disk by dynamically allocated or static
#
#.EXAMPLE
# New-DeployoVHD -Name test -PartitionType MBR -Size 25GB -Dynamic
# 
# Creates a MBR VHD named test that is dynamically allocated
#
#.EXAMPLE
# New-DeployoVHD -Name test -PartitionType GPT -Size 25GB
# 
# Creates a GPT VHD named test that is statically allocated
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
     
    if ($PartitionType -eq 'MBR')
    { 
        $ext = '.vhd'
        $bootType = 'BIOS'
    }
    else
    {
        $ext = '.vhdx' #UEFI disks are for gen2 VMs so VHDX
        $bootType = 'UEFI'
    }

    $vmms = Get-Service vmms
    switch ((Measure-Object -InputObject $vmms).count)
    {
        0 
        {
            echo 'VMMS service does not exist. You must have Hyper-V installed for this to work, aborting.'
            exit
        }
        1 
        {
            if ((Get-Service $vmms.Name).Status -ne 'Running')
            {
                echo 'VMMS is here, but is not running, starting the service...'
                Start-Service $vmms   
            }
        }
    }

    echo 'Creating the disk...'
    $VHDPath = (Get-VMHost).VirtualHardDiskPath 
    if ($Dynamic)
    {
        $disk = New-VHD -Path $VHDPath\$Name$ext -Dynamic -SizeBytes $Size    
    }
    else
    {
        $disk = New-VHD -Path $VHDPath\$Name$ext -SizeBytes $Size 
    }
    Mount-DiskImage -ImagePath $VHDPath\$Name$ext #Create the VHD/X and mount it

    echo 'Working on the partition table...'
    $diskNumber = (Get-DiskImage -ImagePath $VHDPath\$Name$ext).Number
    Initialize-Disk -Number $diskNumber -PartitionStyle $PartitionType 

    Stop-Service -Name ShellHWDetection #We stop the service while working on the VHD/X to prevent dialog popups

    if ($PartitionType -eq 'MBR')
    {
        echo 'MBR setup...'
        $bootPartition = New-Partition -DiskNumber $diskNumber -UseMaximumSize -IsActive -AssignDriveLetter | Format-Volume -FileSystem NTFS 
        $bootDrive = ($bootPartition | Get-Partition).DriveLetter 

        $windowsDrive = $bootDrive 
        #Straightforward MBR setup, one partition covering the full disk, the boot is on the same partition as Windows
    }
    else
    {   
        #Creating an EFI system partition for boot data
        $bootPartition = New-Partition -DiskNumber $diskNumber -Size 100MB -GptType '{c12a7328-f81f-11d2-ba4b-00a0c93ec93b}' 
        $bootPartition | Add-PartitionAccessPath -AssignDriveLetter
        $bootDrive = ($bootPartition | Get-Partition).DriveLetter
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
        assign letter $bootDrive
        exit
        @" | diskpart  
        
        #This is where Windows lives
        $windowsPartition = New-Partition -DiskNumber $diskNumber -UseMaximumSize -GptType '{ebd0a0a2-b9e5-4433-87c0-68b6b72699c7}' -AssignDriveLetter | Format-Volume -FileSystem NTFS 
        $windowsDrive = ($windowsPartition | Get-Partition).DriveLetter    
    }

    Start-Service -Name ShellHWDetection #Done working with the disk so we can start the service back up

    #Quick shortcut while rapidly testing the script, will not exist soon
    New-DeployoOS -windowsDrive $windowsDrive -bootDrive $bootDrive -WimLocation .\IMG\boot.wim -bootType $bootType

    #Workaround II contd. : As mentioned, PS cmdlets will not be able to free letters they assigned to EFI system
    #partitions so we are dropping to diskpart before we dismount to free up the drive letter.
    if ($PartitionType -eq 'GPT')
    {
        "@
        select disk $diskNumber
        select partition $($bootPartition.PartitionNumber)
        remove letter $bootDrive
        exit
        @" | diskpart
    }

    echo 'Done, dismounting...'
    Dismount-DiskImage -ImagePath $VHDPath\$Name$ext
}

##############################################################################
#.SYNOPSIS
# Deploys a WIM on a VHD.
#
#.DESCRIPTION
# Work in progress.
# TODO: Index selection
# TODO: ESD conversion
# TODO: Unattend.xml applying
# TODO: Additional software deployment
#
# This function deploys a WIM over a VHD. Currently in a very crude form,
# be mindfull it expects the DISM tools in .\tools relative to from where 
# you are running the function.
#
#.PARAMETER windowsDrive
# Drive letter (ex. G) for the volume where \Windows is
#
#.PARAMETER bootDrive
# Drive letter (ex. G) for the volume where the boot data should be
#
#.PARAMETER BootType
# BIOS for MBR disks or UEFI for GPT based disks
#
#.EXAMPLE
# New-DeployoOS -windowsDrive Z -bootDrive S -WimLocation .\IMG\win.wim -bootType UEFI
# Will deploy win.wim image on index 1 over the Z drive and configure BCD on S:
#
#.EXAMPLE
# New-DeployoOS -windowsDrive Z -WimLocation .\IMG\win.wim -WimIndex 4 -bootType BIOS
# Will deploy win.wim image on index 4 over the Z drive and configure BCD on Z: as well
##############################################################################
function New-DeployoOS
{ 
    param
    (
        [parameter(Mandatory=$true)]
        [string] $WindowsDrive,
        [parameter(Mandatory=$false)]
        [string] $BootDrive = $WindowsDrive,
        [parameter(Mandatory=$true)]
        [string] $WimLocation,
        [parameter(Mandatory=$false)]
        [int] $WimIndex = 1,
        [parameter(Mandatory=$true)]
        [ValidateSet('BIOS', 'UEFI')]
        [string] $BootType
    )

   .\tools\DISM\imagex.exe /apply $WimLocation $WimIndex ${WindowsDrive}:
   .\tools\BCDBoot\bcdboot.exe ${WindowsDrive}:\Windows /s ${BootDrive}:  /f $BootType
}
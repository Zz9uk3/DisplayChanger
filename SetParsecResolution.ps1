Add-Type @"
using System;
using System.Runtime.InteropServices;

public class DisplayHelper
{
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct DISPLAY_DEVICE
    {
        public int cb;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string DeviceName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
        public string DeviceString;
        public int StateFlags;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
        public string DeviceID;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
        public string DeviceKey;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct DEVMODE
    {
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string dmDeviceName;
        public short dmSpecVersion;
        public short dmDriverVersion;
        public short dmSize;
        public short dmDriverExtra;
        public int dmFields;
        public int dmPositionX;
        public int dmPositionY;
        public int dmDisplayOrientation;
        public int dmDisplayFixedOutput;
        public short dmColor;
        public short dmDuplex;
        public short dmYResolution;
        public short dmTTOption;
        public short dmCollate;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string dmFormName;
        public short dmLogPixels;
        public int dmBitsPerPel;
        public int dmPelsWidth;
        public int dmPelsHeight;
        public int dmDisplayFlags;
        public int dmDisplayFrequency;
        public int dmICMMethod;
        public int dmICMIntent;
        public int dmMediaType;
        public int dmDitherType;
        public int dmReserved1;
        public int dmReserved2;
        public int dmPanningWidth;
        public int dmPanningHeight;
    }

    [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "EnumDisplayDevicesW")]
    public static extern bool EnumDisplayDevices(IntPtr lpDevice, uint iDevNum, ref DISPLAY_DEVICE lpDisplayDevice, uint dwFlags);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "EnumDisplayDevicesW")]
    public static extern bool EnumDisplayDevicesByName(string lpDevice, uint iDevNum, ref DISPLAY_DEVICE lpDisplayDevice, uint dwFlags);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern bool EnumDisplaySettings(string lpszDeviceName, uint iModeNum, ref DEVMODE lpDevMode);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int ChangeDisplaySettingsEx(string lpszDeviceName, ref DEVMODE lpDevMode, IntPtr hwnd, uint dwflags, IntPtr lParam);

    public const uint ENUM_CURRENT_SETTINGS = 0xFFFFFFFF;
    public const uint CDS_UPDATEREGISTRY = 0x01;
    public const int DM_PELSWIDTH = 0x80000;
    public const int DM_PELSHEIGHT = 0x100000;
}
"@

$targetWidth = 1920
$targetHeight = 1080

# Find the Parsec adapter
$adapter = New-Object DisplayHelper+DISPLAY_DEVICE
$adapter.cb = [System.Runtime.InteropServices.Marshal]::SizeOf($adapter)
$adapterName = $null

for ($i = 0; [DisplayHelper]::EnumDisplayDevices([IntPtr]::Zero, $i, [ref]$adapter, 0); $i++) {
    if ($adapter.DeviceString -like "*Parsec*" -and ($adapter.StateFlags -band 1)) {
        $adapterName = $adapter.DeviceName
        break
    }
    $adapter.cb = [System.Runtime.InteropServices.Marshal]::SizeOf($adapter)
}

if (-not $adapterName) {
    Write-Host "ParsecVDA display not found." -ForegroundColor Red
    exit 1
}

# Get current resolution
$devMode = New-Object DisplayHelper+DEVMODE
$devMode.dmSize = [System.Runtime.InteropServices.Marshal]::SizeOf($devMode)

if (-not [DisplayHelper]::EnumDisplaySettings($adapterName, [DisplayHelper]::ENUM_CURRENT_SETTINGS, [ref]$devMode)) {
    Write-Host "Failed to get current display settings." -ForegroundColor Red
    exit 1
}

Write-Host "Found Parsec adapter: $adapterName"
Write-Host "Current resolution: $($devMode.dmPelsWidth)x$($devMode.dmPelsHeight)"

if ($devMode.dmPelsWidth -eq $targetWidth -and $devMode.dmPelsHeight -eq $targetHeight) {
    Write-Host "Resolution is already ${targetWidth}x${targetHeight}." -ForegroundColor Green
    exit 0
}

# Set new resolution
$devMode.dmPelsWidth = $targetWidth
$devMode.dmPelsHeight = $targetHeight
$devMode.dmFields = [DisplayHelper]::DM_PELSWIDTH -bor [DisplayHelper]::DM_PELSHEIGHT

$result = [DisplayHelper]::ChangeDisplaySettingsEx($adapterName, [ref]$devMode, [IntPtr]::Zero, [DisplayHelper]::CDS_UPDATEREGISTRY, [IntPtr]::Zero)

if ($result -eq 0) {
    Write-Host "Resolution changed to ${targetWidth}x${targetHeight}." -ForegroundColor Green
} else {
    Write-Host "Failed to change resolution. Error code: $result" -ForegroundColor Red
    exit 1
}

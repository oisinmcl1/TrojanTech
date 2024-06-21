This script assists in gathering the hardware hash of devices during the Out-of-Box Experience (OOBE) process, which is essential for registering devices with Microsoft Intune's Autopilot service (for MOWI).


Instructions for Obtaining Hardware Hash:

Access Command Prompt: During the Out-of-Box Experience (OOBE) setup process of your Windows 11 device, press Shift + F10 to open the Command Prompt.

Insert the USB drive: Insert this USB drive into the device's USB port.

Run the script: In the Command Prompt window, navigate to the USB drive (D:) using the command D: and press Enter. Then, run the CMD file named GetAutoPilot.cmd by typing .\GetAutoPilot.cmd and pressing Enter.

Wait for completion: Allow the script to run and gather the hardware hash. Once the process is complete, a CSV file named compHash.csv will be generated on the USB drive.

Remove the USB drive: Safely remove the USB drive from the device.

Email the CSV file: Locate the CSV file named compHash.csv on the USB drive. Email this file to the administrator responsible for device registration (MOWI?).


For additional assistance with Windows Autopilot deployment, refer to https://learn.microsoft.com/en-us/autopilot/add-devices.


Oisín Mc Laughlin
24/05/2024

Script just quickly made to ping an ip of your desire and log the response and produce a timestamp

https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.management/test-connection?view=powershell-7.4


.\pingMonitor -ip "Ip address of device to monitor" -interval "Interval of pings in seconds"

If no parameter's provided for ip or interval, you will be prompted in cli.

Press Control-C to stop monitor


Output file will go into same directory as script, output location is also outputted in cli.

Oisín Mc Laughlin
07/08/2024
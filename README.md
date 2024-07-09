# Bluetooth Powerpoint Adapter

Used to control powerpoints via a Bluetooth Connection

## Usage

The exe must be in the same Folder as the presentations
Create a Config.ini following this scheme:
> [Bluetooth]
>> uuid= {Your own BluetoothApplicationId}

Then just connect via any Bluetooth capable device to the device running this adapter
and choosing your specified ApplicationId to reach your socket.

Upon establishing a connection you will be given a comma separated string of all available
presentations. Split that String at the separator.

### Commands
- '0:{presentation Index}' -> Open the presentation with the given index
- '\x01' -> goto the next slide
- '\x02' -> goto the previous slide
- '\x03' -> close the connection and the presentation
- '\x04:{slide Index}' -> jump to the slide with the given index 
### Answers
- '0:{comma-separated string}' -> Connection established. The comma separated String lists all available presentations
- '1' -> Done Action. This will be the message to verify that the requested action has been executed
- '2' -> Opened. This Answer signalizes that the requested presentation has been opened
- '3:{Error Code}' -> An Exception has occured. The request has not been completed successfully.
### Error Codes
- '0' -> This Error shows that the Presentation has already been closed or has crashed
- '1' -> This Error shows that the requested Action is not available due to the Slide Show not being present
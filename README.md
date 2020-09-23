<div align="center">

## Telemtry Control


</div>

### Description

This program controls a remote unit across a range. The OFF button sends a "FF" and expects an "FF" within 1/3 of a second otherwise it times out.

The reason I uploaded this is because I could not find a serial program that worked using MSCOMM and ONCOM event.

I did have trouble at first until I read that Microsoft has Rthreshold default as '0'-- it needs to be '1' for the received data event to be captured

I realise this is a very simple program, but it WORKS.

To check this out place a wire/short between pins 2 and 3 on your DB9 serial port. This will route TX to RX on the same port.

I limited it to COM1 and Com2 but others can be added.

I hope this helps someone trying to use the serial port.
 
### More Info
 
Serial port DB9 pins 2 and 3

Place a wire between pins 2 and 3 on your com1 or com2 DB9 port.

data from the serial port


<span>             |<span>
---                |---
**Submitted On**   |2002-02-14 11:17:04
**By**             |[Terry S\. Wade](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/terry-s-wade.md)
**Level**          |Beginner
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Telemtry\_C552952142002\.zip](https://github.com/Planet-Source-Code/terry-s-wade-telemtry-control__1-31799/archive/master.zip)









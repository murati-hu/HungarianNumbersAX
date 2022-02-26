# Hungarian Numbers Reading and Writing

This is an archive project to provide an ActiveX library
to support reading and writing numbers and fractions in
Hungarian as speech and as text.

The project is not maintained since 2003.

## Development Requirements
 - Visual Studio 6
    - Visual Basic 6.0 language support
    - ActiveX support
 - Minimum DirectX 7 support for number to speech support

## Folders and Structure
 - `Szamok_Irasa_ActiveX` - Number to Text project
 - `Szamok_Olvasasa_ActiveX` - Numbers to Speech project

## Installation
Register the build ActiveX components in Windows:
```
regsvr32.exe -i szamok.dll
regsvr32.exe -i olvasas.dll
```

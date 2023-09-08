# AccUnit
**AccUnit - Test your VBA/Access code**

For more informations see [Wiki](https://github.com/AccessCodeLib/AccUnit/wiki) and https://accunit.access-codelib.net/

### Installation
[How-to: Install-AccUnit](https://github.com/AccessCodeLib/AccUnit/wiki/How%E2%80%90to:-Install-AccUnit)

### Notes
* For 32 and 64 Bit Office/Access
* Run COM regfree (with [Access-Add-In](https://github.com/AccessCodeLib/AccUnit/tree/main/access-add-in))

### Dev state
* Simple test can started by [Access-Add-In](https://github.com/AccessCodeLib/AccUnit/tree/main/access-add-in)
* Row test available
* VbaTestSuite with 'Debug.Print'-Output to VBE Immediate window
* Code coverage tests ([Video](https://accunit.access-codelib.net/videos/examples/CodeCoverageTest.mp4))
* Generate test classes with TestClassGenerator (write TestSuite.TestClassGenerator in VBA immediate window) ([Video](https://accunit.access-codelib.net/videos/examples/NW2-UnitTests.mp4
))
* Tag filter: `TestSuite.Add(...).Filter("abc,xyz").Run`
* Test filter: `TestSuite.Add(...).SelectTests("xyz*").Run`

### Remarks
Examples see [./examples/msaccess/](https://github.com/AccessCodeLib/AccUnit/blob/main/examples/msaccess/)

If the DLL files are located in a network drive, it will not work.

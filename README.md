# AccUnit
**AccUnit - Test your Access/Excel VBA code**

For more informations see [Wiki](https://github.com/AccessCodeLib/AccUnit/wiki) and https://accunit.access-codelib.net/

### Installation
[How-to: Install-AccUnit](https://github.com/AccessCodeLib/AccUnit/wiki/How%E2%80%90to:-Install-AccUnit)

### Notes
* For 32 and 64 Bit Office/Access
* Run COM regfree (with [Access Add-In](https://github.com/AccessCodeLib/AccUnit/tree/main/access-add-in) or [Excel Add-In](https://github.com/AccessCodeLib/AccUnit/tree/main/excel-add-in))

### Dev state
* Run tests
  * Tests can started by [Access-Add-In](https://github.com/AccessCodeLib/AccUnit/tree/main/access-add-in) or [Excel Add-In](https://github.com/AccessCodeLib/AccUnit/tree/main/excel-add-in)
  * Tag filter: `TestSuite.Add(...).Filter("abc,xyz").Run`
  * Test filter: `TestSuite.Add(...).SelectTests("xyz*").Run`
* Write tests
  * Row test available
  * Generate test classes with TestClassGenerator (write TestSuite.TestClassGenerator in VBA immediate window) ([Video](https://accunit.access-codelib.net/videos/examples/NW2-UnitTests.mp4))
* Output test results
  * TestSuite with 'Debug.Print' output to VBE Immediate window
  * Output test results to log/text file
* Other features
  * Code coverage tests ([Video](https://accunit.access-codelib.net/videos/examples/CodeCoverageTest.mp4))

### Remarks
Examples see [./examples/msaccess/](https://github.com/AccessCodeLib/AccUnit/blob/main/examples/msaccess/)

If the DLL files are located in a network drive, it will not work.

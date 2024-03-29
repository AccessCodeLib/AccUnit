﻿namespace AccessCodeLib.AccUnit.Configuration
{
    internal class TestClassTemplate
    {
        const string TestProcedureTemplateCode =
            @"Public Sub {MethodUnderTest}_{StateUnderTest}_{ExpectedBehaviour}({Params})
	' Arrange
	Err.Raise vbObjectError, ""{MethodUnderTest}_{StateUnderTest}_{ExpectedBehaviour}"", ""Test not implemented""
	Const Expected As Variant = ""expected value""
	Dim Actual As Variant
	' Act
	Actual = ""actual value""
	' Assert
	Assert.That Actual, Iz.EqualTo(Expected)
End Sub";

    }
}

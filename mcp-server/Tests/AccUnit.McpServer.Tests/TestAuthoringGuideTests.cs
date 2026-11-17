using System;
using System.IO;
using NUnit.Framework;

namespace AccessCodeLib.AccUnit.McpServer.Tests
{
    public class TestAuthoringGuideTests
    {
        private static string ShippedSkillFile
        {
            get { return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "skills", "accunit-tests", "SKILL.md"); }
        }

        [TestCase("skill accunit-tests")]
        [TestCase("accunit_install_skill")]
        [TestCase("not VBA")]
        [TestCase("no constants or enums of the own VBA project")]
        [TestCase("System.Convert.ToDateTime")]
        [TestCase("arrays next to other values as {\"a\", \"b\"}")]
        [TestCase("Array(...) only as the only value of a row")]
        [TestCase("distinguishes vbNullString")]
        [TestCase("do not read the body of the method under test")]
        [TestCase("Assert.Throws 457")]
        [TestCase("start with accunit_generate_test_class")]
        [TestCase("neither the exported source files nor through other tools")]
        public void ServerInstructions_ContainSkillReferenceAndCoreRules(string text)
        {
            Assert.That(TestAuthoringGuide.Text, Does.Contain(text));
        }

        [Test]
        public void Skill_IsShippedNextToTheServer()
        {
            Assert.That(File.Exists(ShippedSkillFile), Is.True, ShippedSkillFile);
        }

        [Test]
        public void Skill_HasPortableFrontmatter()
        {
            var lines = File.ReadAllLines(ShippedSkillFile);

            Assert.That(lines[0], Is.EqualTo("---"));
            Assert.That(lines[1], Is.EqualTo("name: " + TestAuthoringGuide.SkillName));
            Assert.That(lines[2], Does.StartWith("description: "));
            Assert.That(lines[3], Is.EqualTo("---"));
        }

        [TestCase("'AccUnit:Row(2+8,")]
        [TestCase("'AccUnit:Row(2, Null,")]
        [TestCase("System.Convert.ToDateTime(\"2014-05-16\")")]
        [TestCase(".Name(\"Equal\")")]
        [TestCase("not VBA")]
        [TestCase("Constants and enums of the own VBA project are not available")]
        [TestCase("inputs first, expected value last")]
        [TestCase("Arrays next to other values in VB.NET syntax: `'AccUnit:Row(1, {\"a\", \"b\"}, ...)`")]
        [TestCase("`Array(\"a\", \"b\")` only when the array is the only value of the row")]
        [TestCase("Whole numbers arrive as Long")]
        [TestCase("distinguishes `vbNullString`")]
        [TestCase("accunit_run_tests")]
        [TestCase("Do not read the body of the method under test")]
        [TestCase("`Assert.Throws 457`")]
        [TestCase("for expected errors use `Assert.Throws`")]
        [TestCase("Start here: do not read the source of the module under")]
        [TestCase("Use when the user asks to create unit tests for VBA code")]
        public void Skill_ContainsExampleAndRules(string text)
        {
            Assert.That(File.ReadAllText(ShippedSkillFile), Does.Contain(text));
        }
    }
}

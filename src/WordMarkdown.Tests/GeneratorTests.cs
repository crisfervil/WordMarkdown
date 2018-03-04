using System;
using System.IO;
using Xunit;

namespace WordMarkdown.Tests
{
    public class GeneratorTests
    {
        [Fact]
        public void Generates_JSon_For_Example1()
        {
            var expectedJsonFileName = "..\\..\\assets\\Example1.json";
            var fileName = "..\\..\\assets\\Example1.docx";

            var generator = new WordMarkdown.Generator();
            var actualJson = generator.GetJSon(fileName);
            actualJson = actualJson.Replace("\r", "");

            var expectedJson = File.ReadAllText(expectedJsonFileName);
            expectedJson = expectedJson.Replace("\r","");
            Assert.Equal(expectedJson, actualJson);
        }

        [Fact]
        public void Generates_JSon_For_Example2()
        {
            var expectedJsonFileName = "..\\..\\assets\\Example2.json";
            var fileName = "..\\..\\assets\\Example2.docx";

            var generator = new WordMarkdown.Generator();
            var actualJson = generator.GetJSon(fileName);
            actualJson = actualJson.Replace("\r", "");

            var expectedJson = File.ReadAllText(expectedJsonFileName);
            expectedJson = expectedJson.Replace("\r", "");
            Assert.Equal(expectedJson, actualJson);
        }

        [Fact]
        public void Generates_JSon_For_Example3()
        {
            var expectedJsonFileName = "..\\..\\assets\\Example3.json";
            var fileName = "..\\..\\assets\\Example3.docx";

            var generator = new WordMarkdown.Generator();
            var actualJson = generator.GetJSon(fileName);
            actualJson = actualJson.Replace("\r", "");

            var expectedJson = File.ReadAllText(expectedJsonFileName);
            expectedJson = expectedJson.Replace("\r", "");
            Assert.Equal(expectedJson, actualJson);
        }

        [Fact]
        public void Generates_JSon_For_Example4()
        {
            var expectedJsonFileName = "..\\..\\assets\\Example4.json";
            var fileName = "..\\..\\assets\\Example4.docx";

            var generator = new WordMarkdown.Generator();
            var actualJson = generator.GetJSon(fileName);
            actualJson = actualJson.Replace("\r", "");

            var expectedJson = File.ReadAllText(expectedJsonFileName);
            expectedJson = expectedJson.Replace("\r", "");
            Assert.Equal(expectedJson, actualJson);
        }
    }
}

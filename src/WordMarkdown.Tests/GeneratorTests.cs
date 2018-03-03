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
            var epectedJsonFileName = "assets\\Example1.json";
            var fileName = "assets\\Example1.docx";

            var generator = new WordMarkdown.Generator();
            var actualJson = generator.GetJSon(fileName);

            var expectedJson = File.ReadAllText(epectedJsonFileName);
            Assert.Equal(expectedJson,actualJson);
        }
    }
}

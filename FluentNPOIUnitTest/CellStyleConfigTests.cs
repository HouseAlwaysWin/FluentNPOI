using Xunit;
using FluentNPOI;
using NPOI.SS.UserModel;
using System;
using FluentNPOI.Models;

namespace NPOIPlusUnitTest
{
    public class CellStyleConfigTests
    {
        [Fact]
        public void Constructor_ShouldSetProperties()
        {
            // Arrange
            Action<ICellStyle> setter = (style) => { };

            // Act
            var config = new CellStyleConfig("TestKey", setter);

            // Assert
            Assert.Equal("TestKey", config.Key);
            Assert.Equal(setter, config.StyleSetter);
        }

        [Fact]
        public void Deconstruct_ShouldReturnKeyAndSetter()
        {
            // Arrange
            Action<ICellStyle> setter = (style) => { };
            var config = new CellStyleConfig("TestKey", setter);

            // Act
            var (key, styleSetter) = config;

            // Assert
            Assert.Equal("TestKey", key);
            Assert.Equal(setter, styleSetter);
        }
    }
}


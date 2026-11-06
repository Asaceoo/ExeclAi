using System;
using System.IO;
using ExcelAiAssistant.Core.Configuration;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    public sealed class WritePreferenceStoreTests : IDisposable
    {
        private readonly string filePath;
        private readonly WritePreferenceStore store;

        public WritePreferenceStoreTests()
        {
            filePath = Path.Combine(Path.GetTempPath(), $"xai-write-pref-{Guid.NewGuid():N}.json");
            store = new WritePreferenceStore(filePath);
        }

        [Fact]
        public void Load_returns_auto_when_file_absent()
        {
            Assert.Equal("auto", store.Load().ValueMode);
        }

        [Theory]
        [InlineData("auto")]
        [InlineData("formula")]
        [InlineData("value")]
        public void Round_trips_valid_modes(string mode)
        {
            store.Save(new WritePreference { ValueMode = mode });

            Assert.Equal(mode, new WritePreferenceStore(filePath).Load().ValueMode);
        }

        [Fact]
        public void Load_falls_back_to_auto_for_invalid_mode()
        {
            store.Save(new WritePreference { ValueMode = "banana" });

            Assert.Equal("auto", new WritePreferenceStore(filePath).Load().ValueMode);
        }

        [Fact]
        public void Formula_and_Value_factories_build_correct_values()
        {
            Assert.Equal("formula", WritePreference.Formula().ValueMode);
            Assert.Equal("value", WritePreference.Value().ValueMode);
            Assert.True(WritePreference.Auto().IsValid);
            Assert.False(new WritePreference { ValueMode = "x" }.IsValid);
        }

        public void Dispose()
        {
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
        }
    }
}
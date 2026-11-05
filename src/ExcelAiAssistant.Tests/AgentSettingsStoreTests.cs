using System.IO;
using ExcelAiAssistant.Core.Configuration;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    public sealed class AgentSettingsStoreTests
    {
        private readonly string directory;
        private readonly string filePath;

        public AgentSettingsStoreTests()
        {
            directory = Path.Combine(Path.GetTempPath(), $"excel-ai-agentsettings-{System.Guid.NewGuid():N}");
            filePath = Path.Combine(directory, "agent-settings.json");
        }

        [Fact]
        public void Load_returns_default_15_when_file_absent()
        {
            Assert.Equal(10000, new AgentSettingsStore(filePath).Load().MaxToolCallsPerTurn);
        }

        [Theory]
        [InlineData(1)]
        [InlineData(5)]
        [InlineData(20)]
        [InlineData(50)]
        [InlineData(100)]
        [InlineData(1000)]
        [InlineData(10000)]
        public void Round_trips_valid_values(int maxCalls)
        {
            var store = new AgentSettingsStore(filePath);
            store.Save(new AgentSettings { MaxToolCallsPerTurn = maxCalls });
            Assert.Equal(maxCalls, new AgentSettingsStore(filePath).Load().MaxToolCallsPerTurn);
        }

        [Theory]
        [InlineData(0)]
        [InlineData(-3)]
        [InlineData(10001)]
        [InlineData(90000)]
        public void Load_falls_back_to_default_for_invalid_values(int maxCalls)
        {
            new AgentSettingsStore(filePath).Save(new AgentSettings { MaxToolCallsPerTurn = maxCalls });
            Assert.Equal(10000, new AgentSettingsStore(filePath).Load().MaxToolCallsPerTurn);
        }
    }
}
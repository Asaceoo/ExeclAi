using System;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    public sealed class ComBusyCompensationTests
    {
        private static COMException Busy(int hr)
        {
            // Marshal.ThrowExceptionForHR 生成的异常带有真实 HRESULT。
            try
            {
                Marshal.ThrowExceptionForHR(hr);
                throw new InvalidOperationException("expected exception");
            }
            catch (COMException ex)
            {
                return ex;
            }
        }

        private static MethodInfo Execute() => typeof(ExcelToolHost).GetMethod("Execute", BindingFlags.NonPublic | BindingFlags.Static);
        private static MethodInfo IsComBusy() => typeof(ExcelToolHost).GetMethod("IsComBusy", BindingFlags.NonPublic | BindingFlags.Static);

        [Theory]
        [InlineData(-2147417846)] // RPC_E_SERVERCALL_RETRYLATER
        [InlineData(-2147418111)] // RPC_E_CALL_REJECTED
        [InlineData(-2147417851)] // RPC_E_SERVERCALL_REJECTED
        public void Recognizes_com_busy_hrresults(int hr)
        {
            Assert.True((bool)IsComBusy().Invoke(null, new object[] { Busy(hr) }));
        }

        [Fact]
        public void Ignores_non_busy_exceptions()
        {
            Assert.False((bool)IsComBusy().Invoke(null, new object[] { new InvalidOperationException("x") }));
            Assert.False((bool)IsComBusy().Invoke(null, new object[] { Busy(-2147467259) })); // E_FAIL
        }

        [Fact]
        public void Compensates_busy_failure_once_then_succeeds()
        {
            var calls = 0;
            var result = (ToolResult)Execute().Invoke(null, new object[]
            {
                new Func<ToolResult>(() =>
                {
                    calls++;
                    if (calls == 1)
                    {
                        throw Busy(-2147417846);
                    }

                    return new ToolResult { Ok = true, Data = "ok" };
                })
            });

            Assert.True(result.Ok);
            Assert.Equal(2, calls);
        }

        [Fact]
        public void Reports_failure_when_still_busy_after_compensation()
        {
            var result = (ToolResult)Execute().Invoke(null, new object[]
            {
                new Func<ToolResult>(() => { throw Busy(-2147417846); })
            });

            Assert.False(result.Ok);
            Assert.Equal("EXCEL_COM_ERROR", result.ErrorCode);
        }
    }

    public sealed class RangeSamplerTests
    {
        [Fact]
        public void MergeSample_combines_top_and_bottom()
        {
            var merge = typeof(ExcelReadService).GetMethod("MergeSample", BindingFlags.NonPublic | BindingFlags.Static);
            var top = new object[,] { { "a", 1D }, { "b", 2D } };
            var bottom = new object[,] { { "z", 9D } };

            var result = (object[,])merge.Invoke(null, new object[] { top, bottom, 2 });

            Assert.Equal(3, result.GetLength(0));
            Assert.Equal("a", result[0, 0]);
            Assert.Equal(1D, result[0, 1]);
            Assert.Equal("b", result[1, 0]);
            Assert.Equal("z", result[2, 0]);
            Assert.Equal(9D, result[2, 1]);
        }
    }
}
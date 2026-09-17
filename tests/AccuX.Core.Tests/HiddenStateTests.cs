using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using AccuX.Core.Operations;
using AccuX.Host;
using Xunit;

namespace AccuX.Core.Tests
{
    public class HiddenStateTests
    {
        [Theory]
        [InlineData(true, true)]
        [InlineData(true, false)]
        [InlineData(false, true)]
        [InlineData(false, false)]
        public void ReadHiddenState_WhenAggregateIsMisleading_ReadsEachDimension(bool rows, bool aggregate)
        {
            var visited = new List<int>();
            var result = Read(rows, aggregate, new object[] { true, false, true, false }, visited);

            Assert.Equal(new[] { true, false, true, false }, result);
            Assert.Equal(new[] { 7, 8, 9, 10 }, visited);
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void ReadHiddenState_HandlesIntegerBooleans(bool rows)
        {
            var result = Read(rows, null, new object[] { -1, 0, (short)-1, (short)0 }, new List<int>());

            Assert.Equal(new[] { true, false, true, false }, result);
        }

        [Theory]
        [InlineData(true, true)]
        [InlineData(true, false)]
        [InlineData(false, true)]
        [InlineData(false, false)]
        public void ReadHiddenState_HandlesUniformDimensions(bool rows, bool hidden)
        {
            var result = Read(rows, hidden, new object[] { hidden, hidden }, new List<int>());

            Assert.Equal(new[] { hidden, hidden }, result);
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void ReadHiddenState_WhenIndividualValueIsUnknown_Throws(bool rows)
        {
            var exception = Assert.Throws<HostOperationException>(
                () => Read(rows, true, new object[] { true, null }, new List<int>()));

            Assert.Contains("第 8 " + (rows ? "行" : "列"), exception.Message);
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void ReadHiddenState_WhenIndividualReadFails_Throws(bool rows)
        {
            var exception = Assert.Throws<HostOperationException>(
                () => Read(rows, true, new object[] { new COMException("Read failed") }, new List<int>()));

            Assert.Contains("第 7 " + (rows ? "行" : "列"), exception.Message);
            Assert.IsType<COMException>(exception.InnerException);
        }

        private static bool[] Read(bool rows, object aggregate, object[] values, List<int> visited)
        {
            var method = typeof(ExcelHostBase).GetMethod(
                rows ? "ReadHiddenRows" : "ReadHiddenColumns",
                BindingFlags.Static | BindingFlags.NonPublic);
            var worksheetType = method.GetParameters()[0].ParameterType;
            var rangeType = method.GetParameters()[1].ParameterType;
            var combined = Proxy(rangeType, call =>
            {
                Assert.Equal("get_Hidden", call.MethodName);
                return aggregate;
            });
            var range = Proxy(rangeType, call =>
            {
                if (call.MethodName == (rows ? "get_Row" : "get_Column"))
                {
                    return 7;
                }

                Assert.Equal(rows ? "get_EntireRow" : "get_EntireColumn", call.MethodName);
                return combined;
            });
            var dimensions = Proxy(rangeType, call =>
            {
                Assert.Equal("get_Item", call.MethodName);
                var index = Convert.ToInt32(call.Args[0]);
                visited.Add(index);
                return Proxy(rangeType, itemCall =>
                {
                    Assert.Equal("get_Hidden", itemCall.MethodName);
                    var value = values[index - 7];
                    if (value is Exception exception)
                    {
                        throw exception;
                    }

                    return value;
                });
            });
            var worksheet = Proxy(worksheetType, call =>
            {
                Assert.Equal(rows ? "get_Rows" : "get_Columns", call.MethodName);
                return dimensions;
            });
            try
            {
                return (bool[])method.Invoke(null, new object[] { worksheet, range, values.Length });
            }
            catch (TargetInvocationException ex)
            {
                System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(ex.InnerException).Throw();
                throw;
            }
        }

        private static object Proxy(Type type, Func<IMethodCallMessage, object> handler)
        {
            return new InterfaceProxy(type, handler).GetTransparentProxy();
        }

        private sealed class InterfaceProxy : RealProxy
        {
            private readonly Func<IMethodCallMessage, object> _handler;

            public InterfaceProxy(Type type, Func<IMethodCallMessage, object> handler) : base(type)
            {
                _handler = handler;
            }

            public override IMessage Invoke(IMessage message)
            {
                var call = (IMethodCallMessage)message;
                try
                {
                    return new ReturnMessage(_handler(call), null, 0, call.LogicalCallContext, call);
                }
                catch (Exception ex)
                {
                    return new ReturnMessage(ex, call);
                }
            }
        }
    }
}

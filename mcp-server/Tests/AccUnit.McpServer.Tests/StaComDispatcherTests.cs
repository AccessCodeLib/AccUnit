using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using NUnit.Framework;

namespace AccessCodeLib.AccUnit.McpServer.Tests
{
    public class StaComDispatcherTests
    {
        [Test]
        public void Invoke_RunsWorkOnStaThread()
        {
            using (var dispatcher = new StaComDispatcher())
            {
                var apartment = dispatcher.Invoke(() => Thread.CurrentThread.GetApartmentState());

                Assert.That(apartment, Is.EqualTo(ApartmentState.STA));
            }
        }

        [Test]
        public void Invoke_AlwaysUsesTheSameThread()
        {
            using (var dispatcher = new StaComDispatcher())
            {
                var firstThreadId = dispatcher.Invoke(() => Thread.CurrentThread.ManagedThreadId);
                var secondThreadId = dispatcher.Invoke(() => Thread.CurrentThread.ManagedThreadId);

                Assert.That(secondThreadId, Is.EqualTo(firstThreadId));
                Assert.That(firstThreadId, Is.Not.EqualTo(Thread.CurrentThread.ManagedThreadId));
            }
        }

        [Test]
        public void Invoke_ConcurrentCallers_WorkNeverOverlaps()
        {
            using (var dispatcher = new StaComDispatcher())
            {
                var running = 0;
                var maxRunning = 0;

                var callers = Enumerable.Range(0, 8).Select(_ => Task.Run(() => dispatcher.Invoke(() =>
                {
                    var current = Interlocked.Increment(ref running);
                    maxRunning = Math.Max(maxRunning, current);
                    Thread.Sleep(10);
                    Interlocked.Decrement(ref running);
                }))).ToArray();
                Task.WaitAll(callers);

                Assert.That(maxRunning, Is.EqualTo(1));
            }
        }

        [Test]
        public void Invoke_WorkThrows_ExceptionIsPassedToCaller()
        {
            using (var dispatcher = new StaComDispatcher())
            {
                Assert.Throws<InvalidOperationException>(() =>
                    dispatcher.Invoke(() => { throw new InvalidOperationException("expected"); }));

                // The dispatcher keeps working after an error.
                Assert.That(dispatcher.Invoke(() => 42), Is.EqualTo(42));
            }
        }

        [Test]
        public void Invoke_FromDispatcherThread_RunsDirectly()
        {
            using (var dispatcher = new StaComDispatcher())
            {
                var result = dispatcher.Invoke(() => dispatcher.Invoke(() => "nested"));

                Assert.That(result, Is.EqualTo("nested"));
            }
        }

        [Test]
        public void Invoke_AfterDispose_Throws()
        {
            var dispatcher = new StaComDispatcher();
            dispatcher.Dispose();

            Assert.Throws<ObjectDisposedException>(() => dispatcher.Invoke(() => 1));
        }
    }
}

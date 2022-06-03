using System;
using System.Runtime.InteropServices;

namespace SubscribeAndHandleQBEvent
{
	[ComVisible(false)]  // This ComVisibleAttribute is set to false so that TLBEXP and REGASM will not expose it nor COM-register it.
	public class ReferenceCountedObjectBase
	{
		public ReferenceCountedObjectBase()
		{
			Console.WriteLine("ReferenceCountedObjectBase contructor.");
			// We increment the global count of objects.
            SubscribeAndHandleQBEvent.InterlockedIncrementObjectsCount();
		}

		~ReferenceCountedObjectBase()
		{
			Console.WriteLine("ReferenceCountedObjectBase destructor.");
			// We decrement the global count of objects.
            SubscribeAndHandleQBEvent.InterlockedDecrementObjectsCount();
			// We then immediately test to see if we the conditions
			// are right to attempt to terminate this server application.
            SubscribeAndHandleQBEvent.AttemptToTerminateServer();
		}
	}
}

﻿using System.Collections.Generic;

namespace ASPRequestImpersonator
{
	/// <summary>
	/// This doesn't need to be ComVisible as it will only ever be used by managed code
	/// </summary>
	public interface IManagedRequestStringList : IEnumerable<string> { }
}

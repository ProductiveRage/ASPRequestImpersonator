namespace ASPRequestImpersonator
{
	/// <summary>
	/// This doesn't need to be ComVisible as it will only ever be used by managed code
	/// </summary>
	public interface IManagedRequestImpersonator
	{
		/// <summary>
		/// This will never return null
		/// </summary>
		IManagedRequestDictionary Form { get; }

		/// <summary>
		/// This will never return null
		/// </summary>
		IManagedRequestDictionary QueryString { get; }

		/// <summary>
		/// This will never return null
		/// </summary>
		IManagedRequestDictionary ServerVariables { get; }

		/// <summary>
		/// Try to retrieve data from the internal lists - QueryString takes precedence over Form which takes precedence over ServerVariables. An exception will
		/// be thrown for null name argument. Requests with name values that are not present in any list will receive an empty RequestStringList.
		/// </summary>
		IManagedRequestStringList this[string key] { get; }
	}
}

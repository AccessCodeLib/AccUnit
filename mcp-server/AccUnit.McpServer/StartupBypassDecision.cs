namespace AccessCodeLib.AccUnit.McpServer
{
    /// <summary>
    /// Decides whether the Shift key is held while opening a database, to skip AutoExec and the startup form.
    /// </summary>
    public sealed class StartupBypassDecision
    {
        private StartupBypassDecision(bool pressShiftKey, string warning)
        {
            PressShiftKey = pressShiftKey;
            Warning = warning;
        }

        public bool PressShiftKey { get; }

        public string Warning { get; }

        /// <param name="bypassRequested">The client asked to skip the startup code.</param>
        /// <param name="allowBypassKey">Value of the database property AllowBypassKey; null if it could not be read.</param>
        public static StartupBypassDecision Create(bool bypassRequested, bool? allowBypassKey)
        {
            if (!bypassRequested)
                return new StartupBypassDecision(false, null);

            if (allowBypassKey == false)
                return new StartupBypassDecision(false,
                    "The database disables the bypass key (AllowBypassKey = False). AutoExec and the startup form run.");

            if (allowBypassKey == null)
                return new StartupBypassDecision(true,
                    "AllowBypassKey could not be read. The startup code is skipped only if the bypass key is allowed.");

            return new StartupBypassDecision(true, null);
        }
    }
}

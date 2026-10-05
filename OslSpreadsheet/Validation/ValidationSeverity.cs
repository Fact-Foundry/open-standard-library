namespace OslSpreadsheet.Validation
{
    /// <summary>
    /// How serious a validation issue is.
    /// </summary>
    public enum ValidationSeverity
    {
        /// <summary>
        /// The file deviates from the specification in a way most applications tolerate.
        /// </summary>
        Warning = 0,

        /// <summary>
        /// The file is malformed. Applications will refuse to open it, prompt to repair it, or silently lose data.
        /// </summary>
        Error = 1
    }
}

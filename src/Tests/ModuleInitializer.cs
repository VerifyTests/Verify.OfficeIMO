public static class ModuleInitializer
{
    #region enable

    [ModuleInitializer]
    public static void Initialize() =>
        VerifyOfficeIMO.Initialize();

    #endregion

    [ModuleInitializer]
    public static void InitializeOther()
    {
        VerifierSettings.UniqueForTargetFrameworkAndVersion();
        VerifierSettings.InitializePlugins();
    }
}

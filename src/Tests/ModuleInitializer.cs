public static class ModuleInitializer
{
    #region enable

    [ModuleInitializer]
    public static void Initialize() =>
        VerifyOfficeIMO.Initialize();

    #endregion

    #region RegisterOpenXml

    [ModuleInitializer]
    public static void InitializeOpenXml() =>
        VerifyOfficeIMO.RegisterOpenXml();

    #endregion

    #region Fonts

    // Aptos is Office's default font, but is installed only with Office, so a build server renders a
    // fallback. Registering it makes the png the same everywhere.
    [ModuleInitializer]
    public static void InitializeFonts()
    {
        var directory = Path.Combine(ProjectFiles.ProjectDirectory, "..", "Fonts");
        AddFont(directory, "Aptos", "Aptos_400.ttf", OfficeFontStyle.Regular);
        AddFont(directory, "Aptos", "Aptos_400_Italic.ttf", OfficeFontStyle.Italic);
        AddFont(directory, "Aptos", "Aptos_700.ttf", OfficeFontStyle.Bold);
        AddFont(
            directory,
            "Aptos",
            "Aptos_700_Italic.ttf",
            OfficeFontStyle.Bold | OfficeFontStyle.Italic);
        AddFont(directory, "Aptos Display", "Aptos_Display_700.ttf", OfficeFontStyle.Bold);
    }

    static void AddFont(string directory, string family, string file, OfficeFontStyle style) =>
        VerifyOfficeIMO.Fonts.Add(family, File.ReadAllBytes(Path.Combine(directory, file)), style);

    #endregion

    [ModuleInitializer]
    public static void InitializeOther()
    {
        VerifierSettings.UniqueForTargetFrameworkAndVersion();
        VerifierSettings.InitializePlugins();
    }
}

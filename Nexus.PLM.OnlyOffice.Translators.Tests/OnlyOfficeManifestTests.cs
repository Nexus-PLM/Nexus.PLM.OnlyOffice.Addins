using Nexus.PLM.Addin.Sdk.Translators;
using Nexus.PLM.Addin.Sdk.Translators.Testing;
using Xunit;

namespace Nexus.PLM.OnlyOffice.Translators.Tests;

/// <summary>
/// Contract tests for this repo's <c>translators/translators.json</c> — the whole server-side
/// deliverable here, since the repo ships no .NET by design (templates for OOXML live with the
/// Office add-ins; these conversions take the INPUT formats the Office repo does not).
/// </summary>
public class OnlyOfficeManifestTests : ManifestContractTests
{
    /// <inheritdoc/>
    protected override string ManifestPath =>
        Path.Combine(AppContext.BaseDirectory, "translators", TranslatorManifest.FileName);

    /// <summary>The three keys reserved in the Services ownership table.</summary>
    [Fact]
    public void CarriesTheThreeReservedKeys() =>
        Assert.Equal(
            ["pptx_to_pdf", "xlsx_to_csv", "xlsx_to_pdf"],
            Load().Translators.Select(t => t.Key).OrderBy(k => k).ToArray());

    /// <summary>Ownership is by input format: nothing here reads docx (the Office repo's).</summary>
    [Fact]
    public void NothingReadsDocx() =>
        Assert.DoesNotContain(Load().Translators,
            t => t.InputMimeType.Contains("wordprocessingml"));

    /// <summary>
    /// Every entry leans on the basename contract: soffice writes its output named after the
    /// input's basename into --outdir, and the server's runner creates temp input and expected
    /// output with the same basename in the process temp folder — so the outdir must be
    /// exactly ${ENV:TEMP} and the output extension must match the --convert-to filter.
    /// </summary>
    [Fact]
    public void EveryEntryUsesTheSofficeOutdirBasenameContract()
    {
        foreach (var t in Load().Translators)
        {
            Assert.Equal("${ENV:LIBREOFFICE_CMD}", t.Command);
            var args = t.Args;
            Assert.Equal("--headless", args[0]);
            Assert.Equal("--convert-to", args[1]);
            Assert.Equal(t.OutputFileExtension.TrimStart('.'), args[2]);
            Assert.Equal("--outdir", args[3]);
            Assert.Equal("${tempdir}", args[4]);
            Assert.Equal("${input}", args[5]);
            Assert.NotNull(t.InputFileExtension); // soffice picks its import filter from it
        }
    }
}

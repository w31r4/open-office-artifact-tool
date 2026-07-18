using Xunit;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec.Tests;

public sealed class DefaultTemplateLibraryCodecTests
{
    [Theory]
    [InlineData("artifact-template-business-review")]
    [InlineData("artifact-template-market-trends-report")]
    [InlineData("artifact-template-operating-review")]
    [InlineData("artifact-template-project-kickoff")]
    [InlineData("artifact-template-simple-dark-mode")]
    [InlineData("artifact-template-simple-light-mode")]
    [InlineData("artifact-template-team-alignment")]
    public void RetainedPresentationTemplateImportsWithoutAnUnhandledCodecFailure(string templateId)
    {
        var limits = EffectiveCodecLimits.From(null);
        var result = PptxCodec.Import(ReadReference(templateId, ".pptx"), limits);

        Assert.NotNull(result.Artifact.Presentation);
        Assert.NotEmpty(result.Artifact.Presentation.Slides);
        var exported = PptxCodec.Export(result.Artifact, limits);
        var reimported = PptxCodec.Import(exported.File, limits);
        Assert.Equal(result.Artifact.Presentation.Slides.Count, reimported.Artifact.Presentation.Slides.Count);
    }

    [Theory]
    [InlineData("artifact-template-business-review")]
    [InlineData("artifact-template-market-trends-report")]
    [InlineData("artifact-template-operating-review")]
    [InlineData("artifact-template-project-kickoff")]
    [InlineData("artifact-template-simple-dark-mode")]
    [InlineData("artifact-template-simple-light-mode")]
    [InlineData("artifact-template-team-alignment")]
    public void RetainedPresentationTemplateSupportsOneBoundedSourceSlideMetadataEdit(string templateId)
    {
        var limits = EffectiveCodecLimits.From(null);
        var result = PptxCodec.Import(ReadReference(templateId, ".pptx"), limits);
        var slide = result.Artifact.Presentation.Slides[0];
        var requestedName = $"{slide.Name} · OpenChestnut QA";
        slide.Name = requestedName;

        var exported = PptxCodec.Export(result.Artifact, limits);
        var reimported = PptxCodec.Import(exported.File, limits);
        Assert.Equal(requestedName, reimported.Artifact.Presentation.Slides[0].Name);
    }

    [Theory]
    [InlineData("artifact-template-design-report")]
    [InlineData("artifact-template-experiment-analysis")]
    [InlineData("artifact-template-investment-committee-memo")]
    [InlineData("artifact-template-legal-memorandum")]
    [InlineData("artifact-template-minimal-letterhead")]
    [InlineData("artifact-template-strategy-memorandum")]
    [InlineData("artifact-template-system-design")]
    public void RetainedDocumentTemplateRoundTripsThroughTheSourceBoundCodec(string templateId)
    {
        var limits = EffectiveCodecLimits.From(null);
        var result = DocxCodec.Import(ReadReference(templateId, ".docx"), limits);

        Assert.NotNull(result.Artifact.Document);
        Assert.NotEmpty(result.Artifact.Document.Blocks);
        var exported = DocxCodec.Export(result.Artifact, limits);
        var reimported = DocxCodec.Import(exported.File, limits);
        Assert.Equal(result.Artifact.Document.Blocks.Count, reimported.Artifact.Document.Blocks.Count);
    }

    [Theory]
    [InlineData("artifact-template-design-report")]
    [InlineData("artifact-template-experiment-analysis")]
    [InlineData("artifact-template-investment-committee-memo")]
    [InlineData("artifact-template-legal-memorandum")]
    [InlineData("artifact-template-minimal-letterhead")]
    [InlineData("artifact-template-strategy-memorandum")]
    [InlineData("artifact-template-system-design")]
    public void RetainedDocumentTemplateSupportsOneBoundedDocumentSettingEdit(string templateId)
    {
        var limits = EffectiveCodecLimits.From(null);
        var result = DocxCodec.Import(ReadReference(templateId, ".docx"), limits);
        var requested = !result.Artifact.Document.UpdateFields;
        result.Artifact.Document.UpdateFields = requested;

        var exported = DocxCodec.Export(result.Artifact, limits);
        var reimported = DocxCodec.Import(exported.File, limits);
        Assert.Equal(requested, reimported.Artifact.Document.UpdateFields);
    }

    [Theory]
    [InlineData("artifact-template-analytics-dashboard")]
    [InlineData("artifact-template-financial-budget")]
    [InlineData("artifact-template-operating-calendar")]
    [InlineData("artifact-template-project-tracker")]
    [InlineData("artifact-template-sales-pipeline")]
    [InlineData("artifact-template-three-statement-forecast")]
    public void RetainedWorkbookTemplateRoundTripsThroughTheSourceBoundCodec(string templateId)
    {
        var limits = EffectiveCodecLimits.From(null);
        var result = XlsxCodec.Import(ReadReference(templateId, ".xlsx"), limits);

        Assert.NotNull(result.Artifact.Workbook);
        Assert.NotEmpty(result.Artifact.Workbook.Worksheets);
        var exported = XlsxCodec.Export(result.Artifact, limits);
        var reimported = XlsxCodec.Import(exported.File, limits);
        Assert.Equal(result.Artifact.Workbook.Worksheets.Count, reimported.Artifact.Workbook.Worksheets.Count);
    }

    [Theory]
    [InlineData("artifact-template-analytics-dashboard")]
    [InlineData("artifact-template-financial-budget")]
    [InlineData("artifact-template-operating-calendar")]
    [InlineData("artifact-template-project-tracker")]
    [InlineData("artifact-template-sales-pipeline")]
    [InlineData("artifact-template-three-statement-forecast")]
    public void RetainedWorkbookTemplateSupportsOneBoundedTextCellEdit(string templateId)
    {
        var limits = EffectiveCodecLimits.From(null);
        var result = XlsxCodec.Import(ReadReference(templateId, ".xlsx"), limits);
        var cell = result.Artifact.Workbook.Worksheets
            .SelectMany(sheet => sheet.Cells)
            .FirstOrDefault(candidate =>
                candidate.ValueCase == CellArtifact.ValueOneofCase.StringValue &&
                candidate.Formula.Length == 0 &&
                candidate.StringValue.Length > 0);
        Assert.NotNull(cell);
        var marker = " [OpenChestnut QA]";
        cell!.StringValue += marker;

        var exported = XlsxCodec.Export(result.Artifact, limits);
        var reimported = XlsxCodec.Import(exported.File, limits);
        Assert.Contains(reimported.Artifact.Workbook.Worksheets.SelectMany(sheet => sheet.Cells), candidate =>
            candidate.ValueCase == CellArtifact.ValueOneofCase.StringValue &&
            candidate.StringValue.Contains(marker, StringComparison.Ordinal));
    }

    [Fact]
    public void RetainedFinancialBudgetTemplateImportsWithoutAnUnhandledCodecFailure()
    {
        var limits = EffectiveCodecLimits.From(null);
        var result = XlsxCodec.Import(ReadReference("artifact-template-financial-budget", ".xlsx"), limits);

        Assert.NotNull(result.Artifact.Workbook);
        Assert.NotEmpty(result.Artifact.Workbook.Worksheets);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "partial_shared_formula_preserved" &&
            diagnostic.SourcePath == "Op Build" &&
            diagnostic.SourceIdentity == "C24:N24");
        var exported = XlsxCodec.Export(result.Artifact, limits);
        var reimported = XlsxCodec.Import(exported.File, limits);
        Assert.Equal(result.Artifact.Workbook.Worksheets.Count, reimported.Artifact.Workbook.Worksheets.Count);

        var partialFormula = Assert.Single(result.Artifact.Workbook.Worksheets.Single(sheet => sheet.Name == "Op Build").Cells,
            cell => cell.Row == 23 && cell.Column == 2);
        partialFormula.Formula = "=0";
        var rejected = Assert.Throws<CodecException>(() => XlsxCodec.Export(result.Artifact, limits));
        Assert.Equal("unsupported_cell_formula_edit", rejected.Code);
    }

    private static byte[] ReadReference(string templateId, string extension)
    {
        var root = FindRepositoryRoot();
        return File.ReadAllBytes(Path.Combine(root, "skills", "default-template-library", "skills", templateId, "assets", $"reference{extension}"));
    }

    private static string FindRepositoryRoot()
    {
        foreach (var startingDirectory in new[] { AppContext.BaseDirectory, Environment.CurrentDirectory })
        {
            for (var directory = new DirectoryInfo(startingDirectory); directory is not null; directory = directory.Parent)
            {
                if (Directory.Exists(Path.Combine(directory.FullName, "skills", "default-template-library")))
                    return directory.FullName;
            }
        }

        throw new DirectoryNotFoundException("Could not locate the repository default-template-library directory.");
    }
}

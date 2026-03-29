using SpiOpenAccess.Modules.Programming;

namespace SpiOpenAccess.Tests;

public sealed class ProgrammingModuleTests
{
    [Fact]
    public void RunProgram_EvaluatesAssignmentsAndPrintsOutput()
    {
        var module = new ProgrammingModule();
        const string source = """
            LET a = 10
            LET b = 5
            PRINT a + b
            """;

        var result = module.RunProgram(source);

        Assert.Equal(3, result.StatementCount);
        Assert.Equal("15", result.Output.Single());
    }
}

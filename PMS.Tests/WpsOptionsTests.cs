using PMS.Controllers;
using PMS.Models;
using System.Collections.Generic;
using Xunit;

namespace PMS.Tests;

public class WpsOptionsTests
{
    [Fact]
    public void SanitizeWpsOptions_RemovesFreeTextAndDuplicates()
    {
        var inputs = new List<WpsOption>
        {
            new WpsOption { Id = 0, Wps = "CUSTOM-1" },
            new WpsOption { Id = 1, Wps = "WPS-A" },
            new WpsOption { Id = 2, Wps = "wps-a" }, // duplicate by text case-insensitive
            null,
            new WpsOption { Id = -5, Wps = "" },
            new WpsOption { Id = 3, Wps = "WPS-B" }
        };

        var outp = HomeController.SanitizeWpsOptions(inputs);
        Assert.NotNull(outp);
        Assert.Equal(2, outp.Count);
        Assert.Contains(outp, x => x.Wps == "WPS-A" && x.Id == 1);
        Assert.Contains(outp, x => x.Wps == "WPS-B" && x.Id == 3);
    }
}

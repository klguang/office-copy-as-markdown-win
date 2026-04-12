using System.Windows.Forms;

namespace OfficeCopyAsMarkdown.Tests;

public sealed class ApplicationSettingsTests : IDisposable
{
    private readonly string _testDirectory;

    public ApplicationSettingsTests()
    {
        _testDirectory = Path.Combine(Path.GetTempPath(), "OfficeCopyAsMarkdown.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_testDirectory);
    }

    [Fact]
    public void Load_WhenSettingsFileIsMissing_ReturnsDefaultHotkey()
    {
        var service = CreateService();

        var settings = service.Load();

        Assert.Equal(HotkeyGesture.Default, settings.ResolveHotkey());
    }

    [Fact]
    public void Load_WhenSettingsFileContainsInvalidJson_ReturnsDefaultHotkey()
    {
        File.WriteAllText(GetSettingsPath(), "{ invalid json");
        var service = CreateService();

        var settings = service.Load();

        Assert.Equal(HotkeyGesture.Default, settings.ResolveHotkey());
    }

    [Fact]
    public void HotkeySettings_RoundTripSupportedShortcut()
    {
        Assert.True(HotkeyGesture.TryCreate(Keys.Control | Keys.Alt, Keys.M, out var hotkey, out _));

        var serialized = HotkeySettings.FromGesture(hotkey);

        Assert.True(serialized.TryToGesture(out var roundTripped, out var error));
        Assert.Null(error);
        Assert.Equal(hotkey, roundTripped);
    }

    [Fact]
    public void HotkeyGesture_RejectsBareKey()
    {
        var result = HotkeyGesture.TryCreate(Keys.None, Keys.C, out _, out var error);

        Assert.False(result);
        Assert.Equal("Shortcut must include Ctrl, Alt, or Shift.", error);
    }

    [Fact]
    public void HotkeyGesture_AcceptsRepresentativeShortcut()
    {
        var result = HotkeyGesture.TryCreate(Keys.Control | Keys.Shift, Keys.C, out var hotkey, out var error);

        Assert.True(result);
        Assert.Null(error);
        Assert.Equal(Keys.Control | Keys.Shift, hotkey.Modifiers);
        Assert.Equal(Keys.C, hotkey.Key);
    }

    [Fact]
    public void TryUpdateHotkey_WhenRegistrationFails_KeepsPreviousShortcutActive()
    {
        Assert.True(HotkeyGesture.TryCreate(Keys.Control | Keys.Shift, Keys.C, out var original, out _));
        Assert.True(HotkeyGesture.TryCreate(Keys.Control | Keys.Alt, Keys.M, out var failingHotkey, out _));

        var registrar = new FakeHotkeyRegistrar
        {
            FailingHotkey = failingHotkey
        };

        var state = RunInSta(() =>
        {
            using var window = new HotkeyWindow(original, registrar);
            var updated = window.TryUpdateHotkey(failingHotkey, out var error);
            return (updated, error, window.ActiveHotkey);
        });

        Assert.False(state.updated);
        Assert.Equal(original, state.ActiveHotkey);
        Assert.Contains(failingHotkey.DisplayText, state.error);
        Assert.Equal(original, registrar.LastRegisteredHotkey);
    }

    public void Dispose()
    {
        if (Directory.Exists(_testDirectory))
        {
            Directory.Delete(_testDirectory, recursive: true);
        }
    }

    private ApplicationSettingsService CreateService() => new(GetSettingsPath());

    private string GetSettingsPath() => Path.Combine(_testDirectory, "settings.json");

    private static T RunInSta<T>(Func<T> action)
    {
        T result = default!;
        Exception? captured = null;

        var thread = new Thread(() =>
        {
            try
            {
                result = action();
            }
            catch (Exception ex)
            {
                captured = ex;
            }
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();

        if (captured is not null)
        {
            throw new InvalidOperationException("STA test execution failed.", captured);
        }

        return result;
    }

    private sealed class FakeHotkeyRegistrar : IHotkeyRegistrar
    {
        public HotkeyGesture? CurrentHotkey { get; private set; }

        public HotkeyGesture? LastRegisteredHotkey { get; private set; }

        public HotkeyGesture? FailingHotkey { get; init; }

        public bool Register(IntPtr handle, int id, HotkeyGesture hotkey)
        {
            if (FailingHotkey == hotkey)
            {
                return false;
            }

            CurrentHotkey = hotkey;
            LastRegisteredHotkey = hotkey;
            return true;
        }

        public bool Unregister(IntPtr handle, int id)
        {
            CurrentHotkey = null;
            return true;
        }
    }
}

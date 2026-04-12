namespace OfficeCopyAsMarkdown;

internal sealed class SettingsForm : Form
{
    private readonly Label _statusLabel;
    private readonly TextBox _hotkeyTextBox;
    private HotkeyGesture _selectedHotkey;

    public SettingsForm(HotkeyGesture currentHotkey)
    {
        _selectedHotkey = currentHotkey;

        Text = "Settings";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.CenterParent;
        ClientSize = new Size(420, 165);

        var descriptionLabel = new Label
        {
            AutoSize = false,
            Left = 16,
            Top = 16,
            Width = 388,
            Height = 36,
            Text = "Press the shortcut you want to use for converting the current Office selection to Markdown."
        };

        var hotkeyLabel = new Label
        {
            AutoSize = true,
            Left = 16,
            Top = 63,
            Text = "Shortcut"
        };

        _hotkeyTextBox = new TextBox
        {
            Left = 16,
            Top = 84,
            Width = 388,
            ReadOnly = true,
            Text = currentHotkey.DisplayText,
            TabStop = true
        };
        _hotkeyTextBox.KeyDown += OnHotkeyKeyDown;

        _statusLabel = new Label
        {
            AutoSize = false,
            Left = 16,
            Top = 114,
            Width = 388,
            Height = 20,
            Text = "Use Ctrl, Alt, or Shift plus one non-modifier key."
        };

        var saveButton = new Button
        {
            Text = "Save",
            Left = 248,
            Top = 136,
            Width = 75,
            DialogResult = DialogResult.OK
        };
        saveButton.Click += OnSaveClicked;

        var cancelButton = new Button
        {
            Text = "Cancel",
            Left = 329,
            Top = 136,
            Width = 75,
            DialogResult = DialogResult.Cancel
        };

        AcceptButton = saveButton;
        CancelButton = cancelButton;

        Controls.Add(descriptionLabel);
        Controls.Add(hotkeyLabel);
        Controls.Add(_hotkeyTextBox);
        Controls.Add(_statusLabel);
        Controls.Add(saveButton);
        Controls.Add(cancelButton);
    }

    public HotkeyGesture SelectedHotkey => _selectedHotkey;

    protected override void OnShown(EventArgs e)
    {
        base.OnShown(e);
        _hotkeyTextBox.Focus();
        _hotkeyTextBox.SelectAll();
    }

    private void OnHotkeyKeyDown(object? sender, KeyEventArgs e)
    {
        e.SuppressKeyPress = true;
        e.Handled = true;

        if (HotkeyGesture.IsModifierKey(e.KeyCode))
        {
            _statusLabel.Text = "Add a non-modifier key to complete the shortcut.";
            return;
        }

        if (!HotkeyGesture.TryCreate(e.Modifiers, e.KeyCode, out var hotkey, out var error))
        {
            _statusLabel.Text = error ?? "The shortcut is not supported.";
            return;
        }

        _selectedHotkey = hotkey;
        _hotkeyTextBox.Text = hotkey.DisplayText;
        _statusLabel.Text = "The new shortcut will take effect immediately after saving.";
    }

    private void OnSaveClicked(object? sender, EventArgs e)
    {
        if (!HotkeyGesture.TryCreate(_selectedHotkey.Modifiers, _selectedHotkey.Key, out _, out var error))
        {
            MessageBox.Show(
                this,
                error ?? "The shortcut is not supported.",
                "Office Copy as Markdown",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            DialogResult = DialogResult.None;
        }
    }
}

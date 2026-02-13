using System.Diagnostics;
using System.Drawing;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace TJTaskCenterPreview;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new TaskCenterForm());
    }
}

internal sealed record TaskDefinition(
    string Name,
    string Description,
    IReadOnlyList<string> Command,
    string WorkingDirectory
);

internal sealed class TaskCenterForm : Form
{
    private static readonly Color BgApp = ColorTranslator.FromHtml("#0B1017");
    private static readonly Color BgPanel = ColorTranslator.FromHtml("#131C29");
    private static readonly Color BgCard = ColorTranslator.FromHtml("#1A2535");
    private static readonly Color FgPrimary = ColorTranslator.FromHtml("#F5F8FF");
    private static readonly Color FgMuted = ColorTranslator.FromHtml("#93A4BC");
    private static readonly Color Accent = ColorTranslator.FromHtml("#13C6D7");
    private static readonly Color AccentHover = ColorTranslator.FromHtml("#1DB5FF");
    private static readonly Color StatusRunning = ColorTranslator.FromHtml("#F59E0B");
    private static readonly Color Danger = ColorTranslator.FromHtml("#E05454");
    private static readonly Color ConsoleBg = ColorTranslator.FromHtml("#0F1724");
    private static readonly Color ConsoleInfo = ColorTranslator.FromHtml("#8EC9FF");
    private static readonly Color ConsoleSuccess = ColorTranslator.FromHtml("#66E3A6");
    private static readonly Color ConsoleWarn = ColorTranslator.FromHtml("#F2CC71");
    private static readonly Color ConsoleError = ColorTranslator.FromHtml("#FF8F8F");

    private readonly string _projectDir;
    private readonly List<TaskDefinition> _tasks;
    private readonly List<Button> _taskButtons = [];

    private Label _statusValue = null!;
    private Button _stopButton = null!;
    private RichTextBox _output = null!;
    private Process? _runningProcess;

    public TaskCenterForm()
    {
        _projectDir = ResolveProjectDirectory();
        _tasks = LoadTasks(_projectDir);

        Text = "TJ Task Center (C# Preview)";
        MinimumSize = new Size(960, 620);
        Size = new Size(1160, 720);
        StartPosition = FormStartPosition.CenterScreen;
        BackColor = BgApp;
        Font = new Font("Segoe UI", 9f, FontStyle.Regular);

        TrySetAppIcon();
        BuildUi();
        AppendLog("Task center started.");
    }

    private void BuildUi()
    {
        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            BackColor = BgApp,
            Padding = new Padding(14),
            ColumnCount = 1,
            RowCount = 2,
        };
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 116));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        Controls.Add(root);

        root.Controls.Add(BuildHeaderPanel(), 0, 0);
        root.Controls.Add(BuildBodyPanel(), 0, 1);
    }

    private Control BuildHeaderPanel()
    {
        var headerShell = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = BgPanel,
            Padding = new Padding(14, 10, 14, 12),
        };

        var accentBar = new Panel
        {
            Dock = DockStyle.Top,
            Height = 4,
            BackColor = Accent,
        };
        headerShell.Controls.Add(accentBar);

        var content = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = BgPanel,
            Padding = new Padding(0, 8, 0, 0),
        };
        headerShell.Controls.Add(content);

        var logo = BuildLogoControl();
        logo.Location = new Point(0, 6);
        content.Controls.Add(logo);

        var title = new Label
        {
            Text = "TJ Task Center",
            ForeColor = FgPrimary,
            Font = new Font("Segoe UI Semibold", 20f, FontStyle.Bold),
            AutoSize = true,
            Location = new Point(120, 8),
            BackColor = BgPanel,
        };
        content.Controls.Add(title);

        var subtitle = new Label
        {
            Text = "C# dashboard preview with one-click task execution.",
            ForeColor = FgMuted,
            Font = new Font("Segoe UI", 10f, FontStyle.Regular),
            AutoSize = true,
            Location = new Point(122, 50),
            BackColor = BgPanel,
        };
        content.Controls.Add(subtitle);

        var countLabel = new Label
        {
            Text = $"{_tasks.Count} task{(_tasks.Count == 1 ? string.Empty : "s")} loaded",
            ForeColor = FgMuted,
            Font = new Font("Segoe UI", 10f, FontStyle.Regular),
            AutoSize = true,
            BackColor = BgPanel,
            Anchor = AnchorStyles.Top | AnchorStyles.Right,
        };
        content.Controls.Add(countLabel);
        content.Resize += (_, _) =>
        {
            countLabel.Location = new Point(Math.Max(0, content.Width - countLabel.Width - 8), 16);
        };
        content.PerformLayout();

        return headerShell;
    }

    private Control BuildBodyPanel()
    {
        var body = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            BackColor = BgApp,
            Margin = new Padding(0, 12, 0, 0),
            ColumnCount = 2,
            RowCount = 1,
        };
        body.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 340));
        body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        body.Controls.Add(BuildTasksPanel(), 0, 0);
        body.Controls.Add(BuildConsolePanel(), 1, 0);
        return body;
    }

    private Control BuildTasksPanel()
    {
        var container = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = BgPanel,
            Padding = new Padding(12),
            Margin = new Padding(0),
        };

        var heading = new Label
        {
            Text = "Tasks",
            ForeColor = FgPrimary,
            Font = new Font("Segoe UI Semibold", 12f, FontStyle.Bold),
            AutoSize = true,
            BackColor = BgPanel,
            Location = new Point(12, 10),
        };
        container.Controls.Add(heading);

        var hint = new Label
        {
            Text = "Run any script with one click.",
            ForeColor = FgMuted,
            AutoSize = true,
            BackColor = BgPanel,
            Location = new Point(12, 36),
        };
        container.Controls.Add(hint);

        var taskFlow = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            BackColor = BgPanel,
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            AutoScroll = true,
            Padding = new Padding(0, 58, 0, 0),
        };
        container.Controls.Add(taskFlow);

        foreach (var task in _tasks)
        {
            var card = BuildTaskCard(task);
            taskFlow.Controls.Add(card);
        }

        if (_tasks.Count == 0)
        {
            var noTasks = new Label
            {
                Text = "No tasks found. Add .py files or update tasks.json.",
                ForeColor = FgMuted,
                AutoSize = true,
                BackColor = BgPanel,
                Margin = new Padding(8, 2, 8, 0),
            };
            taskFlow.Controls.Add(noTasks);
        }

        taskFlow.Resize += (_, _) =>
        {
            var desiredWidth = Math.Max(220, taskFlow.ClientSize.Width - 28);
            foreach (Control control in taskFlow.Controls)
            {
                if (control is Panel card)
                {
                    card.Width = desiredWidth;
                }
            }
        };

        return container;
    }

    private Control BuildConsolePanel()
    {
        var right = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = BgPanel,
            Padding = new Padding(12),
            Margin = new Padding(12, 0, 0, 0),
        };

        var statusRow = new Panel
        {
            Dock = DockStyle.Top,
            Height = 34,
            BackColor = BgPanel,
        };
        right.Controls.Add(statusRow);

        var statusText = new Label
        {
            Text = "Status:",
            ForeColor = FgMuted,
            AutoSize = true,
            BackColor = BgPanel,
            Location = new Point(0, 8),
        };
        statusRow.Controls.Add(statusText);

        _statusValue = new Label
        {
            Text = "Ready",
            ForeColor = Accent,
            Font = new Font("Segoe UI Semibold", 9f, FontStyle.Bold),
            AutoSize = true,
            BackColor = BgPanel,
            Location = new Point(statusText.Right + 4, 8),
        };
        statusRow.Controls.Add(_statusValue);

        _stopButton = new Button
        {
            Text = "Stop Current Task",
            Enabled = false,
            Size = new Size(154, 28),
            Anchor = AnchorStyles.Top | AnchorStyles.Right,
        };
        StyleDangerButton(_stopButton);
        _stopButton.Click += (_, _) => StopCurrentTask();
        statusRow.Controls.Add(_stopButton);
        statusRow.Resize += (_, _) =>
        {
            _stopButton.Location = new Point(Math.Max(0, statusRow.Width - _stopButton.Width - 4), 2);
        };

        _output = new RichTextBox
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            BorderStyle = BorderStyle.None,
            BackColor = ConsoleBg,
            ForeColor = FgPrimary,
            Font = new Font("Consolas", 10f, FontStyle.Regular),
            Margin = new Padding(0, 8, 0, 0),
        };
        right.Controls.Add(_output);
        _output.BringToFront();

        return right;
    }

    private Panel BuildTaskCard(TaskDefinition task)
    {
        var card = new Panel
        {
            BackColor = BgCard,
            Width = 280,
            Height = 124,
            Margin = new Padding(0, 0, 0, 8),
            Padding = new Padding(10),
        };

        var title = new Label
        {
            Text = task.Name,
            ForeColor = FgPrimary,
            AutoSize = false,
            Height = 22,
            Dock = DockStyle.Top,
            Font = new Font("Segoe UI Semibold", 10f, FontStyle.Bold),
            BackColor = BgCard,
        };
        card.Controls.Add(title);

        var runButton = new Button
        {
            Text = "Run Task",
            Height = 34,
            Dock = DockStyle.Top,
            Margin = new Padding(0, 6, 0, 6),
            Cursor = Cursors.Hand,
        };
        StyleTaskButton(runButton);
        runButton.Click += (_, _) => RunTask(task);
        _taskButtons.Add(runButton);
        card.Controls.Add(runButton);

        var desc = new Label
        {
            Text = string.IsNullOrWhiteSpace(task.Description) ? "Run script task." : task.Description,
            ForeColor = FgMuted,
            AutoSize = false,
            Height = 40,
            Dock = DockStyle.Fill,
            BackColor = BgCard,
        };
        card.Controls.Add(desc);

        // Ensure vertical order: title, button, description.
        card.Controls.SetChildIndex(title, 0);
        card.Controls.SetChildIndex(runButton, 1);
        card.Controls.SetChildIndex(desc, 2);
        return card;
    }

    private void StyleTaskButton(Button button)
    {
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderSize = 0;
        button.BackColor = Accent;
        button.ForeColor = FgPrimary;
        button.Font = new Font("Segoe UI Semibold", 9f, FontStyle.Bold);

        button.MouseEnter += (_, _) =>
        {
            if (button.Enabled)
            {
                button.BackColor = AccentHover;
            }
        };
        button.MouseLeave += (_, _) =>
        {
            button.BackColor = button.Enabled ? Accent : ColorTranslator.FromHtml("#2A3343");
        };
        button.EnabledChanged += (_, _) =>
        {
            button.BackColor = button.Enabled ? Accent : ColorTranslator.FromHtml("#2A3343");
            button.ForeColor = button.Enabled ? FgPrimary : ColorTranslator.FromHtml("#5A6880");
        };
    }

    private void StyleDangerButton(Button button)
    {
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderSize = 0;
        button.BackColor = ColorTranslator.FromHtml("#3B2326");
        button.ForeColor = ColorTranslator.FromHtml("#FFE6E6");
        button.Font = new Font("Segoe UI Semibold", 9f, FontStyle.Bold);

        button.MouseEnter += (_, _) =>
        {
            if (button.Enabled)
            {
                button.BackColor = ColorTranslator.FromHtml("#C64040");
            }
        };
        button.MouseLeave += (_, _) =>
        {
            button.BackColor = button.Enabled ? Danger : ColorTranslator.FromHtml("#2B2E35");
        };
        button.EnabledChanged += (_, _) =>
        {
            button.BackColor = button.Enabled ? Danger : ColorTranslator.FromHtml("#2B2E35");
            button.ForeColor = button.Enabled ? ColorTranslator.FromHtml("#FFE6E6") : ColorTranslator.FromHtml("#626B78");
        };
    }

    private void RunTask(TaskDefinition task)
    {
        if (_runningProcess is { HasExited: false })
        {
            MessageBox.Show("A task is already running. Stop it first.", "TJ Task Center", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var executable = task.Command.Count > 0 ? task.Command[0] : string.Empty;
        if (string.IsNullOrWhiteSpace(executable))
        {
            AppendLog("[ERROR] Task has no executable command.");
            return;
        }

        var args = task.Command.Skip(1).Select(QuoteArg);
        var psi = new ProcessStartInfo
        {
            FileName = executable,
            Arguments = string.Join(" ", args),
            WorkingDirectory = task.WorkingDirectory,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
        };

        try
        {
            _runningProcess = new Process
            {
                StartInfo = psi,
                EnableRaisingEvents = true,
            };
            _runningProcess.OutputDataReceived += (_, e) =>
            {
                if (!string.IsNullOrWhiteSpace(e.Data))
                {
                    AppendLog(e.Data);
                }
            };
            _runningProcess.ErrorDataReceived += (_, e) =>
            {
                if (!string.IsNullOrWhiteSpace(e.Data))
                {
                    AppendLog($"[ERROR] {e.Data}");
                }
            };
            _runningProcess.Exited += (_, _) =>
            {
                var code = _runningProcess?.ExitCode ?? 1;
                AppendLog(code == 0
                    ? $"[DONE] {task.Name} completed successfully."
                    : $"[DONE] {task.Name} exited with code {code}.");
                SetRunningState(false);
            };

            AppendLog(string.Empty);
            AppendLog($"=== {task.Name} ===");
            AppendLog($"Working directory: {task.WorkingDirectory}");
            AppendLog($"Command: {psi.FileName} {psi.Arguments}");
            AppendLog(string.Empty);
            SetRunningState(true);

            _runningProcess.Start();
            _runningProcess.BeginOutputReadLine();
            _runningProcess.BeginErrorReadLine();
        }
        catch (Exception ex)
        {
            AppendLog($"[ERROR] Failed to start task: {ex.Message}");
            SetRunningState(false);
        }
    }

    private void StopCurrentTask()
    {
        if (_runningProcess is null || _runningProcess.HasExited)
        {
            return;
        }

        try
        {
            AppendLog("[INFO] Stop requested. Terminating task...");
            _runningProcess.Kill(entireProcessTree: true);
        }
        catch (Exception ex)
        {
            AppendLog($"[ERROR] Could not stop task: {ex.Message}");
        }
    }

    private void SetRunningState(bool running)
    {
        if (InvokeRequired)
        {
            BeginInvoke(() => SetRunningState(running));
            return;
        }

        foreach (var button in _taskButtons)
        {
            button.Enabled = !running;
        }

        _stopButton.Enabled = running;
        _statusValue.Text = running ? "Running" : "Ready";
        _statusValue.ForeColor = running ? StatusRunning : Accent;
    }

    private void AppendLog(string message)
    {
        if (InvokeRequired)
        {
            BeginInvoke(() => AppendLog(message));
            return;
        }

        var color = SelectLogColor(message);
        var output = message;
        if (!output.EndsWith(Environment.NewLine, StringComparison.Ordinal))
        {
            output += Environment.NewLine;
        }

        _output.SelectionStart = _output.TextLength;
        _output.SelectionLength = 0;
        _output.SelectionColor = color;
        _output.AppendText(output);
        _output.SelectionColor = FgPrimary;
        _output.ScrollToCaret();
    }

    private static Color SelectLogColor(string message)
    {
        if (message.Contains("[ERROR]", StringComparison.OrdinalIgnoreCase))
        {
            return ConsoleError;
        }
        if (message.Contains("[DONE]", StringComparison.OrdinalIgnoreCase))
        {
            return ConsoleSuccess;
        }
        if (message.Contains("[PROGRESS]", StringComparison.OrdinalIgnoreCase))
        {
            return ConsoleWarn;
        }
        if (message.Contains("[INFO]", StringComparison.OrdinalIgnoreCase))
        {
            return ConsoleInfo;
        }
        return FgPrimary;
    }

    private static string QuoteArg(string value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return "\"\"";
        }

        var needsQuotes = value.Any(char.IsWhiteSpace) || value.Contains('"');
        var escaped = value.Replace("\"", "\\\"");
        return needsQuotes ? $"\"{escaped}\"" : escaped;
    }

    private Control BuildLogoControl()
    {
        var logoPath = Path.Combine(_projectDir, "assets", "tj_logo.png");
        if (File.Exists(logoPath))
        {
            try
            {
                using var stream = new FileStream(logoPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var source = Image.FromStream(stream);
                var copy = new Bitmap(source);
                return new PictureBox
                {
                    Image = copy,
                    SizeMode = PictureBoxSizeMode.Zoom,
                    Size = new Size(106, 72),
                    BackColor = BgPanel,
                };
            }
            catch
            {
                // Falls back to text badge below.
            }
        }

        return new Label
        {
            Text = "TJ",
            ForeColor = Accent,
            Font = new Font("Segoe UI Black", 32f, FontStyle.Bold),
            AutoSize = true,
            BackColor = BgPanel,
        };
    }

    private void TrySetAppIcon()
    {
        var iconPath = Path.Combine(_projectDir, "assets", "tj_icon.ico");
        if (!File.Exists(iconPath))
        {
            return;
        }

        try
        {
            Icon = new Icon(iconPath);
        }
        catch
        {
            // Optional icon, safe to ignore if format is invalid.
        }
    }

    private static List<TaskDefinition> LoadTasks(string projectDir)
    {
        var tasks = new List<TaskDefinition>();
        var configPath = Path.Combine(projectDir, "tasks.json");

        if (File.Exists(configPath))
        {
            try
            {
                using var doc = JsonDocument.Parse(File.ReadAllText(configPath));
                if (doc.RootElement.TryGetProperty("tasks", out var tasksElement) &&
                    tasksElement.ValueKind == JsonValueKind.Array)
                {
                    foreach (var taskElement in tasksElement.EnumerateArray())
                    {
                        var parsed = ParseTask(taskElement, projectDir);
                        if (parsed is not null)
                        {
                            tasks.Add(parsed);
                        }
                    }
                }
            }
            catch
            {
                // Falls back to script discovery.
            }
        }

        var configuredScripts = tasks
            .Select(t => FindScriptName(t.Command))
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .Select(s => s!.ToLowerInvariant())
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var pythonExe = ResolvePythonExecutable();
        foreach (var scriptPath in Directory.GetFiles(projectDir, "*.py").OrderBy(Path.GetFileName))
        {
            var scriptName = Path.GetFileName(scriptPath);
            if (string.Equals(scriptName, "TJTaskCenter.py", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }
            if (configuredScripts.Contains(scriptName.ToLowerInvariant()))
            {
                continue;
            }

            tasks.Add(new TaskDefinition(
                Name: PrettyName(Path.GetFileNameWithoutExtension(scriptName)),
                Description: $"Run {scriptName}",
                Command: [pythonExe, scriptName],
                WorkingDirectory: projectDir));
        }

        return tasks;
    }

    private static TaskDefinition? ParseTask(JsonElement taskElement, string projectDir)
    {
        if (taskElement.ValueKind != JsonValueKind.Object)
        {
            return null;
        }

        var name = taskElement.TryGetProperty("name", out var nameProp) ? nameProp.GetString() ?? string.Empty : string.Empty;
        if (string.IsNullOrWhiteSpace(name))
        {
            return null;
        }

        var description = taskElement.TryGetProperty("description", out var descProp)
            ? descProp.GetString() ?? string.Empty
            : string.Empty;

        if (!taskElement.TryGetProperty("command", out var commandProp))
        {
            return null;
        }

        var command = ParseCommand(commandProp, projectDir);
        if (command.Count == 0)
        {
            return null;
        }

        var cwdRaw = taskElement.TryGetProperty("cwd", out var cwdProp)
            ? cwdProp.GetString() ?? "{project_dir}"
            : "{project_dir}";
        var cwdExpanded = ExpandPlaceholders(cwdRaw, projectDir);
        var cwd = Path.IsPathRooted(cwdExpanded)
            ? cwdExpanded
            : Path.GetFullPath(Path.Combine(projectDir, cwdExpanded));
        if (!Directory.Exists(cwd))
        {
            cwd = projectDir;
        }

        return new TaskDefinition(name.Trim(), description.Trim(), command, cwd);
    }

    private static IReadOnlyList<string> ParseCommand(JsonElement commandProp, string projectDir)
    {
        switch (commandProp.ValueKind)
        {
            case JsonValueKind.Array:
            {
                var list = new List<string>();
                foreach (var item in commandProp.EnumerateArray())
                {
                    if (item.ValueKind != JsonValueKind.String)
                    {
                        continue;
                    }
                    var value = item.GetString();
                    if (string.IsNullOrWhiteSpace(value))
                    {
                        continue;
                    }
                    list.Add(ExpandPlaceholders(value, projectDir));
                }
                return list;
            }
            case JsonValueKind.String:
            {
                var shellCommand = ExpandPlaceholders(commandProp.GetString() ?? string.Empty, projectDir);
                if (string.IsNullOrWhiteSpace(shellCommand))
                {
                    return [];
                }
                return ["cmd.exe", "/c", shellCommand];
            }
            default:
                return [];
        }
    }

    private static string FindScriptName(IReadOnlyList<string> command)
    {
        for (var i = command.Count - 1; i >= 0; i--)
        {
            var part = command[i].Trim().Trim('"', '\'');
            if (part.EndsWith(".py", StringComparison.OrdinalIgnoreCase))
            {
                return Path.GetFileName(part);
            }
        }
        return string.Empty;
    }

    private static string PrettyName(string stem)
    {
        var withSpaces = Regex.Replace(stem.Replace("_", " ").Trim(), "(?<!^)([A-Z])", " $1");
        return string.Join(" ", withSpaces.Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Select(word => char.ToUpperInvariant(word[0]) + word[1..].ToLowerInvariant()));
    }

    private static string ExpandPlaceholders(string value, string projectDir)
    {
        var assetsDir = Path.Combine(projectDir, "assets");
        var python = ResolvePythonExecutable();
        return value
            .Replace("{python}", python, StringComparison.OrdinalIgnoreCase)
            .Replace("{project_dir}", projectDir, StringComparison.OrdinalIgnoreCase)
            .Replace("{assets_dir}", assetsDir, StringComparison.OrdinalIgnoreCase);
    }

    private static string ResolvePythonExecutable()
    {
        return CommandExists("py") ? "py" : "python";
    }

    private static bool CommandExists(string commandName)
    {
        try
        {
            using var probe = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "where",
                    Arguments = commandName,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                }
            };
            probe.Start();
            probe.WaitForExit(1500);
            return probe.ExitCode == 0;
        }
        catch
        {
            return false;
        }
    }

    private static string ResolveProjectDirectory()
    {
        var cursor = new DirectoryInfo(AppContext.BaseDirectory);
        while (cursor is not null)
        {
            var hasTasks = File.Exists(Path.Combine(cursor.FullName, "tasks.json"));
            var hasAssets = Directory.Exists(Path.Combine(cursor.FullName, "assets"));
            if (hasTasks || hasAssets)
            {
                return cursor.FullName;
            }
            cursor = cursor.Parent;
        }
        return Environment.CurrentDirectory;
    }
}

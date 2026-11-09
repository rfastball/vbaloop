using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using Microsoft.Win32;

namespace VbaLoop.Updater;

public partial class MainWindow : Window
{
    private enum Screen { Select, Ready, Progress, Done, Error }
    private Screen _screen;
    private string? _package, _workbook, _backupFolder, _resultWorkbook;
    private PackageResponse? _inspection;
    private bool _busy, _adopt, _preview, _uncertain, _allowBackgroundClose;
    private readonly CliClient _client = new();

    public MainWindow(string[] args)
    {
        InitializeComponent();
        Height = Math.Min(720, SystemParameters.WorkArea.Height - 48);
        ShowSelection();
        Loaded += (_, _) =>
        {
            if (args.Length == 2 && args[0] == "--preview") { ShowPreview(args[1]); return; }
            if (args.Length > 0) LoadFiles(args);
        };
    }

    private void ShowScreen(Screen screen, string title, string subtitle, string status)
    {
        _screen = screen;
        Heading.Text = title;
        Subtitle.Text = subtitle;
        StatusLabel.Text = status;
        SelectPanel.Visibility = Visible(screen == Screen.Select);
        SelectionHint.Visibility = Visible(screen == Screen.Select);
        ReadyPanel.Visibility = Visible(screen == Screen.Ready);
        ProgressPanel.Visibility = Visible(screen == Screen.Progress);
        DonePanel.Visibility = Visible(screen == Screen.Done);
        ErrorPanel.Visibility = Visible(screen == Screen.Error);
        BackButton.Visibility = Visible(screen != Screen.Select && screen != Screen.Progress);
        PrimaryButton.IsEnabled = !_busy && !_preview;
        PrimaryButton.Visibility = Visible(screen != Screen.Progress);
        BackButton.IsEnabled = !_busy && !_preview;
    }

    private static Visibility Visible(bool value) => value ? Visibility.Visible : Visibility.Collapsed;

    private void PickPackage(object sender, RoutedEventArgs e)
    {
        if (_busy || _preview) return;
        var dialog = new OpenFileDialog { Title = "업데이트 패키지 선택", Filter = "vbaloop 텍스트 패키지 (*.txt)|*.txt", CheckFileExists = true };
        if (dialog.ShowDialog(this) == true) LoadFiles([dialog.FileName]);
    }

    private void PickWorkbook(object sender, RoutedEventArgs e)
    {
        if (_busy || _preview) return;
        var dialog = new OpenFileDialog { Title = "업데이트할 업무 파일 선택", Filter = "매크로 사용 Excel 파일|*.xlsm;*.xlsb;*.xlam", CheckFileExists = true };
        if (dialog.ShowDialog(this) == true) LoadFiles([dialog.FileName]);
    }

    private void LoadFiles(string[] files)
    {
        if (_busy || _preview) return;
        try
        {
            string? package = null, workbook = null;
            foreach (string file in files)
            {
                string path = Path.GetFullPath(file);
                if (!File.Exists(path)) throw new IOException("선택한 파일을 찾을 수 없습니다.");
                string ext = Path.GetExtension(path).ToLowerInvariant();
                if (ext == ".txt" && package == null) package = path;
                else if (ext is ".xlsm" or ".xlsb" or ".xlam" && workbook == null) workbook = path;
                else throw new IOException("txt 패키지 하나와 매크로 사용 업무 파일 하나를 선택해 주세요.");
            }
            _package = package ?? _package;
            _workbook = workbook ?? _workbook;
            _inspection = null;
            _adopt = false;
            _uncertain = false;
            _backupFolder = null;
            _resultWorkbook = null;
            AdoptCheck.IsChecked = false;
            ShowSelection();
        }
        catch (Exception ex) { ShowError(ex.Message, ex.ToString()); }
    }

    private void ShowSelection()
    {
        ShowScreen(Screen.Select, "업무 파일 업데이트",
            "원본을 보존하고 업데이트된 새 파일을 만듭니다.", "오프라인 업데이트");
        SelectedPackage.Text = _package == null ? "전달받은 TXT 패키지" : Path.GetFileName(_package);
        SelectedWorkbook.Text = _workbook == null ? "업데이트할 업무 파일" : Path.GetFileName(_workbook);
        SelectedPackage.Style = (Style)FindResource(_package == null ? "EmptyFile" : "FileName");
        SelectedWorkbook.Style = (Style)FindResource(_workbook == null ? "EmptyFile" : "FileName");
        SelectedPackage.ToolTip = SelectedPackagePath.ToolTip = _package;
        SelectedWorkbook.ToolTip = SelectedWorkbookPath.ToolTip = _workbook;
        SelectedPackagePath.Text = _package == null ? ".txt" : Path.GetDirectoryName(_package);
        SelectedWorkbookPath.Text = _workbook == null ? ".xlsm · .xlsb · .xlam" : Path.GetDirectoryName(_workbook);
        PackagePickButton.Content = _package == null ? "패키지 선택" : "변경";
        WorkbookPickButton.Content = _workbook == null ? "Excel 파일 선택" : "변경";
        PrimaryButton.Content = "적용 내용 확인";
        PrimaryButton.IsEnabled = !_preview && _package != null && _workbook != null;
        SelectionHint.Text = (_package != null, _workbook != null) switch
        {
            (true, true) => "아직 파일은 변경되지 않습니다.",
            (true, false) => "대상 Excel 파일을 선택해 주세요.",
            (false, true) => "업데이트 패키지를 선택해 주세요.",
            _ => "두 파일을 선택해 주세요.",
        };
    }

    private async Task InspectAsync()
    {
        if (_busy || _package == null || _workbook == null) return;
        _resultWorkbook = null;
        _inspection = null;
        AdoptCheck.IsChecked = false;
        await RunOperationAsync(async progress =>
        {
            var result = await _client.RunAsync(_package, _workbook, true, false, null, progress);
            _adopt = result.ErrorCode == "adoption_required";
            // --plan --adopt는 등록하지 않는다. 확인용 계획을 만든 뒤 체크박스로 적용 동의를 받는다.
            if (_adopt) result = await _client.RunAsync(_package, _workbook, true, true, null, progress);
            _backupFolder = result.Payload?.BackupFolder;
            if (!result.Ok) { ShowResponseError(result); return; }
            if (result.AlreadyApplied) { ShowDone(result, already: true); return; }
            if (!result.CanApply || !result.CopyMode) throw new InvalidDataException("사본 업데이트 계획을 확인하지 못했습니다. 새 업데이터로 다시 검사해 주세요.");
            _inspection = result;
            ShowReady(result.Payload!);
        }, false);
    }

    private async Task ApplyAsync()
    {
        if (_inspection?.CanApply != true || _package == null || _workbook == null || (_adopt && AdoptCheck.IsChecked != true)) return;
        var inspection = _inspection;
        _resultWorkbook = null;
        _inspection = null; // 실패·중복 클릭은 새 검사 없이 재시도할 수 없다.
        await RunOperationAsync(async progress =>
        {
            var result = await _client.RunAsync(_package, _workbook, false, _adopt, inspection, progress);
            _resultWorkbook = result.OutputWorkbook;
            _backupFolder = result.Payload?.BackupFolder ?? _backupFolder;
            if (result.Applied || result.AlreadyApplied) ShowDone(result, result.AlreadyApplied);
            else if (!result.Ok) ShowResponseError(result);
            else ShowError("결과 사본의 저장을 확인하지 못했습니다. 원본은 변경하지 않았습니다.", "Unexpected apply result", true);
        }, true);
    }

    private async Task RunOperationAsync(Func<IProgress<string>, Task> operation, bool applying)
    {
        _busy = true;
        _allowBackgroundClose = false;
        ShowScreen(Screen.Progress, applying ? "업데이트하고 있습니다" : "적용 내용을 확인하고 있습니다",
            applying ? "작업 사본에 코드를 적용하고 새 결과 파일로 저장합니다. 원본은 보존됩니다." : "작업 사본에서 파일 구성과 버전을 확인합니다.", applying ? "업데이트 진행 중" : "파일 검사 중");
        ProgressSteps.Text = "";
        var progress = new Progress<string>(phase =>
        {
            if (phase == "slow_saving")
            {
                _allowBackgroundClose = true;
                Subtitle.Text = "Excel의 저장 응답이 지연되고 있습니다. 창을 닫아도 저장 작업은 백그라운드에서 계속됩니다.";
            }
            ProgressLabel.Text = phase switch
            {
                "preparing" => "실행 도구를 준비하고 있습니다",
                "checking" => "패키지를 검사하고 있습니다",
                "opening" => "업무 파일을 확인하고 있습니다",
                "backing_up" => "원본 파일을 백업하고 있습니다",
                "applying" => "코드와 폼을 적용하고 있습니다",
                "compiling" => "컴파일을 확인하고 있습니다",
                "saving" => "업데이트된 새 파일을 저장하고 있습니다",
                "saved" => "저장을 완료하고 정리하고 있습니다",
                "slow_saving" => "Excel의 저장 응답을 기다리고 있습니다",
                _ => "작업 중입니다",
            };
            ProgressSteps.Text = phase switch
            {
                "applying" => "원본 백업 완료\n코드 적용 중\n컴파일 확인 대기",
                "compiling" => "원본 백업 완료\n코드 적용 완료\n컴파일 확인 중",
                "saving" or "saved" => "원본 백업 완료\n코드 적용 완료\n컴파일 확인 완료",
                _ => "",
            };
        });
        try { await operation(progress); }
        catch (Exception ex)
        {
            ShowError(applying ? "결과 사본의 저장을 확인하지 못했습니다. 원본은 변경하지 않았습니다. 진단 로그를 저장해 주세요." : ex.Message, ex.ToString(), applying);
        }
        finally
        {
            _busy = false;
            PrimaryButton.IsEnabled = !_preview && (_screen != Screen.Ready || !_adopt || AdoptCheck.IsChecked == true);
            BackButton.IsEnabled = !_preview;
        }
    }

    private void ShowReady(PackageInfo info)
    {
        ShowScreen(Screen.Ready, info.Target + " 업데이트", "적용할 파일과 버전을 확인해 주세요.", "업데이트 준비 완료");
        FromVersion.Text = info.PreviousVersion ?? "첫 적용";
        ToVersion.Text = info.Version;
        WorkbookName.Text = Path.GetFileName(_workbook);
        WorkbookPath.Text = _workbook;
        WorkbookPath.ToolTip = _workbook;
        PackageName.Text = Path.GetFileName(_package);
        PackageName.ToolTip = _package;
        int added = info.Plan.Count(x => x.Action is "add" or "add form");
        int removed = info.Plan.Count(x => x.Action == "remove");
        int replaced = info.Plan.Count(x => x.Action is "replace" or "replace document" or "reimport form");
        ChangeSummary.Text = $"변경 내용   교체 {replaced} · 추가 {added} · 삭제 {removed}";
        Changes.Text = string.Join(Environment.NewLine, info.Plan.Select(x => $"{ActionLabel(x.Action)}   {x.Target}"));
        var warnings = new List<string>();
        if (info.Decision == "reapplied") warnings.Add("현장에서 수정한 코드를 같은 패키지의 코드로 되돌립니다.");
        if (removed > 0) warnings.Add($"패키지에 없는 모듈 {removed}개가 삭제됩니다. 변경 내용을 확인해 주세요.");
        if (info.Warnings.Count > 0) warnings.Add("업무 파일에만 있는 시트 코드는 유지합니다. 상세 변경 내용을 확인해 주세요.");
        WarningText.Text = string.Join("\n", warnings);
        WarningText.Visibility = Visible(warnings.Count > 0);
        AdoptText.Text = $"이 파일을 ‘{info.Target}’ 업무 파일로 등록하고 처음 업데이트하는 데 동의합니다.";
        AdoptCheck.Visibility = Visible(_adopt);
        PrimaryButton.Content = "업데이트 사본 만들기";
        PrimaryButton.IsEnabled = !_adopt && !_preview;
    }

    private static string ActionLabel(string action) => action switch
    {
        "add" => "추가", "add form" => "폼 추가", "remove" => "삭제", "replace" => "교체",
        "replace document" => "문서 코드 교체", "reimport form" => "폼 교체", "leave sheet" => "시트 유지", _ => action,
    };

    private void ShowDone(PackageResponse response, bool already)
    {
        _resultWorkbook = already ? null : response.OutputWorkbook;
        ShowScreen(Screen.Done, already ? "이미 적용된 업데이트입니다" : "업데이트된 새 파일을 만들었습니다",
            already ? "업무 파일의 코드가 패키지와 같습니다. 파일을 변경하지 않았습니다." : "원본은 그대로 보존됩니다. 아래 결과 파일을 사용하세요.", "업데이트 완료");
        DoneDescription.Text = $"{response.Payload?.Target}  {response.Payload?.Version}\n{(already ? _workbook : _resultWorkbook)}";
        ResultFolderButton.Visibility = Visible(!already && _resultWorkbook != null);
        BackupDescription.Text = already ? "새 파일을 만들지 않았습니다." : "결과 폴더에서 파일을 원하는 업무 위치로 복사할 수 있습니다. 상대 경로를 사용하는 매크로·외부 연결은 새 위치에서 확인해 주세요.";
        PrimaryButton.Content = already ? "업무 파일 열기" : "업데이트된 파일 열기";
    }

    private void ShowResponseError(PackageResponse response)
    {
        _resultWorkbook = response.OutputWorkbook;
        ShowError(response.UserMessage, string.Join("\n", response.Errors.Append(response.Hint ?? "")), response.SaveState != "not_saved");
    }

    private void ShowError(string message, string details, bool uncertain = false)
    {
        _inspection = null;
        _uncertain = uncertain;
        ShowScreen(Screen.Error, uncertain ? "결과 사본을 확인해 주세요" : "업데이트를 진행할 수 없습니다",
            "아래 안내를 확인해 주세요.", "확인이 필요합니다");
        ErrorMessage.Text = message;
        ErrorDetails.Text = details;
        ErrorBackupButton.Visibility = Visible(_backupFolder != null && Directory.Exists(_backupFolder));
        PrimaryButton.Content = uncertain ? "결과 폴더 열기" : _package != null && _workbook != null ? "다시 검사" : "파일 선택";
    }

    private async void Primary(object sender, RoutedEventArgs e)
    {
        if (_busy || _preview) return;
        switch (_screen)
        {
            case Screen.Select:
                await InspectAsync();
                break;
            case Screen.Ready: await ApplyAsync(); break;
            case Screen.Done: OpenPath(_resultWorkbook ?? _workbook); break;
            case Screen.Error:
                if (_uncertain) OpenResultFolder(sender, e);
                else if (_package != null && _workbook != null) await InspectAsync();
                else ShowSelection();
                break;
        }
    }

    private void Back(object sender, RoutedEventArgs e)
    {
        if (_busy || _preview) return;
        _package = _workbook = _backupFolder = _resultWorkbook = null;
        _inspection = null;
        _adopt = _uncertain = false;
        ShowSelection();
    }

    private void AdoptionChanged(object sender, RoutedEventArgs e)
    {
        if (PrimaryButton != null && _screen == Screen.Ready) PrimaryButton.IsEnabled = !_busy && !_preview && (!_adopt || AdoptCheck.IsChecked == true);
    }
    private void OpenBackup(object sender, RoutedEventArgs e) => OpenPath(_backupFolder);
    private void OpenResultFolder(object sender, RoutedEventArgs e)
        => OpenPath(_resultWorkbook == null
            ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "vbaloop", "updated")
            : Path.GetDirectoryName(_resultWorkbook));
    private void OpenPath(string? path)
    {
        if (_preview || path == null) return;
        try { Process.Start(new ProcessStartInfo(path) { UseShellExecute = true }); }
        catch (Exception ex) { MessageBox.Show(this, ex.Message, "파일을 열 수 없습니다", MessageBoxButton.OK, MessageBoxImage.Information); }
    }
    private void CopyError(object sender, RoutedEventArgs e)
    {
        try { Clipboard.SetText(ErrorMessage.Text + "\n\n" + ErrorDetails.Text); }
        catch { MessageBox.Show(this, "클립보드를 사용할 수 없습니다. 잠시 후 다시 시도해 주세요."); }
    }

    private async void SaveDiagnostics(object sender, RoutedEventArgs e)
    {
        if (_preview) return;
        var dialog = new SaveFileDialog { Title = "진단 로그 저장", Filter = "진단 로그 ZIP (*.zip)|*.zip",
            FileName = "vbaloop-diagnostics-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".zip", AddExtension = true };
        if (dialog.ShowDialog(this) != true) return;
        DiagnosticsButton.IsEnabled = false;
        try
        {
            await Task.Run(() => _client.ExportDiagnostics(dialog.FileName));
            MessageBox.Show(this, "진단 로그를 저장했습니다. 이 ZIP 파일을 전달해 주세요.\n진행 중에 저장했다면 오류가 난 뒤 한 번 더 저장해 주세요.",
                "진단 로그 저장", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex) { MessageBox.Show(this, ex.Message, "진단 로그를 저장할 수 없습니다", MessageBoxButton.OK, MessageBoxImage.Information); }
        finally { DiagnosticsButton.IsEnabled = true; }
    }
    private void OnDragOver(object sender, DragEventArgs e)
    {
        e.Effects = !_busy && !_preview && e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }
    private void OnDrop(object sender, DragEventArgs e)
    {
        if (!_busy && e.Data.GetData(DataFormats.FileDrop) is string[] paths) LoadFiles(paths);
    }
    private void OnClosing(object? sender, CancelEventArgs e)
    {
        if (!_busy) return;
        if (_allowBackgroundClose && MessageBox.Show(this,
            "창을 닫아도 결과 사본의 저장 작업은 계속됩니다. 원본은 보존됩니다. 다시 실행하기 전에 결과 폴더와 진단 로그를 확인해 주세요. 창을 닫을까요?",
            "저장 응답 대기 중", MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.Yes) return;
        e.Cancel = true;
        if (_allowBackgroundClose) return;
        MessageBox.Show(this, "작업 중에는 창을 닫을 수 없습니다. 완료될 때까지 기다려 주세요.", "업데이트 진행 중", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    // 실제 파일/CLI를 사용하지 않는 시각 검토 경로. 모든 쓰기 동작을 비활성화한다.
    private void ShowPreview(string state)
    {
        _preview = true;
        PreviewLabel.Visibility = Visibility.Visible;
        _workbook = @"D:\업무\회계\2026년 지출관리.xlsm";
        _package = @"D:\업데이트\지출관리-1.3.0.txt";
        var info = new PackageInfo { Target = "지출관리", Version = "1.3.0", PreviousVersion = "1.2.0", Decision = "applied",
            Plan = [new("modMain", "replace"), new("modReport", "replace"), new("Sheet1", "replace document"), new("frmEntry", "reimport form"), new("modExport", "add")] };
        if (state == "ready") ShowReady(info);
        else if (state == "done") ShowDone(new PackageResponse { Payload = info, CopyMode = true,
            OutputWorkbook = @"C:\Users\사용자\AppData\Local\vbaloop\updated\20261104-160000-example\2026년 지출관리.xlsm" }, false);
        else if (state == "progress")
        {
            ShowScreen(Screen.Progress, "업데이트하고 있습니다", "원본을 백업한 뒤 코드를 적용하고 컴파일을 확인합니다.", "업데이트 진행 중");
            ProgressLabel.Text = "코드와 폼을 적용하고 있습니다";
            ProgressSteps.Text = "원본 백업 완료\n코드 적용 중\n컴파일 확인 대기";
        }
        else if (state == "error") ShowError("업무 파일이 열려 있습니다. Excel에서 닫은 뒤 다시 검사해 주세요.", "file_locked — preview");
        else
        {
            _package = _workbook = null;
            ShowSelection();
        }
    }
}

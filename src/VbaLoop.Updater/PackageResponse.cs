using System.IO;
using System.Text.Json;

namespace VbaLoop.Updater;

public sealed record ModuleChange(string Target, string Action);
public sealed class PackageInfo
{
    public string Target { get; set; } = "";
    public string Version { get; set; } = "";
    public string? PreviousVersion { get; set; }
    public string Decision { get; set; } = "";
    public bool Adopting { get; set; }
    public bool Saved { get; set; }
    public string? BackupFolder { get; set; }
    public List<ModuleChange> Plan { get; set; } = [];
    public List<string> Warnings { get; set; } = [];
}

public sealed class PackageResponse
{
    public int Protocol { get; set; }
    public string Command { get; set; } = "";
    public bool Ok { get; set; }
    public int ExitCode { get; set; }
    public string? ErrorCode { get; set; }
    public List<string> Errors { get; set; } = [];
    public string? Hint { get; set; }
    public string SaveState { get; set; } = "unknown";
    public string? PackageHash { get; set; }
    public string? WorkbookHash { get; set; }
    public PackageInfo? Payload { get; set; }

    public bool AlreadyApplied => Ok && Payload?.Decision == "alreadyApplied";
    public bool CanApply => Ok && SaveState == "not_saved" && Payload?.Decision is "applied" or "reapplied"
        && ValidHash(PackageHash) && ValidHash(WorkbookHash);
    public bool Applied => Ok && SaveState == "saved" && Payload?.Saved == true;
    private static bool ValidHash(string? hash) => hash?.Length == 64 && hash.All(Uri.IsHexDigit);

    public static PackageResponse Parse(string json, int processExitCode)
    {
        var result = JsonSerializer.Deserialize<PackageResponse>(json, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
        if (result == null || result.Protocol != 1 || result.Command != "package" || result.ExitCode != processExitCode ||
            result.Ok != (result.ExitCode == 0) || result.SaveState is not ("not_saved" or "saved" or "unknown"))
            throw new InvalidDataException("실행 결과를 확인할 수 없습니다. 업무 파일 상태를 확인한 뒤 다시 검사해 주세요.");
        return result;
    }

    public string UserMessage => SaveState == "unknown"
        ? "저장 결과를 확인할 수 없습니다. 업무 파일과 백업을 확인해 주세요. 자동으로 다시 적용하지 않습니다."
        : SaveState == "saved" && !Ok
        ? "파일은 저장되었지만 마무리 중 문제가 발생했습니다. 업무 파일을 확인해 주세요."
        : ErrorCode switch
        {
            "adoption_required" => "이 파일은 처음 업데이트하는 업무 파일입니다. 아래 내용을 확인해 주세요.",
            "target_mismatch" => "다른 업무용 패키지입니다. 이 업무 파일에 맞는 업데이트 파일을 선택해 주세요.",
            "downgrade" => "현재 업무 파일보다 이전 버전입니다. 더 최신 패키지를 선택해 주세요.",
            "invalid_version" => "업무 파일의 버전을 비교할 수 없습니다. 배포 담당자에게 확인해 주세요.",
            "version_conflict" => "같은 버전 번호에 서로 다른 코드가 있습니다. 버전을 올린 패키지를 받아 주세요.",
            "invalid_package" => "업데이트 파일이 손상되었거나 지원하지 않는 형식입니다. 원본 txt 패키지를 다시 받아 주세요.",
            "files_changed" => "검사 후 파일이 변경되었습니다. 다시 검사하고 적용 내용을 확인해 주세요.",
            "file_locked" => "업무 파일이 열려 있습니다. Excel에서 닫은 뒤 다시 검사해 주세요.",
            "read_only" => "업무 파일이 읽기 전용입니다. 파일 속성을 확인한 뒤 다시 검사해 주세요.",
            "no_write_access" => "업무 파일에 쓸 수 없습니다. 쓰기 권한이 있는 위치를 선택해 주세요.",
            "excel_missing" => "데스크톱 Excel이 필요합니다. Excel 설치 여부를 확인해 주세요.",
            "vbom_denied" => "VBA 프로젝트에 접근할 수 없습니다. Excel의 VBA 프로젝트 접근 허용 설정이나 프로젝트 잠금을 확인해 주세요.",
            "compile_failed" => "컴파일 완료를 확인하지 못해 저장하지 않았습니다. 오류 상세를 배포 담당자에게 전달해 주세요.",
            "backup_failed" => "원본 백업을 만들지 못해 중단했습니다. 저장 공간과 폴더 권한을 확인해 주세요.",
            "encoding_mismatch" => "이 PC에서 저장할 수 없는 문자가 포함되어 있습니다. 배포 담당자에게 확인해 주세요.",
            "incompatible_workbook" => "패키지와 업무 파일의 구성에 차이가 있습니다. 올바른 업무 파일인지 확인해 주세요.",
            "excel_disconnected" => "Excel 연결이 끊겼습니다. 업무 파일을 확인한 뒤 다시 검사해 주세요.",
            "operation_timeout" => "Excel이 오랫동안 응답하지 않아 중단했습니다. 열려 있는 Excel 창과 업무 파일을 확인한 뒤 다시 검사해 주세요.",
            _ => "작업을 완료하지 못했습니다. 오류 상세와 파일 상태를 확인한 뒤 다시 검사해 주세요.",
        };
}

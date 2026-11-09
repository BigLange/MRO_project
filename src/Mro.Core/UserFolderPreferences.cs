using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace Mro.Core;

// Field names recovered from the original record and store, with atomic writes added.
public sealed record UserFolderPreferences(string? MainStandardWorkbookPath = null, string? MainStandardWorkbookDirectory = null,
    string? MainWeighingWorkbookPath = null, string? MainWeighingWorkbookDirectory = null, string? MainOutputDirectory = null,
    string? HelperEntryWorkbookDirectory = null, string? HelperOutputDirectory = null,
    string? MainProductionWorkbookDirectory = null);

public sealed class UserFolderPreferenceStore(string storePath)
{
    public string? LoadWarning { get; private set; }
    public UserFolderPreferences Load()
    {
        LoadWarning = null;
        if (!File.Exists(storePath)) return new();
        try { return JsonSerializer.Deserialize<UserFolderPreferences>(File.ReadAllText(storePath), DataFiles.Json) ?? new(); }
        catch (JsonException) { LoadWarning = "路径偏好文件损坏，已使用默认路径；保存前将备份损坏文件。"; return new(); }
    }
    public UserFolderPreferences Update(Func<UserFolderPreferences, UserFolderPreferences> update)
    {
        var key = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(Path.GetFullPath(storePath))));
        using var mutex = new Mutex(false, "MRO_Recovered_Preferences_" + key);
        var acquired = false;
        try
        {
            try { acquired = mutex.WaitOne(TimeSpan.FromSeconds(10)); } catch (AbandonedMutexException) { acquired = true; }
            if (!acquired) throw new IOException("路径偏好正被另一程序使用，请稍后重试。");
            var old = Load();
            if (LoadWarning is not null) File.Copy(storePath, storePath + ".corrupt-" + Guid.NewGuid().ToString("N"));
            var current = update(old);
            DataFiles.SaveJson(storePath, current);
            return current;
        }
        finally { if (acquired) mutex.ReleaseMutex(); }
    }
    public static string? GetExistingFile(string? path) => !string.IsNullOrWhiteSpace(path) && File.Exists(path) ? path : null;
    public static string? GetExistingDirectory(string? path) => !string.IsNullOrWhiteSpace(path) && Directory.Exists(path) ? path : null;
}

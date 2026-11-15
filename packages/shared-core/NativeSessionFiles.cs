using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml;
using Microsoft.Win32.SafeHandles;

namespace OfficeInEditor.Core;

/// <summary>
/// Owns bounded working/recovery files. This class never writes, replaces, or deletes
/// a source document. Office performs explicit source saves; failures retain recovery.
/// </summary>
internal sealed class NativeSessionFiles
{
    private const int SessionLimit = 16;
    private const long ByteLimit = 128L * 1024 * 1024;
    private const long MetadataAllowance = 4096;
    private static readonly object Gate = new();
    private static readonly Dictionary<string, long> Reservations = new(StringComparer.OrdinalIgnoreCase);
    private static readonly HashSet<string> ActiveDirectories = new(StringComparer.OrdinalIgnoreCase);
    private static readonly Regex SessionPattern = new("^ole-[a-fA-F0-9]{32}$", RegexOptions.CultureInvariant);
    private static readonly Regex BackupPattern = new("^b-([a-f0-9]{32})-([a-f0-9]{32})$", RegexOptions.CultureInvariant);
    private readonly string root, directory, extension, sourceKey;
    private SourceVersion version;
    private bool pendingSave, recoveryNeeded, cleaned;

    private sealed record SourceVersion(uint Volume, ulong FileId, long Length, long LastWrite, string Hash);
    private sealed record BackupMetadata(int Schema, string SourceKey, string Extension, SourceVersion SourceVersion, long Length, string Hash);
    private NativeSessionFiles(string source, string root, string directory, string extension, SourceVersion version)
    {
        SourcePath = source; this.root = root; this.directory = directory; this.extension = extension;
        this.version = version; sourceKey = PathKey(source);
        WorkingPath = Path.Combine(directory, "working" + extension);
    }

    internal string SourcePath { get; }
    internal string WorkingPath { get; }
    internal string PreviousPath => Path.Combine(directory, "previous" + extension);
    private static string DefaultRoot => Path.Combine(RuntimeIdentity.DataRoot, "NativeSessions");

    // dataRoot is an isolated filesystem root for non-Office tests; production callers omit it.
    internal static NativeSessionFiles Create(string sourcePath, string sessionId, string? backupId = null, string? dataRoot = null)
    {
        if (!SessionPattern.IsMatch(sessionId)) throw Fault("invalid_session", "Invalid native session identity.");
        var source = LocalPath(sourcePath);
        var ext = Path.GetExtension(source).ToLowerInvariant();
        if (!Regex.IsMatch(ext, "^\\.[a-z0-9]{1,12}$", RegexOptions.CultureInvariant))
            throw Fault("unsupported_format", "The document extension is invalid.");
        var root = LocalPath(dataRoot ?? DefaultRoot);
        lock (Gate)
        {
            PlainPath(root);
            Directory.CreateDirectory(root);
            PlainPath(root);
            var sourceVersion = Inspect(source, validateWord: ext == ".docx");
            var directory = Path.Combine(root, "s-" + sessionId[4..].ToLowerInvariant());
            if (Directory.Exists(directory) || SessionDirectories(root).Count >= SessionLimit)
                throw Fault("recovery_limit", "Native session recovery capacity is full.");
            string copyFrom = source;
            SourceVersion copyVersion = sourceVersion;
            if (backupId is not null)
            {
                var backup = ReadBackup(root, backupId);
                if (backup.Metadata.SourceKey != PathKey(source) || backup.Metadata.SourceVersion != sourceVersion || backup.Metadata.Extension != ext)
                    throw Fault("document_conflict", "The source no longer matches this recovery snapshot.");
                copyFrom = backup.Snapshot;
                copyVersion = Inspect(copyFrom, validateWord: ext == ".docx");
                if (copyVersion.Length != backup.Metadata.Length || copyVersion.Hash != backup.Metadata.Hash)
                    throw Fault("invalid_backup", "The recovery snapshot has changed.");
            }
            EnsureBudget(root, copyVersion.Length + MetadataAllowance);
            Directory.CreateDirectory(directory);
            var files = new NativeSessionFiles(source, root, directory, ext, sourceVersion);
            try
            {
                files.WriteMetadata(Path.Combine(directory, "session.json"), new { schema = 1, sourceKey = files.sourceKey, extension = ext, sourceVersion });
                CopyChecked(copyFrom, files.WorkingPath, copyVersion);
                ActiveDirectories.Add(directory);
                return files;
            }
            catch
            {
                // Only a not-yet-exposed initial working copy and marker are removed.
                DeleteOwned(root, files.WorkingPath);
                DeleteOwned(root, Path.Combine(directory, "session.json"));
                DeleteEmptyDirectory(root, directory);
                throw;
            }
        }
    }

    internal void VerifySource()
    {
        lock (Gate)
        {
            Alive();
            if (Inspect(SourcePath) != version)
                throw Fault("document_conflict", "The source changed outside this native session.");
        }
    }

    internal string PrepareSave()
    {
        lock (Gate)
        {
            Alive();
            if (pendingSave || recoveryNeeded || File.Exists(PreviousPath))
                throw Fault("document_conflict", "An earlier save still requires recovery review.");
            EnsureBudget(root, version.Length);
            VerifySource();
            CopyChecked(SourcePath, PreviousPath, version);
            pendingSave = true;
            return PreviousPath;
        }
    }

    internal void ConfirmSave()
    {
        lock (Gate)
        {
            Alive();
            if (!pendingSave) throw Fault("invalid_save_state", "No prepared save can be confirmed.");
            // Inspection and XML validation use one read handle, denying concurrent
            // writes/deletes while the resulting source version is captured.
            var saved = Inspect(SourcePath, validateWord: extension == ".docx");
            version = saved;
            pendingSave = false;
            recoveryNeeded = false;
            DeleteOwned(root, PreviousPath);
        }
    }

    internal void MarkRecoveryNeeded() { lock (Gate) { Alive(); recoveryNeeded = true; } }

    internal string Backup(Action<string> writeCopy)
    {
        ArgumentNullException.ThrowIfNull(writeCopy);
        string id, snapshot, metadata;
        lock (Gate)
        {
            Alive();
            var key = Guid.NewGuid().ToString("N");
            id = "b-" + Path.GetFileName(directory)[2..] + "-" + key;
            snapshot = Path.Combine(directory, "backup-" + key + extension);
            metadata = Path.Combine(directory, "backup-" + key + ".json");
            EnsureBudget(root, Math.Max(1, version.Length) + MetadataAllowance);
            // Reserve the remaining space while COM writes outside the managed lock.
            // Other sessions cannot start another write against this same budget.
            Reservations[root] = ByteLimit - UsedBytes(root);
        }
        try
        {
            writeCopy(snapshot);
            lock (Gate)
            {
                PlainPath(snapshot);
                var captured = Inspect(snapshot, validateWord: extension == ".docx");
                Reservations.Remove(root);
                EnsureBudget(root, MetadataAllowance);
                WriteMetadata(metadata, new BackupMetadata(1, sourceKey, extension, version, captured.Length, captured.Hash));
                return id;
            }
        }
        catch
        {
            lock (Gate)
            {
                Reservations.Remove(root);
                // Never remove an already-confirmed recovery snapshot on quota failure.
                DeleteOwned(root, snapshot);
                DeleteOwned(root, metadata);
            }
            throw;
        }
    }

    internal static void ReleaseBackup(string backupId, string? dataRoot = null)
    {
        var root = LocalPath(dataRoot ?? DefaultRoot);
        lock (Gate)
        {
            var backup = ReadBackup(root, backupId);
            DeleteOwned(root, backup.Snapshot);
            DeleteOwned(root, backup.MetadataPath);
            var directory = Path.GetDirectoryName(backup.MetadataPath)!;
            if (!ActiveDirectories.Contains(directory) && !File.Exists(Path.Combine(directory, "working" + backup.Metadata.Extension))
                && !Directory.EnumerateFiles(directory, "backup-*.json").Any())
            {
                DeleteOwned(root, Path.Combine(directory, "session.json"));
                DeleteEmptyDirectory(root, directory);
            }
        }
    }

    // The caller must only choose this after Office confirmed a clean close.
    // Failed or dirty closes leave these known files in place for recovery.
    internal void DisposeClean()
    {
        lock (Gate)
        {
            if (cleaned) return;
            Alive();
            if (pendingSave || recoveryNeeded)
                throw Fault("recovery_required", "Recovery files must be retained after an unconfirmed save.");
            DeleteOwned(root, WorkingPath);
            DeleteOwned(root, PreviousPath);
            if (!Directory.EnumerateFiles(directory, "backup-*.json").Any()) DeleteOwned(root, Path.Combine(directory, "session.json"));
            ActiveDirectories.Remove(directory);
            DeleteEmptyDirectory(root, directory);
            cleaned = true;
        }
    }

    private void Alive()
    {
        if (cleaned) throw Fault("office_session_expired", "Native recovery ownership has ended.");
        PlainPath(directory);
    }

    private static (string Snapshot, string MetadataPath, BackupMetadata Metadata) ReadBackup(string root, string id)
    {
        var match = BackupPattern.Match(id);
        if (!match.Success) throw Fault("invalid_backup", "Invalid recovery identity.");
        var directory = Path.Combine(root, "s-" + match.Groups[1].Value);
        var metadata = Path.Combine(directory, "backup-" + match.Groups[2].Value + ".json");
        PlainPath(metadata);
        if (!File.Exists(metadata) || new FileInfo(metadata).Length > MetadataAllowance)
            throw Fault("invalid_backup", "The recovery metadata is unavailable.");
        BackupMetadata record;
        try { record = JsonSerializer.Deserialize<BackupMetadata>(File.ReadAllBytes(metadata)) ?? throw new JsonException(); }
        catch (JsonException) { throw Fault("invalid_backup", "The recovery metadata is invalid."); }
        if (record.Schema != 1 || !Regex.IsMatch(record.Extension, "^\\.[a-z0-9]{1,12}$", RegexOptions.CultureInvariant)
            || record.Length <= 0 || record.Length > ByteLimit || record.SourceVersion is null)
            throw Fault("invalid_backup", "The recovery metadata is invalid.");
        var snapshot = Path.Combine(directory, "backup-" + match.Groups[2].Value + record.Extension);
        PlainPath(snapshot);
        return (snapshot, metadata, record);
    }

    private void WriteMetadata(string path, object value)
    {
        var bytes = JsonSerializer.SerializeToUtf8Bytes(value);
        if (bytes.Length > MetadataAllowance) throw Fault("recovery_limit", "Recovery metadata exceeds its budget.");
        using var output = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        output.Write(bytes);
        output.Flush(true);
    }

    private static void CopyChecked(string source, string destination, SourceVersion expected)
    {
        PlainPath(source);
        PlainPath(Path.GetDirectoryName(destination)!);
        using var input = OpenRead(source);
        if (Inspect(input, false) != expected) throw Fault("document_conflict", "The document changed before it could be protected.");
        input.Position = 0;
        using var output = new FileStream(destination, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        input.CopyTo(output);
        output.Flush(true);
    }

    private static SourceVersion Inspect(string path, bool validateWord = false)
    {
        PlainPath(path);
        using var input = OpenRead(path);
        return Inspect(input, validateWord);
    }

    private static FileStream OpenRead(string path) => new(path, FileMode.Open, FileAccess.Read, FileShare.Read);

    private static SourceVersion Inspect(FileStream input, bool validateWord)
    {
        if (!GetFileInformationByHandle(input.SafeFileHandle, out var info)
            || (info.Attributes & (uint)(FileAttributes.ReadOnly | FileAttributes.ReparsePoint | FileAttributes.Directory)) != 0
            || info.Links != 1)
            throw Fault("document_conflict", "The document is read-only, linked, or has an unsupported identity.");
        long length = ((long)info.SizeHigh << 32) | info.SizeLow;
        if (length <= 0 || length > ByteLimit) throw Fault("recovery_limit", "The document is empty or exceeds the native recovery budget.");
        input.Position = 0;
        var hash = Convert.ToHexString(SHA256.HashData(input));
        if (validateWord)
        {
            input.Position = 0;
            try
            {
                using var zip = new ZipArchive(input, ZipArchiveMode.Read, leaveOpen: true);
                var entries = zip.Entries.Where(e => e.FullName.Replace('\\', '/') == "word/document.xml").ToArray();
                if (entries.Length != 1 || entries[0].Length > 64L * 1024 * 1024)
                {
                    // #region agent log
                    DebugNdjson.Write("B", "NativeSessionFiles.Inspect", "word_zip_invalid", new
                    {
                        names = zip.Entries.Select(e => e.FullName).Take(12).ToArray(),
                        docs = entries.Length
                    });
                    // #endregion
                    throw new InvalidDataException();
                }
                using var xml = entries[0].Open();
                using var reader = XmlReader.Create(xml, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit,
                    XmlResolver = null, MaxCharactersInDocument = 64L * 1024 * 1024 });
                reader.MoveToContent();
                if (reader.LocalName != "document" || reader.NamespaceURI is not
                    ("http://schemas.openxmlformats.org/wordprocessingml/2006/main" or "http://purl.oclc.org/ooxml/wordprocessingml/main"))
                    throw new InvalidDataException();
                while (reader.Read()) { }
            }
            catch (Exception error) when (error is InvalidDataException or XmlException)
            {
                throw Fault("invalid_saved_document", "Word did not produce a valid document file.");
            }
        }
        return new(info.Volume, ((ulong)info.IndexHigh << 32) | info.IndexLow, length,
            ((long)info.WriteHigh << 32) | info.WriteLow, hash);
    }

    private static void EnsureBudget(string root, long additional)
    {
        PlainPath(root);
        var reserved = Reservations.GetValueOrDefault(root);
        if (additional < 0 || UsedBytes(root) + reserved + additional > ByteLimit)
            throw Fault("recovery_limit", "Native recovery storage is full; existing recovery files were retained.");
    }

    private static List<string> SessionDirectories(string root)
    {
        PlainPath(root);
        if (!Directory.Exists(root)) return [];
        var directories = Directory.EnumerateDirectories(root).ToList();
        foreach (var directory in directories)
        {
            PlainPath(directory);
            if (!Regex.IsMatch(Path.GetFileName(directory), "^s-[a-f0-9]{32}$", RegexOptions.CultureInvariant))
                throw Fault("recovery_limit", "Native recovery storage contains an unknown directory.");
        }
        return directories;
    }

    private static long UsedBytes(string root)
    {
        long total = 0;
        foreach (var directory in SessionDirectories(root))
        {
            if (Directory.EnumerateDirectories(directory).Any()) throw Fault("recovery_limit", "Native recovery storage contains an unknown directory.");
            foreach (var file in Directory.EnumerateFiles(directory))
            {
                PlainPath(file);
                total = checked(total + new FileInfo(file).Length);
            }
        }
        return total;
    }

    private static string LocalPath(string path)
    {
        if (!Path.IsPathFullyQualified(path) || path.StartsWith("\\\\", StringComparison.Ordinal))
            throw Fault("invalid_path", "Native sessions require a local absolute path.");
        return Path.GetFullPath(path);
    }

    private static string PathKey(string path) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(path.ToUpperInvariant())));

    private static void PlainPath(string path)
    {
        for (string? current = Path.GetFullPath(path); current is not null; current = Path.GetDirectoryName(current))
        {
            if ((File.Exists(current) || Directory.Exists(current)) && (File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0)
                throw Fault("document_conflict", "Linked filesystem paths are not supported for native recovery.");
            var parent = Path.GetDirectoryName(current);
            if (parent == current) break;
        }
    }

    private static void DeleteOwned(string root, string path)
    {
        var directory = Path.GetDirectoryName(Path.GetFullPath(path))!;
        if (!string.Equals(Path.GetDirectoryName(directory), Path.GetFullPath(root), StringComparison.OrdinalIgnoreCase)
            || !Regex.IsMatch(Path.GetFileName(directory), "^s-[a-f0-9]{32}$", RegexOptions.CultureInvariant)
            || !Regex.IsMatch(Path.GetFileName(path), "^(session\\.json|(?:working|previous)\\.[a-z0-9]{1,12}|backup-[a-f0-9]{32}\\.(?:json|[a-z0-9]{1,12}))$", RegexOptions.CultureInvariant))
            throw Fault("invalid_path", "Recovery cleanup refused an unowned path.");
        PlainPath(path);
        if (File.Exists(path)) File.Delete(path);
    }

    private static void DeleteEmptyDirectory(string root, string directory)
    {
        PlainPath(directory);
        if (string.Equals(Path.GetDirectoryName(directory), root, StringComparison.OrdinalIgnoreCase)
            && Regex.IsMatch(Path.GetFileName(directory), "^s-[a-f0-9]{32}$", RegexOptions.CultureInvariant)
            && Directory.Exists(directory) && !Directory.EnumerateFileSystemEntries(directory).Any())
            Directory.Delete(directory, false);
    }

    private static CoreFault Fault(string code, string message) => new(code, message);

    [StructLayout(LayoutKind.Sequential)]
    private struct FileInformation
    {
        internal uint Attributes, CreateLow, CreateHigh, AccessLow, AccessHigh, WriteLow, WriteHigh;
        internal uint Volume, SizeHigh, SizeLow, Links, IndexHigh, IndexLow;
    }
    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetFileInformationByHandle(SafeFileHandle handle, out FileInformation information);
}

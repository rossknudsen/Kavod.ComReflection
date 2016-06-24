using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.Contracts;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Win32;

namespace Kavod.ComReflection
{
    public class LibraryRegistration
    {
        private LibraryRegistration(string filePath, string name, Version version, Guid guid)
        {
            Contract.Requires<ArgumentNullException>(!string.IsNullOrEmpty(filePath));
            Contract.Requires<ArgumentNullException>(!string.IsNullOrEmpty(name));

            Name = name;
            FilePath = filePath;
            Guid = guid;

            short minor;
            short major;
            if (short.TryParse(version.MinorVersion, out minor)
                && short.TryParse(version.MajorVersion, out major))
            {
                MajorVersion = major;
                MinorVersion = minor;
            }
        }

        public string FilePath { get; }

        public string Name { get; }

        public short MajorVersion { get; }

        public short MinorVersion { get; }

        public Guid Guid { get; }

        public static IEnumerable<LibraryRegistration> GetRegisteredTypeLibraryEntries()
        {
            // TODO sometimes there is more than one registration per file path i.e. with different GUIDs.
            // need to decide how to cope with this.

            var dictionary = new Dictionary<string, LibraryRegistration>();
            foreach (var e in GetComTypeRegistryEntries())
            {
                if (e.FilePath != null)
                {
                    if (!dictionary.ContainsKey(e.FilePath))
                    {
                        dictionary[e.FilePath] = 
                            new LibraryRegistration(e.FilePath, e.Name, e.Version, e.Guid);
                    }
                }
            }
            return dictionary.Values;
        }

        public static LibraryRegistration GetComTypeRegistryEntry(Guid guid)
        {
            using (var clsidRootKey = Registry.ClassesRoot.OpenSubKey("TypeLib"))
            using (var libKey = clsidRootKey?.OpenSubKey(guid.ToString("B")))
            {
                if (clsidRootKey == null || libKey == null)
                {
                    throw new InvalidOperationException();
                }
                var entry = EnumerateRegistryEntryVersions(libKey, guid).First();
                return new LibraryRegistration(entry.FilePath, entry.Name, entry.Version, guid);
            }
        }

        private static IEnumerable<ComTypeRegistryEntry> GetComTypeRegistryEntries()
        {
            using (var clsidRootKey = Registry.ClassesRoot.OpenSubKey("TypeLib"))
            {
                if (clsidRootKey == null)
                {
                    yield break;
                }

                foreach (var typeLibKey in EnumerateSubKeys(clsidRootKey))
                {
                    var currentTypeLib = GetCurrentKeyName(typeLibKey);
                    Guid clsid;
                    if (!Guid.TryParseExact(currentTypeLib, "B", out clsid))
                    {
                        Debug.WriteLine($"Couldn't parse CLSID = {currentTypeLib}");
                        continue;
                    }
                    
                    foreach (var entry in EnumerateRegistryEntryVersions(typeLibKey, clsid))
                        yield return entry;
                }
            }
        }

        private static IEnumerable<ComTypeRegistryEntry> EnumerateRegistryEntryVersions(RegistryKey typeLibKey, Guid guid)
        {
            // There can be more than one registered item for each GUID.
            foreach (var versionKey in EnumerateSubKeys(typeLibKey))
            {
                var version = GetVersionInfo(GetCurrentKeyName(versionKey));

                var name = (string) versionKey.GetValue("");
                if (string.IsNullOrEmpty(name))
                {
                    continue;
                }

                foreach (var patchKey in EnumerateSubKeys(versionKey))
                {
                    var win32Path = string.Empty;
                    var win64Path = string.Empty;

                    foreach (var win32Key in EnumerateSubKeys(patchKey, k => k.Name.EndsWith("win32")))
                    {
                        if (win32Key.Name.EndsWith("win32"))
                        {
                            win32Path = (string) win32Key.GetValue("") ?? string.Empty;
                        }
                        else if (win32Key.Name.EndsWith("win64"))
                        {
                            win64Path = (string) win32Key.GetValue("") ?? string.Empty;
                        }
                    }

                    if (win64Path == string.Empty && win32Path == string.Empty)
                    {
                        continue;
                    }

                    yield return new ComTypeRegistryEntry()
                    {
                        FilePath = win32Path != string.Empty ? win32Path : win64Path,
                        Name = name,
                        Version = version,
                        Guid = guid
                    };
                }
            }
        }

        private static string GetCurrentKeyName(RegistryKey key)
        {
            Contract.Requires<ArgumentNullException>(key != null);

            // TODO handle if the last char is a '\'
            return key.Name.Substring(key.Name.LastIndexOf('\\') + 1);
        }

        private static IEnumerable<RegistryKey> EnumerateSubKeys(RegistryKey parentKey, Predicate<RegistryKey> matchKey = null)
        {
            Contract.Requires<ArgumentNullException>(parentKey != null);

            if (matchKey == null)
            {
                matchKey = k => true;
            }

            foreach (var subKeyName in parentKey.GetSubKeyNames())
            {
                using (var subKey = parentKey.OpenSubKey(subKeyName))
                {
                    if (subKey != null
                        && matchKey(subKey))
                    {
                        yield return subKey;
                    }
                }
            }
        }

        private static Version GetVersionInfo(string versionInfo)
        {
            Contract.Requires<ArgumentNullException>(versionInfo != null);
            Contract.Requires<ArgumentOutOfRangeException>(versionInfo.Length > 0);

            var regex = Regex.Match(versionInfo, @"^(\w)\.(\w)$");
            var major = regex.Groups[1].Value;
            var minor = regex.Groups[2].Value;
            return new Version()
            {
                MajorVersion = major,
                MinorVersion = minor
            };
        }

        private struct Version
        {
            internal string MajorVersion;

            internal string MinorVersion;
        }

        private struct ComTypeRegistryEntry
        {
            internal string FilePath;

            internal string Name;

            internal Version Version;

            internal Guid Guid;
        }
    }
}
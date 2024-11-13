// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text.Json;

using Microsoft.PowerToys.FilePreviewCommon.Monaco.Formatters;

namespace Microsoft.PowerToys.FilePreviewCommon
{
    public static class MonacoHelper
    {
        public const string VirtualHostName = "PowerToysLocalMonaco";

        private const string DefaultLanguage = "plaintext";

        /// <summary>
        /// Formatters applied before rendering the preview.
        /// </summary>
        public static readonly IReadOnlyCollection<IFormatter> Formatters = new List<IFormatter>
        {
            new JsonFormatter(),
            new XmlFormatter(),
        }.AsReadOnly();

        private static readonly Lazy<Dictionary<string, string>> _languageExtensionMap =
            new(CreateLanguageExtensionMap);

        /// <summary>
        /// Gets the mapping of file extensions to Monaco language IDs.
        /// </summary>
        public static IReadOnlyDictionary<string, string> LanguageExtensionMap => _languageExtensionMap.Value;

        private static readonly Lazy<string> _monacoDirectory = new(GetRuntimeMonacoDirectory);

        /// <summary>
        /// Gets the path of the Monaco assets folder.
        /// </summary>
        public static string MonacoDirectory => _monacoDirectory.Value;

        private static readonly Lazy<string> _indexHtml =
            new(() => File.ReadAllText(Path.Combine(MonacoDirectory, "index.html")));

        /// <summary>
        /// Gets the cached contents of the Monaco index.html file.
        /// </summary>
        public static string IndexHtml => _indexHtml.Value;

        private static string GetRuntimeMonacoDirectory()
        {
            string exePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? string.Empty;

            // If the executable is within "WinUI3Apps", correct the path first.
            if (Path.GetFileName(exePath) == "WinUI3Apps")
            {
                exePath = Path.Combine(exePath, "..");
            }

            string monacoPath = Path.Combine(exePath, "Assets", "Monaco");

            return Directory.Exists(monacoPath) ?
                monacoPath :
                throw new DirectoryNotFoundException($"Monaco assets directory not found at {monacoPath}");
        }

        private static Dictionary<string, string> CreateLanguageExtensionMap()
        {
            string jsonPath = Path.Combine(MonacoDirectory, "monaco_languages.json");

            using var jsonFileStream = new FileStream(jsonPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var jsonDocument = JsonDocument.Parse(jsonFileStream);

            var languageList = jsonDocument.RootElement.GetProperty("list");

            var map = new Dictionary<string, string>();

            foreach (var item in languageList.EnumerateArray())
            {
                if (item.TryGetProperty("id", out var idElement) && idElement.GetString() is string id &&
                    item.TryGetProperty("extensions", out var extensionsElement) && extensionsElement.ValueKind == JsonValueKind.Array)
                {
                    foreach (var ext in extensionsElement.EnumerateArray())
                    {
                        string? extension = ext.GetString();

                        if (extension != null && !map.ContainsKey(extension))
                        {
                            map[extension] = id;
                        }
                    }
                }
            }

            return map;
        }

        /// <summary>
        /// Get the Monaco language ID for the supplied file extension.
        /// </summary>
        /// <param name="fileExtension">The file extension, including the initial period.</param>
        /// <returns>The Monaco language ID for the extension, or "plaintext" if not found.</returns>
        public static string GetLanguage(string fileExtension) =>
            LanguageExtensionMap.TryGetValue(fileExtension, out string? id) ? id : DefaultLanguage;
    }
}

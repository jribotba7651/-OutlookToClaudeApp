using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Newtonsoft.Json;
using OutlookToClaudeApp.Models;

namespace OutlookToClaudeApp.Services
{
    public class ConfigService
    {
        private static readonly string ConfigFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "OutlookToClaudeApp"
        );

        private static readonly string ConfigFile = Path.Combine(ConfigFolder, "config.enc");
        private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("OutlookToClaudeApp-v1.0");

        public ConfigService()
        {
            // Ensure config folder exists
            if (!Directory.Exists(ConfigFolder))
            {
                Directory.CreateDirectory(ConfigFolder);
            }
        }

        public ApiConfig LoadConfig()
        {
            try
            {
                if (!File.Exists(ConfigFile))
                {
                    return new ApiConfig();
                }

                // Read encrypted data
                var encryptedData = File.ReadAllBytes(ConfigFile);

                // Decrypt using Windows DPAPI (Data Protection API)
                var decryptedData = ProtectedData.Unprotect(
                    encryptedData,
                    Entropy,
                    DataProtectionScope.CurrentUser
                );

                // Deserialize JSON
                var json = Encoding.UTF8.GetString(decryptedData);
                return JsonConvert.DeserializeObject<ApiConfig>(json) ?? new ApiConfig();
            }
            catch
            {
                // If decryption fails, return empty config
                return new ApiConfig();
            }
        }

        public void SaveConfig(ApiConfig config)
        {
            try
            {
                // Serialize to JSON
                var json = JsonConvert.SerializeObject(config, Formatting.Indented);
                var data = Encoding.UTF8.GetBytes(json);

                // Encrypt using Windows DPAPI
                var encryptedData = ProtectedData.Protect(
                    data,
                    Entropy,
                    DataProtectionScope.CurrentUser
                );

                // Save to file
                File.WriteAllBytes(ConfigFile, encryptedData);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to save configuration: {ex.Message}", ex);
            }
        }

        public void ClearConfig()
        {
            try
            {
                if (File.Exists(ConfigFile))
                {
                    File.Delete(ConfigFile);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to clear configuration: {ex.Message}", ex);
            }
        }

        public string GetConfigLocation()
        {
            return ConfigFolder;
        }
    }
}

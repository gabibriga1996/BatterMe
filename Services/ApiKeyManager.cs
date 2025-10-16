using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using System.Collections.Generic;
using System.Security.Cryptography.X509Certificates;
using System.Security;

namespace BetterMeV2VSTO.Services
{
    public static class ApiKeyManager
    {
        private const string KeyFileName = "apikey.enc";
        private static readonly byte[] entropy = Encoding.Unicode.GetBytes("BetterMeV2VSTO_Salt_2024");
        private static string _cachedApiKey;

        public static string GetApiKey()
        {
            if (!string.IsNullOrEmpty(_cachedApiKey)) 
                return _cachedApiKey;

            try
            {
                var userDir = GetUserDirectory();
                var keyPath = Path.Combine(userDir, KeyFileName);

                // נסה לקרוא מקובץ מוצפן תחילה
                if (File.Exists(keyPath))
                {
                    var encryptedData = File.ReadAllBytes(keyPath);
                    var decryptedKey = UnprotectData(encryptedData);
                    if (ValidateApiKey(decryptedKey))
                    {
                        _cachedApiKey = decryptedKey;
                        return decryptedKey;
                    }
                }

                // נסה קובץ תצורה כגיבוי (לתאימות לאחור)
                var configKey = GetKeyFromConfig();
                if (ValidateApiKey(configKey))
                {
                    SaveApiKey(configKey); // העבר לאחסון מוצפן
                    _cachedApiKey = configKey;
                    return configKey;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error reading API key: {ex.Message}");
            }

            return null;
        }

        public static bool SaveApiKey(string apiKey)
        {
            if (!ValidateApiKey(apiKey))
                return false;

            try
            {
                var userDir = GetUserDirectory();
                Directory.CreateDirectory(userDir);

                byte[] protectedBytes;
                byte[] sourceBytes = Encoding.Unicode.GetBytes(apiKey);

                using (var aes = Aes.Create())
                {
                    aes.KeySize = 256;
                    aes.GenerateKey();
                    aes.GenerateIV();

                    using (var encryptor = aes.CreateEncryptor())
                    using (var msEncrypt = new MemoryStream())
                    {
                        // שמור את ה-IV בתחילת הקובץ
                        msEncrypt.Write(aes.IV, 0, aes.IV.Length);

                        using (var csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write))
                        using (var swEncrypt = new BinaryWriter(csEncrypt))
                        {
                            swEncrypt.Write(sourceBytes);
                        }

                        protectedBytes = msEncrypt.ToArray();
                    }

                    // הצפן את מפתח ה-AES עם entropy
                    var encryptedKey = EncryptAesKey(aes.Key);
                    
                    // שמור את שניהם
                    var keyPath = Path.Combine(userDir, KeyFileName);
                    File.WriteAllBytes(keyPath, protectedBytes);
                    File.WriteAllBytes(keyPath + ".key", encryptedKey);
                }

                _cachedApiKey = apiKey;
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving API key: {ex.Message}");
                return false;
            }
        }

        private static byte[] EncryptAesKey(byte[] aesKey)
        {
            using (var provider = new RijndaelManaged())
            {
                provider.KeySize = 256;
                provider.BlockSize = 128;
                var salt = entropy;
                var rgb = new Rfc2898DeriveBytes(entropy, salt, 1000);
                provider.Key = rgb.GetBytes(32);
                provider.IV = rgb.GetBytes(16);

                using (var ms = new MemoryStream())
                using (var cs = new CryptoStream(ms, provider.CreateEncryptor(), CryptoStreamMode.Write))
                {
                    cs.Write(aesKey, 0, aesKey.Length);
                    cs.FlushFinalBlock();
                    return ms.ToArray();
                }
            }
        }

        private static byte[] DecryptAesKey(byte[] encryptedKey)
        {
            using (var provider = new RijndaelManaged())
            {
                provider.KeySize = 256;
                provider.BlockSize = 128;
                var salt = entropy;
                var rgb = new Rfc2898DeriveBytes(entropy, salt, 1000);
                provider.Key = rgb.GetBytes(32);
                provider.IV = rgb.GetBytes(16);

                using (var ms = new MemoryStream())
                using (var cs = new CryptoStream(ms, provider.CreateDecryptor(), CryptoStreamMode.Write))
                {
                    cs.Write(encryptedKey, 0, encryptedKey.Length);
                    cs.FlushFinalBlock();
                    return ms.ToArray();
                }
            }
        }

        public static void ClearApiKey()
        {
            _cachedApiKey = null;
            try
            {
                var keyPath = Path.Combine(GetUserDirectory(), KeyFileName);
                if (File.Exists(keyPath))
                    File.Delete(keyPath);
                if (File.Exists(keyPath + ".key"))
                    File.Delete(keyPath + ".key");
            }
            catch { }
        }

        private static string GetUserDirectory()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "BetterMeV2VSTO");
        }

        private static string UnprotectData(byte[] encryptedData)
        {
            try
            {
                var keyPath = Path.Combine(GetUserDirectory(), KeyFileName + ".key");
                if (!File.Exists(keyPath))
                    return null;

                var encryptedKey = File.ReadAllBytes(keyPath);
                var aesKey = DecryptAesKey(encryptedKey);

                using (var aes = Aes.Create())
                {
                    aes.KeySize = 256;
                    
                    // קרא את ה-IV מתחילת הקובץ המוצפן
                    byte[] iv = new byte[aes.BlockSize / 8];
                    Array.Copy(encryptedData, 0, iv, 0, iv.Length);

                    aes.Key = aesKey;
                    aes.IV = iv;

                    using (var decryptor = aes.CreateDecryptor())
                    using (var msDecrypt = new MemoryStream(encryptedData, iv.Length, encryptedData.Length - iv.Length))
                    using (var csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read))
                    using (var srDecrypt = new BinaryReader(csDecrypt))
                    {
                        var decryptedBytes = new byte[encryptedData.Length];
                        int decryptedByteCount = csDecrypt.Read(decryptedBytes, 0, decryptedBytes.Length);
                        return Encoding.Unicode.GetString(decryptedBytes, 0, decryptedByteCount);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error decrypting: {ex.Message}");
                return null;
            }
        }

        private static string GetKeyFromConfig()
        {
            try
            {
                var asmPath = typeof(ApiKeyManager).Assembly.Location;
                var configPath = asmPath + ".config";
                if (File.Exists(configPath))
                {
                    var doc = XDocument.Load(configPath);
                    var appSettings = doc.Root?.Element("appSettings");
                    if (appSettings != null)
                    {
                        foreach (var add in appSettings.Elements("add"))
                        {
                            var keyAttr = add.Attribute("key");
                            if (keyAttr != null && string.Equals(keyAttr.Value, "API_Key", StringComparison.OrdinalIgnoreCase))
                            {
                                return add.Attribute("value")?.Value;
                            }
                        }
                    }
                }
            }
            catch { }
            return null;
        }

        public static bool ValidateApiKey(string key)
        {
            if (string.IsNullOrWhiteSpace(key)) 
                return false;

            // OpenRouter API key format: sk-or-v1-[64 hex chars]
            return System.Text.RegularExpressions.Regex.IsMatch(key, "^sk-or-v1-[a-fA-F0-9]{64}$");
        }
    }
}
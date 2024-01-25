using Microsoft.Extensions.Logging;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System;

namespace ExcelWorkbookPivotTable.ProtectData
{
    public interface ISecureInfo
    {
        string EncryptData(string plainText);

        string DecryptData(string cipherText);
    }

    public class SecureInfo : ISecureInfo
    {
        private readonly ILogger<SecureInfo> _logger;

        #region

        public static string ScrtKey = "fceb0824-95b9-47";

        public static string InitialVectorKey = "54399dbb-6ba6-44";

        #endregion

        public SecureInfo(ILogger<SecureInfo> logger)
        {
            _logger = logger;
        }

        //encrypt data
        public string EncryptData(string plainText)
        {
            string encryptedString = string.Empty;

            try
            {
                var aesKey = Encoding.UTF8.GetBytes(ScrtKey + InitialVectorKey);
                var aesInitialVector = Encoding.UTF8.GetBytes(InitialVectorKey);

                //Create an Aes object with the specified key and IV.
                using (Aes aesAlgorithm = Aes.Create())
                {
                    aesAlgorithm.Key = aesKey;
                    aesAlgorithm.IV = aesInitialVector;

                    //Create an encryptor to perform the stream transform.
                    ICryptoTransform encryptor = aesAlgorithm.CreateEncryptor(aesAlgorithm.Key, aesAlgorithm.IV);

                    //Create the streams used for encryption.
                    using MemoryStream msEncrypt = new MemoryStream();
                    using CryptoStream csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write);
                    using (StreamWriter swEncrypt = new StreamWriter(csEncrypt))
                    {
                        //Write all data to the stream.
                        swEncrypt.Write(plainText);
                    }

                    byte[] encrypted = msEncrypt.ToArray();
                    encryptedString = Convert.ToBase64String(encrypted);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Encryption has failed.");
            }

            return encryptedString;
        }

        //decrypt data
        public string DecryptData(string cipherText)
        {
            string decryptedString = string.Empty;

            try
            {
                var aesKey = Encoding.UTF8.GetBytes(ScrtKey + InitialVectorKey);
                var aesInitialVector = Encoding.UTF8.GetBytes(InitialVectorKey);

                //Create an Aes object with the specified key and IV.
                using (Aes aesAlg = Aes.Create())
                {
                    aesAlg.Key = aesKey;
                    aesAlg.IV = aesInitialVector;

                    // Create a decryptor to perform the stream transform.
                    ICryptoTransform decryptor = aesAlg.CreateDecryptor(aesAlg.Key, aesAlg.IV);

                    //get byte[] from ciper text 
                    byte[] cipherTextByte = Convert.FromBase64String(cipherText);

                    //Create the streams used for decryption.
                    using MemoryStream msDecrypt = new MemoryStream(cipherTextByte);
                    using CryptoStream csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read);
                    using (StreamReader srDecrypt = new StreamReader(csDecrypt))
                    {
                        //Read the decrypted bytes from the decrypting stream and place in a string.
                        decryptedString = srDecrypt.ReadToEnd();
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Decryption has failed.");
            }

            return decryptedString;
        }
    }
}

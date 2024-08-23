using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web;

namespace SCADA_API
{
    public static class Global
    {
        public static string APIPath = (ConfigurationManager.AppSettings["APIPath"]==null)? ConfigurationManager.AppSettings["APIScadaPath"].ToString(): ConfigurationManager.AppSettings["APIPath"].ToString();
        public static string _Auth = "Basic aHlzZHJvZXhwYW5kdGVzdHNjYWRhOmNsYWR0ZWtiYXRhbTIwMjA=";

    }

    public class CryptoAPI
    {
        public static string Hash(string value)
        {
            var item = Encoding.UTF8.GetBytes(value);
            return Convert.ToBase64String(
                System.Security.Cryptography.SHA256.Create()
                .ComputeHash(item)
                );
        }
        public static string EncryptHash(string plainText)
        {
            string EncryptionKey = "C20DA1L8T34ENK9";
            // Create a new Stringbuilder to collect the bytes
            // and create a string.
            StringBuilder sBuilder = new StringBuilder();
            // Return the hexadecimal string.
            using (MD5 md5Hash = MD5.Create())
            {
                byte[] data = md5Hash.ComputeHash(Encoding.UTF8.GetBytes(plainText));

                // Loop through each byte of the hashed data 
                // and format each one as a hexadecimal string.
                for (int i = 0; i < data.Length; i++)
                {
                    sBuilder.Append(data[i].ToString("x2"));
                }

            }
            byte[] clearBytes = Encoding.Unicode.GetBytes(sBuilder.ToString());
            using (Aes encryptor = Aes.Create())
            {
                Rfc2898DeriveBytes pdb = new
                    Rfc2898DeriveBytes(EncryptionKey, new byte[]
                    { 0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76 });
                encryptor.Key = pdb.GetBytes(32);
                encryptor.IV = pdb.GetBytes(16);
                using (MemoryStream ms = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(clearBytes, 0, clearBytes.Length);
                        cs.Close();
                    }
                    plainText = Convert.ToBase64String(ms.ToArray());
                }
            }
            return plainText;
        }
    }

}
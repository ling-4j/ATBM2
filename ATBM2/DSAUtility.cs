using System;
using System.Security.Cryptography;
using System.Text;

namespace ATBM2
{
    public static class DSAUtility
    {
        public static (DSAParameters, DSAParameters) GenerateKeys()
        {
            using (var dsa = new DSACryptoServiceProvider(1024))
            {
                return (dsa.ExportParameters(true), dsa.ExportParameters(false));
            }
        }

        public static byte[] SignData(byte[] data, DSAParameters privateKey)
        {
            using (var dsa = new DSACryptoServiceProvider())
            {
                dsa.ImportParameters(privateKey);
                return dsa.SignData(data);
            }
        }

        public static bool VerifySignature(byte[] data, byte[] signature, DSAParameters publicKey)
        {
            using (var dsa = new DSACryptoServiceProvider())
            {
                dsa.ImportParameters(publicKey);
                return dsa.VerifyData(data, signature);
            }
        }

        public static byte[] ComputeSHA1Hash(byte[] data)
        {
            using (var sha1 = SHA1.Create())
            {
                return sha1.ComputeHash(data);
            }
        }

        public static string BytesToHex(byte[] bytes)
        {
            StringBuilder sb = new StringBuilder();
            foreach (byte b in bytes)
            {
                sb.Append(b.ToString("X2"));
            }
            return sb.ToString();
        }
    }
}

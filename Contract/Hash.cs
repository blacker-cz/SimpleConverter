using System;
using System.Text;
using System.Security.Cryptography;

namespace SimpleConverter.Contract
{
    /// <summary>
    /// Class providing hashing algorithms.
    /// </summary>
    class Hash
    {
        /// <summary>
        /// Compute MD5 hash for input string
        /// </summary>
        /// <param name="input">Input string</param>
        /// <returns>MD5 hash</returns>
        public static string md5(string input)
        {
            MD5 md5Hasher = MD5.Create();

            byte[] data = md5Hasher.ComputeHash(Encoding.UTF8.GetBytes(input));

            md5Hasher.Clear();  // clear to prevent memory leaks

            StringBuilder sBuilder = new StringBuilder();

            // build hexadecimal representation
            for (int i = 0; i < data.Length; i++)
            {
                sBuilder.Append(data[i].ToString("x2"));
            }

            return sBuilder.ToString();
        }
    }
}

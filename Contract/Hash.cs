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
        /// Compute hash for input string.
        /// First compute MD5 then split md5 hash in to 4 32bit parts and xor them.
        /// </summary>
        /// <param name="input">Input string</param>
        /// <returns>Hash in base 36</returns>
        public static string ComputeHash(string input)
        {
            MD5 md5Hasher = MD5.Create();

            byte[] data = md5Hasher.ComputeHash(Encoding.UTF8.GetBytes(input));

            md5Hasher.Clear();  // clear to prevent memory leaks

            uint a, b, c, d, hash;
            a = BitConverter.ToUInt32(data, 0);
            b = BitConverter.ToUInt32(data, 4);
            c = BitConverter.ToUInt32(data, 8);
            d = BitConverter.ToUInt32(data, 12);
            hash = a ^ b;

            string chars = "0123456789abcdefghijklmnopqrstuvwxyz";

            uint r;
            string hashString = "";

            // in r we have the offset of the char that was converted to the new base
            while (hash >= 36)
            {
                r = hash % 36;
                hashString = chars[(int) r] + hashString;
                hash = hash / 36;
            }

            // the last number to convert
            hashString = chars[(int) hash] + hashString;

            return hashString;
        }
    }
}

using System;
using System.Security.Cryptography;

namespace HolyCryptv3.Utils {
    public static class CheckUtils {
        public static bool CheckHashCode(byte[] msg, string hash) {
            SHA384 SHA = SHA384.Create();
            byte[] Hash = SHA.ComputeHash(msg);
            return Convert.ToHexString(Hash) == hash;
            //return this.Encoding.GetString(Hash);
        }
    }
}

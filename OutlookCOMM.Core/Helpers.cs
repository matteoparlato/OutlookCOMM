using System.Collections.Generic;

namespace OutlookCOMM.Core
{
    public static class Helpers
    {
        /// <summary>
        /// Method which separates a string containing multiple addresses using the given delimiter.
        /// </summary>
        /// <param name="addresses">The string containing the addresses</param>
        /// <param name="delimiter">The delimiter used to separate the addresses</param>
        /// <returns>A collection of addresses</returns>
        public static IEnumerable<string> SplitAddressesByDelimiter(string addresses, char delimiter)
        {
            int startIndex = 0;
            int delimiterIndex = 0;

            while (delimiterIndex >= 0)
            {
                delimiterIndex = addresses.IndexOf(delimiter, startIndex);
                string substring = addresses;
                if (delimiterIndex > 0)
                {
                    substring = addresses.Substring(0, delimiterIndex);
                }

                if (!substring.Contains("\"") || substring.IndexOf("\"") != substring.LastIndexOf("\""))
                {
                    yield return substring;
                    addresses = addresses.Substring(delimiterIndex + 1);
                    startIndex = 0;
                }
                else
                {
                    startIndex = delimiterIndex + 1;
                }
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DCon
{
    public static class ColumnLetters
    {

        public static string GetLetter(int columnNumber)
        {
            string result = String.Empty;
            string[] letters = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M",
                                "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"};

            int t = columnNumber % 26;

            if (columnNumber > 26 )
            {
                result = letters[t - 1];
            }
            else if (columnNumber <= 26)
            {
                result = letters[columnNumber-1];
            }
            return result;
        }
    }
}

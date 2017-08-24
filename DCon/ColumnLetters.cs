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
            string[] letters = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M",
                              "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"};

            return letters[columnNumber-1];
        }
    }
}

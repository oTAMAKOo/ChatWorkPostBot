
using System;
using System.Text.RegularExpressions;

namespace ChatWorkPostBot
{
    public static class ExcelUtility
    {
        /// <summary>
        /// アルファベットを返す。
        /// 例えば、引数が1の時はA、2の時はB、27の時はAAとなる。
        /// 引数が不正な場合は<see cref="string.Empty"/>を返す。
        /// </summary>
        public static string ToAlphabet(int index)
        {
            var alphabet = string.Empty;

            if (index < 1){ return alphabet; }

            index--;

            do
            {
                alphabet = Convert.ToChar(index % 26 + 0x41) + alphabet;
            }
            while ((index = index / 26 - 1) != -1);

            return alphabet;
        }

        /// <summary>
        /// AからZのインデックスを返す。
        /// 例えば、引数がAの時は1、Bの時は2、AAの時は27となる。
        /// 引数が不正な場合は-1を返す。
        /// </summary>
        public static int ToAlphabetIndex(string alphabet)
        {
            if (string.IsNullOrEmpty(alphabet)) { return -1; }

            if (new Regex("^[A-Z]+$").IsMatch(alphabet))
            {
                var index = 0;

                for (var i = 0; i < alphabet.Length; i++)
                {
                    var num = Convert.ToChar(alphabet[alphabet.Length - i - 1]) - 64;

                    index += (int)(num * Math.Pow(26, i));
                }
                return index;
            }

            return -1;
        }
    }
}

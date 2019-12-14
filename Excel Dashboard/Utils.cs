using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace Excel_Dashboard
{
    public static class Utils
    {
        public static List<string> FOLIO = new List<string>() { "FOLIO"};
        public static List<string> TICKET = new List<string>() { "FACTURA" };
        public static List<string> NOMBRE_CLIENTE = new List<string>() { "NOMBRE" };
        public static List<string> ZONA = new List<string>() { "LUGAR","ZONA" };
        public static List<string> UNIDAD = new List<string>() { "VEHCULO", "UNIDAD" };
        public static List<string> CHOFER = new List<string>() { "CHOFER" };
        public static List<string> SALIDA = new List<string>() { "SALIDA" };
        public static List<string> ESTATUS = new List<string>() { "ESTATUS" };
        public static List<string> ESTATUS_CARGANDO = new List<string>() { "CARGANDO" };
        public static List<string> ESTATUS_TRAYECTO = new List<string>() { "TRAYECTO" };
        public static List<string> ESTATUS_ENTREGADO = new List<string>() { "TREGADO" };


        public static string STOP_EXCEL_CONSTRUMA = "NOMBREDELCLIENTE";

        private static Random random = new Random();

        public static Font HEADER_FONT = new System.Drawing.Font("Arial Narrow", 12, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
        public static Font CONTENT_FONT = new System.Drawing.Font("Arial Narrow", 12, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
        public static Color HEADER_COLOR = System.Drawing.Color.AntiqueWhite;
        public static Color CONTENT_COLOR = System.Drawing.Color.White;
        public static int COL_NUMBERS = 7;

        public static readonly string CURRENT_PATH = Path.GetDirectoryName(Application.ExecutablePath);
        public static int ROW_HEIGHT = 40;


        public static bool IsLike(string completeString, string conteinedString)
        {
            completeString = NormalizeString(completeString);
            conteinedString = NormalizeString(conteinedString);
            if (completeString.IndexOf(conteinedString) != -1)
            {
                return true;
            }
            return false;
        }

        public static bool IsLikeStringList(string completeString, List<string> conteinedPossibleStringList)
        {
            foreach (string conteinedString in conteinedPossibleStringList)
            {
                if (IsLike(completeString, conteinedString)) return true;
            }

            return false;
        }

        public static void DeleteFileIfExist(string path)
        {
            if (File.Exists(path))
            {
                try
                {
                    File.Delete(path);
                }
                catch (System.IO.IOException ex)
                {
                    return;
                }
            }
        }

        public static void DeleteFolderIfExist(string path)
        {
            if (Directory.Exists(path))
            {
                try
                {
                    Directory.Delete(path, true);
                }
                catch 
                {
                    return;
                }
            }
        }

        public static string RandomString(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            return new string(Enumerable.Repeat(chars, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
        }

        public static void CreateFolder(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }

        public static bool ExistFile(string path)
        {
            if (File.Exists(path))
            {
                return true;
            }
            return false;
        }

        public static double ConvertDynamicToDouble(dynamic dynamicNumber)
        {


            double amount;
            try
            {
                amount = (double)dynamicNumber;
            }
            catch (Exception e)
            {
                amount = 0;
            }
            if (amount == 0)
            {
                string numberString;
                try
                {
                    numberString = (string)dynamicNumber;
                    amount = double.Parse(numberString);
                }
                catch (Exception e)
                {
                    amount = 0;
                }
            }
            return amount;
        }

        public static string ConvertDynamicToString(dynamic dynamicString)
        {
            string newString;
            try
            {
                newString = (string)(dynamicString + "");
            }
            catch (Exception e)
            {
                newString = String.Empty;
            }

            return newString;
        }


        public static string NormalizeString(string chain)
        {
            string newString = RemoveDiacritics(chain);
            newString = RemoveSpecialCharacters(chain).ToUpper();
            return newString;
        }

        public static string NormalizeStringList(List<string> list)
        {
            string newString = String.Empty;
            foreach (string chain in list)
            {
                newString += NormalizeString(chain) + ' ';
            }
            return newString;
        }

        public static string GetStringListDummie(List<string> list)
        {
            string newString = String.Empty;
            foreach (string chain in list)
            {
                newString += chain + ' ';
            }
            return newString;
        }

        public static bool IsEmptyString(string chain)
        {
            if (chain == "" || chain == null)
            {
                return true;
            }
            return false;
        }

        public static bool IsEmptyGuid(Guid chain)
        {
            if (chain == null)
            {
                return true;
            }
            if (chain.ToString() == "00000000-0000-0000-0000-000000000000")
            {
                return true;
            }
            return false;
        }

        public static bool IsEmail(string chain)
        {
            if (IsEmptyString(chain))
            {
                return false;
            }
            if (chain.IndexOf("@") != -1 && chain.IndexOf(".") != -1)
            {
                return true;
            }
            return false;
        }
        
        private static void Message_FormClosed(object sender, FormClosedEventArgs e)
        {

        }


        public static bool StringToBool(string chain)
        {
            if (chain == "1" || chain.ToUpper() == "TRUE") return true;
            return false;
        }



        public static bool FindAndKillProcess(string name)
        {
            foreach (Process clsProcess in Process.GetProcesses())
            {
                if (clsProcess.ProcessName.StartsWith(name))
                {
                    clsProcess.Kill();
                    return true;
                }
            }
            return false;
        }

        static string RemoveDiacritics(string text)
        {
            var normalizedString = text.Normalize(NormalizationForm.FormD);
            var stringBuilder = new StringBuilder();

            foreach (var c in normalizedString)
            {
                var unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
                if (unicodeCategory != UnicodeCategory.NonSpacingMark)
                {
                    stringBuilder.Append(c);
                }
            }

            return stringBuilder.ToString().Normalize(NormalizationForm.FormC);
        }


        public static string GetNetworkPath(string uncPath, string initialString, string rootString)
        {
            try
            {
                // remove the "\\" from the UNC path and split the path
                string path = String.Empty;
                uncPath = uncPath.Replace(@"\\", "");
                string[] uncParts = uncPath.Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
                bool isRoot = false;
                foreach (string part in uncParts)
                {
                    if (isRoot || part.ToUpper() == rootString.ToUpper())
                    {
                        path += $@"\{part}";
                        isRoot = true;
                    }
                }
                if (isRoot) path = initialString + path;
                return path;
            }
            catch (Exception ex)
            {
                return "[ERROR RESOLVING UNC PATH: " + uncPath + ": " + ex.Message + "]";
            }
        }

        public static string RemoveSpecialCharacters(string str)
        {
            StringBuilder sb = new StringBuilder();
            foreach (char c in str)
            {
                if ((c >= '0' && c <= '9') || (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z'))
                {
                    sb.Append(c);
                }
            }
            return sb.ToString();
        }
        
        public static int ExcelColumnNameToNumber(string columnName)
        {
            if (string.IsNullOrEmpty(columnName)) throw new ArgumentNullException("columnName");

            columnName = columnName.ToUpperInvariant();

            int sum = 0;

            for (int i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += (columnName[i] - 'A' + 1);
            }

            return sum;
        }
        public static string ColumnIndexToExcelColumnLetter(int colIndex)
        {
            int div = colIndex;
            string colLetter = String.Empty;
            int mod = 0;

            while (div > 0)
            {
                mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (int)((div - mod) / 26);
            }
            return colLetter;
        }
        public static bool FindPatternMatch(string orignalWord, List<string> possibleMatchList)
        {
            foreach (string word in possibleMatchList)
            {
                if (IsLike(orignalWord, word))
                {
                    return true;
                }
            }
            return false;
        }
    }
}

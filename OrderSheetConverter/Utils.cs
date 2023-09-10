using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using MessageBox = System.Windows.Forms.MessageBox;
using MessageBoxButtons = System.Windows.Forms.MessageBoxButtons;
using MessageBoxImage = System.Windows.Forms.MessageBoxIcon;

namespace Studio.DreamRoom.OrderSheetConverter
{
    internal class Utils
    {
        static internal void ShowErrorMessage(string message, string title)
        {
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxImage.Exclamation);
        }
        static internal void ShowSuccessMessage(string message, string title)
        {
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxImage.Information);
        }
    }

    internal static class IntExtension
    {
        internal static string ToSheetColumn(this int number)
        {
            var array = new LinkedList<int>();

            while (number > 26)
            {
                int value = number % 26;
                if (value == 0)
                {
                    number = number / 26 - 1;
                    array.AddFirst(26);
                }
                else
                {
                    number /= 26;
                    array.AddFirst(value);
                }
            }

            if (number > 0)
            {
                array.AddFirst(number);
            }

            return new string(array.Select(s => (char)('A' + s - 1)).ToArray());
        }

    }
}

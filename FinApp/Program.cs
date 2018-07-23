﻿using System;
using System.Windows.Forms;

namespace FinApp
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form1());
            }
            catch (System.Reflection.TargetInvocationException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}

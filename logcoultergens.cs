/*
 *
 *  < C > Sharp Library Class
 * 
 *  Objetivo : Ferramenta de Log
 * 
 *  Microsoft Visual C# 2010
 *  NET Framework 4.0
 * 
 *  Autor : Renato Igleziaz
 *          renato@esffera.com.br
 * 
 *  Data : 28 / 05 / 2012
 *  
 *  Versão 1.0
 * 
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;

namespace lab_interfaceamento.Drivers.coulter
{
    public class cGENS_logfile
    {
        private int lock_ = 0;
        private string path_ = "";
        private string componente_ = "";
        private string usuario_ = "";

        public void Initialize(string Path, string Componente, string Usuario)
        {
            // inicia
            lock_ = 1;
            path_ = Path;
            componente_ = Componente;
            usuario_ = Usuario;

            if (path_.PadRight(1).ToString() != @"\")
            {
                path_ += @"\";
            }
        }
        public bool Save(string LogTxT)
        {
            // grava dados no log
            if (lock_ != 1)
            {
                return false;
            }

            bool ret = false;

            try
            {
                using (StreamWriter w = File.AppendText(path_ + "labgen_log_coultergens.txt"))
                {
                    w.WriteLine("{0}-{1} - {2} - {3} - {4}",
                                DateTime.Now.ToLongTimeString(),
                                DateTime.Now.ToShortDateString(),
                                componente_,
                                usuario_,
                                LogTxT);
                    w.Close();
                }

                ret = true;
            }
            catch
            {
                ret = false;
            }

            return ret;
        }
    }
}

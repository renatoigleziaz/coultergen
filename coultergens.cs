/*
 * 
 *  Driver Coulter GEN-S
 * 
 *  por Renato Igleziaz
 *  em  25/08/2014
 *
 * 
 *  Fluxograma do serviço de transmissão de dados
 *  COULTER GEN-S
 * 
 *  Nomenclatura de caracteres de controle
 *  DC1  -> Controle de Dispositivo 
 *  CRLF -> Analisa linha dados recebidos
 * 
 *  Processos
 *  
 *  <DC1>08<CRLF>
 *   Inicio de Transmissão
 *   Contém a ID1<space><códigodebarras><CRLF>
 *  
 *  <DC1>0A<CRLF>
 *   Recebimento de dados
 *   Contém WBC<space><valor><space><R><CRLF> 
 *          RBC
 *          HGB
 *          HCT
 *          MCV
 *          MCH
 *          MCHC
 *          RDW
 *          PLT
 *          MPV
 *         
 *  <DC1>05<CRLF>        
 *   Recebimento de dados
 *   Contém LY#<space><valor><space><R><CRLF> 
 *          MO#
 *          NE#
 *          EO#
 *          BA#
 * 
 *  <DC1>05<CRLF>
 *   Recebimento de dados
 *   Contém LY%<space><valor><space><R><CRLF> 
 *          MO%
 *          NE%
 *          EO%
 *          BA%
 * 
 *  <DC1><CRLF>
 *   Fim de Transmissão
 *  
 */

#region imports
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Ports;
using System.IO;
using System.Windows.Forms;
using System.Data;
using System.ComponentModel;
using System.Threading;
using cDataObject;
using VirtualMachine40;
using NCalc;
using lab_rotina.Class;
#endregion

namespace lab_interfaceamento.Drivers.coulter
{
    #region Estrutura os parametros de configuração do equipamento
    public class settings
    {
        public int ID = -1;                 // id do equipamento
        public int BaudRate = 9600;         // velocidade da porta
        public string Parity = "None";      // Pariedade        
        public int DataBits = 0;            // DataBits
        public string StopBits = "One";     // StopBits        
        public string PortName = "COM0";    // endereço de comunicação da porta serial
    }
    #endregion

    #region Log System
    public class gens_logInfo
    {
        public string id;
        public int resultado;
        public DateTime ocorrencia;
        public string ocorrencia_descritivo;
        public string comando_solicitado;
    }

    public class gens_log
    {
        private ArrayList log_element = new ArrayList();

        public void Add(gens_logInfo exame)
        {
            // adiciona um item a coleção

            const string stx = "\x02"; // STX = 0x02
            const string etx = "\x03"; // ETX = 0x03
            const string eot = "\x04"; // EOT = 0x04
            const string enq = "\x05"; // ENQ = 0x05
            const string ack = "\x06"; // ACK = 0x06
            const string nak = "\x15"; // NAK = 0x15
            const string etb = "\x17"; // ETB = 0x17
            const string dc1 = "\x11"; // DC1 = 0x11 device control 1

            // se encontrar qualquer um desses caracteres de controle
            // substitui por exibição "humana"
            if (exame.comando_solicitado == stx)
                exame.comando_solicitado = "stx";
            else if (exame.comando_solicitado == etx)
                exame.comando_solicitado = "etx";
            else if (exame.comando_solicitado == eot)
                exame.comando_solicitado = "eot";
            else if (exame.comando_solicitado == enq)
                exame.comando_solicitado = "enq";
            else if (exame.comando_solicitado == ack)
                exame.comando_solicitado = "ack";
            else if (exame.comando_solicitado == nak)
                exame.comando_solicitado = "nak";
            else if (exame.comando_solicitado == etb)
                exame.comando_solicitado = "etb";
            else if (exame.comando_solicitado == dc1)
                exame.comando_solicitado = "dc1";

            log_element.Add(exame);
        }

        public gens_logInfo ListItem(int index)
        {
            // retorna o item
            return (gens_logInfo)log_element[index];
        }

        public int Count()
        {
            // retorna o total de itens da coleção
            return log_element.Count;
        }

        public void Clear()
        {
            // limpa toda a coleção
            log_element.Clear();
        }
    }
    #endregion

    #region state of the pins
    public class status
    {
        public bool CD;
        public bool CTS;
        public bool DSR;
        public bool DTR;
        public bool RTS;
    }
    #endregion

    #region STRU
    public class retGetParam
    {
        public int position = -1;
        public string body = "";
    }
    #endregion

    public class coultergens
    {
        #region Formato de retorno de Resultado

        public retGetParam GetParamEx(string body, string search)
        {
            // versão 2.0 da GetParam

            retGetParam returnf = new retGetParam();
            const string stx = "\x02"; // STX = 0x02
            const string etx = "\x03"; // ETX = 0x03

            int ret = body.IndexOf(etx + stx);
            int start = ret - 4;
            int fim = ret + 3;

            if (ret > -1)
            {
                for (int x = 0; x < body.Length; x++)
                {
                    if (x < start || x > fim)
                    {
                        returnf.body += body.Substring(x, 1);
                    }
                }
            }
            else
            {
                returnf.body = body;
            }

            returnf.position = returnf.body.IndexOf(search);
            return returnf;
        }

        public retGetParam GetParam(string body, string search)
        {
            retGetParam returnf = new retGetParam();

            int ret = body.IndexOf(search);

            if (ret == -1)
            {
                bool bingo = false;
                int count = 0;

                for (int x = 0; x < search.Length; x++)
                {
                    bingo = false;

                    ret = body.IndexOf(search.Substring(x, 1));
                    if (ret > -1)
                    {
                        bingo = true;
                    }

                    if (!bingo)
                    {
                        break;
                    }

                    count++;
                }

                if (count == 3)
                {
                    ret = body.IndexOf(" ");
                    if (ret == -1)
                    {
                        return returnf;
                    }
                    else
                    {
                        returnf.position = 1;
                        returnf.body = body.Substring(ret);
                        returnf.body = search + returnf.body;
                        return returnf;
                    }
                }
                else
                {
                    returnf.body = body;
                    returnf.position = -1;
                    return returnf;
                }
            }

            returnf.body = body;
            returnf.position = ret;
            return returnf;
        }

        public string GetField(ArrayList input, string position)
        {
            // obtem o conteúdo de um campo, baseado na posição

            string temp = "";
            int i = 0;

            foreach (string buffer in input)
            {
                i = buffer.IndexOf(position);

                if (i > -1)
                {
                    for (int x = (i + position.Length); x < buffer.Length; x++)
                    {
                        if (buffer.Substring(x, 1).Trim().Length > 0)
                        {
                            temp += buffer.Substring(x, 1);
                        }
                    }

                    temp = temp.Replace("R", "");
                    temp = temp.Replace("P", "");
                    temp = temp.Replace("*", "");
                    temp = temp.Replace("A", "");
                    temp = temp.Replace("a", "");
                    temp = temp.Replace("L", "");
                    temp = temp.Replace("l", "");
                    temp = temp.Replace("c", "");
                    temp = temp.Replace("C", "");
                    temp = temp.Replace("H", "");
                    temp = temp.Replace("h", "");
                    temp = temp.Replace("E", "");
                    temp = temp.Replace("e", "");
                    return temp;
                }
            }

            // não achou nada, retorna em branco
            return "";
        }

        #endregion

        #region Ambiente Variaveis
        // internas
        private const string nameofinterface = "Coulter GEN-S";
        private cDataBase mdb = new cDataBase();
        private settings setup = new settings();
        private string SQL = "";
        private DataRow dtEquip = null;
        private Thread thread;
        private bool threadCancelOP = false;
        private Queue<string> recievedData = new Queue<string>();
        private bool vm_state_thread = false;
        private gens_logInfo loginfo = null;
        private ListBox com_listbox = null;
        private Form com_form = null;
        private Label com_bar = null;
        private cGENS_logfile log_file = new cGENS_logfile();
        // internas da thread
        private string stringlido = "";
        private string chr_lido = "";
        private string lido = "";
        private bool is_eof = false;
        // externas
        public status vm_status = new status();
        public gens_log log = new gens_log();
        public bool status = false;
        public SerialPort porta = new SerialPort();
        #endregion

        #region Controle de Dados
        public const string stx = "\x02"; // STX = 0x02
        public const string etx = "\x03"; // ETX = 0x03
        public const string eot = "\x04"; // EOT = 0x04
        public const string enq = "\x05"; // ENQ = 0x05
        public const string ack = "\x06"; // ACK = 0x06
        public const string _lf = "\x0A"; //  LF = 0x0A
        public const string _cr = "\x0D"; //  CR = 0x0D
        public const string nak = "\x15"; // NAK = 0x15
        public const string etb = "\x17"; // ETB = 0x17
        public const string dc1 = "\x11"; // DC1 = 0x11 
        #endregion

        #region Inicializa
        public coultergens(string id_equipamento)
        {
            // banco de dados
            mdb.strProvider = Global.ProviderDb;
            mdb.AppPath = Global.PathApp;
            log_file.Initialize(Global.PathApp, "Interfaceamento", Global.Usuario);

            // carrega ambiente
            SQL = "SELECT * FROM EQUIPAMENTOS WHERE CODIGO=" + id_equipamento;
            cReturnDataTable tmpEquip = mdb.Abrir_DataTable(SQL);
            if (!tmpEquip.Status)
            {
                // retorna erro
                return;
            }

            if (tmpEquip.RecordCount == 0)
            {
                // retorna erro
                return;
            }

            try
            {
                dtEquip = tmpEquip.dt.Rows[0];
                setup.ID = int.Parse(id_equipamento);
                setup.PortName = dtEquip["PORTA"].ToString();
                setup.BaudRate = int.Parse(dtEquip["VELOCIDADE"].ToString());
                setup.Parity = dtEquip["PARIEDADE"].ToString();
                setup.DataBits = int.Parse(dtEquip["DATABITS"].ToString());
                setup.StopBits = dtEquip["STOPBITS"].ToString();
            }
            catch
            {
                // não conseguiu carregar o setup de configuração
                return;
            }

            // OK
            status = true;
        }
        #endregion

        #region Porta de Comunicação

        public bool OpenPort(settings optional = null)
        {
            // vars
            bool error = false;

            if (!status)
                return false;

            // se a porta estiver aberta
            if (porta.IsOpen)
                porta.Close();

            // flags de conexão padrão
            settings now = null;

            if (optional != null)
                now = optional;
            else
                now = setup;

            // velocidade da porta
            porta.BaudRate = now.BaudRate;
            // Pariedade
            porta.Parity = (Parity)Enum.Parse(typeof(Parity), now.Parity);
            // DataBits
            porta.DataBits = now.DataBits;
            // StopBits
            porta.StopBits = (StopBits)Enum.Parse(typeof(StopBits), now.StopBits);
            // endereço de comunicação da porta serial
            porta.PortName = now.PortName;

            try
            {
                porta.Open();
            }
            catch (UnauthorizedAccessException)
            {
                // log
                loginfo = new gens_logInfo();
                loginfo.id = "OpenPort()";
                loginfo.ocorrencia = DateTime.Now;
                loginfo.ocorrencia_descritivo = "UnauthorizedAccessException";
                loginfo.resultado = -1;
                log.Add(loginfo);

                error = true;
            }
            catch (IOException)
            {
                // log
                loginfo = new gens_logInfo();
                loginfo.id = "OpenPort()";
                loginfo.ocorrencia = DateTime.Now;
                loginfo.ocorrencia_descritivo = "IOException";
                loginfo.resultado = -1;
                log.Add(loginfo);

                error = true;
            }
            catch (ArgumentException)
            {
                // log
                loginfo = new gens_logInfo();
                loginfo.id = "OpenPort()";
                loginfo.ocorrencia = DateTime.Now;
                loginfo.ocorrencia_descritivo = "ArgumentException";
                loginfo.resultado = -1;
                log.Add(loginfo);

                error = true;
            }

            if (error)
            {
                // log
                loginfo = new gens_logInfo();
                loginfo.id = "OpenPort()";
                loginfo.ocorrencia = DateTime.Now;
                loginfo.ocorrencia_descritivo = "Não pode abrir a porta de comunicação.";
                loginfo.resultado = -1;
                log.Add(loginfo);
            }
            else
            {
                // registra primeiro pin state
                vm_status.CD = porta.CDHolding;
                vm_status.CTS = porta.CtsHolding;
                vm_status.DSR = porta.DsrHolding;
                vm_status.DTR = porta.DtrEnable;
                vm_status.RTS = porta.RtsEnable;
            }

            if (!porta.IsOpen)
                return false;

            if (error)
                return false;

            // comm event handler
            porta.DataReceived += new SerialDataReceivedEventHandler(port_DataReceived);

            return true;
        }

        public void ClosePort()
        {
            // fecha porta de comunicação

            try
            {
                if (porta.IsOpen)
                    porta.Close();
            }
            catch { }
        }

        private void port_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            // controla eventos de resposta da porta de comunicação

            try
            {
                recievedData.Enqueue(porta.ReadExisting());
            }
            catch { }
        }

        public string ReadBuffer()
        {
            // le o proximo buffer livre

            return (recievedData.Count > 0) ? recievedData.Dequeue() : "";
        }

        #endregion

        #region Multi-Thread / Fluxo de transmissão de dados

        public bool StartCOMM(Form thisForm,
                              Label thisBar = null,
                              ListBox thisListBox = null)
        {
            // inicia os processos de transmissão de dados
            // por meio de multi-processamento

            threadCancelOP = false;
            com_form = thisForm;
            com_listbox = thisListBox;
            com_bar = thisBar;

            try
            {
                thread = new Thread(new ThreadStart(run));
                thread.SetApartmentState(ApartmentState.STA);
                thread.IsBackground = true;
                thread.Priority = ThreadPriority.Normal;
                thread.Start();
            }
            catch
            {
                return false;
            }

            return true;
        }

        public bool CancelCOMM()
        {
            // para um serviço de transmissão

            try
            {
                threadCancelOP = true;
            }
            catch
            {
                return false;
            }

            return true;
        }

        private void AddText(string mensagem)
        {
            if (com_listbox == null)
                return;

            InvokeIfRequired(com_listbox, x =>
            {
                com_listbox.Items.Add(mensagem);

                if (com_listbox.Items.Count > 0)
                {
                    com_listbox.SelectedIndex = (com_listbox.Items.Count - 1);
                }
            });
        }

        private static void InvokeIfRequired(System.Windows.Forms.Control c, Action<System.Windows.Forms.Control> action)
        {
            if (c.InvokeRequired)
            {
                c.Invoke(new Action(() => action(c)));
            }
            else
            {
                action(c);
            }
        }

        public bool ThRead_State()
        {
            // repassa se o serviço ainda está executando
            return vm_state_thread;
        }

        private void run()
        {
            // abre a porta do equipamento
            if (!this.OpenPort(null))
            {
                InvokeIfRequired(com_bar, X => { com_bar.Text = "Erro abertura de porta"; });
                return;
            }
            else
                InvokeIfRequired(com_bar, X => { com_bar.Text = "Serviço Iniciado com Sucesso..."; });

            // marca que a thread iniciou corretamente
            vm_state_thread = true;

            // vars
            bool umavez = false;
            ArrayList body = new ArrayList();
            bool saveline = false;

            // laço principal de leitura
            for (; ; )
            {
                // cancela a operação caso seja solicitado                
                if (threadCancelOP)
                    break;

                // verifica a resposta do equipamento 
                stringlido = this.ReadBuffer();

                if (stringlido.Trim() != "")
                {
                    if (!umavez)
                    {
                        InvokeIfRequired(com_bar, X => { com_bar.Text = "Analisando Buffer..."; });
                        umavez = true;
                    }
                }

                for (int pos = 0; pos < stringlido.Length; pos++)
                {
                    // le caracter a caracter
                    chr_lido = stringlido.Substring(pos, 1);

                    #region (case) Checa byte recebido
                    switch (chr_lido)
                    {
                        case "":
                            {
                                System.Threading.Thread.Sleep(500);
                                break;
                            }
                        case _lf:
                            {
                                // fim de linha
                                is_eof = true;
                                break;
                            }
                        default:
                            {
                                if (lido.Trim().Length == 0)
                                {
                                    // inicia linha
                                    lido = "";
                                    is_eof = false;
                                    InvokeIfRequired(com_bar, X => { com_bar.Text = "Recebendo resultado (não desligue o sistema)..."; });
                                }
                                break;
                            }
                    } // end switch
                    #endregion

                    // acumula linha de mensagem
                    lido = lido + chr_lido;

                    if (is_eof)
                    {
                        // fim de linha
                        log_file.Save("Comando Recebido: " + lido.Replace(_cr, "").Replace(_lf, ""));

                        saveline = false;

                        if (lido.IndexOf("ID1") > -1)
                            saveline = true;
                        else if (lido.IndexOf("WBC") > -1)
                            saveline = true;
                        else if (lido.IndexOf("RBC") > -1)
                            saveline = true;
                        else if (lido.IndexOf("HGB") > -1)
                            saveline = true;
                        else if (lido.IndexOf("HCT") > -1)
                            saveline = true;
                        else if (lido.IndexOf("MCV") > -1)
                            saveline = true;
                        else if (lido.IndexOf("MCH") > -1)
                            saveline = true;
                        else if (lido.IndexOf("MCHC") > -1)
                            saveline = true;
                        else if (GetParamEx(lido, "RDW").position > -1)
                        {
                            lido = GetParamEx(lido, "RDW").body;
                            saveline = true;
                        }
                        else if (GetParam(lido, "PLT").position > -1)
                        {
                            lido = GetParam(lido, "PLT").body;
                            saveline = true;
                        }
                        else if (lido.IndexOf("MPV") > -1)
                            saveline = true;
                        else if (lido.IndexOf("LY#") > -1)
                            saveline = true;
                        else if (lido.IndexOf("MO#") > -1)
                            saveline = true;
                        else if (lido.IndexOf("NE#") > -1)
                            saveline = true;
                        else if (lido.IndexOf("EO#") > -1)
                            saveline = true;
                        else if (lido.IndexOf("BA#") > -1)
                            saveline = true;
                        else if (lido.IndexOf("LY%") > -1)
                            saveline = true;
                        else if (lido.IndexOf("MO%") > -1)
                            saveline = true;
                        else if (lido.IndexOf("NE%") > -1)
                            saveline = true;
                        else if (lido.IndexOf("EO%") > -1)
                            saveline = true;
                        else if (lido.IndexOf("BA%") > -1)
                            saveline = true;
                        else if (lido.IndexOf(dc1 + _cr + _lf) > -1)
                        {
                            // final de transmissão, grava resultado                            
                            this.GravaResultado(body);
                            body.Clear();
                        }

                        if (saveline)
                        {
                            lido = lido.Replace(_cr, "");
                            lido = lido.Replace(_lf, "");

                            // log file da leitura
                            log_file.Save("Decodificado: " + lido);

                            // add no array de dados reconhecidos como válidos
                            body.Add(lido);
                        }
                                                                        
                        // limpa buffer para uma nova linha
                        lido = "";
                        is_eof = false;

                    } // fim do (is_eof==true)

                } // fim (le byte a byte)

                System.Threading.Thread.Sleep(100);

            } // fim (laço principal de leitura)

            // encerra tudo
            vm_state_thread = false;
            this.ClosePort();

        } // fim da thread

        #endregion

        #region Transcreve resultados no banco de dados da interface

        public void GravaResultado(ArrayList _lidoinput)
        {
            // tarefas

            // obtem resultado
            // processa calculos primarios
            // executa as formulas
            // executa os valores de referencia
            // grava resultado do exame

            // specimen ID
            string codigo_barra = this.GetField(_lidoinput, "ID1").Trim();

            // separa a string
            string setor = "";
            string codigoposto = "";
            string amostrareal = "";
            cReturnDataTable tmprec;

            try
            {
                setor = codigo_barra.Substring(0, 2);
                codigoposto = codigo_barra.Substring(2, 2);
                amostrareal = codigo_barra.Substring(4, codigo_barra.Length - 4);
                amostrareal = int.Parse(amostrareal).ToString();
            }
            catch
            {
                setor = "";
                codigoposto = "";
                amostrareal = "";
                AddText("Erro: amostra não encontrada!");
                return;
            }

            if (setor.Trim() == "" || codigoposto.Trim() == "" || amostrareal.Trim() == "")
            {
                AddText("Erro: amostra não encontrada!");
                return;
            }

            // levanta os exames que deverão ser feitos
            SQL = "SELECT ";
            SQL += "AMOSTRAS.CODIGO_AMOSTRA, ";
            SQL += "AMOSTRAS_EXAMES.CODIGO_EXAME, ";
            SQL += "AMOSTRAS_EXAMES.SEQUENCIA, ";
            SQL += "EXAMES.CODIGO, ";
            SQL += "PACIENTE.NASCIMENTO, ";
            SQL += "PACIENTE.SEXO ";
            SQL += "FROM AMOSTRAS ";
            SQL += "INNER JOIN AMOSTRAS_EXAMES ON ";
            SQL += "AMOSTRAS.CODIGO_AMOSTRA=AMOSTRAS_EXAMES.CODIGO_AMOSTRA ";
            SQL += "INNER JOIN EXAMES ON ";
            SQL += "EXAMES.CODIGO_INT=AMOSTRAS_EXAMES.CODIGO_EXAME ";
            SQL += "INNER JOIN PACIENTE ON ";
            SQL += "PACIENTE.CODIGO=AMOSTRAS.CODIGO_PACIENTE ";
            SQL += "WHERE AMOSTRAS.CODIGO_POSTOCOLETA=" + int.Parse(codigoposto).ToString() + " ";
            SQL += "AND AMOSTRAS.AMOSTRA='" + amostrareal + "' ";
            SQL += "AND EXAMES.GER_SETCOD=" + setor + " ";
            tmprec = mdb.Abrir_DataTable(SQL);
            if (!tmprec.Status)
                return;

            cReturnDataTable campos;
            cInsertInto inc = new cInsertInto();
            int seq = 0;
            string campo = "";
            string resultado = "";
            decimal dec_resultado = 0;
            cReturnDataTable valref;
            Misc misc = new Misc();
            calcIdade calcidade = new calcIdade();
            calcIdade.resultIdade resultidade;
            bool bingo = false;
            int idade_value = 0;
            int idade_min = 0;
            int idade_max = 0;
            bool resultado_normal = false;
            bool resultado_alterado = false;
            bool resultado_absurdo = false;
            decimal value = 0;
            decimal min_value = 0;
            decimal max_value = 0;

            foreach (DataRow dt in tmprec.dt.Rows)
            {
                seq = 1;

                #region remove duplicado
                SQL = "DELETE FROM INTERFACE_CAMPOS ";
                SQL += "WHERE CODIGO_EQUIPA=" + setup.ID.ToString() + " ";
                SQL += "AND CODIGO_AMOSTRA='" + dt["CODIGO_AMOSTRA"].ToString() + "' ";
                SQL += "AND CODIGO_EXAME=" + dt["CODIGO_EXAME"].ToString() + " ";
                if (!mdb.SQLExecute(SQL))
                    return;
                #endregion

                #region localiza todos os campos que irão receber dados

                SQL = "SELECT ";
                SQL += "CODIGO_ENVIO, CODIGO_RESPOSTA, ID, ACAO, PRIMEIROCALCULO, DECIMAL ";
                SQL += "FROM EQUIPAMENTOS_EXM_INTERFACE ";
                SQL += "WHERE CODIGO_EQUIPA=" + setup.ID.ToString() + " ";
                SQL += "AND CODIGO_EXAME=" + dt["CODIGO_EXAME"].ToString() + " ";
                SQL += "ORDER BY ORDEM ";
                campos = mdb.Abrir_DataTable(SQL);
                if (!campos.Status)
                    return;

                AddText("Gravando a amostra '" + codigo_barra + "', Exame '" + dt["CODIGO"].ToString() + "'");

                foreach (DataRow dtcampos in campos.dt.Rows)
                {
                    // obtem o resultado de cada campo
                    // e adiciona na base

                    try
                    {
                        if (dtcampos["ACAO"].ToString() == "Recebe como Parametro")
                        {
                            // tenta converter o CODIGO_RESPOSTA em
                            // formato integer para ser usado
                            // no caso da celldyn como indice de resultado
                            campo = dtcampos["CODIGO_RESPOSTA"].ToString();

                            if (campo.Trim().Length == 0)
                            {
                                resultado = "";
                            }
                            else
                            {
                                // obtem resultado
                                resultado = this.GetField(_lidoinput, campo).Trim();
                            }

                            // faz o primeiro ajuste de calculo
                            if (dtcampos["PRIMEIROCALCULO"].ToString() != "")
                            {
                                dec_resultado = this.Calculate(dtcampos["PRIMEIROCALCULO"].ToString().Replace("<result>", resultado));
                                resultado = dec_resultado.ToString();
                            }

                            if (resultado == "")
                                resultado = "0";

                            // formata o resultado
                            resultado = this.FormatEx(resultado.Replace(".", ","), int.Parse(dtcampos["DECIMAL"].ToString()));

                            resultado = resultado.Replace(",", ".");

                            inc.Clear();
                            inc.Add("CODIGO_EQUIPA", setup.ID.ToString(), false);
                            inc.Add("CODIGO_AMOSTRA", dt["CODIGO_AMOSTRA"].ToString(), true);
                            inc.Add("CODIGO_EXAME", dt["CODIGO_EXAME"].ToString(), false);
                            inc.Add("ID_CAMPO", dtcampos["ID"].ToString(), false);
                            inc.Add("R_ORDER", "1", false);
                            inc.Add("R_SEQ", seq.ToString(), false);
                            inc.Add("R_UNIVERSALID", dtcampos["CODIGO_RESPOSTA"].ToString(), true);
                            inc.Add("R_RESULT", resultado.Replace(",", "."), true);
                            inc.Add("R_UNIT", "", true);
                            inc.Add("R_REF", "", true);
                            inc.Add("R_FLAG", "", true);
                            inc.Add("R_STATUS", "", true);
                            inc.Add("R_DATA", DateTime.Now.ToString("MM/dd/yyyy HH:mm"), true);
                            SQL = inc.InsertInto("INTERFACE_CAMPOS");
                            if (!mdb.SQLExecute(SQL))
                                return;

                            seq++;
                        }
                        else if (dtcampos["ACAO"].ToString() == "Calcula")
                        {
                            // grava o espaço para calcular posteriormente

                            inc.Clear();
                            inc.Add("CODIGO_EQUIPA", setup.ID.ToString(), false);
                            inc.Add("CODIGO_AMOSTRA", dt["CODIGO_AMOSTRA"].ToString(), true);
                            inc.Add("CODIGO_EXAME", dt["CODIGO_EXAME"].ToString(), false);
                            inc.Add("ID_CAMPO", dtcampos["ID"].ToString(), false);
                            inc.Add("R_ORDER", "1", false);
                            inc.Add("R_SEQ", seq.ToString(), false);
                            inc.Add("R_UNIVERSALID", dtcampos["CODIGO_RESPOSTA"].ToString(), true);
                            inc.Add("R_RESULT", "", true);
                            inc.Add("R_UNIT", "", true);
                            inc.Add("R_REF", "", true);
                            inc.Add("R_FLAG", "", true);
                            inc.Add("R_STATUS", "", true);
                            inc.Add("R_DATA", DateTime.Now.ToString("MM/dd/yyyy HH:mm"), true);
                            SQL = inc.InsertInto("INTERFACE_CAMPOS");
                            if (!mdb.SQLExecute(SQL))
                                return;

                            seq++;
                        }
                        else if (dtcampos["ACAO"].ToString() == "Observação")
                        {
                            // grava o espaço para a OBS

                            inc.Clear();
                            inc.Add("CODIGO_EQUIPA", setup.ID.ToString(), false);
                            inc.Add("CODIGO_AMOSTRA", dt["CODIGO_AMOSTRA"].ToString(), true);
                            inc.Add("CODIGO_EXAME", dt["CODIGO_EXAME"].ToString(), false);
                            inc.Add("ID_CAMPO", dtcampos["ID"].ToString(), false);
                            inc.Add("R_ORDER", "1", false);
                            inc.Add("R_SEQ", seq.ToString(), false);
                            inc.Add("R_UNIVERSALID", dtcampos["CODIGO_RESPOSTA"].ToString(), true);
                            inc.Add("R_RESULT", "", true);
                            inc.Add("R_UNIT", "", true);
                            inc.Add("R_REF", "", true);
                            inc.Add("R_FLAG", "", true);
                            inc.Add("R_STATUS", "", true);
                            inc.Add("R_DATA", DateTime.Now.ToString("MM/dd/yyyy HH:mm"), true);
                            SQL = inc.InsertInto("INTERFACE_CAMPOS");
                            if (!mdb.SQLExecute(SQL))
                                return;

                            seq++;
                        }
                    }
                    catch { }

                } // end dtcampos
                #endregion

                #region agora com os campos criados, executa as fórmulas e valores de referencia

                SQL = "SELECT ";
                SQL += "EQUIPAMENTOS_EXM_INTERFACE.DECIMAL,";
                SQL += "EQUIPAMENTOS_EXM_INTERFACE.ACAO, ";
                SQL += "EQUIPAMENTOS_EXM_INTERFACE.CALCULO, ";
                SQL += "INTERFACE_CAMPOS.R_RESULT, ";
                SQL += "INTERFACE_CAMPOS.ID_CAMPO, ";
                SQL += "INTERFACE_CAMPOS.R_ORDER, ";
                SQL += "INTERFACE_CAMPOS.R_SEQ ";
                SQL += "FROM INTERFACE_CAMPOS ";
                SQL += "INNER JOIN EQUIPAMENTOS_EXM_INTERFACE ON ";
                SQL += "EQUIPAMENTOS_EXM_INTERFACE.ID=INTERFACE_CAMPOS.ID_CAMPO ";
                SQL += "WHERE INTERFACE_CAMPOS.CODIGO_AMOSTRA='" + dt["CODIGO_AMOSTRA"].ToString() + "' ";
                SQL += "AND INTERFACE_CAMPOS.CODIGO_EQUIPA=" + setup.ID.ToString() + " ";
                SQL += "AND INTERFACE_CAMPOS.CODIGO_EXAME=" + dt["CODIGO_EXAME"].ToString() + " ";
                SQL += "ORDER BY ";
                SQL += "INTERFACE_CAMPOS.CODIGO_EXAME, ";
                SQL += "INTERFACE_CAMPOS.R_ORDER, ";
                SQL += "INTERFACE_CAMPOS.R_SEQ ";
                campos = mdb.Abrir_DataTable(SQL);
                if (!campos.Status)
                    return;

                resultado = "";
                dec_resultado = 0;

                foreach (DataRow dtcalc in campos.dt.Rows)
                {
                    resultado = dtcalc["R_RESULT"].ToString();

                    if (dtcalc["ACAO"].ToString() == "Calcula")
                    {
                        if (dtcalc["CALCULO"].ToString().Trim().Length > 0)
                        {
                            // gera calculo
                            dec_resultado = PreparaFormula(dtcalc["CALCULO"].ToString(),
                                                           dt["CODIGO_AMOSTRA"].ToString(),
                                                           setup.ID.ToString(),
                                                           dt["CODIGO_EXAME"].ToString()
                                                           );

                            resultado = dec_resultado.ToString();

                            // formata resultado
                            resultado = FormatEx(dec_resultado.ToString(), int.Parse(dtcalc["DECIMAL"].ToString()));

                            // atualiza resultado base de dados
                            SQL = "WHERE INTERFACE_CAMPOS.CODIGO_AMOSTRA='" + dt["CODIGO_AMOSTRA"].ToString() + "' ";
                            SQL += "AND INTERFACE_CAMPOS.CODIGO_EQUIPA=" + setup.ID.ToString() + " ";
                            SQL += "AND INTERFACE_CAMPOS.CODIGO_EXAME=" + dt["CODIGO_EXAME"].ToString() + " ";
                            SQL += "AND R_ORDER=" + dtcalc["R_ORDER"].ToString() + " ";
                            SQL += "AND R_SEQ=" + dtcalc["R_SEQ"].ToString() + " ";

                            resultado = resultado.Replace(",", ".");

                            inc.Clear();
                            inc.Add("R_RESULT", resultado, true);
                            inc.Add("R_REF", (dtcalc["CALCULO"].ToString().Length > 100) ? dtcalc["CALCULO"].ToString().Substring(0, 100) : dtcalc["CALCULO"].ToString(), true);
                            SQL = inc.Update("INTERFACE_CAMPOS", SQL);
                            if (!mdb.SQLExecute(SQL))
                                return;
                        }
                    }

                    // valores de referencia

                    if (dt["NASCIMENTO"].ToString() != "")
                    {
                        // obtem a data de nascimento do paciente em dias, meses e anos
                        resultidade = calcidade.CalculateIdade(DateTime.Parse(dt["NASCIMENTO"].ToString()));

                        // valores de referencia apenas para Campos de retorno e Calculos
                        if (dtcalc["ACAO"].ToString() == "Calcula" || dtcalc["ACAO"].ToString() == "Recebe como Parametro")
                        {
                            SQL = "SELECT ";
                            SQL += "* ";
                            SQL += "FROM EQUIPAMENTOS_EXM_VALREF ";
                            SQL += "WHERE ID=" + dtcalc["ID_CAMPO"].ToString() + " ";

                            if (dt["SEXO"].ToString() == "Masculino")
                            {
                                SQL += "AND SEXO IN('A', 'M') ";
                            }
                            else if (dt["SEXO"].ToString() == "Feminino")
                            {
                                SQL += "AND SEXO IN('A', 'F') ";
                            }
                            else
                            {
                                SQL += "AND SEXO='A' ";
                            }

                            SQL += "ORDER BY SEQ ";
                            valref = mdb.Abrir_DataTable(SQL);
                            if (!valref.Status)
                                return;

                            // lemos todos os valores de referencia
                            foreach (DataRow dtvlref in valref.dt.Rows)
                            {
                                bingo = false;

                                idade_value = (resultidade.Anos * 365) + (resultidade.Meses * 30) + resultidade.Dias;
                                idade_min = (int.Parse(dtvlref["ANO_INI"].ToString()) * 365) + (int.Parse(dtvlref["MES_INI"].ToString()) * 30) + int.Parse(dtvlref["DIA_INI"].ToString());
                                idade_max = (int.Parse(dtvlref["ANO_FIM"].ToString()) * 365) + (int.Parse(dtvlref["MES_FIM"].ToString()) * 30) + int.Parse(dtvlref["DIA_FIM"].ToString());

                                if (idade_value >= idade_min && idade_value <= idade_max)
                                    bingo = true;

                                resultado = resultado.Replace(".", ",");

                                if (bingo)
                                {
                                    resultado_normal = false;

                                    value = decimal.Parse(resultado);
                                    min_value = decimal.Parse(dtvlref["NORMAL_MIN"].ToString());
                                    max_value = decimal.Parse(dtvlref["NORMAL_MAX"].ToString());

                                    // verifica resultado normal
                                    if (value >= min_value && value <= max_value)
                                    {
                                        resultado_normal = true;

                                        // atualiza resultado base de dados
                                        SQL = "WHERE INTERFACE_CAMPOS.CODIGO_AMOSTRA='" + dt["CODIGO_AMOSTRA"].ToString() + "' ";
                                        SQL += "AND INTERFACE_CAMPOS.CODIGO_EQUIPA=" + setup.ID.ToString() + " ";
                                        SQL += "AND INTERFACE_CAMPOS.CODIGO_EXAME=" + dt["CODIGO_EXAME"].ToString() + " ";
                                        SQL += "AND R_ORDER=" + dtcalc["R_ORDER"].ToString() + " ";
                                        SQL += "AND R_SEQ=" + dtcalc["R_SEQ"].ToString() + " ";

                                        inc.Clear();
                                        inc.Add("R_REF", dtvlref["NORMAL_MIN"].ToString() + " TO " + dtvlref["NORMAL_MAX"].ToString(), true);
                                        inc.Add("R_FLAG", "N", true);
                                        SQL = inc.Update("INTERFACE_CAMPOS", SQL);
                                        if (!mdb.SQLExecute(SQL))
                                            return;

                                        // atualiza observação
                                        SQL = "WHERE INTERFACE_CAMPOS.CODIGO_AMOSTRA='" + dt["CODIGO_AMOSTRA"].ToString() + "' ";
                                        SQL += "AND INTERFACE_CAMPOS.CODIGO_EQUIPA=" + setup.ID.ToString() + " ";
                                        SQL += "AND INTERFACE_CAMPOS.CODIGO_EXAME=" + dt["CODIGO_EXAME"].ToString() + " ";
                                        SQL += "AND ID_CAMPO=" + dtvlref["CAMPOOBS"].ToString() + " ";

                                        inc.Clear();
                                        inc.Add("R_RESULT", dtvlref["NORMAL_OBS"].ToString(), true);
                                        SQL = inc.Update("INTERFACE_CAMPOS", SQL);
                                        if (!mdb.SQLExecute(SQL))
                                            return;
                                    }

                                    // verifica resultado alterado baixo
                                    value = decimal.Parse(resultado);
                                    min_value = decimal.Parse(dtvlref["NORMAL_MIN"].ToString());

                                    if (value < min_value)
                                    {
                                        if (!resultado_normal)
                                        {
                                            resultado_alterado = true;

                                            // atualiza resultado base de dados
                                            SQL = "WHERE INTERFACE_CAMPOS.CODIGO_AMOSTRA='" + dt["CODIGO_AMOSTRA"].ToString() + "' ";
                                            SQL += "AND INTERFACE_CAMPOS.CODIGO_EQUIPA=" + setup.ID.ToString() + " ";
                                            SQL += "AND INTERFACE_CAMPOS.CODIGO_EXAME=" + dt["CODIGO_EXAME"].ToString() + " ";
                                            SQL += "AND R_ORDER=" + dtcalc["R_ORDER"].ToString() + " ";
                                            SQL += "AND R_SEQ=" + dtcalc["R_SEQ"].ToString() + " ";

                                            inc.Clear();
                                            inc.Add("R_REF", dtvlref["NORMAL_MIN"].ToString() + " TO " + dtvlref["NORMAL_MAX"].ToString(), true);
                                            inc.Add("R_FLAG", "<", true);
                                            SQL = inc.Update("INTERFACE_CAMPOS", SQL);
                                            if (!mdb.SQLExecute(SQL))
                                                return;

                                            // atualiza observação
                                            SQL = "WHERE INTERFACE_CAMPOS.CODIGO_AMOSTRA='" + dt["CODIGO_AMOSTRA"].ToString() + "' ";
                                            SQL += "AND INTERFACE_CAMPOS.CODIGO_EQUIPA=" + setup.ID.ToString() + " ";
                                            SQL += "AND INTERFACE_CAMPOS.CODIGO_EXAME=" + dt["CODIGO_EXAME"].ToString() + " ";
                                            SQL += "AND ID_CAMPO=" + dtvlref["CAMPOOBS"].ToString() + " ";

                                            inc.Clear();
                                            inc.Add("R_RESULT", dtvlref["ALTERADO_MENOS"].ToString(), true);
                                            SQL = inc.Update("INTERFACE_CAMPOS", SQL);
                                            if (!mdb.SQLExecute(SQL))
                                                return;
                                        }
                                    }

                                    // verifica resultado alterado cima
                                    value = decimal.Parse(resultado);
                                    max_value = decimal.Parse(dtvlref["NORMAL_MAX"].ToString());

                                    if (value > max_value)
                                    {
                                        if (!resultado_normal)
                                        {
                                            resultado_alterado = true;

                                            // atualiza resultado base de dados
                                            SQL = "WHERE INTERFACE_CAMPOS.CODIGO_AMOSTRA='" + dt["CODIGO_AMOSTRA"].ToString() + "' ";
                                            SQL += "AND INTERFACE_CAMPOS.CODIGO_EQUIPA=" + setup.ID.ToString() + " ";
                                            SQL += "AND INTERFACE_CAMPOS.CODIGO_EXAME=" + dt["CODIGO_EXAME"].ToString() + " ";
                                            SQL += "AND R_ORDER=" + dtcalc["R_ORDER"].ToString() + " ";
                                            SQL += "AND R_SEQ=" + dtcalc["R_SEQ"].ToString() + " ";

                                            inc.Clear();
                                            inc.Add("R_REF", dtvlref["NORMAL_MIN"].ToString() + " TO " + dtvlref["NORMAL_MAX"].ToString(), true);
                                            inc.Add("R_FLAG", ">", true);
                                            SQL = inc.Update("INTERFACE_CAMPOS", SQL);
                                            if (!mdb.SQLExecute(SQL))
                                                return;

                                            // atualiza observação
                                            SQL = "WHERE INTERFACE_CAMPOS.CODIGO_AMOSTRA='" + dt["CODIGO_AMOSTRA"].ToString() + "' ";
                                            SQL += "AND INTERFACE_CAMPOS.CODIGO_EQUIPA=" + setup.ID.ToString() + " ";
                                            SQL += "AND INTERFACE_CAMPOS.CODIGO_EXAME=" + dt["CODIGO_EXAME"].ToString() + " ";
                                            SQL += "AND ID_CAMPO=" + dtvlref["CAMPOOBS"].ToString() + " ";

                                            inc.Clear();
                                            inc.Add("R_RESULT", dtvlref["ALTERADO_MAIS"].ToString(), true);
                                            SQL = inc.Update("INTERFACE_CAMPOS", SQL);
                                            if (!mdb.SQLExecute(SQL))
                                                return;
                                        }
                                    }

                                    // verifica resultado absurdo
                                    value = decimal.Parse(resultado);
                                    min_value = decimal.Parse(dtvlref["ABS_MIN"].ToString());
                                    max_value = decimal.Parse(dtvlref["ABS_MAX"].ToString());

                                    if (min_value != 0 && max_value != 0)
                                    {
                                        if (value >= min_value && value <= max_value)
                                        {
                                            resultado_absurdo = true;

                                            // atualiza resultado base de dados
                                            SQL = "WHERE INTERFACE_CAMPOS.CODIGO_AMOSTRA='" + dt["CODIGO_AMOSTRA"].ToString() + "' ";
                                            SQL += "AND INTERFACE_CAMPOS.CODIGO_EQUIPA=" + setup.ID.ToString() + " ";
                                            SQL += "AND INTERFACE_CAMPOS.CODIGO_EXAME=" + dt["CODIGO_EXAME"].ToString() + " ";
                                            SQL += "AND R_ORDER=" + dtcalc["R_ORDER"].ToString() + " ";
                                            SQL += "AND R_SEQ=" + dtcalc["R_SEQ"].ToString() + " ";

                                            inc.Clear();
                                            inc.Add("R_REF", dtvlref["NORMAL_MIN"].ToString() + " TO " + dtvlref["NORMAL_MAX"].ToString(), true);
                                            inc.Add("R_FLAG", "A", true);
                                            SQL = inc.Update("INTERFACE_CAMPOS", SQL);
                                            if (!mdb.SQLExecute(SQL))
                                                return;
                                        }
                                    }

                                    // encerra
                                    break;
                                }

                            } // end valores de referencia do campo

                        } // end valores de ref

                    } // end data de nascimento

                } // end leitura de campos

                // registra status geral do resultado do exame
                if (resultado_alterado && !resultado_absurdo)
                {
                    // alterado
                    SQL = "WHERE CODIGO_AMOSTRA='" + dt["CODIGO_AMOSTRA"].ToString() + "' ";
                    SQL += "AND EQUIPAMENTO=" + setup.ID.ToString() + " ";
                    SQL += "AND SEQUENCIA=" + dt["SEQUENCIA"].ToString() + " ";

                    inc.Clear();
                    inc.Add("STATUS", "Alterado", true);
                    SQL = inc.Update("INTERFACE_BOT", SQL);
                    if (!mdb.SQLExecute(SQL))
                        return;
                }
                else
                {
                    if (resultado_absurdo)
                    {
                        // absurdo
                        SQL = "WHERE CODIGO_AMOSTRA='" + dt["CODIGO_AMOSTRA"].ToString() + "' ";
                        SQL += "AND EQUIPAMENTO=" + setup.ID.ToString() + " ";
                        SQL += "AND SEQUENCIA=" + dt["SEQUENCIA"].ToString() + " ";

                        inc.Clear();
                        inc.Add("STATUS", "Absurdo", true);
                        SQL = inc.Update("INTERFACE_BOT", SQL);
                        if (!mdb.SQLExecute(SQL))
                            return;
                    }
                    else
                    {
                        // normal
                        SQL = "WHERE CODIGO_AMOSTRA='" + dt["CODIGO_AMOSTRA"].ToString() + "' ";
                        SQL += "AND EQUIPAMENTO=" + setup.ID.ToString() + " ";
                        SQL += "AND SEQUENCIA=" + dt["SEQUENCIA"].ToString() + " ";

                        inc.Clear();
                        inc.Add("STATUS", "Normal", true);
                        SQL = inc.Update("INTERFACE_BOT", SQL);
                        if (!mdb.SQLExecute(SQL))
                            return;
                    }
                }

                #endregion

            } // end dt amostras           
        }

        public decimal PreparaFormula(string pre_calculo, string codigo_amostra, string codigo_equipa, string codigo_exame)
        {
            // pega o calculo ainda com máscaras e substitui pelos valores verdadeiros

            string retorno = pre_calculo;

            SQL = "SELECT ";
            SQL += "EQUIPAMENTOS_EXM_INTERFACE.VAR_CALCULO, ";
            SQL += "INTERFACE_CAMPOS.R_RESULT ";
            SQL += "FROM INTERFACE_CAMPOS ";
            SQL += "INNER JOIN EQUIPAMENTOS_EXM_INTERFACE ON ";
            SQL += "EQUIPAMENTOS_EXM_INTERFACE.ID=INTERFACE_CAMPOS.ID_CAMPO ";
            SQL += "WHERE INTERFACE_CAMPOS.CODIGO_AMOSTRA='" + codigo_amostra + "' ";
            SQL += "AND INTERFACE_CAMPOS.CODIGO_EQUIPA=" + codigo_equipa + " ";
            SQL += "AND INTERFACE_CAMPOS.CODIGO_EXAME=" + codigo_exame + " ";
            SQL += "ORDER BY  ";
            SQL += "INTERFACE_CAMPOS.R_ORDER, ";
            SQL += "INTERFACE_CAMPOS.R_SEQ ";
            cReturnDataTable tmprec = mdb.Abrir_DataTable(SQL);
            if (!tmprec.Status)
                return 0;

            // troca as variaveis pelos valores reais
            foreach (DataRow dt in tmprec.dt.Rows)
            {
                if (dt["VAR_CALCULO"].ToString().Trim().Length > 0)
                {
                    retorno = retorno.Replace(dt["VAR_CALCULO"].ToString(), dt["R_RESULT"].ToString());
                }
            }

            // calcula realmente
            return (retorno.Length > 0) ? this.Calculate(retorno) : 0;
        }

        public string FormatEx(string input, int qtdecimal)
        {
            // formata resultado de acordo com qtd de casas decimais

            Misc misc = new Misc();
            string format = "0"; // #.###0

            if (qtdecimal > 0)
            {
                for (int x = 1; x < qtdecimal; x++)
                {
                    format = "#" + format;
                }

                format = "0." + format;

                return misc.FormatMoney(input, format);
            }
            else
            {
                return misc.FormatInt(input);
            }
        }

        public decimal Calculate(string formula)
        {
            decimal result = 0;

            try
            {
                // metodo B
                // base projeto: http://ncalc.codeplex.com/
                // funções: http://ncalc.codeplex.com/wikipage?title=functions&referringTitle=Home
                //

                Expression runFormulaB = new Expression(formula);
                result = decimal.Parse(runFormulaB.Evaluate().ToString());
                runFormulaB = null;
            }
            catch { }

            return result;
        }

        #endregion


    }
}

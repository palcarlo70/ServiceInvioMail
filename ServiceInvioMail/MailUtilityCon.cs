using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Net.Mail;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ServiceInvioMail
{
    public class MailUtilityCon
    {
        private Task _proccessSmsQueueTask;

        public MailDto GetMailImpo(int idImpo, string conNection)
        {
            MailDto lst = new MailDto();
            try
            {
                var con = new MaylUtilityDac("System.Data.SqlClient", conNection);
                var ds = con.GetMailImpo(idImpo);
                lst = new MailDto();
                lst = (from DataRow dr in ds.Tables[0].Rows
                       select new MailDto()
                       {
                           Destinatario = dr["ToRiceved"].ToString(),
                           DestinatarioLst = dr["MailLst"].ToString(),
                           Cc = dr["Cc"].ToString(),
                           Ccn = dr["CCn"].ToString(),
                           Oggetto = dr["Oggetto"].ToString(),
                           Messaggio = "",
                           UtenteFiguraInvio = dr["IdirizzoDiInvio"].ToString(),
                           NominativoInvio = dr["DenominazioneInvio"].ToString(),
                           UserMail = dr["MailConfig"].ToString(),
                           Password = dr["Pwd"].ToString(),
                           SsLAuto = !dr.IsNull("SsLAuto") ? Convert.ToBoolean(dr["SsLAuto"].ToString()) : (bool?)null,
                           UseDefaultCredential = !dr.IsNull("UseDefaultCredential") ? Convert.ToBoolean(dr["UseDefaultCredential"].ToString()) : (bool?)null,
                           SmtPostaUscita = dr["SmtPostaUscita"].ToString(),
                           SmtpPort = Convert.ToInt32(dr["SmtpPort"].ToString())
                       }).ToList().FirstOrDefault();

            }
            catch (Exception ex)
            {
                StackTrace st = new StackTrace();
                StackFrame sf = st.GetFrame(0);

                MethodBase currentMethodName = sf.GetMethod();
                string errore = $"Funzione {currentMethodName.Name}; Errore: {ex.Message}";
                var dex = new DataException(ex.Message);

            }
            return lst;
        }

        public bool ControlloCarenzeMagazzino(string conNection, string percorso) //string percorso = Server.MapPath("~/");
        {
            var conn = new MailUtilityCon();
            var lst = conn.GetMaterialiMancantiCon(conNection);
            if (lst.Count > 0)
            {
                _proccessSmsQueueTask = Task.Run(() => DoWorkAsync(lst, percorso, lst.Count, conNection));
            }
            return true;
        }

        public bool ControlloInvioMail(string conNection) //string percorso = Server.MapPath("~/");
        {
            var conn = new MailUtilityCon();
            var lst = conn.GetLstInvioMailDaInviare(conNection);

            if (lst.Count > 0)
            {
                _proccessSmsQueueTask = Task.Run(() => DoWorkSendMailAvvisiAsync(lst, conNection));
            }

            return true;
        }

        public async Task DoWorkSendMailAvvisiAsync(List<AvvisiDaInviare> lst, string conNection)
        {
            try
            {
                var impoMail = GetMailImpo(1, conNection);
                string textBody = "";

                foreach (var l in lst)
                {
                    impoMail.Oggetto = l.Oggetto;
                    impoMail.Messaggio = l.Messaggio;                             
                    impoMail.Destinatario = l.Destinatario;
                    await SendMailAsyncNew(impoMail, textBody, 0, 1, conNection);
                    

                }

                
            }
            catch (Exception e)
            {
                //salvo il log della spedizione effettuata
                var conMail = new MaylUtilityDac("System.Data.SqlClient", conNection);
                conMail.SaveMailLog($"ERRORE: {e.Message} - SOURCE: {e.Source}", 0, 1);

            }

        }

        public async Task DoWorkAsync(List<Articoli> lst, string percorso, int numRecord, string conNection)
        {
            try
            {
                /*<h2 >Articoli carenti INTERNI  </h2>*/

                var impoMail = GetMailImpo(1, conNection);
                string textBody = "<h2 >ARTICOLI LAVORATI carenti </h2> <br/> <table border=" + 1 + " cellpadding=" + 0 + " cellspacing=" + 0 + " width = " + 700 + " style='border: 0.5px;'><tr bgcolor='#4da6ff'><td style='width:15%; text-align: center;'><b>Cod Articolo</b></td> <td style='text-align: center;'> <b> Descrizione</b> </td><td style='text-align: center;'><b>Min Maga</b></td> <td style='text-align: center;'> <b> Q.T.</b> </td></tr>";

                foreach (var l in lst.Where(c => c.ArtInternoEsterno == 1))
                {
                    textBody += "<tr><td>" + l.IdArticolo + "</td><td> " + l.DescriArticolo + "</td><td>" + l.MinMagazzino + "</td><td> " + l.QuantiInMagazzino + "</td> </tr>";
                }

                textBody += "</table> <br/>";


                textBody += "<h2 >MATERIALI di CONSUMO carenti </h2> <br/> <table border=" + 1 + " cellpadding=" + 0 + " cellspacing=" + 0 + " width = " + 700 + " style='border: 0.5px;'><tr bgcolor='#4da6ff'><td style='width:15%; text-align: center;'><b>Cod Articolo</b></td> <td style='text-align: center;'> <b> Descrizione</b> </td><td style='text-align: center;'><b>Min Maga</b></td> <td style='text-align: center;'> <b> Q.T.</b> </td></tr>";

                foreach (var l in lst.Where(c => c.ArtInternoEsterno == 0))
                {
                    textBody += "<tr><td>" + l.IdArticolo + "</td><td> " + l.DescriArticolo + "</td><td>" + l.MinMagazzino + "</td><td> " + l.QuantiInMagazzino + "</td> </tr>";
                }

                textBody += "</table>";

                await SendMailAsyncNew(impoMail, textBody, numRecord,0, conNection);
            }
            catch (Exception e)
            {
                //salvo il log della spedizione effettuata
                var conMail = new MaylUtilityDac("System.Data.SqlClient", conNection);
                conMail.SaveMailLog($"ERRORE: {e.Message} - SOURCE: {e.Source}", 0, 1);
            }

        }


        public async Task SendMailPromemoria(MailDto impoMail, string conNection)
        {
            var conMail = new MaylUtilityDac("System.Data.SqlClient", conNection);
            var mailLog = new MailLogDto();
            mailLog.Data = DateTime.Now;

            try
            {
                var mail = new System.Net.Mail.MailMessage();
                var smtpServer = new SmtpClient();

                mail.SubjectEncoding = System.Text.Encoding.UTF8;
                mail.BodyEncoding = System.Text.Encoding.UTF8;
                mail.IsBodyHtml = true;

                smtpServer.Host = impoMail.SmtPostaUscita;
                if (impoMail.UseDefaultCredential != null)
                    smtpServer.UseDefaultCredentials = (bool)impoMail.UseDefaultCredential;

                smtpServer.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtpServer.Port = impoMail.SmtpPort;

                smtpServer.Credentials = new System.Net.NetworkCredential(impoMail.UserMail, impoMail.Password);
                if (impoMail.SsLAuto != null) smtpServer.EnableSsl = (bool)impoMail.SsLAuto;


                mail.From = new MailAddress(impoMail.UtenteFiguraInvio, impoMail.NominativoInvio);

                List<string> destinatario = new List<string>();

                if (impoMail.Destinatario != null)
                {
                    destinatario = impoMail.Destinatario.Split(';').Select(s => s.Replace(";", "")).ToList();

                }

                if (!string.IsNullOrEmpty(impoMail.DestinatarioLst) && impoMail.DestinatarioLst.Length > 0) destinatario.AddRange(impoMail.DestinatarioLst.Split(';').Select(s => s.Replace(";", "")).ToList());

                if (destinatario.Count > 0)
                {
                    foreach (var dest in destinatario)
                    {
                        if (!string.IsNullOrEmpty(dest.Trim())) mail.To.Add(dest.Trim());
                    }
                }

                var destinatarioCc = impoMail.Cc.Split(';').Select(s => s.Replace(";", "")).ToList();
                foreach (var dest in destinatarioCc)
                {
                    if (!string.IsNullOrEmpty(dest.Trim())) mail.CC.Add(dest.Trim());
                }

                mail.Subject = impoMail.Oggetto;

                mail.Body = impoMail.Messaggio;

                mailLog.Commenti = $"Invio file a {impoMail.Destinatario} Lista {impoMail.DestinatarioLst}; CC {impoMail.Cc}; messaggio: { impoMail.Messaggio}; ";

                mail.Priority = MailPriority.High;

                await smtpServer.SendMailAsync(mail);

                mailLog.Esito = 1;
            }
            catch (Exception ex)
            {
                mailLog.Commenti =
                    $"Invio file a {impoMail.Destinatario} - Lista {impoMail.DestinatarioLst}; Errore: {ex.Message}";
                mailLog.Esito = 0;
            }

            //salvo il log della spedizione effettuata
            conMail.SaveMailLog(mailLog.Commenti, mailLog.Esito, mailLog.Tipo);
        }
        public async Task SendMailAsyncNew(MailDto impoMail, string txtBody, int numRecord, int aggioMailSend, string conNection)
        {
            var conMail = new MaylUtilityDac("System.Data.SqlClient", conNection);
            var mailLog = new MailLogDto();
            mailLog.Data = DateTime.Now;

            try
            {
                var mail = new System.Net.Mail.MailMessage();
                var smtpServer = new SmtpClient();

                mail.SubjectEncoding = System.Text.Encoding.UTF8;
                mail.BodyEncoding = System.Text.Encoding.UTF8;
                mail.IsBodyHtml = true;

                smtpServer.Host = impoMail.SmtPostaUscita;
                if (impoMail.UseDefaultCredential != null)
                    smtpServer.UseDefaultCredentials = (bool)impoMail.UseDefaultCredential;

                smtpServer.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtpServer.Port = impoMail.SmtpPort;

                smtpServer.Credentials = new System.Net.NetworkCredential(impoMail.UserMail, impoMail.Password);
                if (impoMail.SsLAuto != null) smtpServer.EnableSsl = (bool)impoMail.SsLAuto;


                mail.From = new MailAddress(impoMail.UtenteFiguraInvio, impoMail.NominativoInvio);

                List<string> destinatario = new List<string>();

                if (impoMail.Destinatario != null)
                {
                    destinatario = impoMail.Destinatario.Split(';').Select(s => s.Replace(";", "")).ToList();

                }

                if (!string.IsNullOrEmpty(impoMail.DestinatarioLst) && impoMail.DestinatarioLst.Length > 0) destinatario.AddRange(impoMail.DestinatarioLst.Split(';').Select(s => s.Replace(";", "")).ToList());

                if (destinatario.Count > 0)
                {
                    foreach (var dest in destinatario)
                    {
                        if (!string.IsNullOrEmpty(dest.Trim())) mail.To.Add(dest.Trim());
                    }
                }

                var destinatarioCc = impoMail.Cc.Split(';').Select(s => s.Replace(";", "")).ToList();
                foreach (var dest in destinatarioCc)
                {
                    if (!string.IsNullOrEmpty(dest.Trim())) mail.CC.Add(dest.Trim());
                }

                mail.Subject = impoMail.Oggetto;

                var bd = impoMail.Messaggio.Length > 3850 ? impoMail.Messaggio.Substring(0, 3850) : impoMail.Messaggio; //limito la lunghezza del mesaggio
                if (numRecord > 0)
                {
                    bd += "<br /> <br />";
                    bd += $"<b>Numero Articoli trovati:{numRecord}</b>";
                    bd += "<br /> <br />";
                }
                 if(txtBody!= "")
                    bd += txtBody;

                mail.Body = bd;

                mailLog.Commenti = $"Invio file a {impoMail.Destinatario} Lista {impoMail.DestinatarioLst}; CC {impoMail.Cc}; messaggio: { impoMail.Messaggio}; ";

                mail.Priority = MailPriority.High;

                await smtpServer.SendMailAsync(mail);

                mailLog.Esito = 1;
            }
            catch (Exception ex)
            {
                mailLog.Commenti =
                    $"Invio file a {impoMail.Destinatario} - Lista {impoMail.DestinatarioLst}; Errore: {ex.Message}";
                mailLog.Esito = 0;
            }

            if(aggioMailSend== 1)//aggiornamento del campo mail
            {

            }
            //salvo il log della spedizione effettuata
            conMail.SaveMailLog(mailLog.Commenti, mailLog.Esito, mailLog.Tipo);
        }



        public List<Articoli> GetMaterialiMancantiCon(string conAVdb)
        {
            var lst = new List<Articoli>();

            try
            {
                var inDb = new MaylUtilityDac("System.Data.SqlClient", conAVdb);
                var ds = inDb.GetMaterialiMancanti();

                lst = (from DataRow dr in ds.Tables[0].Rows
                       select new Articoli()
                       {
                           IdArticolo = dr["IdArticolo"].ToString(),
                           DescriArticolo = dr["Descrizione"].ToString(),
                           QuantiInMagazzino = !dr.IsNull("Quantita") ? Convert.ToInt32(dr["Quantita"].ToString()) : 0,
                           MinMagazzino = !dr.IsNull("MinMagazzino") ? Convert.ToInt32(dr["MinMagazzino"].ToString()) : 0,
                           ArtInternoEsterno = !dr.IsNull("IdTipo") ? Convert.ToInt32(dr["IdTipo"].ToString()) == 16 ? 1 : 0 : 0
                       }).ToList();
            }
            catch (Exception e)
            {

            }
            return lst;
        }

        public List<AvvisiDaInviare> GetLstInvioMailDaInviare(string conDb)
        {
            var lst = new List<AvvisiDaInviare>();

            try
            {
                var inDb = new MaylUtilityDac("System.Data.SqlClient", conDb);
                var ds = inDb.GetLstInvioMailDaInviare();

                lst = (from DataRow dr in ds.Tables[0].Rows
                       select new AvvisiDaInviare()
                       {
                           IdAvviso = dr["Id"].ToString(),
                           DataInvioProgrammata = !dr.IsNull("DataInvio") ? DateTime.Parse(dr["DataInvio"].ToString()) : (DateTime?)null,
                           Oggetto = dr["Oggetto"].ToString(),
                           Messaggio = dr["Messaggio"].ToString(),
                           Destinatario = dr["Email"].ToString()
                       }).ToList();
            }
            catch (Exception e)
            {

            }
            return lst;
        }


    }
}

using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Net.Mail;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Avalon.EntityDto;

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

        public async Task DoWorkAsync(List<Articoli> lst, string percorso, int numRecord, string conNection)
        {

            try
            {

                var impoMail = GetMailImpo(1, conNection);
                string textBody = " <table border=" + 1 + " cellpadding=" + 0 + " cellspacing=" + 0 + " width = " + 700 + " style='border: 0.5px;'><tr bgcolor='#4da6ff'><td style='width:15%; text-align: center;'><b>Cod Articolo</b></td> <td style='text-align: center;'> <b> Descrizione</b> </td><td style='text-align: center;'><b>Min Maga</b></td> <td style='text-align: center;'> <b> Q.T.</b> </td></tr>";

                foreach (var l in lst)
                {
                    textBody += "<tr><td>" + l.IdArticolo + "</td><td> " + l.DescriArticolo + "</td><td>" + l.MinMagazzino + "</td><td> " + l.QuantiInMagazzino + "</td> </tr>";
                }

                textBody += "</table>";

                await SendMailAsyncNew(impoMail, textBody, numRecord, conNection);
            }
            catch (Exception e)
            {
                //salvo il log della spedizione effettuata
                var conMail = new MaylUtilityDac("System.Data.SqlClient", conNection);
                conMail.SaveMailLog($"ERRORE: {e.Message} - SOURCE: {e.Source}", 0, 1);
            }



        }

        public async Task SendMailAsyncNew(MailDto impoMail, string txtBody, int numRecord, string conNection)
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

                bd += "<br /> <br />";
                bd += $"<b>Numero Articoli trovati:{numRecord}</b>";
                bd += "<br /> <br />";
                bd += txtBody;

                mail.Body = bd;

                mailLog.Commenti = $"Invio file a {impoMail.Destinatario ?? impoMail.DestinatarioLst}; CC {impoMail.Cc}; messaggio: Numero Articoli trovati:{numRecord}; ";

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
                           MinMagazzino = !dr.IsNull("MinMagazzino") ? Convert.ToInt32(dr["MinMagazzino"].ToString()) : 0
                       }).ToList();
            }
            catch (Exception e)
            {

            }
            return lst;
        }


    }
}

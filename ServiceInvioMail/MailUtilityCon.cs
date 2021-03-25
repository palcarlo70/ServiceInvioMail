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
        public Classi.MailDto GetMailImpo(int idImpo, string conNection)
        {
            Classi.MailDto lst = new Classi.MailDto();
            try
            {
                var con = new MaylUtilityDac("System.Data.SqlClient", conNection);
                var ds = con.GetMailImpo(idImpo);
                lst = new Classi.MailDto();
                lst = (from DataRow dr in ds.Tables[0].Rows
                       select new Classi.MailDto()
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
                throw new FaultException<DataException>(dex, new FaultReason(errore), new FaultCode("Sender"), null);
            }
            return lst;
        }

        public async Task<bool> ControlloCarenzeMagazzino(string conNection, string percorso) //string percorso = Server.MapPath("~/");
        {
            var conn = new Avalon.Connection.AnagraficaCon();
            var lst = conn.GetMaterialiMancantiInCommessa(conNection);
            if (lst.Articoli.Count > 0)
            {
                var campi = lst.Articoli.Select(aa => new CampiPdf { Campo1 = aa.IdArticolo, Campo2 = aa.DescriArticolo, Campo3 = aa.QuantiMagazzinoInt.ToString(), Campo4 = aa.MinMagazzino.ToString() }).ToList();

                var crea = new LibreriaPDF.PdfGenerici();
                var fileStampa = crea.StampaListaMateriali(percorso, "LstMateriali", campi);

                var mailImpo = GetMailImpo(1, conNection);
                await SendMailAsyncNew(mailImpo, fileStampa, conNection);
                return true;
            }
            return false;
        }

        public async Task<bool> SendMailAsyncNew(Classi.MailDto impoMail, string fileAllegato, string conNection)
        {
            var conMail = new MaylUtilityDac("System.Data.SqlClient", conNection);
            var mailLog = new Classi.MailLogDto();
            mailLog.Data = DateTime.Now;
            mailLog.Commenti = $"Invio file a {impoMail.Destinatario ?? impoMail.DestinatarioLst}; CC {impoMail.Cc}; messaggio: { impoMail.Messaggio}; ";
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

                var destinatario = impoMail.Destinatario.Split(';').Select(s => s.Replace(";", "")).ToList();
                foreach (var dest in destinatario)
                {
                    mail.To.Add(dest.Trim());
                }

                var destinatarioCc = impoMail.Cc.Split(';').Select(s => s.Replace(";", "")).ToList();
                foreach (var dest in destinatarioCc)
                {
                    mail.CC.Add(dest.Trim());
                }

                mail.Subject = impoMail.Oggetto;
                mail.Body = impoMail.Messaggio.Length > 3850 ? impoMail.Messaggio.Substring(0, 3850) : impoMail.Messaggio; //limito la lunghezza del mesaggio

                mail.Priority = MailPriority.High;

                //string file = @"TxtCreateFileDaLista(lstValor, pr); pr.OutputFileName = $""PresenzeDSGroup_{mese}_{anno}.pdf""; string fileStampa = GenerateReport(mese, anno, 0, pr, null, giornoMax, giornoDal);";
                //allego il file creato
                if (!string.IsNullOrEmpty(fileAllegato)) mail.Attachments.Add(new Attachment(fileAllegato));

                await smtpServer.SendMailAsync(mail);

                mailLog.Esito = 1;
            }
            catch (Exception ex)
            {
                mailLog.Commenti =
                    $"Invio file a {impoMail.Destinatario}; Errore: {ex.Message}";
                mailLog.Esito = 0;
            }
            //salvo il log della spedizione effettuata
            conMail.SaveMailLog(mailLog.Commenti, mailLog.Esito, mailLog.Tipo);
            return true;
        }
    }
}

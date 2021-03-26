using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DBAccess;

namespace ServiceInvioMail
{
    public class MaylUtilityDac : DBWork
    {
        public MaylUtilityDac(string provider, string connectionString) : base(provider, connectionString)
        {

        }

        public DataSet GetMailImpo(int? id)
        {

            DbCommand cmd = CreateCommand("eml.GetEmailImpo", true);

            base.SetParameter(cmd, "id", DbType.Int32, ParameterDirection.Input, (object)id ?? DBNull.Value);

            cmd.CommandType = CommandType.StoredProcedure;
            var lst = base.GetDataSet(cmd);

            return lst;

        }

        public bool SaveMailLog(string commenti, int? esito, int? tipo)
        {
            try
            {
                DbCommand cmd = CreateCommand("eml.InsLogEmail", true);

                base.SetParameter(cmd, "Commenti", DbType.String, ParameterDirection.Input, commenti);
                base.SetParameter(cmd, "Esito", DbType.Int32, ParameterDirection.Input, (object)esito ?? DBNull.Value);
                base.SetParameter(cmd, "Tipo", DbType.Int32, ParameterDirection.Input, (object)tipo ?? DBNull.Value);

                cmd.CommandType = CommandType.StoredProcedure;

                base.GetDataSet(cmd);

            }
            catch (Exception e)
            {
                return false;
            }

            return true;
        }

        public DataSet GetMaterialiMancanti()
        {

            DbCommand cmd = CreateCommand("PR_MaterialeSottoMinimo", true);

            cmd.CommandType = CommandType.StoredProcedure;
            var lst = base.GetDataSet(cmd);

            return lst;
        }
 
    }
}
